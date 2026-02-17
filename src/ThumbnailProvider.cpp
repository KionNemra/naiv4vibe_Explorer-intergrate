#include "ThumbnailProvider.h"

#include <Objbase.h>
#include <Shlwapi.h>
#include <Wincodec.h>
#include <Wincrypt.h>
#include <Windows.h>

#include <algorithm>
#include <memory>
#include <nlohmann/json.hpp>
#include <string>
#include <string_view>
#include <vector>

#pragma comment(lib, "Crypt32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Windowscodecs.lib")

const CLSID CLSID_VibeThumbnailProvider = {0x4d2aa77e,
                                           0xf513,
                                           0x4e30,
                                           {0xa0, 0x34, 0xe6, 0x2c, 0xa8, 0xc2, 0xa9, 0xd8}};

namespace {

template <typename T>
class ComPtr {
 public:
  ComPtr() : ptr_(nullptr) {}
  ~ComPtr() {
    if (ptr_) ptr_->Release();
  }

  T** operator&() { return &ptr_; }
  T* get() const { return ptr_; }
  T* operator->() const { return ptr_; }

 private:
  T* ptr_;
};

std::string ReadAllBytes(IStream* stream) {
  LARGE_INTEGER zero = {};
  ULARGE_INTEGER pos = {};
  if (FAILED(stream->Seek(zero, STREAM_SEEK_SET, &pos))) {
    return {};
  }

  std::string data;
  constexpr size_t kChunk = 8192;
  std::vector<char> buffer(kChunk);
  while (true) {
    ULONG read = 0;
    if (FAILED(stream->Read(buffer.data(), static_cast<ULONG>(buffer.size()), &read))) {
      return {};
    }
    if (read == 0) break;
    data.append(buffer.data(), read);
  }
  return data;
}

std::string StripDataUrlPrefix(const std::string& input) {
  auto comma = input.find(',');
  if (comma == std::string::npos) return input;
  return input.substr(comma + 1);
}

std::wstring AsciiToWide(const std::string& value) {
  if (value.empty()) return {};
  int needed = MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), nullptr, 0);
  if (needed <= 0) return {};
  std::wstring out(needed, L'\0');
  MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), out.data(), needed);
  return out;
}

std::vector<BYTE> DecodeBase64(const std::string& value) {
  std::wstring wide = AsciiToWide(value);
  if (wide.empty()) return {};

  DWORD out_size = 0;
  if (!CryptStringToBinaryW(wide.c_str(), 0, CRYPT_STRING_BASE64, nullptr, &out_size, nullptr, nullptr)) {
    return {};
  }

  std::vector<BYTE> output(out_size);
  if (!CryptStringToBinaryW(wide.c_str(), 0, CRYPT_STRING_BASE64, output.data(), &out_size, nullptr, nullptr)) {
    return {};
  }
  output.resize(out_size);
  return output;
}

HRESULT DecodeImageToBitmap(const std::vector<BYTE>& image_data, UINT cx, HBITMAP* out_bitmap) {
  if (image_data.empty() || !out_bitmap) return E_INVALIDARG;

  IStream* mem_stream = SHCreateMemStream(image_data.data(), static_cast<UINT>(image_data.size()));
  if (!mem_stream) return E_FAIL;

  ComPtr<IWICImagingFactory> factory;
  HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
                                IID_PPV_ARGS(&factory));
  if (FAILED(hr)) return hr;

  ComPtr<IWICBitmapDecoder> decoder;
  hr = factory->CreateDecoderFromStream(mem_stream, nullptr, WICDecodeMetadataCacheOnLoad,
                                        &decoder);
  mem_stream->Release();
  if (FAILED(hr)) return hr;

  ComPtr<IWICBitmapFrameDecode> frame;
  hr = decoder->GetFrame(0, &frame);
  if (FAILED(hr)) return hr;

  UINT width = 0;
  UINT height = 0;
  hr = frame->GetSize(&width, &height);
  if (FAILED(hr) || width == 0 || height == 0) return E_FAIL;

  UINT scaled_width = width;
  UINT scaled_height = height;
  if (std::max(width, height) > cx && cx > 0) {
    if (width >= height) {
      scaled_width = cx;
      scaled_height = (height * cx) / width;
    } else {
      scaled_height = cx;
      scaled_width = (width * cx) / height;
    }
  }

  ComPtr<IWICBitmapScaler> scaler;
  hr = factory->CreateBitmapScaler(&scaler);
  if (FAILED(hr)) return hr;

  hr = scaler->Initialize(frame.get(), scaled_width, scaled_height, WICBitmapInterpolationModeFant);
  if (FAILED(hr)) return hr;

  ComPtr<IWICFormatConverter> converter;
  hr = factory->CreateFormatConverter(&converter);
  if (FAILED(hr)) return hr;

  hr = converter->Initialize(scaler.get(), GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone,
                             nullptr, 0.0f, WICBitmapPaletteTypeCustom);
  if (FAILED(hr)) return hr;

  BITMAPINFO bmi = {};
  bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
  bmi.bmiHeader.biWidth = static_cast<LONG>(scaled_width);
  bmi.bmiHeader.biHeight = -static_cast<LONG>(scaled_height);
  bmi.bmiHeader.biPlanes = 1;
  bmi.bmiHeader.biBitCount = 32;
  bmi.bmiHeader.biCompression = BI_RGB;

  void* bits = nullptr;
  HBITMAP hbmp = CreateDIBSection(nullptr, &bmi, DIB_RGB_COLORS, &bits, nullptr, 0);
  if (!hbmp || !bits) return E_OUTOFMEMORY;

  const UINT stride = scaled_width * 4;
  const UINT image_size = stride * scaled_height;
  hr = converter->CopyPixels(nullptr, stride, image_size, static_cast<BYTE*>(bits));
  if (FAILED(hr)) {
    DeleteObject(hbmp);
    return hr;
  }

  *out_bitmap = hbmp;
  return S_OK;
}

}  // namespace

VibeThumbnailProvider::VibeThumbnailProvider() : ref_count_(1), stream_(nullptr) {
  ModuleAddRef();
}

VibeThumbnailProvider::~VibeThumbnailProvider() {
  if (stream_) stream_->Release();
  ModuleRelease();
}

IFACEMETHODIMP VibeThumbnailProvider::QueryInterface(REFIID riid, void** ppv) {
  if (!ppv) return E_POINTER;

  if (riid == IID_IUnknown || riid == IID_IThumbnailProvider) {
    *ppv = static_cast<IThumbnailProvider*>(this);
  } else if (riid == IID_IInitializeWithStream) {
    *ppv = static_cast<IInitializeWithStream*>(this);
  } else {
    *ppv = nullptr;
    return E_NOINTERFACE;
  }

  AddRef();
  return S_OK;
}

IFACEMETHODIMP_(ULONG) VibeThumbnailProvider::AddRef() {
  return InterlockedIncrement(&ref_count_);
}

IFACEMETHODIMP_(ULONG) VibeThumbnailProvider::Release() {
  ULONG count = InterlockedDecrement(&ref_count_);
  if (count == 0) delete this;
  return count;
}

IFACEMETHODIMP VibeThumbnailProvider::Initialize(IStream* pstream, DWORD) {
  if (!pstream) return E_INVALIDARG;
  if (stream_) return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);

  stream_ = pstream;
  stream_->AddRef();
  return S_OK;
}

IFACEMETHODIMP VibeThumbnailProvider::GetThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) {
  if (!phbmp || !pdwAlpha || !stream_) return E_INVALIDARG;

  *phbmp = nullptr;
  *pdwAlpha = WTSAT_UNKNOWN;

  std::string json_content = ReadAllBytes(stream_);
  if (json_content.empty()) return E_FAIL;

  std::string encoded_image;
  try {
    nlohmann::json parsed = nlohmann::json::parse(json_content);
    if (parsed.contains("thumbnail") && parsed["thumbnail"].is_string() && cx <= 512) {
      encoded_image = StripDataUrlPrefix(parsed["thumbnail"].get<std::string>());
    } else if (parsed.contains("image") && parsed["image"].is_string()) {
      encoded_image = StripDataUrlPrefix(parsed["image"].get<std::string>());
    } else if (parsed.contains("thumbnail") && parsed["thumbnail"].is_string()) {
      encoded_image = StripDataUrlPrefix(parsed["thumbnail"].get<std::string>());
    }
  } catch (...) {
    return E_FAIL;
  }

  if (encoded_image.empty()) return E_FAIL;

  std::vector<BYTE> decoded = DecodeBase64(encoded_image);
  if (decoded.empty()) return E_FAIL;

  HRESULT hr = DecodeImageToBitmap(decoded, cx, phbmp);
  if (FAILED(hr)) return hr;

  *pdwAlpha = WTSAT_ARGB;
  return S_OK;
}
