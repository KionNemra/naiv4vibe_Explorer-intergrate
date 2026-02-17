#include "ThumbnailProvider.h"

#include <Objbase.h>
#include <Shlwapi.h>
#include <Wincodec.h>
#include <Wincrypt.h>
#include <Windows.h>

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <memory>
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

  ComPtr(const ComPtr&) = delete;
  ComPtr& operator=(const ComPtr&) = delete;

  T** operator&() { return &ptr_; }
  T* get() const { return ptr_; }
  T* operator->() const { return ptr_; }

  void Attach(T* value) {
    if (ptr_) ptr_->Release();
    ptr_ = value;
  }

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

size_t SkipJsonWhitespace(std::string_view text, size_t pos) {
  while (pos < text.size() && std::isspace(static_cast<unsigned char>(text[pos]))) {
    ++pos;
  }
  return pos;
}


size_t SkipOptionalUtf8Bom(std::string_view text, size_t pos) {
  if (pos + 3 <= text.size() && static_cast<unsigned char>(text[pos]) == 0xEF &&
      static_cast<unsigned char>(text[pos + 1]) == 0xBB &&
      static_cast<unsigned char>(text[pos + 2]) == 0xBF) {
    return pos + 3;
  }
  return pos;
}

bool ParseHex4(std::string_view text, size_t pos, uint16_t* value) {
  if (!value || pos + 4 > text.size()) return false;

  uint16_t result = 0;
  for (size_t i = 0; i < 4; ++i) {
    const unsigned char ch = static_cast<unsigned char>(text[pos + i]);
    uint16_t nibble = 0;
    if (ch >= '0' && ch <= '9') {
      nibble = static_cast<uint16_t>(ch - '0');
    } else if (ch >= 'A' && ch <= 'F') {
      nibble = static_cast<uint16_t>(ch - 'A' + 10);
    } else if (ch >= 'a' && ch <= 'f') {
      nibble = static_cast<uint16_t>(ch - 'a' + 10);
    } else {
      return false;
    }
    result = static_cast<uint16_t>((result << 4) | nibble);
  }

  *value = result;
  return true;
}

void AppendUtf8Codepoint(uint32_t codepoint, std::string* out) {
  if (!out) return;

  if (codepoint <= 0x7F) {
    out->push_back(static_cast<char>(codepoint));
    return;
  }

  if (codepoint <= 0x7FF) {
    out->push_back(static_cast<char>(0xC0 | (codepoint >> 6)));
    out->push_back(static_cast<char>(0x80 | (codepoint & 0x3F)));
    return;
  }

  if (codepoint <= 0xFFFF) {
    out->push_back(static_cast<char>(0xE0 | (codepoint >> 12)));
    out->push_back(static_cast<char>(0x80 | ((codepoint >> 6) & 0x3F)));
    out->push_back(static_cast<char>(0x80 | (codepoint & 0x3F)));
    return;
  }

  if (codepoint <= 0x10FFFF) {
    out->push_back(static_cast<char>(0xF0 | (codepoint >> 18)));
    out->push_back(static_cast<char>(0x80 | ((codepoint >> 12) & 0x3F)));
    out->push_back(static_cast<char>(0x80 | ((codepoint >> 6) & 0x3F)));
    out->push_back(static_cast<char>(0x80 | (codepoint & 0x3F)));
  }
}

bool DecodeJsonString(std::string_view text, size_t quote_pos, std::string* out, size_t* end_pos) {
  if (!out || !end_pos || quote_pos >= text.size() || text[quote_pos] != '"') return false;

  std::string decoded;
  size_t i = quote_pos + 1;
  while (i < text.size()) {
    char ch = text[i++];
    if (ch == '"') {
      *out = std::move(decoded);
      *end_pos = i;
      return true;
    }

    if (ch != '\\') {
      decoded.push_back(ch);
      continue;
    }

    if (i >= text.size()) return false;
    char escaped = text[i++];
    switch (escaped) {
      case '"':
      case '\\':
      case '/':
        decoded.push_back(escaped);
        break;
      case 'b':
        decoded.push_back('\b');
        break;
      case 'f':
        decoded.push_back('\f');
        break;
      case 'n':
        decoded.push_back('\n');
        break;
      case 'r':
        decoded.push_back('\r');
        break;
      case 't':
        decoded.push_back('\t');
        break;
      case 'u':
        {
          uint16_t code_unit = 0;
          if (!ParseHex4(text, i, &code_unit)) return false;
          i += 4;

          uint32_t codepoint = code_unit;
          if (code_unit >= 0xD800 && code_unit <= 0xDBFF) {
            if (i + 6 > text.size() || text[i] != '\\' || text[i + 1] != 'u') return false;
            uint16_t low_surrogate = 0;
            if (!ParseHex4(text, i + 2, &low_surrogate)) return false;
            if (low_surrogate < 0xDC00 || low_surrogate > 0xDFFF) return false;
            i += 6;

            codepoint = 0x10000 +
                        ((static_cast<uint32_t>(code_unit - 0xD800) << 10) |
                         static_cast<uint32_t>(low_surrogate - 0xDC00));
          } else if (code_unit >= 0xDC00 && code_unit <= 0xDFFF) {
            return false;
          }

          AppendUtf8Codepoint(codepoint, &decoded);
        }
        break;
      default:
        return false;
    }
  }

  return false;
}

bool TryGetJsonStringField(std::string_view json, std::string_view field_name, std::string* value) {
  if (!value || field_name.empty()) return false;

  const auto skip_json_value =
      [&](auto&& self, size_t pos, size_t* end_pos) -> bool {
    if (!end_pos) return false;

    pos = SkipJsonWhitespace(json, pos);
    if (pos >= json.size()) return false;

    if (json[pos] == '"') {
      std::string ignored;
      return DecodeJsonString(json, pos, &ignored, end_pos);
    }

    if (json[pos] == '{') {
      size_t cursor = pos + 1;
      cursor = SkipJsonWhitespace(json, cursor);
      if (cursor < json.size() && json[cursor] == '}') {
        *end_pos = cursor + 1;
        return true;
      }

      while (cursor < json.size()) {
        if (json[cursor] != '"') return false;

        std::string ignored_key;
        size_t key_end = cursor;
        if (!DecodeJsonString(json, cursor, &ignored_key, &key_end)) return false;
        cursor = SkipJsonWhitespace(json, key_end);
        if (cursor >= json.size() || json[cursor] != ':') return false;

        size_t value_end = cursor;
        if (!self(self, cursor + 1, &value_end)) return false;
        cursor = SkipJsonWhitespace(json, value_end);
        if (cursor >= json.size()) return false;
        if (json[cursor] == '}') {
          *end_pos = cursor + 1;
          return true;
        }
        if (json[cursor] != ',') return false;
        cursor = SkipJsonWhitespace(json, cursor + 1);
      }

      return false;
    }

    if (json[pos] == '[') {
      size_t cursor = pos + 1;
      cursor = SkipJsonWhitespace(json, cursor);
      if (cursor < json.size() && json[cursor] == ']') {
        *end_pos = cursor + 1;
        return true;
      }

      while (cursor < json.size()) {
        size_t element_end = cursor;
        if (!self(self, cursor, &element_end)) return false;
        cursor = SkipJsonWhitespace(json, element_end);
        if (cursor >= json.size()) return false;
        if (json[cursor] == ']') {
          *end_pos = cursor + 1;
          return true;
        }
        if (json[cursor] != ',') return false;
        cursor = SkipJsonWhitespace(json, cursor + 1);
      }

      return false;
    }

    if (json.compare(pos, 4, "true") == 0) {
      *end_pos = pos + 4;
      return true;
    }
    if (json.compare(pos, 5, "false") == 0) {
      *end_pos = pos + 5;
      return true;
    }
    if (json.compare(pos, 4, "null") == 0) {
      *end_pos = pos + 4;
      return true;
    }

    size_t cursor = pos;
    if (json[cursor] == '-') ++cursor;
    if (cursor >= json.size()) return false;

    if (json[cursor] == '0') {
      ++cursor;
    } else {
      if (!std::isdigit(static_cast<unsigned char>(json[cursor]))) return false;
      while (cursor < json.size() && std::isdigit(static_cast<unsigned char>(json[cursor]))) {
        ++cursor;
      }
    }

    if (cursor < json.size() && json[cursor] == '.') {
      ++cursor;
      if (cursor >= json.size() || !std::isdigit(static_cast<unsigned char>(json[cursor]))) {
        return false;
      }
      while (cursor < json.size() && std::isdigit(static_cast<unsigned char>(json[cursor]))) {
        ++cursor;
      }
    }

    if (cursor < json.size() && (json[cursor] == 'e' || json[cursor] == 'E')) {
      ++cursor;
      if (cursor < json.size() && (json[cursor] == '+' || json[cursor] == '-')) ++cursor;
      if (cursor >= json.size() || !std::isdigit(static_cast<unsigned char>(json[cursor]))) {
        return false;
      }
      while (cursor < json.size() && std::isdigit(static_cast<unsigned char>(json[cursor]))) {
        ++cursor;
      }
    }

    *end_pos = cursor;
    return true;
  };

  size_t cursor = SkipJsonWhitespace(json, 0);
  cursor = SkipOptionalUtf8Bom(json, cursor);
  cursor = SkipJsonWhitespace(json, cursor);
  if (cursor >= json.size() || json[cursor] != '{') return false;

  cursor = SkipJsonWhitespace(json, cursor + 1);
  if (cursor < json.size() && json[cursor] == '}') return false;

  while (cursor < json.size()) {
    if (json[cursor] != '"') return false;

    std::string key;
    size_t key_end = cursor;
    if (!DecodeJsonString(json, cursor, &key, &key_end)) return false;

    cursor = SkipJsonWhitespace(json, key_end);
    if (cursor >= json.size() || json[cursor] != ':') return false;

    size_t value_start = SkipJsonWhitespace(json, cursor + 1);
    if (value_start >= json.size()) return false;

    if (key == field_name && json[value_start] == '"') {
      size_t value_end = value_start;
      return DecodeJsonString(json, value_start, value, &value_end);
    }

    size_t value_end = value_start;
    if (!skip_json_value(skip_json_value, value_start, &value_end)) return false;

    cursor = SkipJsonWhitespace(json, value_end);
    if (cursor >= json.size()) return false;
    if (json[cursor] == '}') return false;
    if (json[cursor] != ',') return false;
    cursor = SkipJsonWhitespace(json, cursor + 1);
  }

  return false;
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

  ComPtr<IStream> mem_stream;
  mem_stream.Attach(SHCreateMemStream(image_data.data(), static_cast<UINT>(image_data.size())));
  if (!mem_stream.get()) return E_FAIL;

  ComPtr<IWICImagingFactory> factory;
  HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
                                IID_PPV_ARGS(&factory));
  if (FAILED(hr)) return hr;

  ComPtr<IWICBitmapDecoder> decoder;
  hr = factory->CreateDecoderFromStream(mem_stream.get(), nullptr, WICDecodeMetadataCacheOnLoad,
                                        &decoder);
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

    scaled_width = std::max(1u, scaled_width);
    scaled_height = std::max(1u, scaled_height);
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

  std::string thumbnail;
  std::string image;
  const bool has_thumbnail = TryGetJsonStringField(json_content, "thumbnail", &thumbnail);
  const bool has_image = TryGetJsonStringField(json_content, "image", &image);

  std::string encoded_image;
  if (has_thumbnail && cx <= 512) {
    encoded_image = StripDataUrlPrefix(thumbnail);
  } else if (has_image) {
    encoded_image = StripDataUrlPrefix(image);
  } else if (has_thumbnail) {
    encoded_image = StripDataUrlPrefix(thumbnail);
  }

  if (encoded_image.empty()) return E_FAIL;

  std::vector<BYTE> decoded = DecodeBase64(encoded_image);
  if (decoded.empty()) return E_FAIL;

  HRESULT hr = DecodeImageToBitmap(decoded, cx, phbmp);
  if (FAILED(hr)) return hr;

  *pdwAlpha = WTSAT_ARGB;
  return S_OK;
}
