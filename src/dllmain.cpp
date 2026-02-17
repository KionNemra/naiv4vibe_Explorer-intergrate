#include "ThumbnailProvider.h"

#include <Objbase.h>
#include <Shlwapi.h>
#include <Windows.h>

#include <cwchar>
#include <new>

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shlwapi.lib")

namespace {

volatile long g_module_ref_count = 0;
HINSTANCE g_instance = nullptr;

const wchar_t* kThumbnailProviderClsid = L"{4D2AA77E-F513-4E30-A034-E62CA8C2A9D8}";
const wchar_t* kThumbnailProviderName = L"Naiv4Vibe Thumbnail Provider";
class VibeClassFactory final : public IClassFactory {
 public:
  VibeClassFactory() : ref_count_(1) { InterlockedIncrement(&g_module_ref_count); }
  ~VibeClassFactory() { InterlockedDecrement(&g_module_ref_count); }

  IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
    if (!ppv) return E_POINTER;
    if (riid == IID_IUnknown || riid == IID_IClassFactory) {
      *ppv = static_cast<IClassFactory*>(this);
      AddRef();
      return S_OK;
    }
    *ppv = nullptr;
    return E_NOINTERFACE;
  }

  IFACEMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&ref_count_); }

  IFACEMETHODIMP_(ULONG) Release() override {
    ULONG count = InterlockedDecrement(&ref_count_);
    if (count == 0) delete this;
    return count;
  }

  IFACEMETHODIMP CreateInstance(IUnknown* outer, REFIID riid, void** ppv) override {
    if (outer) return CLASS_E_NOAGGREGATION;
    VibeThumbnailProvider* provider = new (std::nothrow) VibeThumbnailProvider();
    if (!provider) return E_OUTOFMEMORY;

    HRESULT hr = provider->QueryInterface(riid, ppv);
    provider->Release();
    return hr;
  }

  IFACEMETHODIMP LockServer(BOOL lock) override {
    if (lock) {
      InterlockedIncrement(&g_module_ref_count);
    } else {
      InterlockedDecrement(&g_module_ref_count);
    }
    return S_OK;
  }

 private:
  long ref_count_;
};

HRESULT SetStringValue(HKEY root, const wchar_t* subkey, const wchar_t* value_name, const wchar_t* value) {
  HKEY key = nullptr;
  LONG result = RegCreateKeyExW(root, subkey, 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr,
                                &key, nullptr);
  if (result != ERROR_SUCCESS) return HRESULT_FROM_WIN32(result);

  result = RegSetValueExW(key, value_name, 0, REG_SZ,
                          reinterpret_cast<const BYTE*>(value),
                          static_cast<DWORD>((wcslen(value) + 1) * sizeof(wchar_t)));
  RegCloseKey(key);
  return HRESULT_FROM_WIN32(result);
}

HRESULT RegisterInprocServer() {
  wchar_t module_path[MAX_PATH] = {};
  if (!GetModuleFileNameW(g_instance, module_path, ARRAYSIZE(module_path))) {
    return HRESULT_FROM_WIN32(GetLastError());
  }

  wchar_t clsid_path[128] = {};
  swprintf_s(clsid_path, _countof(clsid_path), L"CLSID\\%s", kThumbnailProviderClsid);

  HRESULT hr = SetStringValue(HKEY_CLASSES_ROOT, clsid_path, nullptr, kThumbnailProviderName);
  if (FAILED(hr)) return hr;

  wchar_t inproc_path[256] = {};
  swprintf_s(inproc_path, _countof(inproc_path), L"%s\\InprocServer32", clsid_path);
  hr = SetStringValue(HKEY_CLASSES_ROOT, inproc_path, nullptr, module_path);
  if (FAILED(hr)) return hr;

  hr = SetStringValue(HKEY_CLASSES_ROOT, inproc_path, L"ThreadingModel", L"Apartment");
  if (FAILED(hr)) return hr;

  return SetStringValue(HKEY_CLASSES_ROOT,
                        L".naiv4vibe\\ShellEx\\{E357FCCD-A995-4576-B01F-234630154E96}", nullptr,
                        kThumbnailProviderClsid);
}

void UnregisterInprocServer() {
  SHDeleteKeyW(HKEY_CLASSES_ROOT, L".naiv4vibe\\ShellEx\\{E357FCCD-A995-4576-B01F-234630154E96}");

  wchar_t clsid_path[128] = {};
  swprintf_s(clsid_path, _countof(clsid_path), L"CLSID\\%s", kThumbnailProviderClsid);
  SHDeleteKeyW(HKEY_CLASSES_ROOT, clsid_path);
}

}  // namespace

BOOL APIENTRY DllMain(HMODULE hmodule, DWORD reason, LPVOID) {
  if (reason == DLL_PROCESS_ATTACH) {
    g_instance = hmodule;
    DisableThreadLibraryCalls(hmodule);
  }
  return TRUE;
}

void ModuleAddRef() {
  InterlockedIncrement(&g_module_ref_count);
}

void ModuleRelease() {
  InterlockedDecrement(&g_module_ref_count);
}

STDAPI DllCanUnloadNow() {
  return g_module_ref_count == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID clsid, REFIID riid, LPVOID* ppv) {
  if (clsid != CLSID_VibeThumbnailProvider) return CLASS_E_CLASSNOTAVAILABLE;

  VibeClassFactory* factory = new (std::nothrow) VibeClassFactory();
  if (!factory) return E_OUTOFMEMORY;

  HRESULT hr = factory->QueryInterface(riid, ppv);
  factory->Release();
  return hr;
}

STDAPI DllRegisterServer() {
  return RegisterInprocServer();
}

STDAPI DllUnregisterServer() {
  UnregisterInprocServer();
  return S_OK;
}
