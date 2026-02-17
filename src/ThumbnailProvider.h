#pragma once

#include <Unknwn.h>
#include <thumbcache.h>

class VibeThumbnailProvider final : public IThumbnailProvider, public IInitializeWithStream {
 public:
  VibeThumbnailProvider();

  // IUnknown
  IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override;
  IFACEMETHODIMP_(ULONG) AddRef() override;
  IFACEMETHODIMP_(ULONG) Release() override;

  // IInitializeWithStream
  IFACEMETHODIMP Initialize(IStream* pstream, DWORD grfMode) override;

  // IThumbnailProvider
  IFACEMETHODIMP GetThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) override;

 private:
  ~VibeThumbnailProvider();

  long ref_count_;
  IStream* stream_;
};

extern const CLSID CLSID_VibeThumbnailProvider;
