# naiv4vibe_Explorer-intergrate
THIS WILL BE 90% vibe coding project.
创建 VS x64 C++ 项目：ATL In-proc COM DLL（或 WRL/C++ COM 手写也行，但 ATL 最省事）

定义 COM 类 VibeThumbnailProvider：实现 IInitializeWithStream + IThumbnailProvider

在 Initialize 里保存 IStream*；在 GetThumbnail 里：

读 JSON（UTF-8）

取 thumbnail 或 image

处理 data URL 前缀

CryptStringToBinaryW base64 解码

SHCreateMemStream 包成 IStream*

WIC 解码/缩放到 cx，输出 HBITMAP，alpha=ARGB

实现注册表写入（DllRegisterServer）：

HKCR\.naiv4vibe\ShellEx\{E357FCCD-A995-4576-B01F-234630154E96} → {YourCLSID}

提供 scripts/install.ps1 / scripts/uninstall.ps1：

调用对应位数的 regsvr32 /s your.dll

提供简单测试说明：清空缩略图缓存/重启 Explorer 后看效果
