# naiv4vibe_Explorer-intergrate

Windows Explorer `.naiv4vibe` 缩略图处理器（In-proc COM DLL）。

## 实现概览

`VibeThumbnailProvider` 实现：

- `IInitializeWithStream`：接收 Explorer 传入文件流。
- `IThumbnailProvider::GetThumbnail`：
  1. 读取 UTF-8 JSON。
  2. 优先取 `thumbnail`（`cx <= 512`），否则取 `image`。
  3. 对 `thumbnail` / `image` 都兼容去除 `data:image/...;base64,` 前缀（若存在）。
  4. `CryptStringToBinaryW` 解 base64。
  5. `SHCreateMemStream` 建内存流。
  6. WIC 解码 + 等比缩放为 `cx`，输出 `HBITMAP`，`WTSAT_ARGB`。

注册时写入：

- `HKCR\.naiv4vibe\ShellEx\{E357FCCD-A995-4576-B01F-234630154E96} = {4D2AA77E-F513-4E30-A034-E62CA8C2A9D8}`

## 依赖

- CMake 3.20+
- MSVC x64
- `nlohmann/json`（header-only，需能 `#include <nlohmann/json.hpp>`）

## 构建

```powershell
cmake -S . -B build -A x64
cmake --build build --config Release
```

输出 DLL：

- `build\Release\Naiv4VibeThumbnailProvider.dll`

## 安装/卸载

```powershell
# 管理员 PowerShell
./scripts/install.ps1 -DllPath .\build\Release\Naiv4VibeThumbnailProvider.dll
./scripts/uninstall.ps1 -DllPath .\build\Release\Naiv4VibeThumbnailProvider.dll
```

## 调试建议

- 先卸载旧版本再安装新 DLL。
- 清空缩略图缓存后重启 Explorer。
- 观察 `.naiv4vibe` 在大/小图标视图下的缩略图是否正常。
