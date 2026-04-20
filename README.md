# FluentInjector

A Windows DLL injector with a modern WPF UI.

![preview](Assets/FluentInjector.png)

## Requirements

- Windows 11 22H2+ (Build 22621) — acrylic blur requires this
- .NET 8.0 SDK
- Administrator privileges

## Build

```bash
dotnet restore
dotnet build -c Release
```

Self-contained single executable:

```bash
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

Output: `bin/Release/net8.0-windows/win-x64/publish/DllInjector.exe`

## Usage

1. Run as Administrator
2. Select a target process from the list
3. Add one or more DLLs via the drop zone or drag-and-drop
4. Click **Inject**

Config is saved automatically to `%USERPROFILE%\Downloads\FluentInjector_Config.json`.

## How it works

Standard `LoadLibraryW` injection via `CreateRemoteThread`. Opens the target process, writes the DLL path into its memory, spawns a remote thread at `LoadLibraryW`, and waits for completion.

## Notes

- Only processes with a visible window appear in the list
- Injection into protected or anti-cheat processes will fail regardless of privileges
- For educational purposes only
