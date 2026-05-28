# Windows 独立版打包说明

目标产物：

```text
dist/windows/WordStyleStandardizer.exe
```

这个 exe 会内置：

- 图形界面
- 标准样式模板 `assets/标准样式Word文档.docx`
- Word 样式规整脚本逻辑
- Python 运行时和依赖库

## 打包步骤

在 Windows 上双击运行：

```text
packaging/build_windows_exe.bat
```

如果项目位于 WSL 的 `\\wsl.localhost\...` 路径，优先在 PowerShell 中运行：

```powershell
powershell -ExecutionPolicy Bypass -File packaging/build_windows_exe.ps1
```

首次运行会安装：

- `requirements.txt` 中的依赖
- `pyinstaller`

生成后，把下面这个文件发给同事即可：

```text
dist/windows/WordStyleStandardizer.exe
```

## 使用方式

1. 双击 `WordStyleStandardizer.exe`。
2. 选择源 `.docx` 文件。
3. 默认使用内置标准模板。
4. 点击“开始规整”。
5. 程序会在源文件同目录生成：

```text
<源文件名>_标准样式规整.docx
<源文件名>_样式规整校验报告.md
```

如需使用新版标准模板，可在界面中手动选择模板文件。

## 注意事项

- 只支持 `.docx`，旧版 `.doc` 需要先另存为 `.docx`。
- 处理前请关闭正在编辑的源 Word 文件。
- 若杀毒软件拦截自制 exe，需要将该 exe 加入信任或改用公司签名流程。
