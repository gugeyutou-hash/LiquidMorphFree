# LiquidMorph Free

面向 Windows 64 位 PowerPoint 的液态玻璃加载项。欢迎通过 Issue 反馈问题，通过 Pull Request 提交代码、文档和测试。

## 当前公开的代码

- `src/addin` 包含功能区、材质参数、形状选区、动画端点和玻璃移动后的刷新逻辑。
- `src/host` 包含通过 WebView2 调度本地渲染、接收 PNG 和 MP4 的 Windows 宿主。

本仓库公开上述组件的源码。渲染引擎的公开再分发授权仍待核实，因此当前仓库没有随附引擎和完整可安装发行包。单独编译这两个项目可以检查和开发加载项，但生成玻璃图像需要兼容的渲染引擎。

## 构建

使用 Windows x64、.NET Framework 4.8 Targeting Pack、64 位 PowerPoint 和 Office 互操作程序集。下载 Microsoft.Web.WebView2 SDK 1.0.2903.40，将 NuGet 包中下列文件复制到仓库根目录的 `third_party/webview2/`：

- `lib/net462/Microsoft.Web.WebView2.Core.dll`
- `lib/net462/Microsoft.Web.WebView2.WinForms.dll`
- `runtimes/win-x64/native/WebView2Loader.dll`

在仓库根目录用 Windows PowerShell 执行 `.\Build.ps1`。宿主和加载项分别输出到各自项目的 `bin/Release/`。运行宿主还需要 Microsoft Edge WebView2 Runtime。

## 参与共建

请阅读 [贡献指南](CONTRIBUTING.md)。仓库原创代码和文档采用 [MIT 许可证](LICENSE)。Microsoft 组件遵循各自的许可。
