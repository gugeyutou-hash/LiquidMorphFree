# LiquidMorph Free

Windows 版 PowerPoint 液态玻璃加载项。选中形状后可以套用玻璃材质、调节模糊和边光、移动后刷新背景，也可以设置动画起点和终点。首次使用请下载带安装器的完整 ZIP。

[下载 v1.0.1 Windows x64 安装包](https://github.com/gugeyutou-hash/LiquidMorphFree/releases/download/v1.0.1/LiquidMorph-Free-1.0.1-Windows-x64-Source-and-Installer.zip) · [版本说明](https://github.com/gugeyutou-hash/LiquidMorphFree/releases/tag/v1.0.1) · [贡献指南](CONTRIBUTING.md)

## 安装到 PowerPoint

1. 在 PowerPoint 的「文件 → 帐户 → 关于 PowerPoint」确认是 **64 位桌面版**。目标机器还需要 Windows x64、.NET Framework 4.8 或更新的 4.x 版本、Office 互操作组件和 Microsoft Edge WebView2 Runtime。
2. 下载上方文件名以 **Source-and-Installer.zip** 结尾的附件。GitHub 自动生成的 Source code 压缩包仅对应仓库源码。用资源管理器选择「全部提取」，进入能看到 `Install.cmd` 和 `payload.json` 的目录。
3. 保存演示文稿，关闭 PowerPoint，双击 `Install.cmd`。安装器检查依赖和文件校验值，将组件复制到当前账户的 `%LOCALAPPDATA%\LiquidMorphFree\Releases`，并注册 COM 加载项。使用日常运行 PowerPoint 的账户即可。
4. 输出 JSON 中的 `"ok":true` 表示安装成功。重新打开 PowerPoint，在 **LiquidMorph Free** 选项卡下操作。用空白演示文稿放入一张背景图片，在图片上画矩形，单击矩形边框后选择材质。拖动生成的玻璃对象到背景不同的位置，观察停止移动后的刷新结果。

### 命令行或 AI 自动部署

从解压目录执行以下命令。第一条只检查环境和文件，退出码为 0 后再执行安装。自动化程序请使用 64 位 Windows PowerShell 5.1，并读取退出码和 JSON。包内 `AI-DEPLOY.md` 是操作说明，`deploy.json` 和 `payload.json` 分别描述安装入口和文件清单。

```powershell
.\Install.cmd -CheckOnly
if ($LASTEXITCODE -eq 0) {
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1
}
```

### 排错与卸载

| 遇到的问题 | 检查方法 |
| --- | --- |
| 安装检查报错 | 从解压目录运行上面的预检，按输出确认 PowerPoint 位数、WebView2 Runtime 或 Office 互操作组件。校验值不符时重新下载并完整解压到新目录。 |
| 重启后找不到功能区 | 打开「文件 → 选项 → 加载项」，分别查看「禁用项目」与「COM 加载项」，确认 LiquidMorph Free 已启用。重新运行该版本安装器可将当前用户的注册路径指回发行版。 |
| 材质按钮为灰色 | 单击单个矩形的边框，退出文字编辑，等候选区刷新。若仍为灰色，请记录形状类型、PowerPoint 版本和复现步骤。 |
| 渲染返回 exit 3 | 查看弹窗所指工作目录中的 `*.error.txt`。提交 Issue 时附错误文本、所选材质和复现步骤。该目录中的背景图可能含演示文稿内容，请自行决定是否分享。 |

卸载时保存文稿并关闭 PowerPoint，在完整解压目录运行以下命令，注销当前用户的加载项。重新安装可再次启用它。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\register-addin.ps1 -Unregister
```

## 从源码继续开发

**准备工作区。** 安装 Git、.NET Framework 4.8 Targeting Pack，并满足上述 PowerPoint 运行条件。先在 GitHub 上 Fork 本仓库，再克隆自己的 Fork，把下面的 YOUR-ACCOUNT 替换为你的 GitHub 用户名。

```powershell
git clone https://github.com/YOUR-ACCOUNT/LiquidMorphFree.git
cd LiquidMorphFree
git switch -c fix/glass-refresh
```

**准备依赖并编译。** 将完整发行包解压到单独目录。以下命令在仓库根目录的 64 位 Windows PowerShell 中执行。把第一行改为你的解压目录。WebView2 SDK 版本为 1.0.2903.40。

```powershell
$package = 'C:\Downloads\LiquidMorph-1.0.1'
New-Item -ItemType Directory -Force third_party | Out-Null
Copy-Item "$package\third_party\webview2" third_party -Recurse -Force
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build.ps1
```

构建成功会生成 `src/addin/bin/Release/LiquidMorph.Desktop.dll` 和 `src/host/bin/Release/LiquidMorph.Host.exe`。如遇 MSB3644，请安装 4.8 Targeting Pack。若缺少 office 或 Extensibility 引用，检查 Office 互操作程序集是否安装，以及加载项项目文件中的 HintPath 是否对应本机路径。

**把自己的构建加载到 PowerPoint。** 保存文稿、退出 PowerPoint 后，在同一个终端运行以下命令。开发目录承载新编译的组件和发行包中的配套引擎。注册会将当前用户的免费版切换到这个目录。

```powershell
$dev = Join-Path $env:LOCALAPPDATA 'LiquidMorphFree\Development'
New-Item -ItemType Directory -Force $dev | Out-Null
Copy-Item .\src\addin\bin\Release\LiquidMorph.Desktop.dll $dev -Force
Copy-Item .\src\host\bin\Release\* $dev -Force
Copy-Item "$package\src\engine" $dev -Recurse -Force
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\register-addin.ps1 -Dist $dev
```

重新打开 PowerPoint 测试。下一轮修改后，退出 PowerPoint，重新编译并复制产物，再打开 PowerPoint。运行发行包的 Install.cmd 可切回稳定版。

| 想修改的功能 | 从这里开始 |
| --- | --- |
| 功能区布局、图标和按钮回调 | `src/addin/Ribbon.xml`、`Icons/`、`RibbonActions.cs` |
| 材质参数、预设和持久化 | `src/addin/GlassOptions.cs`、`GlassOptionsStore.cs` |
| 选区检查、移动检测、动画命令 | `src/addin/RibbonActions.cs` |
| 背景导出、渲染请求、结果插回幻灯片 | `src/addin/GlassPipeline.cs` |
| WebView2 启动、结果接收和错误记录 | `src/host/Host.cs` |

**验证并贡献。** 用同一张测试幻灯片检查无选区、单个矩形和文字编辑三种状态；在相同背景上比较预设、拖动刷新与动画起终点。将改动提交到自己的分支，运行 `git push -u origin fix/glass-refresh`，在 GitHub 对上游 main 发起 Pull Request，写明复现步骤、环境及验证结果。

## 许可范围

仓库原创加载项、宿主、脚本和文档采用 [MIT 许可证](LICENSE)。项目维护者已确认完整发行包可公开分发，任何人均可使用。引擎及第三方组件的修改、再许可等权利按各自授权处理。发行包保留原始 OPEN-SOURCE.md、Microsoft 许可文件及通知。
