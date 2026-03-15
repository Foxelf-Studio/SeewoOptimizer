# 陈叔叔系统优化工具 (SeewoOptimizer)

![GitHub release (latest by date)](https://img.shields.io/github/v/release/Foxelf-Studio/SeewoOptimizer)
![GitHub All Releases](https://img.shields.io/github/downloads/Foxelf-Studio/SeewoOptimizer/total)

**PEACE & LOVE 川中计算机协会 荣誉出品**  
专为老希沃/多媒体教室环境优化的轻量级工具，集时间同步、音量调节、进程管理于一体，让你告别不准时的系统时间导致的证书错误、被调低系统音量老师不会调回去，导致同学们不能开心快乐地观看课件内视频的问题！

---

## 📋 功能特性

- **⏱️ 智能时间同步**  
  内置多个公共 NTP 服务器（含国内阿里云、国际标准），自动重试，支持管理员权限检测与提权重启。
- **🔊 一键音量调节**  
  可自定义开机音量百分比，或随时调整系统主音量。
- **⚡ 结束 WPS 进程**  
  一键清理 WPS 进程，用于wps2019出现无法打开课件问题。
- **🔄 静默启动与托盘驻留**  
  支持开机自启（通过任务计划）和静默启动，启动后自动隐藏到系统托盘，不影响日常使用。
- **📦 自动更新**  
  连接 GitHub Releases 检查新版本，静默下载并替换，无需手动操作。
- **📝 详细日志**  
  所有操作记录在 `%LOCALAPPDATA%\TimeSyncTool\startup.log`，便于排查问题。

---

## 🚀 下载与安装

### 直接下载
前往 [Releases 页面](https://github.com/Foxelf-Studio/SeewoOptimizer/releases) 下载最新的 `SeewoOptimizer.exe` 单文件，双击运行即可。

### 手动编译
如果你有 .NET Framework 4.7.2 或更高版本开发环境：
```bash
git clone https://github.com/Foxelf-Studio/SeewoOptimizer.git
cd SeewoOptimizer
# 使用 Visual Studio 打开解决方案并生成
```

---

## 🛠️ 使用方法

1. **首次运行**  
   程序默认会尝试自动同步时间、调节音量（60%）并结束 WPS 进程，并检查是否已被加入开机自启，若没有，则按照设置项添加。主窗口会显示实时日志。

2. **隐藏到托盘**  
   点击“隐藏到托盘”按钮或**直接关闭窗口**，程序将最小化到系统托盘，**不会直接结束掉进程**，防止老师关掉程序。右键托盘图标可“显示窗口”或“退出程序”。

3. **设置选项**  
   菜单栏 **文件 → 设置** 可以打开设置窗口：
   - **开机自启**：通过 Windows 任务计划实现用户登录后自动启动。
   - **静默启动**：启动后不显示主窗口，直接进入托盘。
   - **自动调节音量**：勾选后每次启动将音量设为指定值（0-100）。
   - **结束 WPS 进程**：勾选后启动时自动结束 WPS 相关进程。

   设置会自动保存到注册表 `HKCU\Software\TimeSyncTool`。

4. **手动同步**  
   若同步失败，可点击“重新同步”按钮再次尝试。

5. **查看日志**  
   菜单栏 **视图 → 打开日志** 将自动弹出日志文件所在文件夹并选中日志文件。

---

## ⚙️ 配置说明

所有设置存储在注册表中，也可通过程序界面修改。默认值如下：

| 配置项       | 默认值 | 说明                     |
| ------------ | ------ | ------------------------ |
| AutoStart    | true   | 开机自启                 |
| SilentStart  | false  | 静默启动（隐藏窗口）     |
| AutoVolume   | true   | 自动调节音量             |
| VolumeLevel  | 60     | 目标音量百分比（0-100）  |
| KillWps      | true   | 启动时结束 WPS 进程      |

---

## 🔄 自动更新机制

- 程序启动后会在后台检查 GitHub 最新 release。
- 如果发现新版本，托盘区会弹出气泡提示，并开始静默下载。
- 下载完成后，程序自动退出，启动批处理脚本替换旧文件，然后退出。**整个过程无需用户干预，不弹窗，不打扰**。
- 更新完成后，下次启动即为新版本。

---

## 📁 项目结构

```
SeewoOptimizer/
├── Program.cs               # 程序入口、单实例控制、更新逻辑
├── TimeSyncForm.cs          # 主窗体、时间同步核心逻辑
├── SettingsForm.cs          # 设置窗体（未提供代码，但需自行实现）
├── Microsoft.Win32.TaskScheduler.dll  # 任务计划操作依赖（内嵌资源）
└── ...
```

---

## 🤝 贡献

欢迎提交 Issue 或 Pull Request。  
项目采用 MIT 许可证，自由使用、修改、分发。

---

## 📄 许可证

MIT License © 2025 川中计算机协会  
（本工具仅供学习交流，请勿用于商业用途）

---

## 💬 联系我们
 
GitHub Issues: [点击反馈](https://github.com/Foxelf-Studio/SeewoOptimizer/issues)
