using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using TimeSyncTool;

namespace WindowsFormsApp1
{
    static class Program
    {
        private static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TimeSyncTool", "startup.log");

        // GitHub 配置
        private const string GITHUB_API = "https://api.github.com/repos/Foxelf-Studio/SeewoOptimizer/releases/latest";
        private const string GITHUB_USER_AGENT = "SeewoOptimizer-Updater";
        private static readonly string UPDATE_DIR = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TimeSyncTool", "updates");

        [STAThread]
        static void Main()
        {
            try
            {
                // 检查是否有待应用的更新（更新后重启）
                HandlePendingUpdate();

                // 异步检查新版本（不阻塞主线程）
                Task.Run(() => CheckForUpdatesAsync());

                // 原有的启动代码
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
                EnsureRequiredDllsExist();

                WriteLog("========================================");
                WriteLog($"程序启动 - 时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
                WriteLog($"========================================");
                WriteLog($"命令行参数: {string.Join(" ", Environment.GetCommandLineArgs())}");
                WriteLog($"当前目录: {Environment.CurrentDirectory}");
                WriteLog($"程序路径: {Application.ExecutablePath}");
                WriteLog($"程序目录: {Path.GetDirectoryName(Application.ExecutablePath)}");
                WriteLog($"当前版本: {Assembly.GetExecutingAssembly().GetName().Version}");
                WriteLog($"========================================");

                bool createdNew;
                using (new Mutex(true, "TimeSyncTool_UniqueMutex", out createdNew))
                {
                    if (!createdNew)
                    {
                        WriteLog("已有实例运行，退出");
                        MessageBox.Show("程序已在运行中。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    Application.ThreadException += (sender, e) =>
                    {
                        WriteLog($"========== 线程异常 ==========");
                        WriteLog($"异常消息: {e.Exception.Message}");
                        WriteLog($"异常堆栈: {e.Exception.StackTrace}");
                        MessageBox.Show($"发生未处理的异常：{e.Exception.Message}\n\n程序将关闭。",
                            "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    };

                    AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
                    {
                        Exception ex = e.ExceptionObject as Exception;
                        WriteLog($"========== 致命异常 ==========");
                        WriteLog($"异常消息: {ex?.Message ?? "未知错误"}");
                        WriteLog($"异常堆栈: {ex?.StackTrace ?? ""}");
                        MessageBox.Show($"发生致命错误：{ex?.Message ?? "未知错误"}\n\n程序将关闭。",
                            "致命错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    };

                    try
                    {
                        WriteLog("开始运行主窗体");
                        Application.Run(new TimeSyncForm());
                        WriteLog("主窗体运行结束");
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"========== Run 异常 ==========");
                        WriteLog($"异常消息: {ex.Message}");
                        MessageBox.Show($"程序运行错误：{ex.Message}", "错误",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                WriteLog("程序正常退出");
                WriteLog("========================================\n");
            }
            catch (Exception ex)
            {
                WriteLog($"========== Main 顶级异常 ==========");
                WriteLog($"异常类型: {ex.GetType().Name}");
                WriteLog($"异常消息: {ex.Message}");
                WriteLog($"异常堆栈: {ex.StackTrace}");
                MessageBox.Show($"程序启动时发生致命错误：{ex.Message}\n\n程序将关闭。",
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 检查是否有待应用的更新（更新后重启）
        /// </summary>
        private static void HandlePendingUpdate()
        {
            try
            {
                string pendingFile = Path.Combine(UPDATE_DIR, "pending.txt");
                if (File.Exists(pendingFile))
                {
                    WriteLog("检测到待应用的更新");

                    string newVersionPath = File.ReadAllText(pendingFile).Trim();
                    string currentExe = Application.ExecutablePath;

                    if (File.Exists(newVersionPath))
                    {
                        File.Copy(newVersionPath, currentExe, true);
                        File.Delete(newVersionPath);
                        File.Delete(pendingFile);

                        WriteLog("更新文件已应用");
                        MessageBox.Show("更新完成！", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog($"处理待更新文件失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 异步检查 GitHub 上的新版本
        /// </summary>
        private static async Task CheckForUpdatesAsync()
        {
            try
            {
                WriteLog("开始后台检查更新...");

                Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                WriteLog($"当前版本: {currentVersion}");

                if (!Directory.Exists(UPDATE_DIR))
                    Directory.CreateDirectory(UPDATE_DIR);

                using (WebClient client = new WebClient())
                {
                    client.Headers.Add("User-Agent", GITHUB_USER_AGENT);
                    client.Headers.Add("Accept", "application/vnd.github.v3+json");

                    // 异步下载，不阻塞主线程
                    string jsonResponse = await client.DownloadStringTaskAsync(GITHUB_API);
                    WriteLog("GitHub API 响应成功");

                    // 使用 JavaScriptSerializer 解析 JSON
                    var serializer = new JavaScriptSerializer();
                    var releaseInfo = serializer.Deserialize<dynamic>(jsonResponse);

                    string tagName = releaseInfo["tag_name"];
                    string versionStr = tagName.TrimStart('v');
                    Version latestVersion = new Version(versionStr);

                    WriteLog($"GitHub 最新版本: {latestVersion}");

                    if (latestVersion.CompareTo(currentVersion) > 0)
                    {
                        WriteLog("发现新版本，准备询问用户");
                        string releaseNotes = releaseInfo["body"] ?? "无更新说明";

                        // 需要在 UI 线程上显示对话框
                        bool shouldUpdate = false;
                        await Task.Run(() =>
                        {
                            shouldUpdate = MessageBox.Show(
                                $"发现新版本 {latestVersion}\n当前版本 {currentVersion}\n\n" +
                                $"更新内容：\n{releaseNotes}\n\n" +
                                "是否立即下载更新？",
                                "软件更新",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question
                            ) == DialogResult.Yes;
                        });

                        if (shouldUpdate)
                        {
                            var assets = releaseInfo["assets"] as object[];
                            if (assets != null && assets.Length > 0)
                            {
                                dynamic firstAsset = assets[0];
                                string downloadUrl = firstAsset["browser_download_url"];
                                string fileName = firstAsset["name"];

                                WriteLog($"下载地址: {downloadUrl}");
                                await DownloadUpdateAsync(downloadUrl, fileName, latestVersion);
                            }
                        }
                    }
                    else
                    {
                        WriteLog("已是最新版本");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog($"后台检查更新失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 异步下载更新文件
        /// </summary>
        private static async Task DownloadUpdateAsync(string downloadUrl, string fileName, Version newVersion)
        {
            try
            {
                string downloadedFile = Path.Combine(UPDATE_DIR, fileName);
                string currentExe = Application.ExecutablePath;
                string updaterScript = Path.Combine(UPDATE_DIR, "update.bat");

                WriteLog("开始下载更新文件...");

                using (WebClient client = new WebClient())
                {
                    client.DownloadProgressChanged += (s, e) =>
                    {
                        try
                        {
                            if (e.ProgressPercentage % 10 == 0)
                                WriteLog($"下载进度: {e.ProgressPercentage}%");
                        }
                        catch { }
                    };

                    // 异步下载
                    await client.DownloadFileTaskAsync(new Uri(downloadUrl), downloadedFile);

                    WriteLog("下载完成，准备更新脚本");
                    CreateUpdateScript(updaterScript, currentExe, downloadedFile, newVersion);
                    WriteLog("更新脚本创建成功，准备重启应用");

                    Process.Start(new ProcessStartInfo()
                    {
                        FileName = updaterScript,
                        UseShellExecute = true,
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    });

                    Environment.Exit(0);
                }
            }
            catch (Exception ex)
            {
                WriteLog($"下载过程出错: {ex.Message}");
                MessageBox.Show($"下载更新出错：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void CreateUpdateScript(string scriptPath, string currentExe, string newExe, Version newVersion)
        {
            string pendingFile = Path.Combine(UPDATE_DIR, "pending.txt");
            File.WriteAllText(pendingFile, newExe);

            string batchContent = $@"@echo off
title 正在更新软件...
echo 正在更新，请稍候...

:: 等待主程序完全退出
timeout /t 2 /nobreak > nul

:: 尝试替换文件（如果被占用就重试）
:retry
taskkill /f /im {Path.GetFileName(currentExe)} 2>nul
timeout /t 1 /nobreak > nul

if exist ""{currentExe}"" (
    del /f /q ""{currentExe}""
    if errorlevel 1 (
        echo 文件被占用，等待重试...
        timeout /t 2 /nobreak > nul
        goto retry
    )
)

:: 复制新文件
copy /y ""{newExe}"" ""{currentExe}""

:: 启动新版本
start """" ""{currentExe}""

:: 清理临时文件
del /f /q ""{newExe}"" 2>nul
del /f /q ""{pendingFile}"" 2>nul
del /f /q ""%~f0"" 2>nul

exit
";
            File.WriteAllText(scriptPath, batchContent);
            WriteLog($"更新脚本已创建: {scriptPath}");
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                if (args.Name.StartsWith("Microsoft.Win32.TaskScheduler"))
                {
                    WriteLog($"AssemblyResolve 被触发，尝试加载：{args.Name}");
                    string appDir = Path.GetDirectoryName(Application.ExecutablePath);
                    string dllPath = Path.Combine(appDir, "Microsoft.Win32.TaskScheduler.dll");

                    if (File.Exists(dllPath))
                    {
                        Assembly assembly = Assembly.LoadFrom(dllPath);
                        WriteLog($"成功从 {dllPath} 加载程序集");
                        return assembly;
                    }
                    else
                    {
                        WriteLog($"文件不存在：{dllPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog($"程序集解析异常：{ex.Message}");
            }
            return null;
        }

        private static void EnsureRequiredDllsExist()
        {
            string[] dllsToExtract = { "Microsoft.Win32.TaskScheduler.dll" };
            string appDir = Path.GetDirectoryName(Application.ExecutablePath);

            foreach (string dllName in dllsToExtract)
            {
                string dllPath = Path.Combine(appDir, dllName);
                if (File.Exists(dllPath))
                {
                    WriteLog($"DLL已存在：{dllName}");
                    try { Assembly.LoadFrom(dllPath); WriteLog($"预加载程序集成功：{dllName}"); }
                    catch (Exception ex) { WriteLog($"预加载程序集失败：{ex.Message}"); }
                    continue;
                }

                try
                {
                    WriteLog($"正在提取DLL：{dllName}");
                    string resourceName = $"WindowsFormsApp1.{dllName}";
                    using (Stream resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                    {
                        if (resourceStream == null)
                        {
                            resourceName = dllName;
                            using (var fallbackStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                            {
                                if (fallbackStream == null)
                                {
                                    WriteLog($"错误：找不到嵌入式资源 {dllName}");
                                    WriteLog("可用的资源列表：");
                                    foreach (string res in Assembly.GetExecutingAssembly().GetManifestResourceNames())
                                        WriteLog($"  - {res}");
                                    continue;
                                }
                                using (FileStream fileStream = new FileStream(dllPath, FileMode.Create, FileAccess.Write))
                                    fallbackStream.CopyTo(fileStream);
                            }
                        }
                        else
                        {
                            using (FileStream fileStream = new FileStream(dllPath, FileMode.Create, FileAccess.Write))
                                resourceStream.CopyTo(fileStream);
                        }
                        WriteLog($"✓ DLL提取成功：{dllPath}");
                        try { Assembly.LoadFrom(dllPath); WriteLog($"预加载程序集成功：{dllName}"); }
                        catch (Exception ex) { WriteLog($"预加载程序集失败：{ex.Message}"); }
                    }
                }
                catch (Exception ex)
                {
                    WriteLog($"提取DLL {dllName} 时出错：{ex.Message}");
                }
            }
        }

        private static void WriteLog(string message)
        {
            try
            {
                string logDir = Path.GetDirectoryName(LogFilePath);
                if (!string.IsNullOrEmpty(logDir)) Directory.CreateDirectory(logDir);
                File.AppendAllText(LogFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}\n");
            }
            catch { }
        }
    }
}