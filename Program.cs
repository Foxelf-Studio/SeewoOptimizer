using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using TimeSyncTool;

namespace WindowsFormsApp1  // 必须与项目命名空间一致
{
    static class Program
    {
        private static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TimeSyncTool", "startup.log");

        [STAThread]
        static void Main()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            EnsureRequiredDllsExist();

            WriteLog("========================================");
            WriteLog($"程序启动 - 时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            WriteLog($"========================================");
            WriteLog($"命令行参数: {string.Join(" ", Environment.GetCommandLineArgs())}");
            WriteLog($"当前目录: {Environment.CurrentDirectory}");
            WriteLog($"程序路径: {Application.ExecutablePath}");
            WriteLog($"程序目录: {Path.GetDirectoryName(Application.ExecutablePath)}");
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

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            if (args.Name.StartsWith("Microsoft.Win32.TaskScheduler"))
            {
                WriteLog($"AssemblyResolve 被触发，尝试加载：{args.Name}");
                string appDir = Path.GetDirectoryName(Application.ExecutablePath);
                string dllPath = Path.Combine(appDir, "Microsoft.Win32.TaskScheduler.dll");

                if (File.Exists(dllPath))
                {
                    try
                    {
                        Assembly assembly = Assembly.LoadFrom(dllPath);
                        WriteLog($"成功从 {dllPath} 加载程序集");
                        return assembly;
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"加载程序集失败：{ex.Message}");
                    }
                }
                else
                {
                    WriteLog($"文件不存在：{dllPath}");
                }
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
                    try { Assembly.LoadFrom(dllPath); WriteLog($"预加载程序集成功：{dllName}"); } catch (Exception ex) { WriteLog($"预加载程序集失败：{ex.Message}"); }
                    continue;
                }

                try
                {
                    WriteLog($"正在提取DLL：{dllName}");
                    // 根据日志，实际资源名是 WindowsFormsApp1.Microsoft.Win32.TaskScheduler.dll
                    string resourceName = $"WindowsFormsApp1.{dllName}";
                    using (Stream resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                    {
                        if (resourceStream == null)
                        {
                            // 如果找不到，尝试其他可能的前缀
                            resourceName = dllName; // 不带前缀
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
                        try { Assembly.LoadFrom(dllPath); WriteLog($"预加载程序集成功：{dllName}"); } catch (Exception ex) { WriteLog($"预加载程序集失败：{ex.Message}"); }
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