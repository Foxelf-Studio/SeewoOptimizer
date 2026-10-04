using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 构建标识。启动时把身份信息写进日志，用于确认"跑着的到底是哪个 exe"。
    ///
    /// 【为什么需要这个】
    /// 排查"改了代码但还是老行为"时，反复出问题的是：改的是工作区源码，
    /// 测的是桌面上另一个目录里的旧副本，两者行为不一致却无从分辨。
    /// 单靠 AssemblyVersion 无法区分——同一个版本号下可以有很多次构建。
    /// 因此这里输出三重指纹：
    ///   1. BuildId           人工递增的构建号，一眼可辨
    ///   2. 可执行文件 SHA256  精确到字节，相同即完全相同
    ///   3. 文件最后写入时间     粗粒度但直观，能立刻看出"这个文件是不是刚编译的"
    ///
    /// SHA256 只在启动时算一次，600KB 文件耗时在毫秒级，不影响启动。
    /// </summary>
    public static class BuildInfo
    {
        /// <summary>
        /// 构建号。每次交付新 exe 手动递增，绝不与上一个交付物重复。
        /// 格式：yyyy.MM.dd-r序号
        /// </summary>
        public const string BuildId = "2026.10.04-r6";

        private static string _cached;

        /// <summary>程序集版本（主.次.修订）</summary>
        public static string AssemblyVersion
        {
            get
            {
                Version v = Assembly.GetExecutingAssembly().GetName().Version;
                return v == null ? "未知" : v.ToString();
            }
        }

        /// <summary>
        /// 完整构建描述，多行，直接写进日志头部。
        /// </summary>
        public static string Describe()
        {
            if (_cached != null) return _cached;

            string exePath = SafeExePath();

            _cached = string.Format(
                "构建标识: {0} | 程序集版本 {1} | 文件 {2} | SHA256 {3} | 写入时间 {4}",
                BuildId,
                AssemblyVersion,
                string.IsNullOrEmpty(exePath) ? "未知" : Path.GetFileName(exePath),
                ComputeSha256(exePath),
                SafeWriteTime(exePath));

            return _cached;
        }

        private static string SafeExePath()
        {
            try
            {
                return Assembly.GetExecutingAssembly().Location;
            }
            catch
            {
                return null;
            }
        }

        private static string SafeWriteTime(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path)) return "未知";
                return File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            catch
            {
                return "未知";
            }
        }

        /// <summary>计算 exe 的 SHA256，取前 16 位十六进制。失败时返回"计算失败"。</summary>
        private static string ComputeSha256(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path)) return "无文件";

                using (var sha = SHA256.Create())
                using (FileStream fs = File.OpenRead(path))
                {
                    byte[] hash = sha.ComputeHash(fs);
                    var sb = new System.Text.StringBuilder(16);
                    for (int i = 0; i < 8; i++)
                        sb.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
                    return sb.ToString();
                }
            }
            catch
            {
                return "计算失败";
            }
        }
    }
}
