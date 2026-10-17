using System;
using System.Collections.Generic;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 一个更新源（Gitee / GitHub）查到的候选版本。
    ///
    /// 只描述"这一次查询的结果"，不持有连接、不负责下载。
    /// </summary>
    public class UpdateCandidate
    {
        /// <summary>来源名称，仅用于日志（Gitee / GitHub）</summary>
        public string SourceName { get; set; }

        /// <summary>是否为国内源。版本相同时优先用它下载。</summary>
        public bool IsDomestic { get; set; }

        /// <summary>发行版标签解析出的版本号</summary>
        public Version Version { get; set; }

        /// <summary>资产文件名（如 SeewoOpt.exe）</summary>
        public string AssetName { get; set; }

        /// <summary>资产下载地址</summary>
        public string DownloadUrl { get; set; }

        /// <summary>从发行说明里提取的 SHA256，可能为 null</summary>
        public string ExpectedSha256 { get; set; }
    }

    /// <summary>
    /// 多更新源的选源与降级判定。
    ///
    /// 【为什么要有这一层】更新是这个程序唯一"自己改自己"的路径，
    /// 而它现在有两个源。选错源的两个后果都很严重：
    ///   · 只认国内源 → 某次发版忘记同步，全体用户永远停在旧版且毫无提示；
    ///   · 只认官方源 → 教室网络下根本下载不动，等于没有自动更新。
    /// 这些都是纯逻辑，能脱离网络单测，所以从 Program 里抽出来。
    /// </summary>
    public static class UpdateSourcePolicy
    {
        /// <summary>
        /// 在多个源的查询结果里选一个：版本更高的优先；版本相同时优先国内源。
        ///
        /// 两端都可能为 null（对应源不可用），此时返回另一个。
        /// </summary>
        public static UpdateCandidate PickNewer(UpdateCandidate a, UpdateCandidate b)
        {
            if (a == null) return b;
            if (b == null) return a;

            int cmp = a.Version.CompareTo(b.Version);
            if (cmp != 0) return cmp > 0 ? a : b;

            return a.IsDomestic ? a : (b.IsDomestic ? b : a);
        }

        /// <summary>
        /// 组装下载地址候选列表：主地址在前，其余"同版本"的源作为备用。
        ///
        /// 【为什么只收同版本的备用地址】版本不同说明那个源还没同步，
        /// 拿它的地址只会下到旧版——而签名对旧版同样有效，
        /// 最后只能靠版本校验拦下来。与其下完再拒，不如一开始就不试。
        /// </summary>
        public static List<string> BuildDownloadUrls(
            UpdateCandidate best, IEnumerable<UpdateCandidate> all)
        {
            var urls = new List<string>();
            if (best == null || string.IsNullOrEmpty(best.DownloadUrl)) return urls;

            urls.Add(best.DownloadUrl);

            if (all == null) return urls;

            foreach (UpdateCandidate other in all)
            {
                if (other == null || ReferenceEquals(other, best)) continue;
                if (string.IsNullOrEmpty(other.DownloadUrl)) continue;
                if (other.Version == null || best.Version == null) continue;
                if (!other.Version.Equals(best.Version)) continue;
                if (urls.Contains(other.DownloadUrl)) continue;

                urls.Add(other.DownloadUrl);
            }

            return urls;
        }

        /// <summary>把版本号统一成 Major.Minor.Build 三段，缺省段按 0 处理。</summary>
        public static Version NormalizeToThreeParts(Version version)
        {
            if (version == null) return null;
            return new Version(version.Major, version.Minor, Math.Max(version.Build, 0));
        }

        /// <summary>
        /// 下载到的文件版本是否等于目标版本。
        ///
        /// 【为什么必须比对】签名只证明"这个文件确实是作者发的"，
        /// 不证明"是我要的那一版"。任何位于中间的源都可以返回一个
        /// 真实、签名有效的旧版 exe，再配一个写着高版本号的清单——
        /// 不校验版本，就等于把"降级到有已知缺陷的旧版"变成正常路径。
        ///
        /// 比较前统一成三段：AssemblyVersion 可能是 "26.10.17"，
        /// 某些构建会写成 "26.10.17.0"，而 Version.CompareTo 认为
        /// 缺省段是 -1，直接比会把一次正常更新误判成版本不符。
        /// </summary>
        public static bool VersionMatches(Version fileVersion, Version target)
        {
            Version actual = NormalizeToThreeParts(fileVersion);
            Version expected = NormalizeToThreeParts(target);

            if (actual == null || expected == null) return false;
            return actual == expected;
        }
    }
}
