using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text.RegularExpressions;

namespace SeewoOpt.Services
{
    /// <summary>GitHub Release 中单个资产文件的信息</summary>
    [DataContract]
    public class ReleaseAsset
    {
        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "browser_download_url")]
        public string BrowserDownloadUrl { get; set; }

        [DataMember(Name = "size")]
        public long Size { get; set; }
    }

    /// <summary>GitHub Release API 响应中我们关心的字段</summary>
    [DataContract]
    public class ReleaseInfo
    {
        [DataMember(Name = "tag_name")]
        public string TagName { get; set; }

        [DataMember(Name = "body")]
        public string Body { get; set; }

        [DataMember(Name = "assets")]
        public List<ReleaseAsset> Assets { get; set; }
    }

    /// <summary>
    /// GitHub Release 响应解析。
    ///
    /// 历史问题：原实现用正则从 JSON 文本中抓取字段，存在三类实际缺陷——
    ///   1. 取"第一个" browser_download_url，Release 里一旦有多个资产
    ///      （例如额外上传了 .sha256 文本）就会下载到错误文件；
    ///   2. body 字段含转义引号与 \uXXXX 序列，正则 + Regex.Unescape
    ///      处理不了多层转义；
    ///   3. body 匹配用了 Singleline 贪婪模式，一旦 Release 说明文本里
    ///      恰好出现 "tag_name" 字样，会从 JSON 开头一路吞到 body 末尾。
    ///
    /// 现改用 .NET Framework 自带的 DataContractJsonSerializer（位于
    /// System.Runtime.Serialization.dll），正确处理所有转义，且不引入
    /// 任何新的 NuGet 依赖。
    /// </summary>
    public static class ReleaseInfoParser
    {
        /// <summary>可执行文件的扩展名</summary>
        private static readonly string[] ExecutableExtensions = { ".exe", ".msi" };

        /// <summary>
        /// 解析 Release JSON。解析失败返回 false 并记录原因，不抛异常。
        /// </summary>
        public static bool TryParse(string json, out ReleaseInfo release, out string error)
        {
            release = null;
            error = null;

            if (string.IsNullOrWhiteSpace(json))
            {
                error = "响应内容为空";
                return false;
            }

            try
            {
                var settings = new DataContractJsonSerializerSettings
                {
                    // GitHub 响应含大量未知字段，必须允许忽略
                    UseSimpleDictionaryFormat = true
                };

                var serializer = new DataContractJsonSerializer(typeof(ReleaseInfo), settings);

                using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json)))
                {
                    release = (ReleaseInfo)serializer.ReadObject(stream);
                }

                if (release == null)
                {
                    error = "解析结果为 null";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(release.TagName))
                {
                    error = "响应中缺少 tag_name 字段";
                    release = null;
                    return false;
                }

                return true;
            }
            catch (SerializationException ex)
            {
                error = "JSON 结构不符合预期: " + ex.Message;
            }
            catch (Exception ex)
            {
                error = "解析响应时发生异常: " + ex.Message;
            }

            release = null;
            return false;
        }

        /// <summary>
        /// 从资产列表中挑选要下载的可执行文件。
        /// 不再简单取第一个——Release 常常附带校验文件、说明文件等。
        /// </summary>
        public static ReleaseAsset SelectExecutableAsset(List<ReleaseAsset> assets)
        {
            if (assets == null || assets.Count == 0) return null;

            // 优先匹配 .exe，其次 .msi
            foreach (string ext in ExecutableExtensions)
            {
                foreach (ReleaseAsset asset in assets)
                {
                    if (asset == null || string.IsNullOrEmpty(asset.Name)) continue;
                    if (!asset.Name.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) continue;
                    if (string.IsNullOrEmpty(asset.BrowserDownloadUrl)) continue;

                    // 排除明显不是主程序的校验类文件
                    if (IsChecksumLike(asset.Name)) continue;

                    return asset;
                }
            }

            return null;
        }

        /// <summary>判断文件名是否是校验和/签名类附属文件</summary>
        private static bool IsChecksumLike(string fileName)
        {
            string lower = fileName.ToLowerInvariant();
            return lower.EndsWith(".sha256")
                || lower.EndsWith(".sha256sum")
                || lower.EndsWith(".md5")
                || lower.EndsWith(".sig")
                || lower.EndsWith(".asc")
                || lower.EndsWith(".sha1")
                || lower.EndsWith(".txt")
                || lower.EndsWith(".json");
        }

        /// <summary>
        /// 从 Release 说明文本中提取预期 SHA256。
        /// 形如 "SHA256: ABCDEF..." 或 "SHA256 ABCDEF..."。
        /// </summary>
        public static string TryExtractExpectedSha256(string releaseBody)
        {
            if (string.IsNullOrEmpty(releaseBody)) return null;

            Match match = Regex.Match(
                releaseBody,
                @"SHA256[:\s=]+([A-Fa-f0-9]{64})",
                RegexOptions.IgnoreCase);

            return match.Success ? match.Groups[1].Value.ToUpperInvariant() : null;
        }

        /// <summary>
        /// 解析版本号。容忍 "v" 前缀与 "-beta" 等预发布后缀，失败返回 null。
        /// </summary>
        public static Version ParseVersion(string tagName)
        {
            if (string.IsNullOrWhiteSpace(tagName)) return null;

            string versionStr = tagName.Trim().TrimStart('v', 'V');

            int dashIndex = versionStr.IndexOf('-');
            if (dashIndex > 0)
                versionStr = versionStr.Substring(0, dashIndex);

            string[] parts = versionStr.Split('.');
            if (parts.Length == 0 || parts.Length > 4) return null;

            foreach (string part in parts)
            {
                if (!int.TryParse(part, out _))
                    return null;
            }

            return Version.TryParse(versionStr, out Version parsed) ? parsed : null;
        }
    }
}
