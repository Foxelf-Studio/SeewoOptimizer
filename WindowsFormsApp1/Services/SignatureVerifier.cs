using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace SeewoOpt.Services
{
    /// <summary>签名验证结果</summary>
    public class SignatureCheckResult
    {
        /// <summary>文件是否带有有效 Authenticode 签名</summary>
        public bool IsSigned { get; set; }

        /// <summary>签名者主体名称（如 CN=Foxelf Studio）</summary>
        public string SignerSubject { get; set; }

        /// <summary>签名者证书指纹（SHA-1）</summary>
        public string Thumbprint { get; set; }

        /// <summary>证书颁发者</summary>
        public string Issuer { get; set; }

        /// <summary>证书到期时间</summary>
        public DateTime? NotAfter { get; set; }

        /// <summary>不通过时的原因说明</summary>
        public string FailureReason { get; set; }

        /// <summary>可读的结论描述</summary>
        public string Describe()
        {
            if (!IsSigned)
                return "未签名" + (string.IsNullOrEmpty(FailureReason) ? "" : "：" + FailureReason);

            return string.Format("已签名 {0}（指纹 {1}，颁发者 {2}，有效期至 {3:yyyy-MM-dd}）",
                SignerSubject, Thumbprint, Issuer, NotAfter);
        }
    }

    /// <summary>
    /// Authenticode 签名验证。
    ///
    /// 为什么需要它：原实现从 GitHub Release 说明文本里正则提取 SHA256 作为
    /// "预期哈希"，但该哈希与被校验的 exe 来自同一个 HTTP 响应、同一信任域。
    /// 攻击者只要能发布 Release，就能同时替换 exe 和说明里的哈希，校验形同虚设。
    /// 该哈希只能防"下载损坏"，防不了任何恶意篡改。
    ///
    /// 真正的信任锚是代码签名证书：验证签名者指纹固定后，
    /// 私钥被保存在发证机构，攻击者仅凭仓库权限无法伪造。
    ///
    /// 降级策略：未配置期望指纹时不强制校验签名，保持向后兼容；
    /// 配置后则严格校验，不通过则拒绝更新。
    /// </summary>
    public static class SignatureVerifier
    {
        /// <summary>
        /// 期望的签名者证书指纹（SHA-1，去掉空格）。
        /// 留空表示未启用强制签名校验。
        ///
        /// 配置方式：与代码一同发布，把从证书管理工具导出的指纹填到这里。
        /// 例如自签名证书可用 certmgr.msc 或 PowerShell 查看。
        /// </summary>
        private const string EXPECTED_THUMBPRINT = "";

        /// <summary>是否已配置期望指纹（决定是否强制校验）</summary>
        public static bool EnforcementEnabled
        {
            get { return !string.IsNullOrWhiteSpace(EXPECTED_THUMBPRINT); }
        }

        /// <summary>
        /// 验证文件的 Authenticode 签名。
        /// 无论成功失败都不抛异常。
        /// </summary>
        public static SignatureCheckResult Verify(string filePath)
        {
            var result = new SignatureCheckResult();

            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    result.FailureReason = "文件不存在";
                    return result;
                }

                // X509Certificate2(string) 构造器会自动提取 PE 文件中的签名证书。
                // 注意：.NET Framework 中 X509Certificate2.CreateFromSignedFile 的
                // 返回类型是基类 X509Certificate，无法直接赋给 X509Certificate2，
                // 因此这里用构造器（它返回 X509Certificate2 本身）。
                using (var certificate = new X509Certificate2(filePath))
                {
                    if (certificate == null)
                    {
                        result.FailureReason = "文件未包含 Authenticode 签名";
                        return result;
                    }

                    result.IsSigned = true;
                    result.SignerSubject = certificate.Subject;
                    result.Thumbprint = certificate.Thumbprint;
                    result.Issuer = certificate.Issuer;
                    result.NotAfter = certificate.NotAfter;
                }

                // 配了期望指纹则必须严格比对
                if (EnforcementEnabled)
                {
                    string expected = NormalizeThumbprint(EXPECTED_THUMBPRINT);
                    string actual = NormalizeThumbprint(result.Thumbprint);

                    if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
                    {
                        result.IsSigned = false;
                        result.FailureReason = string.Format(
                            "签名者指纹不匹配（期望 {0}，实际 {1}）", expected, actual);
                    }
                }
            }
            catch (Exception ex)
            {
                // CreateFromSignedFile 对未签名文件抛 CryptographicException
                result.IsSigned = false;
                result.FailureReason = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 更新前的签名准入判断。
        /// 未启用强制校验时只要文件存在即放行（保持现有行为）。
        /// </summary>
        public static bool IsAcceptable(string filePath, out string description)
        {
            SignatureCheckResult check = Verify(filePath);
            description = check.Describe();

            if (!EnforcementEnabled)
            {
                // 未配置期望指纹：不强制要求签名，避免未签名版本无法自更新
                return true;
            }

            return check.IsSigned;
        }

        /// <summary>规范化指纹：去空格、转大写</summary>
        private static string NormalizeThumbprint(string thumbprint)
        {
            if (string.IsNullOrEmpty(thumbprint)) return string.Empty;
            return thumbprint.Replace(" ", string.Empty).ToUpperInvariant();
        }
    }
}
