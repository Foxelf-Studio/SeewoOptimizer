using System;
using System.IO;
using System.Security.Cryptography;

namespace SeewoOpt.Services
{
    /// <summary>签名验证结果</summary>
    public class SignatureCheckResult
    {
        /// <summary>签名是否有效</summary>
        public bool IsValid { get; set; }

        /// <summary>失败时的原因说明</summary>
        public string FailureReason { get; set; }

        /// <summary>可读的结论描述</summary>
        public string Describe()
        {
            if (IsValid)
                return "签名有效（RSA-3072 / PKCS#1 v1.5 / SHA-256）";

            return "签名无效" + (string.IsNullOrEmpty(FailureReason) ? "" : "：" + FailureReason);
        }
    }

    /// <summary>
    /// 发布文件签名验证（RSA-3072 / PKCS#1 v1.5 / SHA-256）。
    ///
    /// 【为什么不用 Authenticode】
    /// Authenticode 需要向 CA 机构购买代码签名证书（每年数百至数千元），
    /// 其成本主要来自 CA 对申请者身份的核验与时间戳存证服务。
    /// 但本工具真正要防的威胁只有一个：仓库账号被入侵后，
    /// 攻击者借发布权限投递恶意 exe。
    /// 防住这个威胁只需要"只有发布者能生成合法更新"这一能力，
    /// 而这一能力 RSA 私钥即可提供，无需付费。
    ///
    /// 【信任模型】
    /// - 私钥（private_key.pem）只存在于发布者本机，绝不提交仓库、不上传任何在线服务
    /// - 公钥（下方 PUBLIC_KEY_BASE64）硬编码在本文件中，随程序分发给所有用户
    /// - 发布流程：用私钥对 exe 签名生成 exe.sig → exe 与 exe.sig 一起上传 Release
    /// - 验证流程：程序用内置公钥验证签名 → 通过才替换自身
    ///
    /// 攻击者即使获得 GitHub 账号完全控制权，没有私钥也无法生成能通过验证的更新。
    /// 这与 Authenticode 防的是同一个威胁，安全性等价，成本为零。
    ///
    /// 【为什么用 PKCS#1 v1.5 而不是 PSS】
    /// 实测 .NET Framework 4.7.2 的 RSA 实现是 RSACryptoServiceProvider（旧式 CSP），
    /// 其 PSS 验签路径会抛"指定的填充模式对于此算法无效"（签名可以，验签不行）。
    /// 另实测本机 RSACng 类型不可用。PKCS#1 v1.5 验签在 net472 上工作正常，
    /// OpenSSL 也支持，是 net472 上的可行选择。
    /// 对代码签名这类一次性发布场景，PKCS#1 v1.5 的安全性无实质问题。
    ///
    /// 【为什么私钥可以不放在 HSM/TPM】
    /// 攻击者要利用的窗口期很短（仅在拥有推送权限期间），且签名是发布前
    /// 一次性动作而非持续行为。当前私钥保存在 G:\SeewoBuild\signing\，
    /// 已排除出版本控制。若要进一步加固，可移至 U 盾或仅在发布时挂载。
    /// </summary>
    public static class SignatureVerifier
    {
        /// <summary>
        /// 发布者公钥（RSA-3072，SubjectPublicKeyInfo / X.509 格式，Base64）。
        /// 重新生成密钥对后需替换本常量，导出命令：
        ///   openssl rsa -in private_key.pem -pubout -outform DER | base64 -w 0
        /// </summary>
        private const string PUBLIC_KEY_BASE64 =
            "MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIBigKCAYEAoH6IPFXuLtpBcuuJaWNq" +
            "DiskX1VjcFiDEOZvfBl+SVppZ29YL+YtRp0WN37ZTAHJJuEK+Pam7S2ylTWaZ35P" +
            "DRq0gN4MauSKQJf75MZSazvBjh7g22t/SB1Kr/RzSxHqv4pejXoZtwHbtA8uovcP" +
            "kFF/l8Cs7ZnLMwP9VZUTpD8qC/HimWHt8sGVkF+837FB30o5yYbEGQZHuksKLI1Y" +
            "119K0vs98cT7sj+voy6Y6wX6cMseeHUTBPfAeWGo9LtS6FXwTaDCoPCu1PfVwHF6" +
            "bzg8BcJ2eDZqE2Nnln6ukcc8F4dnxCDURPIrF2WVklSvho/HhJn+ft6veNRFf7bY" +
            "2WKKZbpEpxlr7Acw1oPZyg6lhDg9oLZDVyGCySiTz806ZDYzLiQSjg4JvUnAenYk" +
            "JqwlSOKBPqSbeI6nj6NtwTCIyGcreizVSCwDAVrZra7+gNLkR2B8x9CRyg2TKUjE" +
            "9j+3dG18AJGQTWLo4YsvgSTZhwZeviJeKTsT8rbG/xUHAgMBAAE=";


        /// <summary>签名文件的扩展名</summary>
        public const string SIGNATURE_EXTENSION = ".sig";

        /// <summary>公钥是否已配置</summary>
        public static bool EnforcementEnabled
        {
            get { return !string.IsNullOrWhiteSpace(PUBLIC_KEY_BASE64); }
        }

        /// <summary>
        /// 验证文件的发布签名。
        /// 签名针对文件原始字节本身而非其哈希——文件改动一个字节验证即失败。
        /// 无论成功失败都不抛异常。
        /// </summary>
        public static SignatureCheckResult Verify(string filePath, string signaturePath)
        {
            var result = new SignatureCheckResult();

            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    result.FailureReason = "待更新文件不存在";
                    return result;
                }

                if (string.IsNullOrEmpty(signaturePath) || !File.Exists(signaturePath))
                {
                    result.FailureReason = "签名文件不存在（后缀 " + SIGNATURE_EXTENSION + "）";
                    return result;
                }

                byte[] fileBytes = File.ReadAllBytes(filePath);
                byte[] signatureBytes = File.ReadAllBytes(signaturePath);

                bool ok;
                using (RSA rsa = RSA.Create())
                {
                    ImportPublicKey(rsa);
                    ok = rsa.VerifyData(
                        fileBytes,
                        signatureBytes,
                        HashAlgorithmName.SHA256,
                        RSASignaturePadding.Pkcs1);
                }

                result.IsValid = ok;
                if (!ok)
                    result.FailureReason = "签名与文件内容不匹配，文件可能已被篡改";
            }
            catch (CryptographicException ex)
            {
                result.FailureReason = "加密验证失败: " + ex.Message;
            }
            catch (Exception ex)
            {
                result.FailureReason = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 更新前的签名准入判断。未通过则拒绝更新。
        /// </summary>
        public static bool IsAcceptable(string filePath, out string description)
        {
            string signaturePath = filePath + SIGNATURE_EXTENSION;
            SignatureCheckResult check = Verify(filePath, signaturePath);
            description = check.Describe();
            return check.IsValid;
        }

        /// <summary>供发布脚本参考：拼出配套签名文件路径</summary>
        public static string GetSignaturePath(string filePath)
        {
            return filePath + SIGNATURE_EXTENSION;
        }

        /// <summary>导出公钥 Base64（供重新生成密钥时核对）</summary>
        public static string GetPublicKeyBase64()
        {
            return PUBLIC_KEY_BASE64;
        }

        /// <summary>
        /// 把 SubjectPublicKeyInfo 格式的公钥导入 RSA 实例。
        ///
        /// .NET Framework 4.7.2 没有 RSA.ImportSubjectPublicKeyInfo
        /// （那是 .NET Core 2.1+ 才引入的），因此手动解析 DER 结构：
        ///   SubjectPublicKeyInfo ::= SEQUENCE {
        ///       algorithmAlgorithmIdentifier  AlgorithmIdentifier,
        ///       subjectPublicKey              BIT STRING   -- 内含 RSAPublicKey
        ///   }
        ///   RSAPublicKey ::= SEQUENCE {
        ///       modulus                       INTEGER,
        ///       publicExponent                INTEGER
        ///   }
        /// 只取 modulus 与 exponent 导入 RSAParameters，algorithm 部分不参与验签。
        /// </summary>
        private static void ImportPublicKey(RSA rsa)
        {
            byte[] der = Convert.FromBase64String(PUBLIC_KEY_BASE64);

            int derLength = der.Length;
            int offset = 0;

            // 外层 SEQUENCE
            ReadTagAndLength(der, ref offset, out byte outerTag, out int outerContentLength);
            if (outerTag != 0x30)
                throw new CryptographicException("公钥格式错误：缺少外层 SEQUENCE");
            derLength = outerContentLength;

            // AlgorithmIdentifier（跳过）
            ReadTagAndLength(der, ref offset, out byte algTag, out int algLength);
            if (algTag != 0x30)
                throw new CryptographicException("公钥格式错误：缺少 AlgorithmIdentifier");
            offset += algLength;

            // subjectPublicKey 是 BIT STRING（0x03），首字节是未使用的 bit 数，必须为 0
            ReadTagAndLength(der, ref offset, out byte bitStringTag, out int bitStringLength);
            if (bitStringTag != 0x03)
                throw new CryptographicException("公钥格式错误：缺少 BIT STRING");
            if (der[offset] != 0x00)
                throw new CryptographicException("公钥格式错误：BIT STRING 未按整字节对齐");
            offset++;

            // BIT STRING 内部是 RSAPublicKey 的 DER SEQUENCE
            ReadTagAndLength(der, ref offset, out byte rsaSeqTag, out int rsaSeqLength);
            if (rsaSeqTag != 0x30)
                throw new CryptographicException("公钥格式错误：缺少 RSAPublicKey SEQUENCE");

            // modulus
            ReadTagAndLength(der, ref offset, out byte modTag, out int modLength);
            if (modTag != 0x02)
                throw new CryptographicException("公钥格式错误：modulus 不是 INTEGER");
            byte[] modulus = TrimLeadingZeros(der, offset, modLength);
            offset += modLength;

            // publicExponent
            ReadTagAndLength(der, ref offset, out byte expTag, out int expLength);
            if (expTag != 0x02)
                throw new CryptographicException("公钥格式错误：exponent 不是 INTEGER");
            byte[] exponent = TrimLeadingZeros(der, offset, expLength);
            offset += expLength;

            rsa.ImportParameters(new RSAParameters
            {
                Modulus = modulus,
                Exponent = exponent
            });
        }

        /// <summary>读取 DER 的标签与长度，返回内容起点相对于当前 offset 的长度</summary>
        private static void ReadTagAndLength(byte[] data, ref int offset,
            out byte tag, out int contentLength)
        {
            if (offset >= data.Length)
                throw new CryptographicException("公钥格式错误：数据意外结束");

            tag = data[offset++];

            if (offset >= data.Length)
                throw new CryptographicException("公钥格式错误：缺少长度字段");

            byte first = data[offset++];
            contentLength = first;

            // 长格式：首字节低 7 位表示后续长度字节数
            if ((first & 0x80) != 0)
            {
                int byteCount = first & 0x7F;
                if (byteCount == 0 || byteCount > 4 || offset + byteCount > data.Length)
                    throw new CryptographicException("公钥格式错误：长度字段非法");

                contentLength = 0;
                for (int i = 0; i < byteCount; i++)
                    contentLength = (contentLength << 8) | data[offset++];
            }
        }

        /// <summary>去除 DER INTEGER 前导 0x00（RSA 补符号位用）</summary>
        private static byte[] TrimLeadingZeros(byte[] data, int offset, int length)
        {
            int start = offset;
            int end = offset + length;

            while (start < end - 1 && data[start] == 0x00)
                start++;

            byte[] result = new byte[end - start];
            Array.Copy(data, start, result, 0, result.Length);
            return result;
        }
    }
}
