using System;
using System.Security.Cryptography;
using System.Text;

internal static class DpapiHelper
{
    // ユーザー単位で暗号化
    private static readonly byte[] Entropy =
        Encoding.UTF8.GetBytes("ClipboardUrlSender");

    public static string Encrypt(string plain)
    {
        if (string.IsNullOrEmpty(plain))
            return "";

        var bytes = Encoding.UTF8.GetBytes(plain);
        var encrypted = ProtectedData.Protect(
            bytes,
            Entropy,
            DataProtectionScope.CurrentUser);

        return Convert.ToBase64String(encrypted);
    }

    public static string Decrypt(string encrypted)
    {
        if (string.IsNullOrEmpty(encrypted))
            return "";

        try
        {
            var bytes = Convert.FromBase64String(encrypted);
            var decrypted = ProtectedData.Unprotect(
                bytes,
                Entropy,
                DataProtectionScope.CurrentUser);

            return Encoding.UTF8.GetString(decrypted);
        }
        catch
        {
            // 復号失敗（別ユーザー/別PCなど）
            return "";
        }
    }
}