namespace Inforest.Infrastructure.Security;

/// <summary>
/// Compatibilidad temporal con ClsSeguridad.cls (XOR+César).
/// Solo para lectura/transición.
/// </summary>
internal static class LegacyPasswordCipher
{
    private const string DefaultPassword = "123456789abcde";

    public static string Encrypt(string plainText, string? password = null)
    {
        var key = string.IsNullOrEmpty(password) ? DefaultPassword : password;
        return Transform(plainText, key, encrypt: true);
    }

    public static string Decrypt(string cipherText, string? password = null)
    {
        var key = string.IsNullOrEmpty(password) ? DefaultPassword : password;
        return Transform(cipherText, key, encrypt: false);
    }

    private static string Transform(string input, string password, bool encrypt)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        var buffer = new char[input.Length];

        for (var index = 0; index < input.Length; index++)
        {
            var c1 = input[index];
            var c2 = password[index % password.Length];
            var value = encrypt
                ? c1 + c2 + 64
                : c1 - c2 - 64;

            if (encrypt && value > 255)
                value -= 256;

            if (!encrypt && value < 0)
                value += 256;

            buffer[index] = (char)value;
        }

        return new string(buffer);
    }
}
