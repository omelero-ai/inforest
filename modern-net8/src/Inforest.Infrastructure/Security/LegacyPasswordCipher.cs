namespace Inforest.Infrastructure.Security;

/// <summary>
/// Compatibilidad temporal con la conversión legacy de credenciales.
/// Legacy real para frmAcceso: modProcedimiento.bas (Encapsula/Desencapsula).
/// </summary>
internal static class LegacyPasswordCipher
{
    private const char FallbackCharacter = '´';
    private const string PlainAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 Ñ.*,-·$&/()=?¿!#@ºª\\";
    private const string CipherAlphabet = "!·¿ª$%&/(,.-)=?º{Ç}*<;+[]^¡5\\@>:#_¬|q\u00951njklmr0abcdefghpoz";

    public static string Encrypt(string plainText, string? password = null)
    {
        EnsurePasswordOverrideIsNotUsed(password);
        return Transform(plainText, encrypt: true);
    }

    public static string Decrypt(string cipherText, string? password = null)
    {
        EnsurePasswordOverrideIsNotUsed(password);
        return Transform(cipherText, encrypt: false);
    }

    private static string Transform(string input, bool encrypt)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        var buffer = new char[input.Length];

        for (var index = 0; index < input.Length; index++)
        {
            var mapped = MapCharacter(char.ToUpperInvariant(input[index]), encrypt);
            buffer[index] = mapped;
        }

        return new string(buffer);
    }

    private static char MapCharacter(char value, bool encrypt)
    {
        var source = encrypt ? PlainAlphabet : CipherAlphabet;
        var target = encrypt ? CipherAlphabet : PlainAlphabet;

        var position = source.IndexOf(value);
        return position >= 0
            ? target[position]
            : FallbackCharacter;
    }

    private static void EnsurePasswordOverrideIsNotUsed(string? password)
    {
        if (!string.IsNullOrEmpty(password))
            throw new ArgumentException("La conversión de credenciales legacy de frmAcceso no admite clave personalizada.", nameof(password));
    }
}
