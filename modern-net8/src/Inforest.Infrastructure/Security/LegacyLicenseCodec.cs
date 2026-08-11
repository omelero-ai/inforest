namespace Inforest.Infrastructure.Security;

/// <summary>
/// Compatibilidad con Codificar/DeCodificar de modSeguridadInfhotel.bas.
/// </summary>
internal static class LegacyLicenseCodec
{
    public static string EncodeAsciiTriplets(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        return string.Concat(value.Trim().ToUpperInvariant().Select(character => $"0{(int)character:000}".Substring(1)));
    }

    public static string DecodeAsciiTriplets(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var chunks = Enumerable.Range(0, value.Length / 3)
            .Select(index => value.Substring(index * 3, 3));

        return string.Concat(chunks.Select(chunk => (char)int.Parse(chunk)));
    }
}
