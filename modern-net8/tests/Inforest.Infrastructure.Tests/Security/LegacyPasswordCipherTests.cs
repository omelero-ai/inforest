using Inforest.Infrastructure.Security;

namespace Inforest.Infrastructure.Tests.Security;

public sealed class LegacyPasswordCipherTests
{
    [Fact]
    public void EncryptDecrypt_RoundTripConPasswordLegacy()
    {
        const string plainText = "SECRETO";

        var encrypted = LegacyPasswordCipher.Encrypt(plainText);
        var decrypted = LegacyPasswordCipher.Decrypt(encrypted);

        Assert.NotEqual(plainText, encrypted);
        Assert.Equal(plainText, decrypted);
    }
}
