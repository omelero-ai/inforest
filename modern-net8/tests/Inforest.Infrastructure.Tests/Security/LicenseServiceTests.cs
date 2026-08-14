using Inforest.Infrastructure.Security;

namespace Inforest.Infrastructure.Tests.Security;

public sealed class LicenseServiceTests
{
    [Fact]
    public void ParseExpiryDate_ZeroDate_ReturnsNull()
    {
        var result = LicenseService.ParseExpiryDate("00000000RESTO");

        Assert.Null(result);
    }

    [Fact]
    public void GetLicensedCount_InfoRest_LeePosicionLegacyCorrecta()
    {
        var licencia = "00000000000" + "15" + "0000000000";

        var result = LicenseService.GetLicensedCount(licencia, "INFOREST");

        Assert.Equal(15, result);
    }
}
