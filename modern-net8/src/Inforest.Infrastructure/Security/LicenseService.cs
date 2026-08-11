using Dapper;
using Inforest.Application.Interfaces;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Security;

/// <summary>
/// Estrategia inicial de licenciamiento:
/// - contrato Legacy THARDKEY/usp_Seg_cLientes
/// - validación física de HardKey pospuesta
/// </summary>
internal sealed class LicenseService : ILicenseService
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ILogger<LicenseService> _logger;

    public LicenseService(IDbConnectionFactory connectionFactory, ILogger<LicenseService> logger)
    {
        _connectionFactory = connectionFactory;
        _logger = logger;
    }

    public async Task<LicenseValidationResult> ValidarLicenciaAsync(LicenseValidationRequest request, CancellationToken cancellationToken = default)
    {
        if (!request.Habilitado)
            return LicenseValidationResult.Valida("CONFIG_DISABLED", "Validación de licencia deshabilitada por configuración.");

        if (string.IsNullOrWhiteSpace(request.RucEmpresa))
            return LicenseValidationResult.Invalida("LICENCIA_RUC_REQUERIDO", "Se requiere Inforest:Licensing:RucEmpresa para validar licencia.", "LEGACY_THARDKEY");

        try
        {
            await using var securityConnection = await _connectionFactory.CreateOpenConnectionAsync("INFSEGURIDAD", cancellationToken);
            var encodedRuc = LegacyLicenseCodec.EncodeAsciiTriplets(request.RucEmpresa);

            const string sql = """
                SELECT id, tLicencia, ISNULL(tCliente, '') AS tCliente
                FROM dbo.THARDKEY
                WHERE id = @Id
                """;

            var record = await securityConnection.QueryFirstOrDefaultAsync<HardKeyRecord>(new CommandDefinition(
                sql,
                new { Id = encodedRuc },
                cancellationToken: cancellationToken));

            if (record is null || string.IsNullOrWhiteSpace(record.tLicencia))
                return LicenseValidationResult.Invalida("LICENCIA_NO_ENCONTRADA", "Archivo de validación no encontrado en THARDKEY.", "LEGACY_THARDKEY");

            var decoded = LegacyLicenseCodec.DecodeAsciiTriplets(record.tLicencia);
            if (!decoded.EndsWith(request.RucEmpresa.Trim(), StringComparison.OrdinalIgnoreCase))
                return LicenseValidationResult.Invalida("LICENCIA_RUC_INVALIDO", "La licencia no corresponde al RUC configurado.", "LEGACY_THARDKEY");

            var expiry = ParseExpiryDate(decoded);
            if (expiry is not null)
            {
                var serverDate = await securityConnection.ExecuteScalarAsync<DateTime>(new CommandDefinition(
                    "SELECT GETDATE()",
                    cancellationToken: cancellationToken));

                var daysRemaining = (expiry.Value.Date - serverDate.Date).Days;
                if (daysRemaining < -4)
                    return LicenseValidationResult.Invalida("LICENCIA_VENCIDA", "La licencia ha caducado.", "LEGACY_THARDKEY");

                var warning = daysRemaining switch
                {
                    < 0 => $"La licencia está en prórroga ({Math.Abs(daysRemaining)} día(s)).",
                    <= 4 => $"La licencia vence en {daysRemaining} día(s).",
                    _ => request.OmitirHardwareLegacy
                        ? "Validación física de HardKey pendiente de migración; se valida contrato SQL legado."
                        : null
                };

                var maxLicenses = GetLicensedCount(decoded, request.ModuloEjecutable);
                if (maxLicenses == 0)
                    return LicenseValidationResult.Invalida("LICENCIA_SIN_CUPO", "No hay licencias disponibles para el módulo.", "LEGACY_THARDKEY");

                if (maxLicenses != 99)
                {
                    await using var mainConnection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
                    var concurrencyResult = await mainConnection.QueryFirstOrDefaultAsync<int>(new CommandDefinition(
                        "usp_Seg_cLientes",
                        new
                        {
                            tModulo = request.ModuloEjecutable.Trim().ToUpperInvariant(),
                            tDbName = request.BaseDatos.Trim().ToUpperInvariant(),
                            nCantidad = maxLicenses
                        },
                        commandType: System.Data.CommandType.StoredProcedure,
                        cancellationToken: cancellationToken));

                    if (concurrencyResult == 0)
                        return LicenseValidationResult.Invalida("LICENCIA_EXCESO_CONCURRENCIA", "Se superó el máximo de licencias concurrentes para el módulo.", "LEGACY_THARDKEY");
                }

                return LicenseValidationResult.Valida("LEGACY_THARDKEY", warning, daysRemaining);
            }

            return LicenseValidationResult.Valida("LEGACY_THARDKEY", "Licencia de vigencia ilimitada detectada.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error validando licencia para {Modulo}", request.ModuloEjecutable);
            return LicenseValidationResult.Invalida("LICENCIA_ERROR_VALIDACION", "No se pudo validar la licencia legacy.", "LEGACY_THARDKEY");
        }
    }

    internal static DateTime? ParseExpiryDate(string decodedLicense)
    {
        if (string.IsNullOrWhiteSpace(decodedLicense) || decodedLicense.Length < 8)
            return null;

        var value = decodedLicense[..8];
        if (value == "00000000")
            return null;

        var day = int.Parse(value[..2]);
        var month = int.Parse(value.Substring(2, 2));
        var year = int.Parse(value.Substring(4, 4));
        return new DateTime(year, month, day);
    }

    internal static int GetLicensedCount(string decodedLicense, string executableName)
    {
        if (string.IsNullOrWhiteSpace(decodedLicense))
            return 0;

        var key = executableName.Trim().ToUpperInvariant();
        var start = key switch
        {
            "INFOREST" => 11,
            "ADICION" => 13,
            "CHEFCONTROL" => 15,
            "DESPACHADOR" => 17,
            "ANFITRIONA" => 19,
            "TRANFERENCIA" => 21,
            "ALMACEN" => 23,
            "COSTOS" => 25,
            "INFHOTEL" => 27,
            "EVENTOS" => 29,
            "PROMOCIONES" => 31,
            _ => -1
        };

        if (start < 0 || decodedLicense.Length < start + 2)
            return 0;

        return int.TryParse(decodedLicense.Substring(start, 2), out var value)
            ? value
            : 0;
    }

    private sealed record HardKeyRecord(string id, string tLicencia, string tCliente);
}
