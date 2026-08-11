using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Security;

/// <summary>
/// Auditoría mínima equivalente sobre INFSEGURIDAD.
/// Legacy: ups_Aud_RegistraIngSalUsuario + ups_Aud_RegistraMovimientos.
/// </summary>
internal sealed class AuditoriaService : IAuditoriaService
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ILogger<AuditoriaService> _logger;

    public AuditoriaService(IDbConnectionFactory connectionFactory, ILogger<AuditoriaService> logger)
    {
        _connectionFactory = connectionFactory;
        _logger = logger;
    }

    public async Task<Result<long>> RegistrarIngresoAsync(RegistroAccesoAuditoriaRequest request, CancellationToken cancellationToken = default)
    {
        try
        {
            await using var connection = await _connectionFactory.CreateOpenConnectionAsync("INFSEGURIDAD", cancellationToken);
            var parameters = CreateAccessParameters(request);
            parameters.Add("@nCorrelativo", dbType: System.Data.DbType.Int64, direction: System.Data.ParameterDirection.Output);

            await connection.ExecuteAsync(new CommandDefinition(
                "ups_Aud_RegistraIngSalUsuario",
                parameters,
                commandType: System.Data.CommandType.StoredProcedure,
                cancellationToken: cancellationToken));

            var correlativo = parameters.Get<long>("@nCorrelativo");
            return Result.Ok(correlativo);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "No se pudo registrar ingreso de auditoría para {Usuario}", request.Usuario);
            return Result.Fail<long>("No se pudo registrar el ingreso de auditoría.", "SEGURIDAD_AUDITORIA_INGRESO");
        }
    }

    public async Task<Result> RegistrarSalidaAsync(RegistroAccesoAuditoriaRequest request, CancellationToken cancellationToken = default)
    {
        try
        {
            await using var connection = await _connectionFactory.CreateOpenConnectionAsync("INFSEGURIDAD", cancellationToken);
            var parameters = CreateAccessParameters(request);
            parameters.Add("@nCorrelativo", request.CorrelativoAcceso, dbType: System.Data.DbType.Int64, direction: System.Data.ParameterDirection.Output);

            await connection.ExecuteAsync(new CommandDefinition(
                "ups_Aud_RegistraIngSalUsuario",
                parameters,
                commandType: System.Data.CommandType.StoredProcedure,
                cancellationToken: cancellationToken));

            return Result.Ok();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "No se pudo registrar salida de auditoría para {Usuario}", request.Usuario);
            return Result.Fail("No se pudo registrar la salida de auditoría.", "SEGURIDAD_AUDITORIA_SALIDA");
        }
    }

    public async Task<Result> RegistrarMovimientoAsync(MovimientoAuditoriaRequest request, CancellationToken cancellationToken = default)
    {
        try
        {
            await using var connection = await _connectionFactory.CreateOpenConnectionAsync("INFSEGURIDAD", cancellationToken);
            var parameters = new DynamicParameters();
            parameters.Add("@vch_nCorrelativo", DateTime.UtcNow.Ticks % int.MaxValue);
            parameters.Add("@vch_nCorrelativoCampo", 1);
            parameters.Add("@vch_tModulo", request.Modulo);
            parameters.Add("@vch_tBaseDatos", request.BaseDatos);
            parameters.Add("@vch_tTabla", request.Tabla);
            parameters.Add("@vch_tAccion", request.Accion);
            parameters.Add("@vch_tUsuario", request.Usuario);
            parameters.Add("@vch_tId_RegistroUnico", request.IdRegistroUnico);
            parameters.Add("@vch_tId_RegistroDetalle", request.IdRegistroDetalle);
            parameters.Add("@vch_tCampo", request.Campo);
            parameters.Add("@vch_tValorAnterior", request.ValorAnterior);
            parameters.Add("@vch_tValorActual", request.ValorActual);

            await connection.ExecuteAsync(new CommandDefinition(
                "ups_Aud_RegistraMovimientos",
                parameters,
                commandType: System.Data.CommandType.StoredProcedure,
                cancellationToken: cancellationToken));

            return Result.Ok();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "No se pudo registrar movimiento de auditoría para {Tabla}", request.Tabla);
            return Result.Fail("No se pudo registrar el movimiento de auditoría.", "SEGURIDAD_AUDITORIA_MOVIMIENTO");
        }
    }

    private static DynamicParameters CreateAccessParameters(RegistroAccesoAuditoriaRequest request)
    {
        var parameters = new DynamicParameters();
        parameters.Add("@tTipo", request.Tipo);
        parameters.Add("@tBaseDatos", request.BaseDatos);
        parameters.Add("@tModulo", request.Modulo);
        parameters.Add("@tUsuario", request.Usuario);
        parameters.Add("@nCorrelativoUsuario", request.CorrelativoAcceso);
        return parameters;
    }
}
