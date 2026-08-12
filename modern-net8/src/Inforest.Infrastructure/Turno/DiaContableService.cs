using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Turno;

namespace Inforest.Infrastructure.Turno;

internal sealed class DiaContableService : IDiaContableService
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public DiaContableService(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    public async Task<DateTime> ObtenerDiaContableAsync(bool esAutomatico, string horaCierre, string usuario, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        var parameters = new DynamicParameters();
        parameters.Add("@lDiaContable", esAutomatico);
        parameters.Add("@sHoraCierre", horaCierre);
        parameters.Add("@tUsuario", usuario);
        parameters.Add("@fDiaContable", dbType: System.Data.DbType.Date, direction: System.Data.ParameterDirection.Output);

        await _spExecutor.ExecuteAsync(connection, "usp_GenObtieneDiaContable", parameters, cancellationToken: ct);

        var fecha = parameters.Get<DateTime?>("@fDiaContable");
        return (fecha ?? DateTime.Today).Date;
    }

    public async Task<DateTime> ObtenerFechaServidorAsync(CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = "SELECT CAST(GETDATE() AS DATE) AS fFecha";
        return await connection.QueryFirstAsync<DateTime>(new CommandDefinition(sql, cancellationToken: ct));
    }

    public async Task InsertarDiaContableAsync(DateTime fecha, string usuario, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        await _spExecutor.ExecuteAsync(
            connection,
            "usp_GenInsertarDiaContable",
            new { tUsuario = usuario, fDiaContable = fecha.Date },
            cancellationToken: ct);
    }
}
