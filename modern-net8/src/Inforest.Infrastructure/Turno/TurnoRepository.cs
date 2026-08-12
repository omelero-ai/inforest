using System.Reflection;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Turno;
using Inforest.Domain.Entities.Configuracion;
using TurnoEntity = Inforest.Domain.Entities.Configuracion.Turno;

namespace Inforest.Infrastructure.Turno;

internal sealed class TurnoRepository : ITurnoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public TurnoRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    public async Task<TurnoEntity?> ObtenerTurnoActualAsync(string codigoCaja, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT TOP 1
                tTurno AS CodigoTurno,
                tCaja AS CodigoCaja,
                ISNULL(tUsuario, '') AS CodigoUsuario,
                fInicial AS FechaApertura,
                fFinal AS FechaCierre,
                CAST(ISNULL(fInicial, GETDATE()) AS date) AS FechaDiaContable,
                CAST(ISNULL(nMontoIN, 0) AS decimal(18,2)) AS MontoInicial,
                CAST(ISNULL(nMontoFN, 0) AS decimal(18,2)) AS MontoFinal,
                CAST(ISNULL(lCierre, 0) AS bit) AS Cerrado
            FROM MTURNO
            WHERE tCaja = @caja AND lCierre = 0
            ORDER BY fInicial DESC
            """;

        var record = await connection.QueryFirstOrDefaultAsync<TurnoRecord>(sql, new { caja = codigoCaja });
        return record is null ? null : Map(record);
    }

    public async Task<bool> InsertarAsync(TurnoEntity turno, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO MTURNO (tTurno, tCaja, fInicial, tUsuario, lCierre, nMontoIN)
            VALUES (@CodigoTurno, @CodigoCaja, @FechaApertura, @CodigoUsuario, 0, @MontoInicial)
            """;

        var affected = await connection.ExecuteAsync(sql, new
        {
            turno.CodigoTurno,
            turno.CodigoCaja,
            FechaApertura = turno.FechaApertura,
            CodigoUsuario = turno.CodigoUsuario,
            turno.MontoInicial
        });

        return affected > 0;
    }

    public async Task<bool> CerrarAsync(string codigoTurno, decimal montoFinal, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE MTURNO
            SET lCierre = 1,
                fFinal = GETDATE(),
                nMontoFN = @monto
            WHERE tTurno = @codigo AND ISNULL(lCierre, 0) = 0
            """;

        var affected = await connection.ExecuteAsync(sql, new { codigo = codigoTurno, monto = montoFinal });
        return affected > 0;
    }

    public async Task<IReadOnlyList<TurnoEntity>> ObtenerHistorialAsync(string codigoCaja, DateTime desde, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT
                tTurno AS CodigoTurno,
                tCaja AS CodigoCaja,
                ISNULL(tUsuario, '') AS CodigoUsuario,
                fInicial AS FechaApertura,
                fFinal AS FechaCierre,
                CAST(ISNULL(fInicial, GETDATE()) AS date) AS FechaDiaContable,
                CAST(ISNULL(nMontoIN, 0) AS decimal(18,2)) AS MontoInicial,
                CAST(ISNULL(nMontoFN, 0) AS decimal(18,2)) AS MontoFinal,
                CAST(ISNULL(lCierre, 0) AS bit) AS Cerrado
            FROM MTURNO
            WHERE tCaja = @caja AND fInicial >= @desde
            ORDER BY fInicial DESC
            """;

        var records = await connection.QueryAsync<TurnoRecord>(sql, new { caja = codigoCaja, desde });
        return records.Select(Map).ToList();
    }

    private static TurnoEntity Map(TurnoRecord record)
    {
        var turno = Domain.Entities.Configuracion.Turno.Abrir(
            record.CodigoTurno,
            record.CodigoCaja,
            record.CodigoUsuario,
            record.FechaDiaContable,
            record.MontoInicial);

        SetProperty(turno, nameof(Domain.Entities.Configuracion.Turno.FechaApertura), record.FechaApertura);

        if (record.Cerrado)
        {
            turno.Cerrar(record.MontoFinal);
            if (record.FechaCierre.HasValue)
                SetProperty(turno, nameof(Domain.Entities.Configuracion.Turno.FechaCierre), record.FechaCierre);
        }

        return turno;
    }

    private static void SetProperty<T>(object target, string propertyName, T value)
    {
        var property = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException($"No se encontró la propiedad {propertyName}.");
        property.SetValue(target, value);
    }
}
