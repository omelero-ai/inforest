using Dapper;
using Inforest.Application.Delivery;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Delivery;

namespace Inforest.Infrastructure.Delivery;

/// <summary>
/// Repositorio Dapper para tarjetas RFID/proximidad.
/// Legacy: <c>FrmTarjetaAproximidad.frm</c>, <c>FrmTarjetaAproximidadDetalle.frm</c>,
/// tablas <c>TTARJETASRFID</c> y <c>TMOVIMIENTOTARJETASRFID</c>.
/// </summary>
internal sealed class TarjetaProximidadRepository : ITarjetaProximidadRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public TarjetaProximidadRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<TarjetaProximidad>> ObtenerTodasAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT t1.Id,
                   t1.CodidoRFID AS CodigoRfid,
                   t1.fRegistro AS FechaRegistro,
                   t1.UsuarioRegistro,
                   t1.FechaModificion AS FechaModificacion,
                   t1.UsuarioModificion AS UsuarioModificacion,
                   t1.Estado,
                   CAST(ISNULL(t1.MontoDisponible, 0) AS decimal(18,2)) AS MontoDisponible,
                   t1.Descripcion,
                   t1.CodigoCliente,
                   LTRIM(RTRIM(ISNULL(t2.tApellido, '') + ' ' + ISNULL(t2.tNombre, ''))) AS NombreCliente
            FROM TTARJETASRFID t1
            INNER JOIN TDELIVERY t2 ON t1.CodigoCliente = t2.tCodigoDelivery
            ORDER BY t1.fRegistro, t1.CodidoRFID
            """;

        return (await conn.QueryAsync<TarjetaProximidad>(new CommandDefinition(sql, cancellationToken: ct))).AsList();
    }

    public async Task<TarjetaProximidad?> ObtenerPorCodigoAsync(string codigoRfid, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT TOP 1
                   t1.Id,
                   t1.CodidoRFID AS CodigoRfid,
                   t1.fRegistro AS FechaRegistro,
                   t1.UsuarioRegistro,
                   t1.FechaModificion AS FechaModificacion,
                   t1.UsuarioModificion AS UsuarioModificacion,
                   t1.Estado,
                   CAST(ISNULL(t1.MontoDisponible, 0) AS decimal(18,2)) AS MontoDisponible,
                   t1.Descripcion,
                   t1.CodigoCliente,
                   LTRIM(RTRIM(ISNULL(t2.tApellido, '') + ' ' + ISNULL(t2.tNombre, ''))) AS NombreCliente
            FROM TTARJETASRFID t1
            INNER JOIN TDELIVERY t2 ON t1.CodigoCliente = t2.tCodigoDelivery
            WHERE t1.CodidoRFID = @CodigoRfid
            """;

        return await conn.QueryFirstOrDefaultAsync<TarjetaProximidad>(
            new CommandDefinition(sql, new { CodigoRfid = codigoRfid }, cancellationToken: ct));
    }

    public async Task<bool> InsertarAsync(TarjetaProximidad tarjeta, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO TTARJETASRFID
                (CodidoRFID, fRegistro, UsuarioRegistro, FechaModificion, UsuarioModificion,
                 Estado, MontoDisponible, Descripcion, CodigoCliente)
            VALUES
                (@CodigoRfid, @FechaRegistro, @UsuarioRegistro, @FechaModificacion, @UsuarioModificacion,
                 @Estado, @MontoDisponible, @Descripcion, @CodigoCliente)
            """;

        return await conn.ExecuteAsync(new CommandDefinition(sql, tarjeta, cancellationToken: ct)) > 0;
    }

    public async Task<bool> ActualizarAsync(TarjetaProximidad tarjeta, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE TTARJETASRFID
               SET Descripcion = @Descripcion,
                   MontoDisponible = @MontoDisponible,
                   CodigoCliente = @CodigoCliente,
                   Estado = @Estado,
                   UsuarioModificion = @UsuarioModificacion,
                   FechaModificion = @FechaModificacion
             WHERE CodidoRFID = @CodigoRfid
            """;

        return await conn.ExecuteAsync(new CommandDefinition(sql, tarjeta, cancellationToken: ct)) > 0;
    }

    public async Task<IReadOnlyList<MovimientoTarjetaProximidad>> ObtenerUltimosMovimientosAsync(string codigoRfid, int cantidad = 10, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT TOP (@Cantidad)
                   CAST(fRegistro AS datetime2) AS FechaRegistro,
                   CAST(ISNULL(MontoIngreso, 0) AS decimal(18,2)) AS MontoIngreso,
                   CAST(ISNULL(MontoSalida, 0) AS decimal(18,2)) AS MontoSalida,
                   CAST(ISNULL(MontoAnterior, 0) AS decimal(18,2)) AS MontoAnterior,
                   CAST(ISNULL(MontoFinal, 0) AS decimal(18,2)) AS MontoFinal
            FROM TMOVIMIENTOTARJETASRFID
            WHERE CodidoRFID = @CodigoRfid
            ORDER BY fRegistro DESC
            """;

        return (await conn.QueryAsync<MovimientoTarjetaProximidad>(
            new CommandDefinition(sql, new { CodigoRfid = codigoRfid, Cantidad = cantidad }, cancellationToken: ct))).AsList();
    }
}
