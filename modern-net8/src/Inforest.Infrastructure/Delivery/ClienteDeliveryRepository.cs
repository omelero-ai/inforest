using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Infrastructure.Delivery;

/// <summary>
/// Repositorio Dapper para <see cref="ClienteDelivery"/>.
/// <para>
/// Legacy: tabla <c>TDELIVERY</c>, <c>TDELIVERYCLIENTE</c>, <c>TDELIVERYINVITADO</c>.
/// Vistas: <c>vDelivery</c>, <c>vDistrito</c>, <c>vZona</c>.
/// </para>
/// </summary>
internal sealed class ClienteDeliveryRepository : IClienteDeliveryRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public ClienteDeliveryRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    /// <inheritdoc />
    public async Task<ClienteDelivery?> ObtenerPorCodigoAsync(string codigoDelivery, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoDelivery AS CodigoDelivery, tTipoCliente AS TipoCliente,
                   tApellido AS Apellido, tNombre AS Nombre, tDireccion AS Direccion,
                   tTelefono AS Telefono, tReferencia AS Referencia,
                   tZona AS CodigoZona, tDistrito AS CodigoDistrito,
                   tCodigoCliente AS CodigoCliente, tCodigoTarjeta AS CodigoTarjeta,
                   tNumeroTarjeta AS NumeroTarjeta, tTipoIdentidad AS TipoIdentidad,
                   tIdentidad AS NumeroIdentidad, tObservacion AS Observacion,
                   tEmail AS Email, fNacimiento AS FechaNacimiento, lActivo AS Activo
            FROM TDELIVERY
            WHERE tCodigoDelivery = @CodigoDelivery
            """;
        return await conn.QueryFirstOrDefaultAsync<ClienteDelivery>(sql, new { CodigoDelivery = codigoDelivery });
    }

    /// <inheritdoc />
    public async Task<ClienteDelivery?> ObtenerPorTelefonoAsync(string telefono, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT TOP 1
                   tCodigoDelivery AS CodigoDelivery, tTipoCliente AS TipoCliente,
                   tApellido AS Apellido, tNombre AS Nombre, tDireccion AS Direccion,
                   tTelefono AS Telefono, tReferencia AS Referencia,
                   tZona AS CodigoZona, tDistrito AS CodigoDistrito,
                   lActivo AS Activo
            FROM TDELIVERY
            WHERE tTelefono = @Telefono AND lActivo = 1
            """;
        return await conn.QueryFirstOrDefaultAsync<ClienteDelivery>(sql, new { Telefono = telefono });
    }

    /// <inheritdoc />
    public async Task<ClienteDelivery?> ObtenerPorIdentidadAsync(string tipoIdentidad, string numeroIdentidad, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT TOP 1
                   tCodigoDelivery AS CodigoDelivery, tTipoCliente AS TipoCliente,
                   tApellido AS Apellido, tNombre AS Nombre, tDireccion AS Direccion,
                   tTelefono AS Telefono, tZona AS CodigoZona, tDistrito AS CodigoDistrito,
                   tTipoIdentidad AS TipoIdentidad, tIdentidad AS NumeroIdentidad,
                   lActivo AS Activo
            FROM TDELIVERY
            WHERE tTipoIdentidad = @TipoIdentidad AND tIdentidad = @NumeroIdentidad
            """;
        return await conn.QueryFirstOrDefaultAsync<ClienteDelivery>(
            sql, new { TipoIdentidad = tipoIdentidad, NumeroIdentidad = numeroIdentidad });
    }

    /// <inheritdoc />
    public async Task<IEnumerable<ClienteDelivery>> BuscarAsync(
        string? apellido, string? nombre, string? telefono, string? codigoDistrito, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigoDelivery AS CodigoDelivery, tTipoCliente AS TipoCliente,
                   tApellido AS Apellido, tNombre AS Nombre, tDireccion AS Direccion,
                   tTelefono AS Telefono, tZona AS CodigoZona, tDistrito AS CodigoDistrito,
                   lActivo AS Activo
            FROM vDelivery
            WHERE (@Apellido IS NULL OR tApellido LIKE '%' + @Apellido + '%')
              AND (@Nombre IS NULL OR tNombre LIKE '%' + @Nombre + '%')
              AND (@Telefono IS NULL OR tTelefono LIKE '%' + @Telefono + '%')
              AND (@Distrito IS NULL OR tDistrito = @Distrito)
              AND lActivo = 1
            ORDER BY tApellido, tNombre
            """;
        return await conn.QueryAsync<ClienteDelivery>(sql,
            new { Apellido = apellido, Nombre = nombre, Telefono = telefono, Distrito = codigoDistrito });
    }

    /// <inheritdoc />
    public async Task<IEnumerable<ClienteDelivery>> ObtenerFrecuentesAsync(int top = 20, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        var sql = $"""
            SELECT TOP {top}
                   tCodigoDelivery AS CodigoDelivery, tTipoCliente AS TipoCliente,
                   tApellido AS Apellido, tNombre AS Nombre, tDireccion AS Direccion,
                   tTelefono AS Telefono, tZona AS CodigoZona, tDistrito AS CodigoDistrito,
                   lActivo AS Activo
            FROM vEstadoFrecuente
            WHERE lActivo = 1
            ORDER BY tApellido, tNombre
            """;
        return await conn.QueryAsync<ClienteDelivery>(sql);
    }

    /// <inheritdoc />
    public async Task InsertarAsync(ClienteDelivery cliente, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            INSERT INTO TDELIVERY
                (tCodigoDelivery, tTipoCliente, tApellido, tNombre, tDireccion,
                 tTelefono, tReferencia, tZona, tDistrito, tCodigoCliente,
                 tCodigoTarjeta, tNumeroTarjeta, tTipoIdentidad, tIdentidad,
                 tObservacion, tEmail, fNacimiento, lActivo)
            VALUES
                (@CodigoDelivery, @TipoCliente, @Apellido, @Nombre, @Direccion,
                 @Telefono, @Referencia, @CodigoZona, @CodigoDistrito, @CodigoCliente,
                 @CodigoTarjeta, @NumeroTarjeta, @TipoIdentidad, @NumeroIdentidad,
                 @Observacion, @Email, @FechaNacimiento, @Activo)
            """;
        await conn.ExecuteAsync(sql, cliente);
    }

    /// <inheritdoc />
    public async Task ActualizarAsync(ClienteDelivery cliente, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            UPDATE TDELIVERY SET
                tApellido = @Apellido, tNombre = @Nombre, tDireccion = @Direccion,
                tTelefono = @Telefono, tReferencia = @Referencia, tZona = @CodigoZona,
                tDistrito = @CodigoDistrito, tObservacion = @Observacion, tEmail = @Email
            WHERE tCodigoDelivery = @CodigoDelivery
            """;
        await conn.ExecuteAsync(sql, cliente);
    }

    /// <inheritdoc />
    public async Task ActualizarFotoAsync(string codigoDelivery, byte[] foto, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);

        var parameters = new DynamicParameters();
        parameters.Add("@tCodigo", codigoDelivery, DbType.StringFixedLength, ParameterDirection.Input, 10);
        parameters.Add("@oFoto", foto, DbType.Binary, ParameterDirection.Input, foto.Length);

        await _spExecutor.ExecuteAsync(conn, "sp_UpdFotoDelivery", parameters, cancellationToken: ct);
    }
}
