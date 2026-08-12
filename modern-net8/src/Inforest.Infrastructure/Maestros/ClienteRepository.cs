using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio Dapper de clientes.
/// Legacy: TCLIENTE / frmCliente.frm. Database: Inforest.
/// </summary>
internal sealed class ClienteRepository : IClienteRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public ClienteRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<Cliente>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoCliente AS CodigoCliente,
                   tEmpresa AS Empresa,
                   tIdentidad AS Identidad,
                   tDireccion AS Direccion,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro,
                   tCorreo AS Correo,
                   tTipoIdentidad AS TipoIdentidad,
                   tEnlace AS Enlace,
                   tTipoCliente AS TipoCliente,
                   tUbigeo AS Ubigeo,
                   tUrbanizacion AS Urbanizacion,
                   tTipoResponsable AS TipoResponsable
            FROM TCLIENTE
            ORDER BY tEmpresa
            """;
        return (await conn.QueryAsync<Cliente>(sql)).AsList();
    }

    public async Task<Cliente?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoCliente AS CodigoCliente,
                   tEmpresa AS Empresa,
                   tIdentidad AS Identidad,
                   tDireccion AS Direccion,
                   lActivo AS Activo,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro,
                   tCorreo AS Correo,
                   tTipoIdentidad AS TipoIdentidad,
                   tEnlace AS Enlace,
                   tTipoCliente AS TipoCliente,
                   tUbigeo AS Ubigeo,
                   tUrbanizacion AS Urbanizacion,
                   tTipoResponsable AS TipoResponsable
            FROM TCLIENTE
            WHERE tCodigoCliente = @Codigo
            """;
        return await conn.QueryFirstOrDefaultAsync<Cliente>(sql, new { Codigo = codigo });
    }

    public async Task<bool> InsertarAsync(Cliente cliente, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO TCLIENTE (tCodigoCliente, tEmpresa, tIdentidad, tDireccion, lActivo, tUsuario, fRegistro,
                                  tCorreo, tTipoIdentidad, tEnlace, tTipoCliente, tUbigeo, tUrbanizacion, tTipoResponsable)
            VALUES (@CodigoCliente, @Empresa, @Identidad, @Direccion, @Activo, @Usuario, @FechaRegistro,
                    @Correo, @TipoIdentidad, @Enlace, @TipoCliente, @Ubigeo, @Urbanizacion, @TipoResponsable)
            """;
        return await conn.ExecuteAsync(sql, cliente) > 0;
    }

    public async Task<bool> ActualizarAsync(Cliente cliente, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE TCLIENTE
               SET tEmpresa = @Empresa,
                   tIdentidad = @Identidad,
                   tDireccion = @Direccion,
                   lActivo = @Activo,
                   tUsuario = @Usuario,
                   fRegistro = @FechaRegistro,
                   tCorreo = @Correo,
                   tTipoIdentidad = @TipoIdentidad,
                   tEnlace = @Enlace,
                   tTipoCliente = @TipoCliente,
                   tUbigeo = @Ubigeo,
                   tUrbanizacion = @Urbanizacion,
                   tTipoResponsable = @TipoResponsable
             WHERE tCodigoCliente = @CodigoCliente
            """;
        return await conn.ExecuteAsync(sql, cliente) > 0;
    }
}
