using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Motorizado;
using Inforest.Domain.Repositories;

namespace Inforest.Infrastructure.Motorizado;

/// <summary>
/// Repositorio Dapper para <see cref="Entities.Motorizado.Motorizado"/> y <see cref="AsignacionMotorizado"/>.
/// <para>
/// Legacy: tabla <c>TMOTORIZADODATOS</c>, vista <c>vMotorizado</c>.
/// SP: <c>sp_UpdFotoDelivery</c> (actualización de foto — pendiente en known-gaps).
/// </para>
/// Reglas: BR-DEL-002, BR-DEL-005, BR-DEL-006.
/// </summary>
internal sealed class MotorizadoRepository : IMotorizadoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public MotorizadoRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    /// <inheritdoc />
    public async Task<Domain.Entities.Motorizado.Motorizado?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigo AS Codigo, tDocumentoIdentidad AS DocumentoIdentidad,
                   ISNULL(nTarifaLV,0) AS TarifaLunesViernes,
                   ISNULL(nTarifaSD,0) AS TarifaSabadoDomingo,
                   ISNULL(nTarifaES,0) AS TarifaEspecial
            FROM TMOTORIZADODATOS
            WHERE tCodigo = @Codigo
            """;
        return await conn.QueryFirstOrDefaultAsync<Domain.Entities.Motorizado.Motorizado>(
            sql, new { Codigo = codigo });
    }

    /// <inheritdoc />
    public async Task<IEnumerable<Domain.Entities.Motorizado.Motorizado>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT tCodigo AS Codigo, tDocumentoIdentidad AS DocumentoIdentidad,
                   ISNULL(nTarifaLV,0) AS TarifaLunesViernes,
                   ISNULL(nTarifaSD,0) AS TarifaSabadoDomingo,
                   ISNULL(nTarifaES,0) AS TarifaEspecial
            FROM TMOTORIZADODATOS
            ORDER BY tCodigo
            """;
        return await conn.QueryAsync<Domain.Entities.Motorizado.Motorizado>(sql);
    }

    /// <inheritdoc />
    public async Task InsertarAsync(Domain.Entities.Motorizado.Motorizado motorizado, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            INSERT INTO TMOTORIZADODATOS (tCodigo, tDocumentoIdentidad, nTarifaLV, nTarifaSD, nTarifaES)
            VALUES (@Codigo, @DocumentoIdentidad, @TarifaLunesViernes, @TarifaSabadoDomingo, @TarifaEspecial)
            """;
        await conn.ExecuteAsync(sql, motorizado);
    }

    /// <inheritdoc />
    public async Task ActualizarTarifasAsync(
        string codigo, decimal tarifaLV, decimal tarifaSD, decimal tarifaES,
        string usuarioModificacion, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        // Actualiza tarifa activa en TPARAMETRO (BR-DEL-005) y tarifas en TMOTORIZADODATOS
        const string sqlMotorizado = """
            UPDATE TMOTORIZADODATOS SET nTarifaLV = @LV, nTarifaSD = @SD, nTarifaES = @ES
            WHERE tCodigo = @Codigo
            """;
        const string sqlParametro = """
            UPDATE TPARAMETRO SET
                tTarifaActualMotorizado = @Codigo,
                tUsuarioTarifa = @Usuario,
                fRegistroTarifa = GETDATE()
            """;
        await conn.ExecuteAsync(sqlMotorizado, new { Codigo = codigo, LV = tarifaLV, SD = tarifaSD, ES = tarifaES });
        await conn.ExecuteAsync(sqlParametro, new { Codigo = codigo, Usuario = usuarioModificacion });
    }

    /// <inheritdoc />
    public async Task<AsignacionMotorizado?> ObtenerAsignacionActivaAsync(string codigoPedido, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT TOP 1
                   tCodigoPedido AS CodigoPedido, tMotorizado AS CodigoMotorizado,
                   fAsignacion AS FechaAsignacion, fSalida AS FechaSalida, fLlegada AS FechaLlegada,
                   tUsuarioAsignacion AS UsuarioAsignacion, lReasignacion AS EsReasignacion
            FROM vMotorizado
            WHERE tCodigoPedido = @CodigoPedido
              AND fLlegada IS NULL
            ORDER BY fAsignacion DESC
            """;
        return await conn.QueryFirstOrDefaultAsync<AsignacionMotorizado>(sql, new { CodigoPedido = codigoPedido });
    }

    /// <inheritdoc />
    public async Task RegistrarAsignacionAsync(AsignacionMotorizado asignacion, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            INSERT INTO MPEDIDO (tCodigoPedido, tMotorizado, fAsignacion, tUsuarioAsignacion, lReasignacion)
            VALUES (@CodigoPedido, @CodigoMotorizado, @FechaAsignacion, @UsuarioAsignacion, @EsReasignacion)
            ON DUPLICATE KEY UPDATE tMotorizado = @CodigoMotorizado
            """;
        // SQL Server version:
        const string sqlUpdate = """
            UPDATE MPEDIDO SET
                tMotorizado = @CodigoMotorizado,
                fAsignacion = @FechaAsignacion,
                tUsuarioAsignacion = @UsuarioAsignacion
            WHERE tCodigoPedido = @CodigoPedido
            """;
        await conn.ExecuteAsync(sqlUpdate, asignacion);
    }

    /// <inheritdoc />
    public async Task RegistrarSalidaAsync(string codigoPedido, string codigoMotorizado, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        // Legacy: frmLlegadaSalida.frm — registra fSalida en MPEDIDO
        const string sql = "UPDATE MPEDIDO SET fSalida = GETDATE() WHERE tCodigoPedido = @CodigoPedido";
        await conn.ExecuteAsync(sql, new { CodigoPedido = codigoPedido });
    }

    /// <inheritdoc />
    public async Task RegistrarLlegadaAsync(string codigoPedido, string codigoMotorizado, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(ct);
        // Legacy: frmLlegadaSalida.frm — registra fLlegada en MPEDIDO
        const string sql = "UPDATE MPEDIDO SET fLlegada = GETDATE() WHERE tCodigoPedido = @CodigoPedido";
        await conn.ExecuteAsync(sql, new { CodigoPedido = codigoPedido });
    }
}
