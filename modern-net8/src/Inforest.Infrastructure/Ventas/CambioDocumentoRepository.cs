using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;

namespace Inforest.Infrastructure.Ventas;

/// <summary>
/// Repositorio Dapper para el proceso de cambio de tipo de documento.
/// Legacy: frmCambio.frm — operaciones directas SQL. BR-CAMBIO-001..005.
/// Tablas: MDOCUMENTO, DDOCUMENTO, DPEDIDO, DPAGODOCUMENTO, TTIPODOCUMENTOIMPRESORA, vTipoDocumento, vTipoDocumentoImpresora.
/// </summary>
internal sealed class CambioDocumentoRepository : ICambioDocumentoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public CambioDocumentoRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    /// <inheritdoc />
    public async Task<IReadOnlyList<TipoDocumentoDisponible>> ObtenerTiposDocumentoAsync(
        string codigoCaja, string pais, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);

        // Ecuador incluye también facturación electrónica con número de autorización
        string sql;
        if (pais == "002")
        {
            sql = """
                SELECT tTipoEmision, Descripcion, Prefijo, tSerie, tUltimoNumero,
                       timpresora, lCliente, lResumen
                FROM vTipoDocumentoImpresora
                WHERE tCaja = @Caja
                  AND Transporte = 0
                  AND lNotaCredito = 0
                  AND lActivo = 1
                  AND (ISNULL(tNumeroAutorizacion,'') <> '' OR lFacturacionElectronica = 1)
                UNION
                SELECT tTipoEmision, Descripcion, Prefijo, tSerie, tUltimoNumero,
                       timpresora, lCliente, lResumen
                FROM vTipoDocumentoImpresora
                WHERE tCaja = @Caja
                  AND Transporte = 0
                  AND lNotaCredito = 0
                  AND lFacturacionElectronica = 1
                  AND lActivo = 1
                ORDER BY tTipoEmision
                """;
        }
        else
        {
            sql = """
                SELECT tTipoEmision, Descripcion, Prefijo, tSerie, tUltimoNumero,
                       timpresora, lCliente, lResumen
                FROM vTipoDocumentoImpresora
                WHERE tCaja = @Caja
                  AND Transporte = 0
                  AND lNotaCredito = 0
                  AND lActivo = 1
                ORDER BY tTipoEmision
                """;
        }

        var rows = await connection.QueryAsync<TipoDocumentoRow>(
            new CommandDefinition(sql, new { Caja = codigoCaja }, cancellationToken: ct));

        return rows.Select(r => new TipoDocumentoDisponible(
            TipoEmision: r.tTipoEmision,
            Descripcion: r.Descripcion,
            Prefijo: r.Prefijo,
            Serie: r.tSerie,
            UltimoNumero: r.tUltimoNumero,
            CodigoImpresora: r.timpresora,
            RequiereCliente: r.lCliente,
            EsResumen: r.lResumen)).ToList();
    }

    /// <inheritdoc />
    public async Task<(decimal MontoDocumento, decimal MontoMaximo)> ObtenerMontosValidacionAsync(
        string codigoDocumento, string tipoEmision, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);

        var montoMax = await connection.ExecuteScalarAsync<decimal>(
            new CommandDefinition(
                "SELECT ISNULL(nMontoMaximo,0) FROM vTipoDocumento WHERE Codigo = @Tipo",
                new { Tipo = tipoEmision },
                cancellationToken: ct));

        var montoDoc = await connection.ExecuteScalarAsync<decimal>(
            new CommandDefinition(
                "SELECT ISNULL(nventa,0) FROM MDOCUMENTO WHERE tDocumento = @Doc",
                new { Doc = codigoDocumento },
                cancellationToken: ct));

        return (montoDoc, montoMax);
    }

    /// <inheritdoc />
    public async Task EjecutarCambioAsync(
        EjecutarCambioDocumentoParams p, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        using var tx = connection.BeginTransaction();

        try
        {
            // 1. Actualizar correlativo en TTIPODOCUMENTOIMPRESORA
            await connection.ExecuteAsync(
                new CommandDefinition(
                    """
                    UPDATE TTIPODOCUMENTOIMPRESORA
                    SET tUltimoNumero = @Correlativo
                    WHERE tTipoEmision = @TipoEmision AND tCaja = @Caja
                    """,
                    new { Correlativo = p.NuevoCorrelatvo, TipoEmision = p.TipoEmision, Caja = p.CodigoCaja },
                    transaction: tx, cancellationToken: ct));

            // 2. Reasignar pagos al nuevo documento
            await connection.ExecuteAsync(
                new CommandDefinition(
                    "UPDATE DPAGODOCUMENTO SET tdocumento = @Nuevo WHERE tDocumento = @Origen",
                    new { Nuevo = p.CodigoDocumentoNuevo, Origen = p.CodigoDocumentoOrigen },
                    transaction: tx, cancellationToken: ct));

            // 3. Reasignar detalle de pedido al nuevo documento
            await connection.ExecuteAsync(
                new CommandDefinition(
                    "UPDATE DPEDIDO SET tDocumento = @Nuevo WHERE tDocumento = @Origen",
                    new { Nuevo = p.CodigoDocumentoNuevo, Origen = p.CodigoDocumentoOrigen },
                    transaction: tx, cancellationToken: ct));

            // 4. Copiar DDOCUMENTO al nuevo número
            await connection.ExecuteAsync(
                new CommandDefinition(
                    """
                    INSERT INTO DDOCUMENTO
                        (tDocumento, tItem, tCodigoPedido, tCodigoProducto,
                         nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3,
                         nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial,
                         nImpuesto1, nImpuesto2, nImpuesto3, nVenta)
                    SELECT @Nuevo, tItem, tCodigoPedido, tCodigoProducto,
                           nPrecioNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3,
                           nPrecioVenta, nRecargo, nDescuento, nCantidad, nPrecioOficial,
                           nImpuesto1, nImpuesto2, nImpuesto3, nVenta
                    FROM DDOCUMENTO
                    WHERE tDocumento = @Origen
                    """,
                    new { Nuevo = p.CodigoDocumentoNuevo, Origen = p.CodigoDocumentoOrigen },
                    transaction: tx, cancellationToken: ct));

            // 5. Copiar MDOCUMENTO al nuevo número (con nuevo tipo, cliente, turno, caja, etc.)
            await connection.ExecuteAsync(
                new CommandDefinition(
                    """
                    INSERT INTO MDOCUMENTO
                        (tDocumento, tTipoDocumento, tCodigoCliente, tCortesia,
                         nNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3,
                         nVenta, nRecargo, nDescuento, nPrecioOficial, nPropina, nTotal,
                         nAbono, nVuelto, tEstadoDocumento, tClientePago, tTurno, fPago,
                         tCaja, tSalon, tUsuario, tUsuarioAutoriza, fRegistro, tObservacion,
                         fDiaContable, tautorizacion, tcodigocontrol, lreplica)
                    SELECT @Nuevo, @TipoEmision, @Cliente, tCortesia,
                           nNeto, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3,
                           nVenta, nRecargo, nDescuento, nPrecioOficial, nPropina, nTotal,
                           nAbono, nVuelto, tEstadoDocumento, tClientePago, @Turno, fPago,
                           @Caja, @Salon, @Usuario, @Usuario, GETDATE(), tObservacion,
                           fDiaContable, @Autorizacion, @CodigoControl, 1
                    FROM MDOCUMENTO
                    WHERE tDocumento = @Origen
                    """,
                    new
                    {
                        Nuevo = p.CodigoDocumentoNuevo,
                        TipoEmision = p.TipoEmision,
                        Cliente = p.CodigoClienteNuevo,
                        Turno = p.CodigoTurno,
                        Caja = p.CodigoCaja,
                        Salon = p.CodigoSalon,
                        Usuario = p.CodigoUsuario,
                        Autorizacion = p.Autorizacion,
                        CodigoControl = p.CodigoControl,
                        Origen = p.CodigoDocumentoOrigen
                    },
                    transaction: tx, cancellationToken: ct));

            // 6. Marcar documento origen como estado '04' (BR-CAMBIO-004)
            await connection.ExecuteAsync(
                new CommandDefinition(
                    """
                    UPDATE MDOCUMENTO
                    SET tEstadoDocumento = '04',
                        tObservacion = @Motivo,
                        tUsuarioAnulado = @Usuario,
                        fRegistroAnulado = GETDATE(),
                        lreplica = 1
                    WHERE tDocumento = @Origen
                    """,
                    new { Motivo = p.Motivo, Usuario = p.CodigoUsuario, Origen = p.CodigoDocumentoOrigen },
                    transaction: tx, cancellationToken: ct));

            tx.Commit();
        }
        catch
        {
            tx.Rollback();
            throw;
        }
    }

    /// <inheritdoc />
    public async Task ActualizarEmisionAsync(
        string codigoNuevo, bool esConsumo, string descripcionConsumo,
        CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        await connection.ExecuteAsync(
            new CommandDefinition(
                """
                UPDATE MDOCUMENTO
                SET tEmision = @Emision,
                    tConsumo = @Consumo,
                    lreplica = 1
                WHERE tDocumento = @Nuevo
                """,
                new
                {
                    Emision = esConsumo ? "C" : "D",
                    Consumo = esConsumo ? descripcionConsumo : string.Empty,
                    Nuevo = codigoNuevo
                },
                cancellationToken: ct));
    }

    // ── Private row types ────────────────────────────────────────────────────
#pragma warning disable CS8618
    private sealed class TipoDocumentoRow
    {
        public string tTipoEmision { get; set; }
        public string Descripcion { get; set; }
        public string Prefijo { get; set; }
        public string tSerie { get; set; }
        public string tUltimoNumero { get; set; }
        public string timpresora { get; set; }
        public bool lCliente { get; set; }
        public bool lResumen { get; set; }
    }
#pragma warning restore CS8618
}
