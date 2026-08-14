using Dapper;
using Inforest.Application.Almacen;
using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Configuracion;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Almacen;

/// <summary>
/// Gateway que crea un pedido MPEDIDO + DPEDIDO en INFOREST a partir de un requerimiento de almacén.
/// Legacy: InsertaProducto() + spIns_MPEDIDO en frmImportacionRequerimientos.frm.
/// BR-IMPORT-003: Si un producto no tiene enlace, cancela el pedido (estado '03').
/// BR-IMPORT-004: El requerimiento se marca como importado tras la creación exitosa.
/// </summary>
internal sealed class ImportacionPedidoGateway : IImportacionPedidoGateway
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;
    private readonly IProductoMaestroRepository _productoRepository;
    private readonly IParametroRepository _parametroRepository;

    public ImportacionPedidoGateway(
        IDbConnectionFactory connectionFactory,
        ISpExecutor spExecutor,
        IProductoMaestroRepository productoRepository,
        IParametroRepository parametroRepository)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
        _productoRepository = productoRepository;
        _parametroRepository = parametroRepository;
    }

    /// <inheritdoc/>
    public async Task<Result<ImportarRequerimientoResult>> CrearPedidoDesdeRequerimientoAsync(
        ImportacionPedidoContexto contexto,
        CancellationToken ct = default)
    {
        var config = await _parametroRepository.ObtenerConfiguracionAsync(ct)
            ?? throw new InvalidOperationException("No se pudo obtener la configuración del sistema.");

        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);

        // 1. Crear encabezado MPEDIDO (spIns_MPEDIDO)
        var pedidoParams = new DynamicParameters();
        pedidoParams.Add("@tCliente", string.Empty);
        pedidoParams.Add("@tTipoPedido", contexto.TipoPedido);
        pedidoParams.Add("@lPrioridad", false);
        pedidoParams.Add("@tTipoAtencion", string.Empty);
        pedidoParams.Add("@tMesa", string.Empty);
        pedidoParams.Add("@tMozo", string.Empty);
        pedidoParams.Add("@tMotorizado", "0000");
        pedidoParams.Add("@tCaja", contexto.CodigoCaja);
        pedidoParams.Add("@tSalon", contexto.CodigoSalon);
        pedidoParams.Add("@tTurno", contexto.CodigoTurno);
        pedidoParams.Add("@tObservacion", contexto.Observacion);
        pedidoParams.Add("@nTiempo", 0);
        pedidoParams.Add("@tUsuario", contexto.CodigoUsuario.Length > 15
            ? contexto.CodigoUsuario[^15..] : contexto.CodigoUsuario);
        pedidoParams.Add("@nAdulto", 0);
        pedidoParams.Add("@nNino", 0);
        pedidoParams.Add("@nMesa", 0);
        pedidoParams.Add("@tPuntoVenta", string.Empty);
        pedidoParams.Add("@tHabitacion", string.Empty);
        pedidoParams.Add("@tReserva", string.Empty);
        pedidoParams.Add("@tPasajero", string.Empty);
        pedidoParams.Add("@tCompania", string.Empty);
        pedidoParams.Add("@tContacto", string.Empty);
        pedidoParams.Add("@nDescuento", 0m);
        pedidoParams.Add("@tDescuento", string.Empty);
        pedidoParams.Add("@tObservacionDescuento", string.Empty);
        pedidoParams.Add("@tAutorizaDescuento", string.Empty);
        pedidoParams.Add("@nTiempoDelivery", 0);
        pedidoParams.Add("@tTienda", string.Empty);
        pedidoParams.Add("@fDiaContable", contexto.FechaDiaContable);
        pedidoParams.Add("@fProgramacion", dbType: System.Data.DbType.DateTime, direction: System.Data.ParameterDirection.Input, value: (object?)null);
        pedidoParams.Add("@tCodigoInvitado", string.Empty);
        pedidoParams.Add("@tcodigoPariente", string.Empty);
        pedidoParams.Add("@tEntregarA", string.Empty);
        pedidoParams.Add("@nTiempoAntesEnvio", 0);
        pedidoParams.Add("@nMontoMaximo", 0);
        pedidoParams.Add("@tPedido", dbType: System.Data.DbType.String, size: 10, direction: System.Data.ParameterDirection.Output);
        pedidoParams.Add("@codigoOrigenVentas", string.Empty);

        await _spExecutor.ExecuteAsync(conn, "spIns_MPEDIDO", pedidoParams, cancellationToken: ct);

        var codigoPedido = pedidoParams.Get<string>("@tPedido");
        if (string.IsNullOrWhiteSpace(codigoPedido))
            return Result.Fail<ImportarRequerimientoResult>(
                "El SP spIns_MPEDIDO no devolvió un código de pedido.",
                "IMPORT_REQ_PEDIDO_NO_GENERADO");

        // 2. Insertar líneas DPEDIDO
        var nOrden = 0;
        var productosInsertados = 0;

        foreach (var item in contexto.Detalle)
        {
            var producto = await _productoRepository.ObtenerPorCodigoAsync(item.CodigoProductoInforRest, ct);
            if (producto is null)
            {
                // BR-IMPORT-003: Si un producto no existe, cancelar el pedido
                await CancelarPedidoAsync(conn, codigoPedido, contexto.CodigoTurno, contexto.CodigoUsuario, contexto.Rq, ct);
                return Result.Fail<ImportarRequerimientoResult>(
                    $"El producto '{item.CodigoProductoInforRest}' no existe en INFOREST. Pedido '{codigoPedido}' cancelado.",
                    "IMPORT_REQ_PRODUCTO_NO_ENCONTRADO");
            }

            nOrden++;
            var sitem = await ObtenerProximoItemAsync(conn, codigoPedido);

            var (precioOficial, imp1, imp2, imp3, precioNeto) = CalcularPrecios(
                producto, config, contexto.TipoPedido);

            // Legacy: INSERT INTO DPEDIDO (...)
            const string sqlDetalle = """
                INSERT INTO DPEDIDO
                (tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo,
                 nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta,
                 nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3,
                 lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, tEstadoItem, tComanda,
                 fRegistro, tMozoD, tUsuarioD, nInsumo, nGasto, nManoObra, nOrden,
                 tUnidadNegocio, tOferta, tsubalmacen, fdiacontable, tcajad)
                VALUES
                (@codigoPedido, @tipoPedido, @item, @codigoProducto, @grupo, @subgrupo,
                 @precioNeto, 0, 0, @precioOficial, @precioImpuesto1, @precioImpuesto2, @precioImpuesto3, @precioVenta,
                 @cantidad, @venta, @montoImp1, @montoImp2, @montoImp3,
                 0, @area, @imprimeArea, @combinacion, @numeroCombinacion, 'N', '',
                 GETDATE(), '', @usuario, 0, 0, 0, @orden,
                 @unidadNegocio, '', '', @fechaDiaContable, @caja)
                """;

            var cantidad = (decimal)item.Cantidad;
            await conn.ExecuteAsync(sqlDetalle, new
            {
                codigoPedido,
                tipoPedido = contexto.TipoPedido,
                item = sitem,
                codigoProducto = producto.CodigoProducto,
                grupo = producto.Grupo,
                subgrupo = producto.SubGrupo ?? string.Empty,
                precioNeto,
                precioOficial,
                precioImpuesto1 = (decimal)config.Impuesto1,
                precioImpuesto2 = (decimal)config.Impuesto2,
                precioImpuesto3 = (decimal)config.Impuesto3,
                precioVenta = precioOficial,
                cantidad,
                venta = precioOficial * cantidad,
                montoImp1 = imp1 * cantidad,
                montoImp2 = imp2 * cantidad,
                montoImp3 = imp3 * cantidad,
                area = producto.Area ?? string.Empty,
                imprimeArea = producto.ImprimeArea ? -1 : 0,
                combinacion = producto.Combinacion ? -1 : 0,
                numeroCombinacion = producto.NumeroCombinacion,
                usuario = contexto.CodigoUsuario.Length > 15
                    ? contexto.CodigoUsuario[^15..] : contexto.CodigoUsuario,
                orden = nOrden,
                unidadNegocio = producto.UnidadNegocio ?? string.Empty,
                fechaDiaContable = contexto.FechaDiaContable.ToString("yyyyMMdd"),
                caja = contexto.CodigoCaja
            });

            productosInsertados++;
        }

        return Result.Ok(new ImportarRequerimientoResult(codigoPedido, productosInsertados));
    }

    // ── Price calculation ────────────────────────────────────────────────────

    /// <summary>
    /// Calcula precio oficial e impuestos según canal de venta (TipoPedido).
    /// Legacy: InsertaProducto() en frmImportacionRequerimientos.frm.
    /// BR-IMPORT-003.
    /// </summary>
    private static (decimal precioOficial, decimal imp1, decimal imp2, decimal imp3, decimal precioNeto)
        CalcularPrecios(ProductoMaestro producto, ConfiguracionSistema config, string tipoPedido)
    {
        bool lImp1, lImp2, lImp3;
        decimal precioOficial;

        switch (tipoPedido)
        {
            case "02": // Delivery
                if (producto.PrecioDelivery == 0)
                {
                    precioOficial = producto.PrecioVenta + (producto.PrecioVenta * (decimal)config.nDelivery / 100m);
                    lImp1 = producto.Impuesto1;
                    lImp2 = producto.Impuesto2;
                    lImp3 = producto.Impuesto3;
                }
                else
                {
                    precioOficial = producto.PrecioDelivery;
                    lImp1 = producto.Impuesto4;
                    lImp2 = producto.Impuesto5;
                    lImp3 = producto.Impuesto6;
                }
                break;

            case "03": // Llevar
                if (producto.PrecioLlevar == 0)
                {
                    precioOficial = producto.PrecioVenta + (producto.PrecioVenta * (decimal)config.nLlevar / 100m);
                    lImp1 = producto.Impuesto1;
                    lImp2 = producto.Impuesto2;
                    lImp3 = producto.Impuesto3;
                }
                else
                {
                    precioOficial = producto.PrecioLlevar;
                    lImp1 = producto.Impuesto7;
                    lImp2 = producto.Impuesto8;
                    lImp3 = producto.Impuesto9;
                }
                break;

            case "04":
                precioOficial = producto.PrecioCanal4 == 0
                    ? producto.PrecioVenta + (producto.PrecioVenta * (decimal)config.nLlevar / 100m)
                    : producto.PrecioCanal4;
                lImp1 = producto.Impuesto10;
                lImp2 = producto.Impuesto11;
                lImp3 = producto.Impuesto12;
                break;

            case "05":
                precioOficial = producto.PrecioCanal5 == 0
                    ? producto.PrecioVenta + (producto.PrecioVenta * (decimal)config.nLlevar / 100m)
                    : producto.PrecioCanal5;
                lImp1 = producto.Impuesto13;
                lImp2 = producto.Impuesto14;
                lImp3 = producto.Impuesto15;
                break;

            default: // 01 = Salón y otros
                precioOficial = producto.PrecioVenta;
                lImp1 = producto.Impuesto1;
                lImp2 = producto.Impuesto2;
                lImp3 = producto.Impuesto3;
                break;
        }

        // Cálculo de impuestos (precio incluye impuestos — método Perú/Ecuador por defecto)
        // BR: nValor = 1 + (%imp/100); nNeto = nPrecioVenta / nValor
        var pct = 0m;
        if (lImp1) pct += (decimal)config.Impuesto1;
        if (lImp2) pct += (decimal)config.Impuesto2;
        if (lImp3) pct += (decimal)config.Impuesto3;

        decimal imp1 = lImp1 ? Round(precioOficial / (1m + pct / 100m) * (decimal)config.Impuesto1 / 100m) : 0m;
        decimal imp2 = lImp2 ? Round(precioOficial / (1m + pct / 100m) * (decimal)config.Impuesto2 / 100m) : 0m;
        decimal imp3 = lImp3 ? Round(precioOficial / (1m + pct / 100m) * (decimal)config.Impuesto3 / 100m) : 0m;
        decimal precioNeto = precioOficial - imp1 - imp2 - imp3;

        return (precioOficial, imp1, imp2, imp3, precioNeto);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static async Task<string> ObtenerProximoItemAsync(System.Data.IDbConnection conn, string codigoPedido)
    {
        // Legacy: Lib.Correlativo(Calcular("select max(tItem) as codigo from DPEDIDO where tCodigoPedido = '...'", Cn), 3)
        const string sql = "SELECT ISNULL(MAX(tItem), '000') AS Codigo FROM DPEDIDO WHERE tCodigoPedido = @codigo";
        var maxItem = await conn.QueryFirstOrDefaultAsync<string>(sql, new { codigo = codigoPedido }) ?? "000";
        return int.TryParse(maxItem, out var n) ? (n + 1).ToString("D3") : "001";
    }

    private async Task CancelarPedidoAsync(
        System.Data.IDbConnection conn,
        string codigoPedido,
        string turno,
        string usuario,
        string rq,
        CancellationToken ct)
    {
        // Legacy: update MPEDIDO set tEstadoPedido='03', tTurnoAnulado, ... where tCodigoPedido='...'
        const string sql = """
            UPDATE MPEDIDO
            SET tEstadoPedido = '03',
                tTurnoAnulado = @turno,
                tMotivoAnulacion = '000',
                tObservacionAnulado = @observacion,
                tUsuarioAnulado = @usuario,
                fRegAnulado = GETDATE()
            WHERE tCodigoPedido = @codigoPedido
            """;
        await conn.ExecuteAsync(sql, new
        {
            codigoPedido,
            turno,
            observacion = $"Error Rq. {rq}",
            usuario = usuario.Length > 15 ? usuario[^15..] : usuario
        });
    }

    private static decimal Round(decimal v) => Math.Round(v, 4, MidpointRounding.AwayFromZero);
}
