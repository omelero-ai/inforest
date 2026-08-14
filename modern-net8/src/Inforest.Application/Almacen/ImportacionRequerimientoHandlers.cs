using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Almacen;

namespace Inforest.Application.Almacen;

// ── Queries ──────────────────────────────────────────────────────────────────

/// <summary>
/// Lista requerimientos de almacén pendientes de importar al POS.
/// Legacy: cmdProcesa_Click en frmImportacionRequerimientos.frm.
/// BR-IMPORT-001.
/// </summary>
public sealed record ObtenerRequerimientosPendientesQuery(
    DateTime FechaInicio,
    DateTime FechaFin);

public sealed class ObtenerRequerimientosPendientesHandler
{
    private readonly IRequerimientoAlmacenRepository _repository;

    public ObtenerRequerimientosPendientesHandler(IRequerimientoAlmacenRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<RequerimientoAlmacen>>> HandleAsync(
        ObtenerRequerimientosPendientesQuery query, CancellationToken ct = default)
    {
        if (query.FechaFin < query.FechaInicio)
            return Result.Fail<IReadOnlyList<RequerimientoAlmacen>>(
                "La fecha final no puede ser anterior a la fecha inicial.",
                "REQ_RANGO_FECHAS_INVALIDO");

        var items = await _repository.ObtenerPendientesAsync(query.FechaInicio, query.FechaFin, ct);
        return Result.Ok<IReadOnlyList<RequerimientoAlmacen>>(items);
    }
}

/// <summary>
/// Obtiene el detalle de artículos de un requerimiento con enlace a INFOREST.
/// Legacy: Form_Load / Asignar en frmImportacionRequerimientoDetalle.frm.
/// BR-IMPORT-003.
/// </summary>
public sealed record ObtenerDetalleRequerimientoQuery(
    string Rq,
    string BaseDatosInforest = "INFOREST");

public sealed class ObtenerDetalleRequerimientoHandler
{
    private readonly IRequerimientoAlmacenRepository _repository;

    public ObtenerDetalleRequerimientoHandler(IRequerimientoAlmacenRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<DetalleRequerimientoAlmacen>>> HandleAsync(
        ObtenerDetalleRequerimientoQuery query, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.Rq))
            return Result.Fail<IReadOnlyList<DetalleRequerimientoAlmacen>>(
                "El código de requerimiento es obligatorio.",
                "REQ_CODIGO_REQUERIDO");

        var detalle = await _repository.ObtenerDetalleAsync(query.Rq, query.BaseDatosInforest, ct);
        return Result.Ok<IReadOnlyList<DetalleRequerimientoAlmacen>>(detalle);
    }
}

// ── Commands ─────────────────────────────────────────────────────────────────

/// <summary>
/// Resultado de importar un requerimiento como pedido.
/// </summary>
public sealed record ImportarRequerimientoResult(
    string CodigoPedido,
    int ProductosImportados);

/// <summary>
/// Importa un requerimiento de almacén como pedido en INFOREST.
/// Legacy: cmdOpcion_Click(Index=2) en frmImportacionRequerimientos.frm.
/// BR-IMPORT-001: Solo requerimientos en estado '02' y área habilitada.
/// BR-IMPORT-002: No importar si ya existe un pedido para este requerimiento.
/// BR-IMPORT-003: Todos los artículos deben tener enlace; si falla uno, cancelar pedido.
/// BR-IMPORT-004: Marcar el requerimiento como importado con el código del pedido.
/// </summary>
public sealed record ImportarRequerimientoCommand(
    string Rq,
    string CodigoCaja,
    string CodigoTurno,
    string CodigoSalon,
    string CodigoUsuario,
    string TipoPedido,
    DateTime FechaDiaContable,
    string BaseDatosInforest = "INFOREST");

public sealed class ImportarRequerimientoHandler
{
    private readonly IRequerimientoAlmacenRepository _requerimientoRepository;
    private readonly IImportacionPedidoGateway _importacionGateway;

    public ImportarRequerimientoHandler(
        IRequerimientoAlmacenRepository requerimientoRepository,
        IImportacionPedidoGateway importacionGateway)
    {
        _requerimientoRepository = requerimientoRepository;
        _importacionGateway = importacionGateway;
    }

    public async Task<Result<ImportarRequerimientoResult>> HandleAsync(
        ImportarRequerimientoCommand command, CancellationToken ct = default)
    {
        // BR-IMPORT-002: Verificar que no tenga pedido ya generado
        var pedidoExistente = await _requerimientoRepository.ObtenerPedidoAsociadoAsync(command.Rq, ct);
        if (!string.IsNullOrWhiteSpace(pedidoExistente))
            return Result.Fail<ImportarRequerimientoResult>(
                $"El requerimiento '{command.Rq}' ya tiene un pedido generado: {pedidoExistente}.",
                "REQ_YA_IMPORTADO");

        // Obtener detalle del requerimiento con enlace a INFOREST
        var detalle = await _requerimientoRepository.ObtenerDetalleAsync(
            command.Rq, command.BaseDatosInforest, ct);

        if (detalle.Count == 0)
            return Result.Fail<ImportarRequerimientoResult>(
                $"El requerimiento '{command.Rq}' no tiene productos.",
                "REQ_SIN_PRODUCTOS");

        // BR-IMPORT-003: Verificar que todos tienen enlace antes de crear el pedido
        var sinEnlace = detalle.Where(d => !d.TieneEnlace).ToList();
        if (sinEnlace.Count > 0)
            return Result.Fail<ImportarRequerimientoResult>(
                $"Los siguientes artículos no tienen enlace con INFOREST: {string.Join(", ", sinEnlace.Select(d => d.CodArt))}.",
                "REQ_PRODUCTO_SIN_ENLACE");

        // Crear pedido encabezado (MPEDIDO) y detalle (DPEDIDO)
        var contexto = new ImportacionPedidoContexto(
            Rq: command.Rq,
            CodigoCaja: command.CodigoCaja,
            CodigoTurno: command.CodigoTurno,
            CodigoSalon: command.CodigoSalon,
            CodigoUsuario: command.CodigoUsuario,
            TipoPedido: command.TipoPedido,
            FechaDiaContable: command.FechaDiaContable,
            Observacion: $"Req: {command.Rq}",
            Detalle: detalle.Select(d => new ItemImportacionPedido(
                CodigoProductoInforRest: d.EnlaceInforRest!,
                Cantidad: d.Cantidad)).ToList());

        var resultado = await _importacionGateway.CrearPedidoDesdeRequerimientoAsync(contexto, ct);

        if (!resultado.EsExitoso)
            return Result.Fail<ImportarRequerimientoResult>(resultado.MensajeError!, resultado.CodigoError!);

        // BR-IMPORT-004: Marcar el requerimiento como importado
        await _requerimientoRepository.MarcarImportadoAsync(command.Rq, resultado.Valor!.CodigoPedido, ct);

        return Result.Ok(resultado.Valor!);
    }
}
