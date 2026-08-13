using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Impresion;

// ─────────────────────────────────────────────────────────────────────────────
// Commands / Queries
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Comando para imprimir la pre-cuenta de un pedido en una impresora específica.
/// Legacy: frmPrecuentaImpresora.frm + modProcedimiento.bas ImprimirTicket. BR-008.
/// </summary>
public sealed record ImprimirPrecuentaCommand(
    string CodigoPedido,
    string CodigoImpresora);

/// <summary>
/// Query para obtener las impresoras disponibles de una caja.
/// Legacy: frmPrecuentaImpresora.frm Form_Load — SELECT tImpresora, tDescripcion FROM TIMPRESORA WHERE tCaja=sCaja
/// </summary>
public sealed record ObtenerImpresorasPorCajaQuery(string CodigoCaja);

// ─────────────────────────────────────────────────────────────────────────────
// Handler: Imprimir Precuenta
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Imprime la pre-cuenta del pedido indicado en la impresora seleccionada.
/// Legacy: frmPrecuentaImpresora.frm cmdImpresora_Click + modProcedimiento.bas. BR-008.
/// </summary>
public sealed class ImprimirPrecuentaHandler
{
    private readonly IPedidoRepository _pedidoRepository;
    private readonly IImpresoraService _impresoraService;

    public ImprimirPrecuentaHandler(
        IPedidoRepository pedidoRepository,
        IImpresoraService impresoraService)
    {
        _pedidoRepository = pedidoRepository;
        _impresoraService = impresoraService;
    }

    public async Task<Result> HandleAsync(ImprimirPrecuentaCommand command, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(command.CodigoPedido))
            return Result.Fail("El código de pedido es requerido.", "IMPRESION_PEDIDO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(command.CodigoImpresora))
            return Result.Fail("Seleccione una impresora.", "IMPRESION_IMPRESORA_REQUERIDA");

        var pedido = await _pedidoRepository.ObtenerPorCodigoAsync(command.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail("No se encontró el pedido.", "IMPRESION_PEDIDO_NO_ENCONTRADO");

        if (pedido.Detalles.Count == 0)
            return Result.Fail("El pedido no tiene ítems para imprimir.", "IMPRESION_PEDIDO_SIN_DETALLES");

        var contenido = BuildPrecuenta(pedido);
        await _impresoraService.ImprimirTicketAsync(contenido, command.CodigoImpresora, ct);
        return Result.Ok();
    }

    /// <summary>
    /// Construye el texto de la pre-cuenta.
    /// Legacy: modProcedimiento.bas — formato de pre-cuenta estándar.
    /// </summary>
    private static string BuildPrecuenta(Pedido pedido)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("========== PRE-CUENTA ==========");
        sb.AppendLine($"Pedido : {pedido.CodigoPedido}");
        sb.AppendLine($"Mesa   : {pedido.CodigoMesa}");
        sb.AppendLine($"Fecha  : {DateTime.Now:dd/MM/yyyy HH:mm}");
        sb.AppendLine("--------------------------------");

        foreach (var d in pedido.Detalles)
        {
            sb.AppendLine($"{d.Cantidad,5:##0.##}  {d.CodigoProducto,-15}  {d.PrecioVenta * d.Cantidad,10:###,##0.00}");
        }

        var subtotal = pedido.Detalles.Sum(d => d.PrecioVenta * d.Cantidad);
        sb.AppendLine("--------------------------------");
        sb.AppendLine($"{"TOTAL",-22}  {subtotal,10:###,##0.00}");
        sb.AppendLine("================================");
        sb.AppendLine("        ** PRE-CUENTA **");
        sb.AppendLine("   No es comprobante de pago");

        return sb.ToString();
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Interface: impresoras disponibles por caja (needed by form)
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>DTO de impresora disponible. Legacy: TIMPRESORA.</summary>
public sealed record ImpresoraDisponible(string CodigoImpresora, string Descripcion);

/// <summary>
/// Repositorio de impresoras por caja.
/// Legacy: SELECT tImpresora, tDescripcion FROM TIMPRESORA WHERE tCaja=sCaja
/// </summary>
public interface IImpresoraRepository
{
    Task<IReadOnlyList<ImpresoraDisponible>> ObtenerImpresorasPorCajaAsync(string codigoCaja, CancellationToken ct = default);
}

/// <summary>
/// Handler para obtener impresoras de una caja (alimenta el combo de frmPrecuentaImpresora).
/// Legacy: frmPrecuentaImpresora.frm Form_Load. BR-008.
/// </summary>
public sealed class ObtenerImpresorasPorCajaHandler
{
    private readonly IImpresoraRepository _repository;

    public ObtenerImpresorasPorCajaHandler(IImpresoraRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<ImpresoraDisponible>>> HandleAsync(
        ObtenerImpresorasPorCajaQuery query,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoCaja))
            return Result.Fail<IReadOnlyList<ImpresoraDisponible>>(
                "El código de caja es requerido.", "IMPRESION_CAJA_REQUERIDA");

        var impresoras = await _repository.ObtenerImpresorasPorCajaAsync(query.CodigoCaja, ct);
        return Result.Ok(impresoras);
    }
}
