namespace Inforest.Domain.Entities.Delivery;

/// <summary>
/// Pedido delivery visible en la pantalla de reasignación de motorizado.
/// <para>
/// Legacy: <c>frmReasignacionMotorizado.frm</c> consulta <c>MPEDIDO</c> +
/// <c>vMotorizado</c> usando los campos <c>tMotorizadoN</c> y
/// <c>nTarifaMotorizadoN</c>.
/// </para>
/// </summary>
public sealed record PedidoReasignacionMotorizado
{
    public string CodigoPedido { get; init; } = string.Empty;
    public string? MotorizadoDescripcion { get; init; }
    public decimal? TarifaMotorizado { get; init; }
}
