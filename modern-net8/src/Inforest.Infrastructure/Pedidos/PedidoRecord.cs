namespace Inforest.Infrastructure.Pedidos;

internal sealed class PedidoRecord
{
    public string CodigoPedido { get; init; } = string.Empty;
    public string? CodigoMesa { get; init; }
    public string? CodigoTurno { get; init; }
    public string? CodigoUsuario { get; init; }
    public string? TipoPedido { get; init; }
    public string? EstadoPedido { get; init; }
    public DateTime FechaRegistro { get; init; }
    public DateTime? FechaDiaContable { get; init; }
    public int NumeroAdultos { get; init; }
    public int NumeroNinos { get; init; }
    public string? Observacion { get; init; }
}
