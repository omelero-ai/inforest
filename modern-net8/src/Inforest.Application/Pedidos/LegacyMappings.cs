using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Pedidos;

public static class LegacyMappings
{
    public static string MapCanal(CanalVenta canal)
        => canal switch
        {
            CanalVenta.Local => "01",
            CanalVenta.Delivery => "02",
            CanalVenta.Llevar => "03",
            CanalVenta.Canal4 => "04",
            CanalVenta.Canal5 => "05",
            _ => "01"
        };

    public static EstadoPedido MapEstado(string? estadoLegacy)
        => (estadoLegacy ?? string.Empty).Trim().ToUpperInvariant() switch
        {
            "02" => EstadoPedido.EnProceso,
            "03" => EstadoPedido.Cerrado,
            "04" => EstadoPedido.Anulado,
            "AN" => EstadoPedido.Anulado,
            "CE" => EstadoPedido.Cerrado,
            _ => EstadoPedido.Abierto
        };
    public static string MapEstado(EstadoPedido estado)
        => estado switch
        {
            EstadoPedido.EnProceso => "02",
            EstadoPedido.Cerrado => "CE",
            EstadoPedido.Anulado => "AN",
            _ => "01"
        };
}
