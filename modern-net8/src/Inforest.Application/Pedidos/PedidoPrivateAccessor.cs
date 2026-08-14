using System.Reflection;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Pedidos;

internal static class PedidoPrivateAccessor
{
    public static void SetObservacion(Pedido pedido, string? observacion)
        => SetProperty(pedido, nameof(Pedido.Observacion), string.IsNullOrWhiteSpace(observacion) ? null : observacion.Trim());

    public static void SetFechaDiaContable(Pedido pedido, DateTime fechaDiaContable)
        => SetProperty(pedido, nameof(Pedido.FechaDiaContable), fechaDiaContable.Date);

    public static void SetEstado(Pedido pedido, EstadoPedido estado)
        => SetProperty(pedido, nameof(Pedido.Estado), estado);

    public static void ConfigurarDetalle(
        DetallePedido detalle,
        string? codigoGrupo,
        string? codigoSubGrupo,
        string? area,
        string? observacion)
    {
        SetProperty(detalle, nameof(DetallePedido.CodigoGrupo), string.IsNullOrWhiteSpace(codigoGrupo) ? null : codigoGrupo.Trim());
        SetProperty(detalle, nameof(DetallePedido.CodigoSubGrupo), string.IsNullOrWhiteSpace(codigoSubGrupo) ? null : codigoSubGrupo.Trim());
        SetProperty(detalle, nameof(DetallePedido.Area), string.IsNullOrWhiteSpace(area) ? null : area.Trim());
        SetProperty(detalle, nameof(DetallePedido.Observacion), string.IsNullOrWhiteSpace(observacion) ? null : observacion.Trim());
    }

    private static void SetProperty<T>(object target, string propertyName, T value)
    {
        var property = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException($"No se encontró la propiedad {propertyName} en {target.GetType().Name}.");

        property.SetValue(target, value);
    }
}
