using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Delivery;

/// <summary>
/// Configuración de comportamiento de un tipo de pedido en delivery y motorizado.
/// <para>
/// Legacy: tabla <c>TTIPOPEDIDODETALLE</c>.
/// Determina qué campos son requeridos según el tipo de atención.
/// </para>
/// Reglas: BR-DEL-001, BR-DEL-010.
/// </summary>
public class TipoPedidoDelivery : Entity
{
    /// <summary>Código del tipo de pedido. Legacy: <c>tcodigoTipoPedido</c>.</summary>
    public string CodigoTipoPedido { get; private set; } = string.Empty;

    /// <summary>Activa mozo para este tipo de pedido. Legacy: <c>lActivaMozo</c>.</summary>
    public bool ActivaMozo { get; private set; }

    /// <summary>Activa motorizado para este tipo de pedido. Legacy: <c>lActivaMotorizado</c>. Regla BR-DEL-001.</summary>
    public bool ActivaMotorizado { get; private set; }

    /// <summary>Obliga ingreso de mesa. Legacy: <c>lObligaMesa</c>.</summary>
    public bool ObligaMesa { get; private set; }

    /// <summary>Obliga ingreso de PAX. Legacy: <c>lObligaPax</c>.</summary>
    public bool ObligaPax { get; private set; }

    /// <summary>Obliga ingreso de mozo. Legacy: <c>lObligaMozo</c>.</summary>
    public bool ObligaMozo { get; private set; }

    /// <summary>Obliga asignación de motorizado. Legacy: <c>lObligaMotorizado</c>. Regla BR-DEL-001.</summary>
    public bool ObligaMotorizado { get; private set; }

    /// <summary>Habilita canal Central de Pedidos. Legacy: <c>lCanalCentralPedidos</c>. Regla BR-DEL-010.</summary>
    public bool CanalCentralPedidos { get; private set; }

    /// <summary>Habilita canal Delivery. Legacy: <c>lCanalDelivery</c>. Regla BR-DEL-010.</summary>
    public bool CanalDelivery { get; private set; }

    /// <summary>Obliga ingreso de fecha estimada de entrega. Legacy: <c>lObligaIngresoFechaEntrega</c>.</summary>
    public bool ObligaFechaEntrega { get; private set; }

    /// <summary>Obliga selección de cliente frecuente. Legacy: <c>lObligaClienteFrecuente</c>. Regla BR-DEL-001.</summary>
    public bool ObligaClienteFrecuente { get; private set; }

    // Dapper constructor
    private TipoPedidoDelivery() { }

    public static TipoPedidoDelivery Crear(
        string codigoTipoPedido,
        bool activaMotorizado,
        bool obligaMotorizado,
        bool canalDelivery,
        bool canalCentralPedidos,
        bool obligaClienteFrecuente,
        bool obligaFechaEntrega = false)
    {
        if (string.IsNullOrWhiteSpace(codigoTipoPedido))
            throw new DomainException("El código de tipo de pedido es obligatorio.", "TIPO_PEDIDO_CODIGO_REQUERIDO");

        return new TipoPedidoDelivery
        {
            CodigoTipoPedido = codigoTipoPedido,
            ActivaMotorizado = activaMotorizado,
            ObligaMotorizado = obligaMotorizado,
            CanalDelivery = canalDelivery,
            CanalCentralPedidos = canalCentralPedidos,
            ObligaClienteFrecuente = obligaClienteFrecuente,
            ObligaFechaEntrega = obligaFechaEntrega
        };
    }

    /// <summary>
    /// Valida que un pedido delivery cumpla los requisitos del tipo.
    /// Reglas BR-DEL-001, BR-DEL-010.
    /// </summary>
    public void ValidarPedidoDelivery(string? codigoClienteDelivery, string? codigoMotorizado, DateTime? fechaEstimada)
    {
        if (ObligaClienteFrecuente && string.IsNullOrWhiteSpace(codigoClienteDelivery))
            throw new DomainException(
                $"El tipo de pedido '{CodigoTipoPedido}' requiere cliente frecuente.",
                "DELIVERY_CLIENTE_FRECUENTE_REQUERIDO");

        if (ObligaMotorizado && string.IsNullOrWhiteSpace(codigoMotorizado))
            throw new DomainException(
                $"El tipo de pedido '{CodigoTipoPedido}' requiere motorizado asignado.",
                "DELIVERY_MOTORIZADO_REQUERIDO");

        if (ObligaFechaEntrega && !fechaEstimada.HasValue)
            throw new DomainException(
                $"El tipo de pedido '{CodigoTipoPedido}' requiere fecha estimada de entrega.",
                "DELIVERY_FECHA_ENTREGA_REQUERIDA");
    }
}
