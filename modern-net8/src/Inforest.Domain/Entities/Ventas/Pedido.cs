using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Ventas;

/// <summary>
/// Canal de venta del sistema INFOREST.
/// Legacy: modDeclaracion.bas — variables CANAL_VENTA_LOCAL, CANAL_VENTA_DELIVERY, etc.
/// Regla BR-001: Hasta 5 canales con precios independientes por producto.
/// </summary>
public enum CanalVenta
{
    Local = 1,
    Delivery = 2,
    Llevar = 3,
    Canal4 = 4,
    Canal5 = 5
}

/// <summary>
/// Estado de una mesa/pedido.
/// Legacy: modDeclaracion.bas — estados de mesa en MPEDIDO.tEstado
/// </summary>
public enum EstadoPedido
{
    Libre = 0,
    Abierto = 1,
    EnProceso = 2,
    Cerrado = 3,
    Anulado = 4
}

/// <summary>
/// Cabecera de pedido — equivalente a tabla MPEDIDO.
/// Legacy: MPEDIDO en 1. Estructura.sql
/// </summary>
public class Pedido : Entity
{
    public string CodigoPedido { get; private set; } = string.Empty;
    public string? CodigoMesa { get; private set; }
    public string? CodigoTurno { get; private set; }
    public string? CodigoUsuario { get; private set; }
    public CanalVenta CanalVenta { get; private set; }
    public EstadoPedido Estado { get; private set; }
    public DateTime FechaRegistro { get; private set; }
    public DateTime? FechaDiaContable { get; private set; }
    public int NumeroPersonas { get; private set; }
    public string? Observacion { get; private set; }

    private readonly List<DetallePedido> _detalles = [];
    public IReadOnlyList<DetallePedido> Detalles => _detalles.AsReadOnly();

    // EF/Dapper constructor
    private Pedido() { }

    public static Pedido Crear(
        string codigoPedido,
        string? codigoMesa,
        string codigoTurno,
        string codigoUsuario,
        CanalVenta canalVenta,
        int numeroPersonas = 1)
    {
        if (string.IsNullOrWhiteSpace(codigoPedido))
            throw new DomainException("El código de pedido es obligatorio.", "PEDIDO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoTurno))
            throw new DomainException("El turno es obligatorio para crear un pedido.", "PEDIDO_TURNO_REQUERIDO");

        if (numeroPersonas < 0)
            throw new DomainException("El número de personas no puede ser negativo.", "PEDIDO_PERSONAS_INVALIDO");

        return new Pedido
        {
            CodigoPedido = codigoPedido,
            CodigoMesa = codigoMesa,
            CodigoTurno = codigoTurno,
            CodigoUsuario = codigoUsuario,
            CanalVenta = canalVenta,
            Estado = EstadoPedido.Abierto,
            FechaRegistro = DateTime.Now,
            NumeroPersonas = numeroPersonas
        };
    }

    public void AgregarDetalle(DetallePedido detalle)
    {
        if (Estado == EstadoPedido.Cerrado || Estado == EstadoPedido.Anulado)
            throw new DomainException("No se puede modificar un pedido cerrado o anulado.", "PEDIDO_ESTADO_INVALIDO");

        _detalles.Add(detalle);
    }

    public void Anular(string motivo, string usuarioAnulacion)
    {
        if (Estado == EstadoPedido.Cerrado)
            throw new DomainException("No se puede anular un pedido ya cerrado.", "PEDIDO_YA_CERRADO");

        Estado = EstadoPedido.Anulado;
    }
}
