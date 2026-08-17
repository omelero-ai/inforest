namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Estados operativos de una mesa en salón.
/// Legacy: TMESA.tEstadoMesa — códigos '01'..'06' / frmMesa.frm / frmMesas.frm. BR-003, BR-MESAS-001.
/// </summary>
public enum EstadoMesa
{
    /// <summary>'01' — Mesa libre y limpia. Seleccionable para nuevo pedido.</summary>
    Libre = 0,
    /// <summary>'02' — Mesa con pedido activo en curso.</summary>
    Ocupada = 1,
    /// <summary>'03' — Mesa reservada. No seleccionable operativamente.</summary>
    Reservada = 2,
    /// <summary>'04' — Mesa sucia (precuenta emitida o limpieza pendiente). Seleccionable para nuevo pedido.</summary>
    Sucia = 3,
    /// <summary>'05' — Mesa bloqueada administrativamente. No seleccionable.</summary>
    Bloqueada = 4,
    /// <summary>'06' — Mesa fuera de servicio. No seleccionable.</summary>
    FueraDeServicio = 5,
    /// <summary>Estado de transición interno: mesa en cuenta (precuenta enviada, aún no cerrada).</summary>
    EnCuenta = 6
}
