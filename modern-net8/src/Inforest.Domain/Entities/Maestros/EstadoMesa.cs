namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Estados operativos de una mesa en salón.
/// Legacy: TMESA.tEstadoMesa / frmMesa.frm. BR-003.
/// </summary>
public enum EstadoMesa
{
    Libre = 0,
    Ocupada = 1,
    Reservada = 2,
    EnCuenta = 3
}
