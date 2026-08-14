using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Reservas;

/// <summary>
/// Estado de una reserva.
/// Legacy: TRESERVA.tEstadoReserva — '01'=Pendiente, '02'=Atendida, '03'=Anulada
/// Regla BR-RESERVA-001
/// </summary>
public enum EstadoReserva
{
    Pendiente = 1,   // '01'
    Atendida  = 2,   // '02'
    Anulada   = 3    // '03'
}

/// <summary>
/// Reserva de mesa/comedor — equivalente a tabla TRESERVA.
/// Legacy: frmReserva.frm, frmReservaDetalle.frm, TRESERVA
/// Reglas: BR-RESERVA-001, BR-RESERVA-002, BR-RESERVA-003, BR-RESERVA-004
/// </summary>
public class Reserva : Entity
{
    public string CodigoReserva { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }
    public DateTime FechaReserva { get; private set; }
    public string HoraReserva { get; private set; } = string.Empty;
    public string? CodigoCliente { get; private set; }
    public string? NombreCliente { get; private set; }
    public string? Telefono { get; private set; }
    public string? CodigoMotivo { get; private set; }
    public string? CodigoMesero { get; private set; }
    public int Pax { get; private set; }
    public string? CodigoMesa { get; private set; }
    public EstadoReserva Estado { get; private set; }
    public string? Observacion { get; private set; }
    public DateTime? FechaDiaContable { get; private set; }

    // EF/Dapper constructor
    private Reserva() { }

    /// <summary>
    /// Crea una nueva reserva.
    /// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 1 (Sw=True, INSERT INTO TRESERVA)
    /// BR-RESERVA-001: tEstadoReserva='01' (Pendiente) al crear
    /// BR-RESERVA-002: Cliente requerido
    /// BR-RESERVA-003: Pax requerido > 0
    /// </summary>
    public static Reserva Crear(
        string codigoReserva,
        DateTime fechaReserva,
        string horaReserva,
        string? codigoCliente,
        string? nombreCliente,
        string? telefono,
        string? codigoMotivo,
        string? codigoMesero,
        int pax,
        string? codigoMesa,
        string? observacion,
        DateTime? fechaDiaContable)
    {
        if (string.IsNullOrWhiteSpace(codigoReserva))
            throw new DomainException("El código de reserva es requerido.", "RESERVA_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoCliente) && string.IsNullOrWhiteSpace(nombreCliente))
            throw new DomainException("El cliente es requerido para la reserva.", "RESERVA_CLIENTE_REQUERIDO");

        if (pax <= 0)
            throw new DomainException("El número de PAX debe ser mayor a cero.", "RESERVA_PAX_INVALIDO");

        return new Reserva
        {
            CodigoReserva    = codigoReserva.Trim(),
            FechaRegistro    = DateTime.Now,
            FechaReserva     = fechaReserva.Date,
            HoraReserva      = horaReserva?.Trim() ?? string.Empty,
            CodigoCliente    = string.IsNullOrWhiteSpace(codigoCliente) ? null : codigoCliente.Trim(),
            NombreCliente    = string.IsNullOrWhiteSpace(nombreCliente) ? null : nombreCliente.Trim(),
            Telefono         = string.IsNullOrWhiteSpace(telefono) ? null : telefono.Trim(),
            CodigoMotivo     = string.IsNullOrWhiteSpace(codigoMotivo) ? null : codigoMotivo.Trim(),
            CodigoMesero     = string.IsNullOrWhiteSpace(codigoMesero) ? null : codigoMesero.Trim(),
            Pax              = pax,
            CodigoMesa       = string.IsNullOrWhiteSpace(codigoMesa) ? null : codigoMesa.Trim(),
            Estado           = EstadoReserva.Pendiente,
            Observacion      = string.IsNullOrWhiteSpace(observacion) ? null : observacion.Trim(),
            FechaDiaContable = fechaDiaContable
        };
    }

    /// <summary>
    /// Modifica datos de la reserva.
    /// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 1 (Sw=False, UPDATE TRESERVA)
    /// BR-RESERVA-004: No se puede modificar si ya está Atendida o Anulada
    /// </summary>
    public void Modificar(
        DateTime fechaReserva,
        string horaReserva,
        string? codigoCliente,
        string? nombreCliente,
        string? telefono,
        string? codigoMotivo,
        string? codigoMesero,
        int pax,
        string? codigoMesa,
        string? observacion)
    {
        if (Estado == EstadoReserva.Atendida)
            throw new DomainException("No se puede modificar una reserva ya atendida.", "RESERVA_YA_ATENDIDA");

        if (Estado == EstadoReserva.Anulada)
            throw new DomainException("No se puede modificar una reserva anulada.", "RESERVA_YA_ANULADA");

        if (string.IsNullOrWhiteSpace(codigoCliente) && string.IsNullOrWhiteSpace(nombreCliente))
            throw new DomainException("El cliente es requerido para la reserva.", "RESERVA_CLIENTE_REQUERIDO");

        if (pax <= 0)
            throw new DomainException("El número de PAX debe ser mayor a cero.", "RESERVA_PAX_INVALIDO");

        FechaReserva  = fechaReserva.Date;
        HoraReserva   = horaReserva?.Trim() ?? string.Empty;
        CodigoCliente = string.IsNullOrWhiteSpace(codigoCliente) ? null : codigoCliente.Trim();
        NombreCliente = string.IsNullOrWhiteSpace(nombreCliente) ? null : nombreCliente.Trim();
        Telefono      = string.IsNullOrWhiteSpace(telefono) ? null : telefono.Trim();
        CodigoMotivo  = string.IsNullOrWhiteSpace(codigoMotivo) ? null : codigoMotivo.Trim();
        CodigoMesero  = string.IsNullOrWhiteSpace(codigoMesero) ? null : codigoMesero.Trim();
        Pax           = pax;
        CodigoMesa    = string.IsNullOrWhiteSpace(codigoMesa) ? null : codigoMesa.Trim();
        Observacion   = string.IsNullOrWhiteSpace(observacion) ? null : observacion.Trim();
    }

    /// <summary>
    /// Anula la reserva.
    /// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 2 (UPDATE TRESERVA SET tEstadoReserva='03')
    /// BR-RESERVA-004: No se puede anular si ya fue anulada
    /// </summary>
    public void Anular()
    {
        if (Estado == EstadoReserva.Anulada)
            throw new DomainException("La reserva ya fue anulada.", "RESERVA_YA_ANULADA");

        Estado = EstadoReserva.Anulada;
    }

    /// <summary>
    /// Marca la reserva como atendida (vinculada a un pedido).
    /// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 4 (spIns_MPEDIDO_RESERVA)
    /// BR-RESERVA-004: Estado '01' → '02'
    /// </summary>
    public void MarcarAtendida()
    {
        if (Estado != EstadoReserva.Pendiente)
            throw new DomainException("Solo se puede atender una reserva pendiente.", "RESERVA_NO_PENDIENTE");

        Estado = EstadoReserva.Atendida;
    }

    /// <summary>Reconstruye entidad desde persistencia.</summary>
    public static Reserva Reconstruir(
        string codigoReserva,
        DateTime fechaRegistro,
        DateTime fechaReserva,
        string horaReserva,
        string? codigoCliente,
        string? nombreCliente,
        string? telefono,
        string? codigoMotivo,
        string? codigoMesero,
        int pax,
        string? codigoMesa,
        EstadoReserva estado,
        string? observacion,
        DateTime? fechaDiaContable)
    {
        return new Reserva
        {
            CodigoReserva    = codigoReserva,
            FechaRegistro    = fechaRegistro,
            FechaReserva     = fechaReserva,
            HoraReserva      = horaReserva,
            CodigoCliente    = codigoCliente,
            NombreCliente    = nombreCliente,
            Telefono         = telefono,
            CodigoMotivo     = codigoMotivo,
            CodigoMesero     = codigoMesero,
            Pax              = pax,
            CodigoMesa       = codigoMesa,
            Estado           = estado,
            Observacion      = observacion,
            FechaDiaContable = fechaDiaContable
        };
    }
}
