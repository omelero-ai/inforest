using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Configuracion;

/// <summary>
/// Legacy: TDIACONTABLE, clsDiaContable.cls. BR-005.
/// </summary>
public class DiaContable : Entity
{
    public DateTime Fecha { get; private set; }
    public bool Apertura { get; private set; }
    public bool Cierre { get; private set; }
    public string? Usuario { get; private set; }
    public DateTime FechaRegistro { get; private set; }
    public string? UsuarioCierre { get; private set; }
    public DateTime? FechaRegistroCierre { get; private set; }

    private DiaContable() { }

    public static DiaContable Crear(DateTime fecha, string usuario)
    {
        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario es obligatorio para aperturar el día contable.", "DIA_CONTABLE_USUARIO_REQUERIDO");

        var dia = new DiaContable
        {
            Fecha = fecha.Date,
            Apertura = false,
            Cierre = false
        };

        dia.Abrir(usuario);
        return dia;
    }

    public void Abrir(string usuario)
    {
        if (Apertura && !Cierre)
            throw new DomainException("El día contable ya se encuentra aperturado.", "DIA_CONTABLE_YA_ABIERTO");

        if (Cierre)
            throw new DomainException("No se puede reabrir un día contable cerrado.", "DIA_CONTABLE_YA_CERRADO");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario es obligatorio para aperturar el día contable.", "DIA_CONTABLE_USUARIO_REQUERIDO");

        Apertura = true;
        Cierre = false;
        Usuario = usuario.Trim();
        FechaRegistro = DateTime.Now;
        UsuarioCierre = null;
        FechaRegistroCierre = null;
    }

    public void Cerrar(string usuarioCierre)
    {
        if (!Apertura)
            throw new DomainException("No se puede cerrar un día contable no aperturado.", "DIA_CONTABLE_NO_APERTURADO");

        if (Cierre)
            throw new DomainException("El día contable ya se encuentra cerrado.", "DIA_CONTABLE_YA_CERRADO");

        if (string.IsNullOrWhiteSpace(usuarioCierre))
            throw new DomainException("El usuario de cierre es obligatorio.", "DIA_CONTABLE_USUARIO_CIERRE_REQUERIDO");

        Cierre = true;
        UsuarioCierre = usuarioCierre.Trim();
        FechaRegistroCierre = DateTime.Now;
    }
}
