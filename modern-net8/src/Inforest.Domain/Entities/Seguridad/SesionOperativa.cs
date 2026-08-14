using Inforest.Domain.Common;

namespace Inforest.Domain.Entities.Seguridad;

/// <summary>
/// Contexto de sesión operativa por usuario/caja/terminal.
/// Legacy: frmAcceso.frm + modAuditoriaIntegral.bas + USUARIO.INI.
/// Regla BR-006 y BR-014.
/// </summary>
public sealed class SesionOperativa
{
    private readonly List<PermisoAcceso> _permisos;

    private SesionOperativa(
        Guid id,
        string codigoUsuario,
        string nombreUsuario,
        string codigoGrupoUsuario,
        string modulo,
        string codigoCaja,
        string codigoTerminal,
        string baseDatos,
        long correlativoAcceso,
        DateTime fechaIngreso,
        IEnumerable<PermisoAcceso> permisos)
    {
        Id = id;
        CodigoUsuario = codigoUsuario;
        NombreUsuario = nombreUsuario;
        CodigoGrupoUsuario = codigoGrupoUsuario;
        Modulo = modulo;
        CodigoCaja = codigoCaja;
        CodigoTerminal = codigoTerminal;
        BaseDatos = baseDatos;
        CorrelativoAcceso = correlativoAcceso;
        FechaIngreso = fechaIngreso;
        _permisos = permisos.ToList();
        EstaActiva = true;
    }

    public Guid Id { get; }
    public string CodigoUsuario { get; }
    public string NombreUsuario { get; }
    public string CodigoGrupoUsuario { get; }
    public string Modulo { get; }
    public string CodigoCaja { get; }
    public string CodigoTerminal { get; }
    public string BaseDatos { get; }
    public long CorrelativoAcceso { get; }
    public DateTime FechaIngreso { get; }
    public DateTime? FechaSalida { get; private set; }
    public bool EstaActiva { get; private set; }
    public IReadOnlyCollection<PermisoAcceso> Permisos => _permisos;

    public static SesionOperativa Iniciar(
        string codigoUsuario,
        string nombreUsuario,
        string codigoGrupoUsuario,
        string modulo,
        string codigoCaja,
        string codigoTerminal,
        string baseDatos,
        long correlativoAcceso,
        DateTime fechaIngreso,
        IEnumerable<PermisoAcceso> permisos)
    {
        if (string.IsNullOrWhiteSpace(codigoUsuario))
            throw new ArgumentException("El código de usuario es obligatorio.", nameof(codigoUsuario));

        if (string.IsNullOrWhiteSpace(modulo))
            throw new ArgumentException("El módulo es obligatorio.", nameof(modulo));

        return new SesionOperativa(
            Guid.NewGuid(),
            codigoUsuario.Trim().ToUpperInvariant(),
            nombreUsuario.Trim(),
            codigoGrupoUsuario.Trim().ToUpperInvariant(),
            modulo.Trim().ToUpperInvariant(),
            codigoCaja.Trim().ToUpperInvariant(),
            codigoTerminal.Trim().ToUpperInvariant(),
            baseDatos.Trim().ToUpperInvariant(),
            correlativoAcceso,
            fechaIngreso,
            permisos);
    }

    public bool TienePermiso(string codigoAcceso)
        => _permisos.Any(p => string.Equals(p.CodigoAcceso, codigoAcceso, StringComparison.OrdinalIgnoreCase));

    public Result Cerrar(DateTime fechaSalida)
    {
        if (!EstaActiva)
            return Result.Fail("La sesión ya está cerrada.", "SEGURIDAD_SESION_CERRADA");

        FechaSalida = fechaSalida;
        EstaActiva = false;
        return Result.Ok();
    }
}
