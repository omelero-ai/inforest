using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Cocina;

/// <summary>
/// Mensaje operativo de cocina/KDS administrado desde mantenimiento POS.
/// Legacy: frmMensajeCocina.frm, frmMensajeCocinaDetalle.frm, TMENSAJECOCINA.
/// Reglas BR-MSGCOC-001, BR-MSGCOC-002, BR-MSGCOC-003.
/// </summary>
public sealed class MensajeCocina
{
    public const int MaximoMensajesActivos = 30;
    public const int LongitudMaximaMensaje = 95;

    public string Codigo { get; private set; } = string.Empty;
    public string UsuarioRegistro { get; private set; } = string.Empty;
    public string UsuarioModificacion { get; private set; } = string.Empty;
    public string Mensaje { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }
    public DateTime FechaModificacion { get; private set; }
    public string Caja { get; private set; } = string.Empty;
    public bool Activo { get; private set; }
    public bool ActivarAlerta { get; private set; }

    private MensajeCocina()
    {
    }

    public static MensajeCocina Crear(
        string codigo,
        string usuario,
        string mensaje,
        string caja,
        bool activo,
        bool activarAlerta,
        DateTime? fechaRegistro = null,
        DateTime? fechaModificacion = null,
        string? usuarioModificacion = null)
    {
        ValidarCodigo(codigo);
        ValidarUsuario(usuario);
        ValidarCaja(caja);
        var mensajeNormalizado = NormalizarMensaje(mensaje);

        return new MensajeCocina
        {
            Codigo = codigo.Trim(),
            UsuarioRegistro = usuario.Trim(),
            UsuarioModificacion = string.IsNullOrWhiteSpace(usuarioModificacion) ? usuario.Trim() : usuarioModificacion.Trim(),
            Mensaje = mensajeNormalizado,
            Caja = caja.Trim(),
            Activo = activo,
            ActivarAlerta = activarAlerta,
            FechaRegistro = fechaRegistro ?? DateTime.Now,
            FechaModificacion = fechaModificacion ?? fechaRegistro ?? DateTime.Now
        };
    }

    public void Actualizar(string usuario, string mensaje, string caja, bool activo, bool activarAlerta)
    {
        ValidarUsuario(usuario);
        ValidarCaja(caja);

        UsuarioModificacion = usuario.Trim();
        Mensaje = NormalizarMensaje(mensaje);
        Caja = caja.Trim();
        Activo = activo;
        ActivarAlerta = activarAlerta;
        FechaModificacion = DateTime.Now;
    }

    private static string NormalizarMensaje(string mensaje)
    {
        if (string.IsNullOrWhiteSpace(mensaje))
            throw new DomainException("Ingrese el mensaje.", "MENSAJE_COCINA_REQUERIDO");

        var valor = mensaje.Trim();
        if (valor.Length > LongitudMaximaMensaje)
            throw new DomainException($"El mensaje no puede exceder {LongitudMaximaMensaje} caracteres.", "MENSAJE_COCINA_LONGITUD_INVALIDA");

        if (valor.Contains('&'))
            throw new DomainException("Caracter (&) no permitido, favor corregir.", "MENSAJE_COCINA_CARACTER_INVALIDO");

        if (valor.Contains('\''))
            throw new DomainException("Caracter (') no permitido, favor corregir.", "MENSAJE_COCINA_CARACTER_INVALIDO");

        if (valor.Contains('"'))
            throw new DomainException("Caracter (\") no permitido, favor corregir.", "MENSAJE_COCINA_CARACTER_INVALIDO");

        return valor.ToUpperInvariant();
    }

    private static void ValidarCodigo(string codigo)
    {
        if (string.IsNullOrWhiteSpace(codigo))
            throw new DomainException("El código del mensaje es obligatorio.", "MENSAJE_COCINA_CODIGO_REQUERIDO");

        if (codigo.Trim().Length != 8)
            throw new DomainException("El código del mensaje debe tener 8 caracteres.", "MENSAJE_COCINA_CODIGO_INVALIDO");
    }

    private static void ValidarUsuario(string usuario)
    {
        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario es obligatorio.", "MENSAJE_COCINA_USUARIO_REQUERIDO");
    }

    private static void ValidarCaja(string caja)
    {
        if (string.IsNullOrWhiteSpace(caja))
            throw new DomainException("La caja es obligatoria.", "MENSAJE_COCINA_CAJA_REQUERIDA");
    }
}
