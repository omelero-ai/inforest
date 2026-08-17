using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Delivery;

/// <summary>
/// Tarjeta RFID/proximidad asociada a un cliente delivery.
/// Legacy: <c>FrmTarjetaAproximidad.frm</c>, <c>FrmTarjetaAproximidadDetalle.frm</c>, tabla <c>TTARJETASRFID</c>.
/// Reglas: BR-RFID-001, BR-RFID-002, BR-RFID-003, BR-RFID-004.
/// </summary>
public sealed class TarjetaProximidad : Entity
{
    private static readonly Dictionary<string, string> EstadosPermitidos = new(StringComparer.OrdinalIgnoreCase)
    {
        ["FREE"] = "Free",
        ["ASIGNADA"] = "Asignada",
        ["BLOQUEADO"] = "Bloqueado"
    };

    private TarjetaProximidad()
    {
    }

    public int Id { get; private set; }
    public string CodigoRfid { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }
    public string UsuarioRegistro { get; private set; } = string.Empty;
    public DateTime FechaModificacion { get; private set; }
    public string UsuarioModificacion { get; private set; } = string.Empty;
    public string Estado { get; private set; } = string.Empty;
    public decimal MontoDisponible { get; private set; }
    public string Descripcion { get; private set; } = string.Empty;
    public string CodigoCliente { get; private set; } = string.Empty;
    public string NombreCliente { get; private set; } = string.Empty;

    public static TarjetaProximidad Crear(
        string codigoRfid,
        string descripcion,
        decimal montoDisponible,
        string codigoCliente,
        string estado,
        string usuarioRegistro,
        DateTime? fechaRegistro = null)
    {
        var tarjeta = new TarjetaProximidad
        {
            CodigoRfid = NormalizarCodigo(codigoRfid),
            Descripcion = NormalizarDescripcion(descripcion),
            CodigoCliente = NormalizarCodigoCliente(codigoCliente),
            Estado = NormalizarEstado(estado),
            MontoDisponible = ValidarMonto(montoDisponible),
            UsuarioRegistro = NormalizarUsuario(usuarioRegistro, "RFID_USUARIO_REGISTRO_REQUERIDO"),
            FechaRegistro = fechaRegistro ?? DateTime.Now
        };

        tarjeta.UsuarioModificacion = tarjeta.UsuarioRegistro;
        tarjeta.FechaModificacion = tarjeta.FechaRegistro;
        return tarjeta;
    }

    public void Actualizar(
        string descripcion,
        decimal montoDisponible,
        string codigoCliente,
        string estado,
        string usuarioModificacion,
        DateTime? fechaModificacion = null)
    {
        Descripcion = NormalizarDescripcion(descripcion);
        CodigoCliente = NormalizarCodigoCliente(codigoCliente);
        Estado = NormalizarEstado(estado);
        MontoDisponible = ValidarMonto(montoDisponible);
        UsuarioModificacion = NormalizarUsuario(usuarioModificacion, "RFID_USUARIO_MODIFICACION_REQUERIDO");
        FechaModificacion = fechaModificacion ?? DateTime.Now;
    }

    public void AsignarNombreCliente(string? nombreCliente)
        => NombreCliente = (nombreCliente ?? string.Empty).Trim();

    private static string NormalizarCodigo(string codigo)
    {
        var valor = (codigo ?? string.Empty).Trim();
        if (valor.Length == 0)
            throw new DomainException("El código de tarjeta es obligatorio.", "RFID_CODIGO_REQUERIDO");

        return valor;
    }

    private static string NormalizarDescripcion(string descripcion)
    {
        var valor = (descripcion ?? string.Empty).Trim();
        if (valor.Length == 0)
            throw new DomainException("La descripción de la tarjeta es obligatoria.", "RFID_DESCRIPCION_REQUERIDA");

        return valor;
    }

    private static string NormalizarCodigoCliente(string codigoCliente)
    {
        var valor = (codigoCliente ?? string.Empty).Trim();
        if (valor.Length == 0)
            throw new DomainException("Debe asociar un cliente delivery a la tarjeta.", "RFID_CLIENTE_REQUERIDO");

        return valor;
    }

    private static string NormalizarUsuario(string usuario, string errorCode)
    {
        var valor = (usuario ?? string.Empty).Trim();
        if (valor.Length == 0)
            throw new DomainException("El usuario es obligatorio para registrar la operación.", errorCode);

        return valor;
    }

    private static string NormalizarEstado(string estado)
    {
        var clave = (estado ?? string.Empty).Trim().ToUpperInvariant();
        if (!EstadosPermitidos.TryGetValue(clave, out var valor))
            throw new DomainException(
                "El estado de la tarjeta debe ser Free, Asignada o Bloqueado.",
                "RFID_ESTADO_INVALIDO");

        return valor;
    }

    private static decimal ValidarMonto(decimal montoDisponible)
    {
        if (montoDisponible < 0)
            throw new DomainException("El monto disponible no puede ser negativo.", "RFID_MONTO_INVALIDO");

        return decimal.Round(montoDisponible, 2, MidpointRounding.AwayFromZero);
    }
}

/// <summary>
/// Movimiento reciente de una tarjeta RFID/proximidad.
/// Legacy: impresión de últimos 10 movimientos en <c>FrmTarjetaAproximidad.frm</c>.
/// </summary>
public sealed record MovimientoTarjetaProximidad(
    DateTime FechaRegistro,
    decimal MontoIngreso,
    decimal MontoSalida,
    decimal MontoAnterior,
    decimal MontoFinal);
