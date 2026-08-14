using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Cocina;

/// <summary>
/// Orden de producción para cocina/KDS.
/// Legacy: MPEDIDO/DPEDIDO + modKDS.bas.
/// Reglas BR-010 y BR-011.
/// </summary>
public sealed class OrdenProduccion
{
    private readonly List<ItemProduccion> _items;

    public string CodigoPedido { get; }
    public string CodigoCaja { get; }
    public string? CodigoMesa { get; }
    public string? NombreMesa { get; }
    public string? Destino { get; }
    public string? NombreServidor { get; }
    public string? UsuarioAccion { get; }
    public IReadOnlyList<ItemProduccion> Items => _items.AsReadOnly();

    private OrdenProduccion(
        string codigoPedido,
        string codigoCaja,
        string? codigoMesa,
        string? nombreMesa,
        string? destino,
        string? nombreServidor,
        string? usuarioAccion,
        List<ItemProduccion> items)
    {
        CodigoPedido = codigoPedido;
        CodigoCaja = codigoCaja;
        CodigoMesa = codigoMesa;
        NombreMesa = nombreMesa;
        Destino = destino;
        NombreServidor = nombreServidor;
        UsuarioAccion = usuarioAccion;
        _items = items;
    }

    public static OrdenProduccion Crear(
        string codigoPedido,
        string codigoCaja,
        string? codigoMesa,
        string? nombreMesa,
        string? destino,
        string? nombreServidor,
        string? usuarioAccion,
        IEnumerable<ItemProduccion>? items = null)
    {
        if (string.IsNullOrWhiteSpace(codigoPedido))
            throw new DomainException("El código del pedido es obligatorio.", "COCINA_PEDIDO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(codigoCaja))
            throw new DomainException("La caja es obligatoria.", "COCINA_CAJA_REQUERIDA");

        return new OrdenProduccion(
            codigoPedido.Trim(),
            codigoCaja.Trim(),
            Limpiar(codigoMesa),
            Limpiar(nombreMesa),
            Limpiar(destino),
            Limpiar(nombreServidor),
            Limpiar(usuarioAccion),
            items?.ToList() ?? []);
    }

    public int ObtenerIdNumerico()
    {
        var baseValue = CodigoPedido.Length > 2
            ? CodigoPedido[2..]
            : CodigoPedido;

        var digits = new string(baseValue.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, out var id)
            ? id
            : 0;
    }

    public string ObtenerNombreMesa()
        => !string.IsNullOrWhiteSpace(NombreMesa)
            ? NombreMesa!
            : CodigoMesa ?? string.Empty;

    public static string ObtenerIniciales(string? nombre)
    {
        if (string.IsNullOrWhiteSpace(nombre))
            return string.Empty;

        var partes = nombre
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        return string.Concat(partes.Select(parte => parte[0]));
    }

    private static string? Limpiar(string? valor)
        => string.IsNullOrWhiteSpace(valor) ? null : valor.Trim();
}
