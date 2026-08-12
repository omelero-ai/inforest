namespace Inforest.Domain.Entities.Hardware;

/// <summary>
/// Documento fiscal para impresora fiscal Epson (Argentina).
/// Legacy: modImpresoraFiscal.bas — zonas de formulario (SetZone), ítems de venta.
/// Regla BR-PERIPH-005.
/// </summary>
public sealed class FiscalDocumento
{
    public string LetraComprobante { get; }
    public string RazonSocial { get; }
    public string RazonComercial { get; }
    public string Cuit { get; }
    public string Direccion { get; }
    public string Telefono { get; }
    public string NumeroFactura { get; }
    public DateOnly FechaEmision { get; }
    public string? CuitComprador { get; }
    public string? RazonSocialComprador { get; }
    public IReadOnlyList<FiscalItem> Items { get; }
    public decimal Total { get; }

    private FiscalDocumento(
        string letraComprobante, string razonSocial, string razonComercial,
        string cuit, string direccion, string telefono,
        string numeroFactura, DateOnly fechaEmision,
        string? cuitComprador, string? razonSocialComprador,
        IReadOnlyList<FiscalItem> items, decimal total)
    {
        LetraComprobante = letraComprobante;
        RazonSocial = razonSocial;
        RazonComercial = razonComercial;
        Cuit = cuit;
        Direccion = direccion;
        Telefono = telefono;
        NumeroFactura = numeroFactura;
        FechaEmision = fechaEmision;
        CuitComprador = cuitComprador;
        RazonSocialComprador = razonSocialComprador;
        Items = items;
        Total = total;
    }

    public static FiscalDocumento Crear(
        string letraComprobante, string razonSocial, string razonComercial,
        string cuit, string direccion, string telefono,
        string numeroFactura, DateOnly fechaEmision,
        string? cuitComprador, string? razonSocialComprador,
        IReadOnlyList<FiscalItem> items, decimal total)
    {
        if (string.IsNullOrWhiteSpace(letraComprobante)) throw new ArgumentException("Letra de comprobante requerida.", nameof(letraComprobante));
        if (items == null || items.Count == 0) throw new ArgumentException("El documento fiscal debe tener al menos un ítem.", nameof(items));
        if (total <= 0) throw new ArgumentOutOfRangeException(nameof(total), "El total debe ser mayor a cero.");

        return new FiscalDocumento(letraComprobante, razonSocial, razonComercial,
            cuit, direccion, telefono, numeroFactura, fechaEmision,
            cuitComprador, razonSocialComprador, items, total);
    }
}

/// <summary>
/// Ítem de documento fiscal para impresora Epson.
/// Legacy: zonas 61 (cantidad), 62 (descripción), 63 (precio), 64 (tasa IVA), 65 (tasa ajuste).
/// </summary>
public sealed record FiscalItem(
    decimal Cantidad,
    string Descripcion,
    decimal PrecioUnitario,
    decimal TasaIva,
    decimal TasaAjuste = 0);
