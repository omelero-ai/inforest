namespace Inforest.Domain.Entities.Hardware;

/// <summary>
/// Documento de Facturación Electrónica.
/// Legacy: clsTrama.cls, clsxml.cls — XML de comprobante electrónico por país.
/// Regla BR-PERIPH-006.
/// </summary>
public sealed class DocumentoFE
{
    public string Id { get; }
    public string TipoDocumento { get; }
    public string Serie { get; }
    public string Numero { get; }
    public string CodigoPais { get; }
    public string NumeroDocumentoEmisor { get; }
    public string RazonSocialEmisor { get; }
    public string NumeroDocumentoReceptor { get; }
    public string RazonSocialReceptor { get; }
    public IReadOnlyList<ItemDocumentoFE> Items { get; }
    public decimal SubTotal { get; }
    public decimal Impuesto { get; }
    public decimal Total { get; }
    public DateTime FechaEmision { get; }
    public string? NumeroAutorizacion { get; private set; }
    public string? QrData { get; private set; }
    public string? XmlGenerado { get; private set; }
    public EstadoDocumentoFE Estado { get; private set; }

    private DocumentoFE(
        string id,
        string tipoDocumento,
        string serie,
        string numero,
        string codigoPais,
        string numeroDocumentoEmisor,
        string razonSocialEmisor,
        string numeroDocumentoReceptor,
        string razonSocialReceptor,
        IReadOnlyList<ItemDocumentoFE> items,
        decimal subTotal,
        decimal impuesto,
        decimal total,
        DateTime fechaEmision)
    {
        Id = id;
        TipoDocumento = tipoDocumento;
        Serie = serie;
        Numero = numero;
        CodigoPais = codigoPais;
        NumeroDocumentoEmisor = numeroDocumentoEmisor;
        RazonSocialEmisor = razonSocialEmisor;
        NumeroDocumentoReceptor = numeroDocumentoReceptor;
        RazonSocialReceptor = razonSocialReceptor;
        Items = items;
        SubTotal = subTotal;
        Impuesto = impuesto;
        Total = total;
        FechaEmision = fechaEmision;
        Estado = EstadoDocumentoFE.Pendiente;
    }

    public static DocumentoFE Crear(
        string id,
        string tipoDocumento,
        string serie,
        string numero,
        string codigoPais,
        string numeroDocumentoEmisor,
        string razonSocialEmisor,
        string numeroDocumentoReceptor,
        string razonSocialReceptor,
        IReadOnlyList<ItemDocumentoFE> items,
        decimal subTotal,
        decimal impuesto,
        decimal total,
        DateTime fechaEmision)
    {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Id requerido.", nameof(id));
        if (string.IsNullOrWhiteSpace(tipoDocumento)) throw new ArgumentException("TipoDocumento requerido.", nameof(tipoDocumento));
        if (items == null || items.Count == 0) throw new ArgumentException("El documento debe tener al menos un ítem.", nameof(items));
        if (total <= 0) throw new ArgumentOutOfRangeException(nameof(total), "El total debe ser mayor a cero.");

        return new DocumentoFE(id, tipoDocumento, serie, numero, codigoPais,
            numeroDocumentoEmisor, razonSocialEmisor,
            numeroDocumentoReceptor, razonSocialReceptor,
            items, subTotal, impuesto, total, fechaEmision);
    }

    /// <summary>Registra la autorización recibida del proveedor FE.</summary>
    public void RegistrarAutorizacion(string numeroAutorizacion, string? qrData, string? xml)
    {
        NumeroAutorizacion = numeroAutorizacion;
        QrData = qrData;
        XmlGenerado = xml;
        Estado = EstadoDocumentoFE.Autorizado;
    }

    /// <summary>Marca el documento como rechazado por el proveedor FE.</summary>
    public void RegistrarRechazo(string motivoRechazo)
    {
        Estado = EstadoDocumentoFE.Rechazado;
    }
}

/// <summary>Ítem de un documento de FE.</summary>
public sealed record ItemDocumentoFE(
    string Codigo,
    string Descripcion,
    decimal Cantidad,
    decimal PrecioUnitario,
    decimal PrecioTotal,
    decimal Impuesto,
    string UnidadMedida = "NIU");

/// <summary>Estado del documento electrónico.</summary>
public enum EstadoDocumentoFE
{
    Pendiente,
    Enviado,
    Autorizado,
    Rechazado,
    Anulado
}
