using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Almacen;

/// <summary>
/// Requerimiento de almacén pendiente de importar al POS como pedido.
/// Legacy: vRequerimiento (ALMACEN DB) + MREQUERIMIENTO.
/// frmImportacionRequerimientos.frm / frmImportacionRequerimientoDetalle.frm.
/// BR-IMPORT-001, BR-IMPORT-002, BR-IMPORT-003, BR-IMPORT-004.
/// </summary>
public sealed class RequerimientoAlmacen
{
    /// <summary>Código del requerimiento (Rq).</summary>
    public string Rq { get; }

    /// <summary>Fecha del requerimiento.</summary>
    public DateTime Fecha { get; }

    /// <summary>Área de origen.</summary>
    public string Area { get; }

    /// <summary>Responsable del requerimiento.</summary>
    public string Responsable { get; }

    /// <summary>Estado del requerimiento (CodEstado='02' = aprobado/pendiente importar).</summary>
    public string CodEstado { get; }

    /// <summary>
    /// Código de pedido generado a partir de este requerimiento.
    /// Null si aún no fue importado.
    /// Legacy: MREQUERIMIENTO.tPedido
    /// </summary>
    public string? CodigoPedido { get; }

    /// <summary>
    /// BR-IMPORT-002: Un requerimiento no puede importarse dos veces.
    /// Retorna true si ya tiene un pedido generado.
    /// </summary>
    public bool YaImportado => !string.IsNullOrWhiteSpace(CodigoPedido);

    private RequerimientoAlmacen(
        string rq,
        DateTime fecha,
        string area,
        string responsable,
        string codEstado,
        string? codigoPedido)
    {
        Rq = rq;
        Fecha = fecha;
        Area = area;
        Responsable = responsable;
        CodEstado = codEstado;
        CodigoPedido = codigoPedido;
    }

    /// <summary>
    /// Reconstruye un requerimiento de almacén a partir de datos de BD.
    /// BR-IMPORT-001: Solo requerimientos con CodEstado='02' pueden importarse.
    /// </summary>
    public static RequerimientoAlmacen Crear(
        string rq,
        DateTime fecha,
        string area,
        string responsable,
        string codEstado,
        string? codigoPedido = null)
    {
        if (string.IsNullOrWhiteSpace(rq))
            throw new DomainException("El código de requerimiento no puede estar vacío.", "REQ_CODIGO_REQUERIDO");

        return new RequerimientoAlmacen(
            rq.Trim(),
            fecha,
            area?.Trim() ?? string.Empty,
            responsable?.Trim() ?? string.Empty,
            codEstado?.Trim() ?? string.Empty,
            codigoPedido?.Trim());
    }

    /// <summary>
    /// BR-IMPORT-001: Valida que el requerimiento puede ser importado.
    /// Solo los que están en estado '02' (aprobado) pueden importarse.
    /// </summary>
    public bool PuedeImportarse()
        => CodEstado == "02" && !YaImportado;
}
