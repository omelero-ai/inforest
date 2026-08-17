namespace Inforest.Domain.Entities.Configuracion;

/// <summary>
/// Motivo de anulación de ítems/documentos en el POS.
/// Legacy: TTABLA WHERE TTABLA = 'MOTIVOANULACION' / vMotivoAnulacion.
/// BR-MOTIVO-001.
/// </summary>
public sealed record MotivoAnulacion(
    /// <summary>Código del motivo (3 chars). Columna: TCODIGO / Codigo.</summary>
    string Codigo,
    /// <summary>Descripción larga. Columna: tDetallado / Descripcion.</summary>
    string Descripcion,
    /// <summary>Descripción resumida. Columna: tResumido.</summary>
    string Resumido,
    /// <summary>Activo en catálogo. BR-MOTIVO-001: solo se muestran activos.</summary>
    bool Activo,
    /// <summary>Valor numérico opcional (nValor).</summary>
    double? ValorNumerico = null
);
