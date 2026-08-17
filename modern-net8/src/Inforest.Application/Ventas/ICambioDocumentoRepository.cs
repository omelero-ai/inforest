namespace Inforest.Application.Ventas;

/// <summary>
/// Contrato de acceso a datos para el proceso de cambio de tipo de documento.
/// Legacy: frmCambio.frm — operaciones directas SQL sobre MDOCUMENTO/DDOCUMENTO/DPEDIDO/DPAGODOCUMENTO/TTIPODOCUMENTOIMPRESORA.
/// </summary>
public interface ICambioDocumentoRepository
{
    /// <summary>
    /// Obtiene los tipos de documento disponibles para la caja indicada, filtrando los
    /// no aptos para cambio (notas de crédito, transporte, inactivos).
    /// Legacy: vTipoDocumentoImpresora — Form_Load filtro.
    /// </summary>
    Task<IReadOnlyList<TipoDocumentoDisponible>> ObtenerTiposDocumentoAsync(string codigoCaja, string pais, CancellationToken ct = default);

    /// <summary>
    /// Devuelve el monto del documento y el monto máximo permitido por el tipo de emisión destino.
    /// Legacy: Calcular(nMontoMaximo de vTipoDocumento) + Calcular(nventa de MDOCUMENTO).
    /// </summary>
    Task<(decimal MontoDocumento, decimal MontoMaximo)> ObtenerMontosValidacionAsync(
        string codigoDocumento, string tipoEmision, CancellationToken ct = default);

    /// <summary>
    /// Ejecuta el cambio de documento: copia DDOCUMENTO+MDOCUMENTO al nuevo número,
    /// actualiza DPEDIDO+DPAGODOCUMENTO, marca el documento original como '04'.
    /// Legacy: múltiples Cn.Execute frmCambio.frm cmdOpcion_Click(0). BR-CAMBIO-004.
    /// </summary>
    Task EjecutarCambioAsync(EjecutarCambioDocumentoParams parametros, CancellationToken ct = default);

    /// <summary>
    /// Actualiza el campo tEmision y tConsumo del nuevo documento.
    /// Legacy: Update MDOCUMENTO set tEmision / tConsumo — frmCambio.frm.
    /// </summary>
    Task ActualizarEmisionAsync(string codigoNuevo, bool esConsumo, string descripcionConsumo, CancellationToken ct = default);
}

/// <summary>Tipo de documento disponible para cambio, cargado desde vTipoDocumentoImpresora.</summary>
public sealed record TipoDocumentoDisponible(
    string TipoEmision,
    string Descripcion,
    string Prefijo,
    string Serie,
    string UltimoNumero,
    string CodigoImpresora,
    bool RequiereCliente,
    bool EsResumen);

/// <summary>Parámetros para ejecutar el cambio de documento.</summary>
public sealed record EjecutarCambioDocumentoParams(
    string CodigoDocumentoOrigen,
    string CodigoDocumentoNuevo,
    string TipoEmision,
    string CodigoCaja,
    string CodigoTurno,
    string CodigoSalon,
    string CodigoUsuario,
    string CodigoClienteNuevo,
    string NuevoCorrelatvo,
    string Motivo,
    string Autorizacion,
    string CodigoControl);
