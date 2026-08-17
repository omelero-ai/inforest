using Inforest.Application.Interfaces;
using Inforest.Domain.Common;

namespace Inforest.Application.Ventas;

// ── Commands ────────────────────────────────────────────────────────────────

/// <summary>
/// Comando para cambiar el tipo de documento de un comprobante ya emitido.
/// Legacy: frmCambio.frm — cmdOpcion_Click(0). BR-CAMBIO-001..005.
/// </summary>
public sealed record CambiarDocumentoCommand(
    /// <summary>Documento origen a cambiar (ej. "B-00001-000000042").</summary>
    string CodigoDocumentoOrigen,
    /// <summary>Tipo de emisión destino (ej. "01" factura, "03" boleta).</summary>
    string TipoEmisionDestino,
    /// <summary>Nuevo número de documento generado por correlativo.</summary>
    string CodigoDocumentoNuevo,
    /// <summary>Código de cliente (requerido si el tipo lo exige). BR-CAMBIO-001.</summary>
    string CodigoClienteNuevo,
    /// <summary>Motivo del cambio. BR-CAMBIO-003.</summary>
    string Motivo,
    /// <summary>Si se emite por consumo (cubre para boleta de consumo).</summary>
    bool EsConsumo,
    /// <summary>Descripción del consumo cuando EsConsumo=true.</summary>
    string DescripcionConsumo,
    /// <summary>Turno activo al momento del cambio (sTurno global Legacy).</summary>
    string CodigoTurno = "",
    /// <summary>Salón activo al momento del cambio (sSalon global Legacy).</summary>
    string CodigoSalon = "",
    /// <summary>Código Bolivia: número de autorización. BR-CAMBIO-005.</summary>
    string Autorizacion = "",
    /// <summary>Código Bolivia: código de control. BR-CAMBIO-005.</summary>
    string CodigoControl = "");

// ── DTOs / Queries ───────────────────────────────────────────────────────────

/// <summary>Obtiene los tipos de documento disponibles para cambio en la caja actual.</summary>
public sealed record ObtenerTiposDocumentoCambioQuery(string CodigoCaja, string Pais);

/// <summary>Resultado del handler, contiene el código del nuevo documento.</summary>
public sealed record CambiarDocumentoResultado(string CodigoDocumentoNuevo);

// ── Handlers ────────────────────────────────────────────────────────────────

/// <summary>
/// Obtiene los tipos de documento disponibles para cambiar en la caja indicada.
/// Legacy: Form_Load frmCambio.frm → vTipoDocumentoImpresora filtro.
/// </summary>
public sealed class ObtenerTiposDocumentoCambioHandler
{
    private readonly ICambioDocumentoRepository _repo;

    public ObtenerTiposDocumentoCambioHandler(ICambioDocumentoRepository repo)
        => _repo = repo;

    public async Task<Result<IReadOnlyList<TipoDocumentoDisponible>>> HandleAsync(
        ObtenerTiposDocumentoCambioQuery query,
        CancellationToken ct = default)
    {
        try
        {
            var tipos = await _repo.ObtenerTiposDocumentoAsync(query.CodigoCaja, query.Pais, ct);
            return Result.Ok(tipos);
        }
        catch (Exception ex)
        {
            return Result.Fail<IReadOnlyList<TipoDocumentoDisponible>>(ex.Message, "CAMBIO_TIPOS_ERROR");
        }
    }
}

/// <summary>
/// Ejecuta el cambio de tipo de documento.
/// Legacy: cmdOpcion_Click(0) frmCambio.frm.
/// <para>BR-CAMBIO-001: Si el tipo requiere cliente, el código de cliente no puede estar vacío.</para>
/// <para>BR-CAMBIO-002: El monto del documento no puede superar el nMontoMaximo del tipo destino.</para>
/// <para>BR-CAMBIO-003: Se requiere motivo del cambio.</para>
/// <para>BR-CAMBIO-004: El documento origen queda en estado '04' (cambiado).</para>
/// <para>BR-CAMBIO-005: Bolivia requiere autorización/dosificación (a cargo del caller).</para>
/// </summary>
public sealed class CambiarDocumentoHandler
{
    private readonly ICambioDocumentoRepository _repo;
    private readonly ISessionService _sessionService;

    public CambiarDocumentoHandler(
        ICambioDocumentoRepository repo,
        ISessionService sessionService)
    {
        _repo = repo;
        _sessionService = sessionService;
    }

    public async Task<Result<CambiarDocumentoResultado>> HandleAsync(
        CambiarDocumentoCommand command,
        CancellationToken ct = default)
    {
        // BR-CAMBIO-003: Motivo requerido
        if (string.IsNullOrWhiteSpace(command.Motivo))
            return Result.Fail<CambiarDocumentoResultado>(
                "Se requiere el motivo del cambio de documento.", "CAMBIO_MOTIVO_REQUERIDO");

        // BR-CAMBIO-001: Cliente requerido si el tipo lo exige (validado externamente por ObtenerTipos)
        if (string.IsNullOrWhiteSpace(command.CodigoDocumentoNuevo))
            return Result.Fail<CambiarDocumentoResultado>(
                "El código del nuevo documento es requerido.", "CAMBIO_NUEVO_REQUERIDO");

        var sesion = _sessionService.SesionActual;
        if (sesion is null)
            return Result.Fail<CambiarDocumentoResultado>("No existe sesión activa.", "CAMBIO_SIN_SESION");

        try
        {
            // BR-CAMBIO-002: Validar monto máximo
            var (montoDoc, montoMax) = await _repo.ObtenerMontosValidacionAsync(
                command.CodigoDocumentoOrigen, command.TipoEmisionDestino, ct);

            if (montoMax > 0 && montoDoc > montoMax)
                return Result.Fail<CambiarDocumentoResultado>(
                    $"El monto del documento ({montoDoc:0.00}) supera el máximo permitido ({montoMax:0.00}) para el tipo de documento.",
                    "CAMBIO_MONTO_EXCEDE_MAXIMO");

            // BR-CAMBIO-004: Ejecutar cambio (copia + marca origen como '04')
            var parametros = new EjecutarCambioDocumentoParams(
                CodigoDocumentoOrigen: command.CodigoDocumentoOrigen,
                CodigoDocumentoNuevo: command.CodigoDocumentoNuevo,
                TipoEmision: command.TipoEmisionDestino,
                CodigoCaja: sesion.CodigoCaja,
                CodigoTurno: command.CodigoTurno,
                CodigoSalon: command.CodigoSalon,
                CodigoUsuario: sesion.CodigoUsuario,
                CodigoClienteNuevo: command.CodigoClienteNuevo,
                NuevoCorrelatvo: ObtenerCorrelativo(command.CodigoDocumentoNuevo),
                Motivo: command.Motivo,
                Autorizacion: command.Autorizacion,
                CodigoControl: command.CodigoControl);

            await _repo.EjecutarCambioAsync(parametros, ct);

            // Marcar tipo emisión / consumo
            await _repo.ActualizarEmisionAsync(
                command.CodigoDocumentoNuevo,
                command.EsConsumo,
                command.DescripcionConsumo,
                ct);

            return Result.Ok(new CambiarDocumentoResultado(command.CodigoDocumentoNuevo));
        }
        catch (Exception ex)
        {
            return Result.Fail<CambiarDocumentoResultado>(ex.Message, "CAMBIO_ERROR");
        }
    }

    /// <summary>Extrae la parte numérica del correlativo del código de documento.</summary>
    private static string ObtenerCorrelativo(string codigoDocumento)
    {
        // Formato: P-SSSSS-NNNNNNNNN  (P=prefijo, S=serie, N=número)
        if (codigoDocumento.Length >= 15)
            return codigoDocumento[^9..];
        return codigoDocumento;
    }
}
