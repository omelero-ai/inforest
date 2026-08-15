namespace Inforest.Application.Interfaces.Hardware;

/// <summary>
/// Abstracción para el servicio de PinPad (cobro con tarjeta).
/// Legacy: DLL3500.bas — caja_pinpad.dll (fiOpenPort, fiClosePort, fiStartOperation, fiGetStatus).
/// Regla BR-PERIPH-001. ADR-009.
/// </summary>
public interface IPinPadService
{
    /// <summary>
    /// Inicializa la comunicación con el PinPad.
    /// Legacy: fiOpenPort(ConfigFileName)
    /// </summary>
    Task<PinPadResult> InicializarAsync(string? configPath = null, CancellationToken cancellationToken = default);

    /// <summary>
    /// Ejecuta una operación de venta con tarjeta.
    /// Legacy: fiStartOperation(OP_VENTA / OP_VENTA_CASH / OP_VENTA_CUOTAS, timeout, response)
    /// </summary>
    Task<PinPadResult> EjecutarVentaAsync(decimal monto, int cuotas, int timeoutSegundos = 60, CancellationToken cancellationToken = default);

    /// <summary>
    /// Anula una operación de venta previamente procesada.
    /// Legacy: fiStartOperation(OP_ANULACION, timeout, response)
    /// </summary>
    Task<PinPadResult> AnularAsync(string numeroOperacion, int timeoutSegundos = 60, CancellationToken cancellationToken = default);

    /// <summary>
    /// Realiza el cierre de lote del PinPad.
    /// Legacy: fiStartOperation(OP_CIERRE, timeout, response)
    /// </summary>
    Task<PinPadResult> CerrarAsync(int timeoutSegundos = 60, CancellationToken cancellationToken = default);

    /// <summary>
    /// Verifica la comunicación con el PinPad.
    /// Legacy: fiStartOperation(OP_TEST_COMUNICACIONES, timeout, response)
    /// </summary>
    Task<PinPadResult> TestComunicacionAsync(int timeoutSegundos = 30, CancellationToken cancellationToken = default);

    /// <summary>
    /// Simula cuotas para una operación de venta en cuotas.
    /// Legacy: fiStartOperation(OP_SIMULACION_CUOTAS, timeout, response)
    /// </summary>
    Task<PinPadResult> SimularCuotasAsync(decimal monto, int cuotas, int timeoutSegundos = 30, CancellationToken cancellationToken = default);

    /// <summary>
    /// Cierra el puerto del PinPad.
    /// Legacy: fiClosePort()
    /// </summary>
    Task CerrarPuertoAsync(CancellationToken cancellationToken = default);
}

/// <summary>
/// Resultado de una operación de PinPad.
/// Legacy: RET_OK = 0, RET_NOK = -1, RET_RUNNING = 1, ERR_COM = -2, etc. (DLL3500.bas)
/// </summary>
public sealed record PinPadResult(
    int CodigoRetorno,
    string Mensaje,
    bool EsExitoso,
    string? DatosAutorizacion = null,
    string? NumeroOperacion = null)
{
    /// <summary>RET_OK = 0: operación exitosa.</summary>
    public static PinPadResult Ok(string mensaje, string? datosAutorizacion = null, string? numeroOperacion = null)
        => new(0, mensaje, true, datosAutorizacion, numeroOperacion);

    /// <summary>RET_NOK = -1: operación rechazada.</summary>
    public static PinPadResult Rechazado(string mensaje)
        => new(-1, mensaje, false);

    /// <summary>Error de comunicación.</summary>
    public static PinPadResult ErrorComunicacion(int codigo, string mensaje)
        => new(codigo, mensaje, false);
}
