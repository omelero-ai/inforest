using System.Runtime.InteropServices;
using Inforest.Application.Interfaces.Hardware;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Adaptador PinPad via P/Invoke a caja_pinpad.dll.
/// Legacy: DLL3500.bas — fiOpenPort, fiClosePort, fiStartOperation, fiGetStatus.
/// Regla BR-PERIPH-001. ADR-009.
/// </summary>
internal sealed class PinPadService : IPinPadService
{
    // Códigos de retorno Legacy (DLL3500.bas)
    private const int RET_OK = 0;
    private const int RET_NOK = -1;
    private const int RET_RUNNING = 1;
    private const int ERR_COM = -2;
    private const int ERR_PPAD_NO_RESP = -3;
    private const int ERR_SOCKET = -5;
    private const int ERR_HOST_NO_RESP = -6;

    // Operaciones Legacy (DLL3500.bas)
    private const string OP_VENTA = "01";
    private const string OP_VENTA_CASH = "02";
    private const string OP_VENTA_CUOTAS = "03";
    private const string OP_ANULACION = "04";
    private const string OP_CIERRE = "07";
    private const string OP_SIMULACION_CUOTAS = "10";
    private const string OP_TEST_COMUNICACIONES = "11";

    private readonly PinPadOptions _options;
    private readonly ILogger<PinPadService> _logger;

    public PinPadService(IOptions<PinPadOptions> options, ILogger<PinPadService> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    // P/Invoke declarations — Legacy: DLL3500.bas Declare Function ...
    [DllImport("caja_pinpad.dll", CharSet = CharSet.Ansi)]
    private static extern int fiOpenPort([MarshalAs(UnmanagedType.LPStr)] string configFileName);

    [DllImport("caja_pinpad.dll")]
    private static extern int fiClosePort();

    [DllImport("caja_pinpad.dll", CharSet = CharSet.Ansi)]
    private static extern int fiStartOperation(
        [MarshalAs(UnmanagedType.LPStr)] string pucTipoOperation,
        int iTimeOut,
        [MarshalAs(UnmanagedType.LPStr)] ref string pucResponse);

    [DllImport("caja_pinpad.dll", CharSet = CharSet.Ansi)]
    private static extern int fiGetStatus(
        [MarshalAs(UnmanagedType.LPStr)] ref string sBuffer,
        int sizeBuffer);

    public Task<PinPadResult> InicializarAsync(string configPath, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("PinPad: inicializando con config {ConfigPath}", configPath);
        try
        {
            var ret = fiOpenPort(configPath);
            var result = InterpretarRetorno(ret, "Inicialización");
            _logger.LogInformation("PinPad: inicialización retorno={Retorno}", ret);
            return Task.FromResult(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "PinPad: error al inicializar");
            return Task.FromResult(PinPadResult.ErrorComunicacion(ERR_COM, ex.Message));
        }
    }

    public Task<PinPadResult> EjecutarVentaAsync(decimal monto, int cuotas, int timeoutSegundos = 60, CancellationToken cancellationToken = default)
    {
        var operacion = cuotas > 1 ? OP_VENTA_CUOTAS : OP_VENTA;
        return EjecutarOperacionAsync(operacion, timeoutSegundos, $"Venta monto={monto} cuotas={cuotas}");
    }

    public Task<PinPadResult> AnularAsync(string numeroOperacion, int timeoutSegundos = 60, CancellationToken cancellationToken = default)
        => EjecutarOperacionAsync(OP_ANULACION, timeoutSegundos, $"Anulación op={numeroOperacion}");

    public Task<PinPadResult> CerrarAsync(int timeoutSegundos = 60, CancellationToken cancellationToken = default)
        => EjecutarOperacionAsync(OP_CIERRE, timeoutSegundos, "Cierre de lote");

    public Task<PinPadResult> TestComunicacionAsync(int timeoutSegundos = 30, CancellationToken cancellationToken = default)
        => EjecutarOperacionAsync(OP_TEST_COMUNICACIONES, timeoutSegundos, "Test de comunicación");

    public Task<PinPadResult> SimularCuotasAsync(decimal monto, int cuotas, int timeoutSegundos = 30, CancellationToken cancellationToken = default)
        => EjecutarOperacionAsync(OP_SIMULACION_CUOTAS, timeoutSegundos, $"Simulación cuotas monto={monto} cuotas={cuotas}");

    public Task CerrarPuertoAsync(CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("PinPad: cerrando puerto");
        try { fiClosePort(); }
        catch (Exception ex) { _logger.LogWarning(ex, "PinPad: error al cerrar puerto"); }
        return Task.CompletedTask;
    }

    private Task<PinPadResult> EjecutarOperacionAsync(string operacion, int timeoutSegundos, string descripcion)
    {
        _logger.LogInformation("PinPad: ejecutando {Descripcion}", descripcion);
        try
        {
            var respuesta = string.Empty;
            var ret = fiStartOperation(operacion, timeoutSegundos, ref respuesta);
            var result = InterpretarRetorno(ret, respuesta);
            _logger.LogInformation("PinPad: {Descripcion} retorno={Retorno}", descripcion, ret);
            return Task.FromResult(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "PinPad: error en {Descripcion}", descripcion);
            return Task.FromResult(PinPadResult.ErrorComunicacion(ERR_COM, ex.Message));
        }
    }

    private static PinPadResult InterpretarRetorno(int codigo, string mensaje) => codigo switch
    {
        RET_OK => PinPadResult.Ok(string.IsNullOrEmpty(mensaje) ? "OK" : mensaje),
        RET_NOK => PinPadResult.Rechazado(string.IsNullOrEmpty(mensaje) ? "Operación rechazada" : mensaje),
        ERR_COM => PinPadResult.ErrorComunicacion(ERR_COM, "Error de comunicación"),
        ERR_PPAD_NO_RESP => PinPadResult.ErrorComunicacion(ERR_PPAD_NO_RESP, "PinPad no responde"),
        ERR_SOCKET => PinPadResult.ErrorComunicacion(ERR_SOCKET, "Error de socket"),
        ERR_HOST_NO_RESP => PinPadResult.ErrorComunicacion(ERR_HOST_NO_RESP, "Host no responde"),
        _ => PinPadResult.ErrorComunicacion(codigo, $"Error desconocido: {codigo}")
    };
}
