using Inforest.Application.Interfaces.Hardware;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Hardware;

namespace Inforest.Application.Caja;

public sealed record ObtenerTerminalesPinPadQuery(string CodigoCaja);

public sealed class ObtenerTerminalesPinPadHandler
{
    private readonly IPinPadTerminalRepository _repository;

    public ObtenerTerminalesPinPadHandler(IPinPadTerminalRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<TerminalPinPad>>> HandleAsync(ObtenerTerminalesPinPadQuery query, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoCaja))
            return Result.Fail<IReadOnlyList<TerminalPinPad>>("La caja es obligatoria para consultar terminales PinPad.", "PINPAD_CAJA_REQUERIDA");

        var terminales = await _repository.ObtenerActivosPorCajaAsync(query.CodigoCaja, ct);
        return Result.Ok(terminales);
    }
}

public sealed record ProcesarPagoPinPadCommand(decimal Monto, int Cuotas = 1, string? CodigoTerminal = null);

public sealed record PagoPinPadProcesado(
    string Referencia,
    string? CodigoTerminal,
    string? NumeroOperacion,
    string? DatosAutorizacion);

public sealed class ProcesarPagoPinPadHandler
{
    private readonly IPinPadService _pinPadService;

    public ProcesarPagoPinPadHandler(IPinPadService pinPadService)
        => _pinPadService = pinPadService;

    public async Task<Result<PagoPinPadProcesado>> HandleAsync(ProcesarPagoPinPadCommand command, CancellationToken ct = default)
    {
        if (command.Monto <= 0)
            return Result.Fail<PagoPinPadProcesado>("El monto a procesar en PinPad debe ser mayor a cero.", "PINPAD_MONTO_INVALIDO");

        if (command.Cuotas <= 0)
            return Result.Fail<PagoPinPadProcesado>("La cantidad de cuotas debe ser mayor o igual a uno.", "PINPAD_CUOTAS_INVALIDAS");

        var inicializacion = await _pinPadService.InicializarAsync(cancellationToken: ct);
        if (!inicializacion.EsExitoso)
            return Result.Fail<PagoPinPadProcesado>(inicializacion.Mensaje, "PINPAD_INICIALIZACION_FALLIDA");

        try
        {
            var resultado = await _pinPadService.EjecutarVentaAsync(command.Monto, command.Cuotas, cancellationToken: ct);
            if (!resultado.EsExitoso)
                return Result.Fail<PagoPinPadProcesado>(resultado.Mensaje, "PINPAD_OPERACION_RECHAZADA");

            var referencia = string.IsNullOrWhiteSpace(resultado.NumeroOperacion)
                ? resultado.DatosAutorizacion
                : resultado.NumeroOperacion;

            if (string.IsNullOrWhiteSpace(referencia))
                return Result.Fail<PagoPinPadProcesado>("El PinPad no devolvió una referencia válida.", "PINPAD_REFERENCIA_INVALIDA");

            return Result.Ok(new PagoPinPadProcesado(
                referencia.Trim(),
                command.CodigoTerminal,
                resultado.NumeroOperacion,
                resultado.DatosAutorizacion));
        }
        finally
        {
            await _pinPadService.CerrarPuertoAsync(ct);
        }
    }
}
