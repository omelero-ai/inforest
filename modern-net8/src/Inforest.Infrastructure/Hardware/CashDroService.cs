using System.Diagnostics;
using Inforest.Application.Interfaces.Hardware;
using Inforest.Infrastructure.Exceptions;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Inforest.Infrastructure.Hardware;

/// <summary>
/// Adaptador CashDro — lanza MotorCashDrow.exe con Process.Start.
/// Legacy: modProcedimientoNuevo.bas — IniciarMotorCashDrow(Argumentos), ShellExecute a MotorCashDrow.exe.
/// Regla BR-PERIPH-002. ADR-011.
/// </summary>
internal sealed class CashDroService : ICashDroService
{
    private readonly CashDroOptions _options;
    private readonly ILogger<CashDroService> _logger;

    public CashDroService(IOptions<CashDroOptions> options, ILogger<CashDroService> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    public Task IniciarMotorAsync(string argumentos, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("CashDro: iniciando motor con argumentos={Argumentos}", argumentos);

        var rutaMotor = _options.RutaMotor;
        if (!File.Exists(rutaMotor))
        {
            // Legacy: MsgBox "La ruta del motor cashdro no se ha encontrado, favor de informar al administrador"
            var msg = $"La ruta del motor CashDro no se ha encontrado: '{rutaMotor}'. Favor de informar al administrador de sistemas.";
            _logger.LogError("CashDro: {Mensaje}", msg);
            throw new InfrastructureException(msg, "CASHDRO_NOT_FOUND");
        }

        try
        {
            var psi = new ProcessStartInfo(rutaMotor, argumentos)
            {
                UseShellExecute = true,
                WorkingDirectory = Path.GetDirectoryName(rutaMotor) ?? string.Empty
            };
            Process.Start(psi);
            _logger.LogInformation("CashDro: motor iniciado correctamente");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CashDro: error al iniciar motor");
            throw new InfrastructureException($"CashDro: {ex.Message}", "CASHDRO_ERROR");
        }

        return Task.CompletedTask;
    }

    public Task AbrirCajonAsync(CancellationToken cancellationToken = default)
        => IniciarMotorAsync(_options.ArgumentosAbrirCajon, cancellationToken);
}
