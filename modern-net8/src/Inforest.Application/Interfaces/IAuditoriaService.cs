using Inforest.Domain.Common;

namespace Inforest.Application.Interfaces;

/// <summary>
/// Registra auditoría de acceso y movimientos en INFSEGURIDAD.
/// Legacy: modAuditoriaIntegral.bas + MMOVIMIENTO/MMOVIMIENTOACCESO.
/// </summary>
public interface IAuditoriaService
{
    Task<Result<long>> RegistrarIngresoAsync(RegistroAccesoAuditoriaRequest request, CancellationToken cancellationToken = default);

    Task<Result> RegistrarSalidaAsync(RegistroAccesoAuditoriaRequest request, CancellationToken cancellationToken = default);

    Task<Result> RegistrarMovimientoAsync(MovimientoAuditoriaRequest request, CancellationToken cancellationToken = default);
}

public sealed record RegistroAccesoAuditoriaRequest(
    string Tipo,
    string BaseDatos,
    string Modulo,
    string Usuario,
    long CorrelativoAcceso);

public sealed record MovimientoAuditoriaRequest(
    string Modulo,
    string BaseDatos,
    string Tabla,
    string Accion,
    string Usuario,
    string IdRegistroUnico,
    string IdRegistroDetalle,
    string Campo,
    string ValorAnterior,
    string ValorActual);
