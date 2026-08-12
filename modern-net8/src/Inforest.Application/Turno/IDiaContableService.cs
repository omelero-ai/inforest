namespace Inforest.Application.Turno;

public interface IDiaContableService
{
    Task<DateTime> ObtenerDiaContableAsync(bool esAutomatico, string horaCierre, string usuario, CancellationToken ct = default);
    Task<DateTime> ObtenerFechaServidorAsync(CancellationToken ct = default);
    Task InsertarDiaContableAsync(DateTime fecha, string usuario, CancellationToken ct = default);
}
