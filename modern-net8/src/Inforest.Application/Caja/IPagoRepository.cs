using Inforest.Domain.Entities.Caja;

namespace Inforest.Application.Caja;

public sealed record ResumenTurnoCaja(
    string CodigoTurno,
    decimal TotalVentas,
    decimal TotalPagos,
    decimal TotalPropinas,
    int DocumentosCobrados);

public interface IPagoRepository
{
    Task<bool> InsertarPagoAsync(PagoDocumento pago, CancellationToken ct = default);
    Task<IReadOnlyList<PagoDocumento>> ObtenerPagosPorDocumentoAsync(string codDoc, CancellationToken ct = default);
    Task<ResumenTurnoCaja> ObtenerResumenTurnoAsync(string codigoTurno, CancellationToken ct = default);
}

public interface IMedioPagoRepository
{
    Task<IReadOnlyList<MedioPago>> ObtenerActivosAsync(CancellationToken ct = default);
}
