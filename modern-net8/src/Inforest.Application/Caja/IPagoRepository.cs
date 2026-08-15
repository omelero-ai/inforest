using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Entities.Hardware;

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

public interface IPinPadTerminalRepository
{
    Task<IReadOnlyList<TerminalPinPad>> ObtenerActivosPorCajaAsync(string codigoCaja, CancellationToken ct = default);
}
