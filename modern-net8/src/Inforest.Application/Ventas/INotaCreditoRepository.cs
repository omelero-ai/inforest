using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Ventas;

/// <summary>
/// Repositorio de notas de crédito.
/// Legacy: MNOTACREDITO.
/// </summary>
public interface INotaCreditoRepository
{
    Task<NotaCredito?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default);
    Task<IReadOnlyList<NotaCredito>> ObtenerPorRangoFechaAsync(DateTime desde, DateTime hasta, CancellationToken ct = default);
    Task<decimal> ObtenerTotalEmitidoParaDocumentoAsync(string codigoDocumento, CancellationToken ct = default);
    Task<bool> InsertarAsync(NotaCredito notaCredito, CancellationToken ct = default);
    Task<bool> AnularAsync(string codigo, string usuario, CancellationToken ct = default);
}
