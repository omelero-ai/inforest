using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Ventas;

public interface IDocumentoRepository
{
    Task<Documento?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default);
    Task<bool> InsertarAsync(Documento doc, IReadOnlyList<DetalleDocumento> detalles, CancellationToken ct = default);
    Task<bool> MarcarCobradoAsync(string codigo, decimal abono, CancellationToken ct = default);
    Task<bool> AnularAsync(string codigo, string usuario, string? motivo = null, CancellationToken ct = default);
}
