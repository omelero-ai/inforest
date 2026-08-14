using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Application.Configuracion;

/// <summary>
/// Legacy: TPARAMETRO/TCAJA cargados al iniciar desde modPuntoVenta.bas. BR-SQL-003/BR-SQL-004
/// </summary>
public interface IParametroRepository
{
    Task<ConfiguracionSistema?> ObtenerConfiguracionAsync(CancellationToken ct = default);
    Task<IReadOnlyList<ConfiguracionCaja>> ObtenerTodasCajasAsync(CancellationToken ct = default);
    Task<ConfiguracionCaja?> ObtenerConfiguracionCajaAsync(string codigoCaja, CancellationToken ct = default);
}
