using Inforest.Application.Interfaces;
using Microsoft.Extensions.Configuration;

namespace Inforest.Infrastructure.Delivery;

/// <summary>
/// Adaptador de integración Rappi.
/// <para>
/// Legacy: <c>frmOrdenesConsola.frm</c> y <c>frmEnvioDatosRappi.frm</c>
/// en <c>Despachador.exe</c>.
/// Habilitado por flag <c>lOrdenesRappi</c> en <c>TCAJA</c>.
/// </para>
/// Regla BR-DEL-004.
/// <para>
/// NOTA: La integración concreta con la API de Rappi es un gap conocido
/// (ver docs/migration/known-gaps.md — GAP-DEL-004).
/// Esta implementación provee el contrato y la estructura base.
/// </para>
/// </summary>
internal sealed class RappiOrderAdapter : IRappiOrderService
{
    private readonly IConfiguration _configuration;

    public RappiOrderAdapter(IConfiguration configuration)
        => _configuration = configuration;

    /// <inheritdoc />
    /// <remarks>
    /// La habilitación se lee de la configuración. En runtime se debe consultar
    /// el flag <c>lOrdenesRappi</c> de <c>TCAJA</c> (BR-DEL-004).
    /// </remarks>
    public bool EstaHabilitado
        => _configuration.GetValue<bool>("Rappi:Habilitado", false);

    /// <inheritdoc />
    public Task<IEnumerable<OrdenExterna>> ObtenerOrdenesPendientesAsync(CancellationToken cancellationToken = default)
    {
        // GAP-DEL-004: integración real con API Rappi pendiente de implementación.
        // Retorna lista vacía hasta que se implemente el conector HTTP.
        return Task.FromResult<IEnumerable<OrdenExterna>>([]);
    }

    /// <inheritdoc />
    public Task<bool> EnviarOrdenAsync(string codigoPedido, CancellationToken cancellationToken = default)
    {
        // GAP-DEL-004: integración real con API Rappi pendiente de implementación.
        return Task.FromResult(false);
    }
}
