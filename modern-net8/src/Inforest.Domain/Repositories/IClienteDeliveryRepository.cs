using Inforest.Domain.Entities.Delivery;

namespace Inforest.Domain.Repositories;

/// <summary>
/// Repositorio de clientes delivery.
/// <para>
/// Legacy: tabla <c>TDELIVERY</c>, vistas <c>vDelivery</c>, <c>vDistrito</c>, <c>vZona</c>.
/// Formularios: <c>frmNuevoDelivery.frm</c>, <c>frmClienteDelivery.frm</c>.
/// </para>
/// </summary>
public interface IClienteDeliveryRepository
{
    Task<ClienteDelivery?> ObtenerPorCodigoAsync(string codigoDelivery, CancellationToken cancellationToken = default);
    Task<ClienteDelivery?> ObtenerPorTelefonoAsync(string telefono, CancellationToken cancellationToken = default);
    Task<ClienteDelivery?> ObtenerPorIdentidadAsync(string tipoIdentidad, string numeroIdentidad, CancellationToken cancellationToken = default);
    Task<IEnumerable<ClienteDelivery>> BuscarAsync(string? apellido, string? nombre, string? telefono, string? codigoDistrito, CancellationToken cancellationToken = default);
    Task<IEnumerable<ClienteDelivery>> ObtenerFrecuentesAsync(int top = 20, CancellationToken cancellationToken = default);
    Task InsertarAsync(ClienteDelivery cliente, CancellationToken cancellationToken = default);
    Task ActualizarAsync(ClienteDelivery cliente, CancellationToken cancellationToken = default);
    Task ActualizarFotoAsync(string codigoDelivery, byte[] foto, CancellationToken cancellationToken = default);
}
