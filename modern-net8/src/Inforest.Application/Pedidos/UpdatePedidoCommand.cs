using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Pedidos;

/// <summary>
/// Legacy: spUpd_MPEDIDO. BR-SQL-002.
/// </summary>
public record UpdatePedidoCommand(
    string CodigoPedido,
    string? NuevoEstado,
    string? Observacion,
    string UsuarioModifica
);

public class UpdatePedidoHandler
{
    private readonly IPedidoRepository _pedidoRepository;

    public UpdatePedidoHandler(IPedidoRepository pedidoRepository)
        => _pedidoRepository = pedidoRepository;

    public async Task<Result> HandleAsync(UpdatePedidoCommand command, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(command.CodigoPedido))
            return Result.Fail("El código del pedido es obligatorio.", "PEDIDO_CODIGO_REQUERIDO");

        var pedido = await _pedidoRepository.ObtenerPorCodigoAsync(command.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail("No se encontró el pedido solicitado.", "PEDIDO_NO_ENCONTRADO");

        PedidoPrivateAccessor.SetObservacion(pedido, command.Observacion);

        if (!string.IsNullOrWhiteSpace(command.NuevoEstado))
            PedidoPrivateAccessor.SetEstado(pedido, LegacyMappings.MapEstado(command.NuevoEstado));

        await _pedidoRepository.ActualizarAsync(pedido, ct);
        return Result.Ok();
    }
}
