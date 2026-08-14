using System.Threading;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Pedidos;

/// <summary>
/// Contexto temporal para transportar metadatos Legacy que no existen aún en la entidad <see cref="Pedido"/>.
/// </summary>
public sealed record PedidoRepositoryMetadata(
    string CodigoCaja,
    string CodigoTurno,
    string CodigoMesa,
    string CodigoMozo,
    string CodigoUsuario,
    CanalVenta Canal,
    int NumeroAdultos,
    int NumeroNinos,
    string? Observacion,
    DateTime FechaDiaContable);

public static class PedidoRepositoryContext
{
    private static readonly AsyncLocal<PedidoRepositoryMetadata?> CurrentHolder = new();

    public static PedidoRepositoryMetadata? Current => CurrentHolder.Value;

    public static void Set(PedidoRepositoryMetadata metadata)
        => CurrentHolder.Value = metadata;

    public static void Clear()
        => CurrentHolder.Value = null;
}
