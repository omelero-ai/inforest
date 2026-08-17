using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio de junta de mesas (TPEDIDOMESA).
/// Legacy: frmJuntaMesa.frm — gestiona mesas asignadas a un pedido.
/// BR-JUNTA-001, BR-JUNTA-002.
/// </summary>
public sealed class JuntaMesaRepository : IJuntaMesaRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public JuntaMesaRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    /// <summary>
    /// Obtiene los códigos de mesas asignadas a un pedido.
    /// Legacy: SELECT tMesa FROM TPEDIDOMESA WHERE tCodigoPedido='...'
    /// </summary>
    public async Task<IReadOnlyList<string>> ObtenerMesasJuntadasAsync(
        string codigoPedido,
        CancellationToken ct = default)
    {
        using var cn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = "SELECT tMesa FROM TPEDIDOMESA WHERE tCodigoPedido = @CodigoPedido";
        var mesas = await cn.QueryAsync<string>(sql, new { CodigoPedido = codigoPedido });
        return mesas.ToList().AsReadOnly();
    }

    /// <summary>
    /// Reemplaza las mesas asignadas a un pedido en una transacción.
    /// Legacy:
    ///   UPDATE TMESA SET tEstadoMesa='01' WHERE tCodigoMesa IN (SELECT tMesa FROM TPEDIDOMESA WHERE tCodigoPedido='...')
    ///   DELETE FROM TPEDIDOMESA WHERE tCodigoPedido='...'
    ///   INSERT INTO TPEDIDOMESA SELECT '...', '...' -- para cada mesa
    ///   UPDATE TMESA SET tEstadoMesa='06' WHERE tCodigoMesa='...' -- para cada mesa
    /// BR-JUNTA-001.
    /// </summary>
    public async Task<bool> ActualizarJuntaMesasAsync(
        string codigoPedido,
        IEnumerable<string> codigosMesas,
        CancellationToken ct = default)
    {
        using var cn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        using var tx = cn.BeginTransaction();
        try
        {
            // Liberar mesas previamente juntas (estado '01' = Lista/Libre)
            await cn.ExecuteAsync(
                "UPDATE TMESA SET tEstadoMesa='01' WHERE tCodigoMesa IN (SELECT tMesa FROM TPEDIDOMESA WHERE tCodigoPedido=@CodigoPedido)",
                new { CodigoPedido = codigoPedido },
                transaction: tx);

            // Eliminar junta anterior
            await cn.ExecuteAsync(
                "DELETE FROM TPEDIDOMESA WHERE tCodigoPedido=@CodigoPedido",
                new { CodigoPedido = codigoPedido },
                transaction: tx);

            // Insertar nuevas mesas y marcar estado '06' (FueraDeServicio = usada en junta)
            foreach (var codigoMesa in codigosMesas)
            {
                await cn.ExecuteAsync(
                    "INSERT INTO TPEDIDOMESA (tCodigoPedido, tMesa) VALUES (@CodigoPedido, @CodigoMesa)",
                    new { CodigoPedido = codigoPedido, CodigoMesa = codigoMesa },
                    transaction: tx);

                await cn.ExecuteAsync(
                    "UPDATE TMESA SET tEstadoMesa='06' WHERE tCodigoMesa=@CodigoMesa",
                    new { CodigoMesa = codigoMesa },
                    transaction: tx);
            }

            tx.Commit();
            return true;
        }
        catch
        {
            tx.Rollback();
            return false;
        }
    }
}
