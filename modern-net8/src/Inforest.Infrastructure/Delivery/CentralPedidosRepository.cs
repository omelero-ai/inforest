using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;
using Inforest.Infrastructure.Exceptions;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace Inforest.Infrastructure.Delivery;

/// <summary>
/// Repositorio para operaciones con la base <c>CENTRALDELIVERY</c>.
/// <para>
/// Legacy: conexión separada configurada en INI sección <c>[CentralDelivery]</c>
/// (sRutaCD, sMDBCD). Correspondencia con <c>frmCentralPedidos.frm</c>.
/// SPs: <c>sp_Inforest_PedidosCentralPedido</c>, <c>sp_CD_Modificar_EstadoDelivery_Cabecera</c>.
/// </para>
/// Reglas: BR-DEL-003, BR-DEL-008, BR-DEL-009.
/// </summary>
internal sealed class CentralPedidosRepository : ICentralPedidosRepository
{
    private const string CentralDeliveryDb = "CENTRALDELIVERY";

    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public CentralPedidosRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    /// <inheritdoc />
    public async Task<IEnumerable<PedidoCentral>> ObtenerPedidosCentralAsync(
        DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(CentralDeliveryDb, ct);

        // Legacy SP: sp_Inforest_PedidosCentralPedido
        return await _spExecutor.QueryAsync<PedidoCentral>(
            conn,
            "sp_Inforest_PedidosCentralPedido",
            new { fFechaIni = fechaInicio, fFechaFin = fechaFin },
            cancellationToken: ct);
    }

    /// <inheritdoc />
    public async Task ModificarEstadoDeliveryAsync(
        string codigoPedido, string nuevoEstado, string usuario, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync(CentralDeliveryDb, ct);

        // Legacy SP: sp_CD_Modificar_EstadoDelivery_Cabecera (BR-DEL-009)
        await _spExecutor.ExecuteAsync(
            conn,
            "sp_CD_Modificar_EstadoDelivery_Cabecera",
            new { tCodigoPedido = codigoPedido, tEstado = nuevoEstado, tUsuario = usuario },
            cancellationToken: ct);
    }

    /// <inheritdoc />
    public async Task<bool> VerificarConexionAsync(CancellationToken ct = default)
    {
        try
        {
            using var conn = await _connectionFactory.CreateOpenConnectionAsync(CentralDeliveryDb, ct);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
