using System.Data;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Infrastructure.Ventas;

/// <summary>
/// Repositorio Dapper para correlativos de documento.
/// Legacy: TTIPODOCUMENTOIMPRESORA. BR-POS-008.
/// </summary>
internal sealed class CorrelativoRepository : ICorrelativoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public CorrelativoRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<CorrelativoDocumento>> ObtenerPorCajaAsync(
        string codigoCaja,
        CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                tCaja             AS CodigoCaja,
                tImpresora        AS CodigoImpresora,
                tTipoEmision      AS TipoEmision,
                tDescripcion      AS Descripcion,
                tFormulario       AS Formulario,
                tSerie            AS Serie,
                tUltimoNumero     AS UltimoNumero,
                tNumeroAutorizacion AS NumeroAutorizacion,
                fInicio           AS FechaInicio,
                fCaducidad        AS FechaCaducidad,
                ISNULL(lResumen,  0) AS Resumen,
                ISNULL(lImpuesto1,0) AS Impuesto1,
                ISNULL(lImpuesto2,0) AS Impuesto2,
                ISNULL(lImpuesto3,0) AS Impuesto3,
                ISNULL(lEquivaDolares,0) AS EquivaleDolares,
                ISNULL(lImprimeImageCab,0) AS ImprimeImageCabecera,
                ISNULL(lImprimeImagePie,0) AS ImprimeImagePie,
                ISNULL(lFacturacionElectronica,0) AS FacturacionElectronica,
                tUsuario          AS UsuarioRegistro,
                fRegistro         AS FechaRegistro
            FROM TTIPODOCUMENTOIMPRESORA
            WHERE tCaja = @CodigoCaja
            ORDER BY tTipoEmision
            """;

        var rows = await connection.QueryAsync<CorrelativoRow>(
            new CommandDefinition(sql, new { CodigoCaja = codigoCaja }, cancellationToken: ct));

        return rows.Select(MapRow).ToList().AsReadOnly();
    }

    public async Task<CorrelativoDocumento?> ObtenerPorCajaYTipoAsync(
        string codigoCaja,
        string tipoEmision,
        CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT TOP 1
                tCaja             AS CodigoCaja,
                tImpresora        AS CodigoImpresora,
                tTipoEmision      AS TipoEmision,
                tDescripcion      AS Descripcion,
                tFormulario       AS Formulario,
                tSerie            AS Serie,
                tUltimoNumero     AS UltimoNumero,
                tNumeroAutorizacion AS NumeroAutorizacion,
                fInicio           AS FechaInicio,
                fCaducidad        AS FechaCaducidad,
                ISNULL(lResumen,  0) AS Resumen,
                ISNULL(lImpuesto1,0) AS Impuesto1,
                ISNULL(lImpuesto2,0) AS Impuesto2,
                ISNULL(lImpuesto3,0) AS Impuesto3,
                ISNULL(lEquivaDolares,0) AS EquivaleDolares,
                ISNULL(lImprimeImageCab,0) AS ImprimeImageCabecera,
                ISNULL(lImprimeImagePie,0) AS ImprimeImagePie,
                ISNULL(lFacturacionElectronica,0) AS FacturacionElectronica,
                tUsuario          AS UsuarioRegistro,
                fRegistro         AS FechaRegistro
            FROM TTIPODOCUMENTOIMPRESORA
            WHERE tCaja = @CodigoCaja AND tTipoEmision = @TipoEmision
            """;

        var row = await connection.QueryFirstOrDefaultAsync<CorrelativoRow>(
            new CommandDefinition(sql, new { CodigoCaja = codigoCaja, TipoEmision = tipoEmision }, cancellationToken: ct));

        return row is null ? null : MapRow(row);
    }

    public async Task<IReadOnlyList<CorrelativoDocumento>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync(ct);
        const string sql = """
            SELECT
                tCaja             AS CodigoCaja,
                tImpresora        AS CodigoImpresora,
                tTipoEmision      AS TipoEmision,
                tDescripcion      AS Descripcion,
                tFormulario       AS Formulario,
                tSerie            AS Serie,
                tUltimoNumero     AS UltimoNumero,
                tNumeroAutorizacion AS NumeroAutorizacion,
                fInicio           AS FechaInicio,
                fCaducidad        AS FechaCaducidad,
                ISNULL(lResumen,  0) AS Resumen,
                ISNULL(lImpuesto1,0) AS Impuesto1,
                ISNULL(lImpuesto2,0) AS Impuesto2,
                ISNULL(lImpuesto3,0) AS Impuesto3,
                ISNULL(lEquivaDolares,0) AS EquivaleDolares,
                ISNULL(lImprimeImageCab,0) AS ImprimeImageCabecera,
                ISNULL(lImprimeImagePie,0) AS ImprimeImagePie,
                ISNULL(lFacturacionElectronica,0) AS FacturacionElectronica,
                tUsuario          AS UsuarioRegistro,
                fRegistro         AS FechaRegistro
            FROM TTIPODOCUMENTOIMPRESORA
            ORDER BY tCaja, tTipoEmision
            """;

        var rows = await connection.QueryAsync<CorrelativoRow>(
            new CommandDefinition(sql, cancellationToken: ct));

        return rows.Select(MapRow).ToList().AsReadOnly();
    }

    private static CorrelativoDocumento MapRow(CorrelativoRow r) =>
        CorrelativoDocumento.Crear(
            r.CodigoCaja,
            r.CodigoImpresora,
            r.TipoEmision,
            r.Descripcion,
            r.Formulario,
            r.Serie,
            r.UltimoNumero,
            r.NumeroAutorizacion,
            r.FechaInicio,
            r.FechaCaducidad,
            r.Resumen,
            r.Impuesto1,
            r.Impuesto2,
            r.Impuesto3,
            r.EquivaleDolares,
            r.ImprimeImageCabecera,
            r.ImprimeImagePie,
            r.FacturacionElectronica,
            r.UsuarioRegistro,
            r.FechaRegistro);

    private sealed record CorrelativoRow(
        string CodigoCaja,
        string CodigoImpresora,
        string TipoEmision,
        string? Descripcion,
        string? Formulario,
        string? Serie,
        string? UltimoNumero,
        string? NumeroAutorizacion,
        DateTime? FechaInicio,
        DateTime? FechaCaducidad,
        bool Resumen,
        bool Impuesto1,
        bool Impuesto2,
        bool Impuesto3,
        bool EquivaleDolares,
        bool ImprimeImageCabecera,
        bool ImprimeImagePie,
        bool FacturacionElectronica,
        string? UsuarioRegistro,
        DateTime? FechaRegistro);
}
