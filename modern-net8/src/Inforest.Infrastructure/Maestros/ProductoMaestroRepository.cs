using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Infrastructure.Maestros;

/// <summary>
/// Repositorio Dapper de productos completos.
/// Legacy: TPRODUCTO / frmProducto.frm. Database: Inforest.
/// </summary>
internal sealed class ProductoMaestroRepository : IProductoMaestroRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public ProductoMaestroRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<IReadOnlyList<ProductoMaestro>> ObtenerTodosAsync(CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoProducto AS CodigoProducto,
                   tGrupo AS Grupo,
                   tSubGrupo AS SubGrupo,
                   tTipoProducto AS TipoProducto,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   tMoneda AS Moneda,
                   lImpuesto1 AS Impuesto1,
                   lImpuesto2 AS Impuesto2,
                   lImpuesto3 AS Impuesto3,
                   lImpuesto4 AS Impuesto4,
                   lImpuesto5 AS Impuesto5,
                   lImpuesto6 AS Impuesto6,
                   lImpuesto7 AS Impuesto7,
                   lImpuesto8 AS Impuesto8,
                   lImpuesto9 AS Impuesto9,
                   lImpuesto10 AS Impuesto10,
                   lImpuesto11 AS Impuesto11,
                   lImpuesto12 AS Impuesto12,
                   lImpuesto13 AS Impuesto13,
                   lImpuesto14 AS Impuesto14,
                   lImpuesto15 AS Impuesto15,
                   tDescargo AS Descargo,
                   nPrecioVenta AS PrecioVenta,
                   nPrecioDelivery AS PrecioDelivery,
                   nPrecioLlevar AS PrecioLlevar,
                   nPrecioCanal4 AS PrecioCanal4,
                   nPrecioCanal5 AS PrecioCanal5,
                   tCortesia AS Cortesia,
                   lModificable AS Modificable,
                   tArea AS Area,
                   lImprimeArea AS ImprimeArea,
                   lEspecial AS Especial,
                   lFijo AS Fijo,
                   lActivo AS Activo,
                   lCombinacion AS Combinacion,
                   nCombinacion AS NumeroCombinacion,
                   nBoton AS NumeroBoton,
                   tIcono AS Icono,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro,
                   tBarra AS Barra,
                   lPropiedad AS Propiedad,
                   lDescuento AS Descuento,
                   lLocal AS Local,
                   lDelivery AS Delivery,
                   lLlevar AS Llevar,
                   lCanal4 AS Canal4,
                   lCanal5 AS Canal5,
                   tUnidadNegocio AS UnidadNegocio,
                   lMultiArea AS MultiArea,
                   tCodigoProductoSunat AS CodigoProductoSunat,
                   nPrecioVentaRefer AS PrecioVentaRefer
            FROM TPRODUCTO
            ORDER BY tDetallado
            """;
        return (await conn.QueryAsync<ProductoMaestro>(sql)).AsList();
    }

    public async Task<ProductoMaestro?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT tCodigoProducto AS CodigoProducto,
                   tGrupo AS Grupo,
                   tSubGrupo AS SubGrupo,
                   tTipoProducto AS TipoProducto,
                   tDetallado AS Detallado,
                   tResumido AS Resumido,
                   tMoneda AS Moneda,
                   lImpuesto1 AS Impuesto1,
                   lImpuesto2 AS Impuesto2,
                   lImpuesto3 AS Impuesto3,
                   lImpuesto4 AS Impuesto4,
                   lImpuesto5 AS Impuesto5,
                   lImpuesto6 AS Impuesto6,
                   lImpuesto7 AS Impuesto7,
                   lImpuesto8 AS Impuesto8,
                   lImpuesto9 AS Impuesto9,
                   lImpuesto10 AS Impuesto10,
                   lImpuesto11 AS Impuesto11,
                   lImpuesto12 AS Impuesto12,
                   lImpuesto13 AS Impuesto13,
                   lImpuesto14 AS Impuesto14,
                   lImpuesto15 AS Impuesto15,
                   tDescargo AS Descargo,
                   nPrecioVenta AS PrecioVenta,
                   nPrecioDelivery AS PrecioDelivery,
                   nPrecioLlevar AS PrecioLlevar,
                   nPrecioCanal4 AS PrecioCanal4,
                   nPrecioCanal5 AS PrecioCanal5,
                   tCortesia AS Cortesia,
                   lModificable AS Modificable,
                   tArea AS Area,
                   lImprimeArea AS ImprimeArea,
                   lEspecial AS Especial,
                   lFijo AS Fijo,
                   lActivo AS Activo,
                   lCombinacion AS Combinacion,
                   nCombinacion AS NumeroCombinacion,
                   nBoton AS NumeroBoton,
                   tIcono AS Icono,
                   tUsuario AS Usuario,
                   fRegistro AS FechaRegistro,
                   tBarra AS Barra,
                   lPropiedad AS Propiedad,
                   lDescuento AS Descuento,
                   lLocal AS Local,
                   lDelivery AS Delivery,
                   lLlevar AS Llevar,
                   lCanal4 AS Canal4,
                   lCanal5 AS Canal5,
                   tUnidadNegocio AS UnidadNegocio,
                   lMultiArea AS MultiArea,
                   tCodigoProductoSunat AS CodigoProductoSunat,
                   nPrecioVentaRefer AS PrecioVentaRefer
            FROM TPRODUCTO
            WHERE tCodigoProducto = @Codigo
            """;
        return await conn.QueryFirstOrDefaultAsync<ProductoMaestro>(sql, new { Codigo = codigo });
    }

    public async Task<bool> InsertarAsync(ProductoMaestro producto, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            INSERT INTO TPRODUCTO (
                tCodigoProducto, tGrupo, tSubGrupo, tTipoProducto, tDetallado, tResumido, tMoneda,
                lImpuesto1, lImpuesto2, lImpuesto3, lImpuesto4, lImpuesto5, lImpuesto6, lImpuesto7, lImpuesto8, lImpuesto9,
                lImpuesto10, lImpuesto11, lImpuesto12, lImpuesto13, lImpuesto14, lImpuesto15,
                tDescargo, nPrecioVenta, nPrecioDelivery, nPrecioLlevar, nPrecioCanal4, nPrecioCanal5,
                tCortesia, lModificable, tArea, lImprimeArea, lEspecial, lFijo, lActivo, lCombinacion, nCombinacion,
                nBoton, tIcono, tUsuario, fRegistro, tBarra, lPropiedad, lDescuento, lLocal, lDelivery, lLlevar,
                lCanal4, lCanal5, tUnidadNegocio, lMultiArea, tCodigoProductoSunat, nPrecioVentaRefer)
            VALUES (
                @CodigoProducto, @Grupo, @SubGrupo, @TipoProducto, @Detallado, @Resumido, @Moneda,
                @Impuesto1, @Impuesto2, @Impuesto3, @Impuesto4, @Impuesto5, @Impuesto6, @Impuesto7, @Impuesto8, @Impuesto9,
                @Impuesto10, @Impuesto11, @Impuesto12, @Impuesto13, @Impuesto14, @Impuesto15,
                @Descargo, @PrecioVenta, @PrecioDelivery, @PrecioLlevar, @PrecioCanal4, @PrecioCanal5,
                @Cortesia, @Modificable, @Area, @ImprimeArea, @Especial, @Fijo, @Activo, @Combinacion, @NumeroCombinacion,
                @NumeroBoton, @Icono, @Usuario, @FechaRegistro, @Barra, @Propiedad, @Descuento, @Local, @Delivery, @Llevar,
                @Canal4, @Canal5, @UnidadNegocio, @MultiArea, @CodigoProductoSunat, @PrecioVentaRefer)
            """;
        return await conn.ExecuteAsync(sql, producto) > 0;
    }

    public async Task<bool> ActualizarAsync(ProductoMaestro producto, CancellationToken ct = default)
    {
        using var conn = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            UPDATE TPRODUCTO
               SET tGrupo = @Grupo,
                   tSubGrupo = @SubGrupo,
                   tTipoProducto = @TipoProducto,
                   tDetallado = @Detallado,
                   tResumido = @Resumido,
                   tMoneda = @Moneda,
                   lImpuesto1 = @Impuesto1,
                   lImpuesto2 = @Impuesto2,
                   lImpuesto3 = @Impuesto3,
                   lImpuesto4 = @Impuesto4,
                   lImpuesto5 = @Impuesto5,
                   lImpuesto6 = @Impuesto6,
                   lImpuesto7 = @Impuesto7,
                   lImpuesto8 = @Impuesto8,
                   lImpuesto9 = @Impuesto9,
                   lImpuesto10 = @Impuesto10,
                   lImpuesto11 = @Impuesto11,
                   lImpuesto12 = @Impuesto12,
                   lImpuesto13 = @Impuesto13,
                   lImpuesto14 = @Impuesto14,
                   lImpuesto15 = @Impuesto15,
                   tDescargo = @Descargo,
                   nPrecioVenta = @PrecioVenta,
                   nPrecioDelivery = @PrecioDelivery,
                   nPrecioLlevar = @PrecioLlevar,
                   nPrecioCanal4 = @PrecioCanal4,
                   nPrecioCanal5 = @PrecioCanal5,
                   tCortesia = @Cortesia,
                   lModificable = @Modificable,
                   tArea = @Area,
                   lImprimeArea = @ImprimeArea,
                   lEspecial = @Especial,
                   lFijo = @Fijo,
                   lActivo = @Activo,
                   lCombinacion = @Combinacion,
                   nCombinacion = @NumeroCombinacion,
                   nBoton = @NumeroBoton,
                   tIcono = @Icono,
                   tUsuario = @Usuario,
                   fRegistro = @FechaRegistro,
                   tBarra = @Barra,
                   lPropiedad = @Propiedad,
                   lDescuento = @Descuento,
                   lLocal = @Local,
                   lDelivery = @Delivery,
                   lLlevar = @Llevar,
                   lCanal4 = @Canal4,
                   lCanal5 = @Canal5,
                   tUnidadNegocio = @UnidadNegocio,
                   lMultiArea = @MultiArea,
                   tCodigoProductoSunat = @CodigoProductoSunat,
                   nPrecioVentaRefer = @PrecioVentaRefer
             WHERE tCodigoProducto = @CodigoProducto
            """;
        return await conn.ExecuteAsync(sql, producto) > 0;
    }
}
