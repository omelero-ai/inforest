using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Maestro completo de producto operativo.
/// Legacy: TPRODUCTO — precios por canal, 15 flags de impuesto y visibilidad operativa. BR-001, BR-002, BR-014.
/// </summary>
public class ProductoMaestro : Entity
{
    public string CodigoProducto { get; private set; } = string.Empty;
    public string Grupo { get; private set; } = string.Empty;
    public string? SubGrupo { get; private set; }
    public string? TipoProducto { get; private set; }
    public string Detallado { get; private set; } = string.Empty;
    public string Resumido { get; private set; } = string.Empty;
    public string? Moneda { get; private set; }
    public bool Impuesto1 { get; private set; }
    public bool Impuesto2 { get; private set; }
    public bool Impuesto3 { get; private set; }
    public bool Impuesto4 { get; private set; }
    public bool Impuesto5 { get; private set; }
    public bool Impuesto6 { get; private set; }
    public bool Impuesto7 { get; private set; }
    public bool Impuesto8 { get; private set; }
    public bool Impuesto9 { get; private set; }
    public bool Impuesto10 { get; private set; }
    public bool Impuesto11 { get; private set; }
    public bool Impuesto12 { get; private set; }
    public bool Impuesto13 { get; private set; }
    public bool Impuesto14 { get; private set; }
    public bool Impuesto15 { get; private set; }
    public string? Descargo { get; private set; }
    public decimal PrecioVenta { get; private set; }
    public decimal PrecioDelivery { get; private set; }
    public decimal PrecioLlevar { get; private set; }
    public decimal PrecioCanal4 { get; private set; }
    public decimal PrecioCanal5 { get; private set; }
    public bool Cortesia { get; private set; }
    public bool Modificable { get; private set; }
    public string? Area { get; private set; }
    public bool ImprimeArea { get; private set; }
    public bool Especial { get; private set; }
    public bool Fijo { get; private set; }
    public bool Activo { get; private set; }
    public bool Combinacion { get; private set; }
    public int NumeroCombinacion { get; private set; }
    public int NumeroBoton { get; private set; }
    public string? Icono { get; private set; }
    public string Usuario { get; private set; } = string.Empty;
    public DateTime FechaRegistro { get; private set; }
    public string? Barra { get; private set; }
    public bool Propiedad { get; private set; }
    public bool Descuento { get; private set; }
    public bool Local { get; private set; }
    public bool Delivery { get; private set; }
    public bool Llevar { get; private set; }
    public bool Canal4 { get; private set; }
    public bool Canal5 { get; private set; }
    public string? UnidadNegocio { get; private set; }
    public bool MultiArea { get; private set; }
    public string? CodigoProductoSunat { get; private set; }
    public decimal PrecioVentaRefer { get; private set; }

    private ProductoMaestro() { }

    public static ProductoMaestro Crear(
        string codigoProducto,
        string grupo,
        string detallado,
        string usuario,
        decimal precioVenta,
        string? subGrupo = null,
        string? tipoProducto = null,
        string? resumido = null,
        string? moneda = null,
        IEnumerable<bool>? impuestos = null,
        string? descargo = null,
        decimal precioDelivery = 0m,
        decimal precioLlevar = 0m,
        decimal precioCanal4 = 0m,
        decimal precioCanal5 = 0m,
        bool cortesia = false,
        bool modificable = true,
        string? area = null,
        bool imprimeArea = false,
        bool especial = false,
        bool fijo = false,
        bool activo = true,
        bool combinacion = false,
        int numeroCombinacion = 0,
        int numeroBoton = 0,
        string? icono = null,
        string? barra = null,
        bool propiedad = false,
        bool descuento = true,
        bool local = true,
        bool delivery = true,
        bool llevar = true,
        bool canal4 = false,
        bool canal5 = false,
        string? unidadNegocio = null,
        bool multiArea = false,
        string? codigoProductoSunat = null,
        decimal precioVentaRefer = 0m,
        DateTime? fechaRegistro = null)
    {
        Validar(codigoProducto, grupo, detallado, usuario, precioVenta, precioDelivery, precioLlevar, precioCanal4, precioCanal5, precioVentaRefer);
        var producto = new ProductoMaestro
        {
            CodigoProducto = codigoProducto.Trim(),
            Grupo = grupo.Trim(),
            SubGrupo = Limpiar(subGrupo),
            TipoProducto = Limpiar(tipoProducto),
            Detallado = detallado.Trim(),
            Resumido = string.IsNullOrWhiteSpace(resumido) ? detallado.Trim() : resumido.Trim(),
            Moneda = Limpiar(moneda),
            Descargo = Limpiar(descargo),
            PrecioVenta = precioVenta,
            PrecioDelivery = precioDelivery,
            PrecioLlevar = precioLlevar,
            PrecioCanal4 = precioCanal4,
            PrecioCanal5 = precioCanal5,
            Cortesia = cortesia,
            Modificable = modificable,
            Area = Limpiar(area),
            ImprimeArea = imprimeArea,
            Especial = especial,
            Fijo = fijo,
            Activo = activo,
            Combinacion = combinacion,
            NumeroCombinacion = numeroCombinacion,
            NumeroBoton = numeroBoton,
            Icono = Limpiar(icono),
            Usuario = usuario.Trim(),
            FechaRegistro = fechaRegistro ?? DateTime.Now,
            Barra = Limpiar(barra),
            Propiedad = propiedad,
            Descuento = descuento,
            Local = local,
            Delivery = delivery,
            Llevar = llevar,
            Canal4 = canal4,
            Canal5 = canal5,
            UnidadNegocio = Limpiar(unidadNegocio),
            MultiArea = multiArea,
            CodigoProductoSunat = Limpiar(codigoProductoSunat),
            PrecioVentaRefer = precioVentaRefer
        };

        producto.AsignarImpuestos(impuestos);
        return producto;
    }

    public void ActualizarBasico(
        string grupo,
        string? subGrupo,
        string detallado,
        string? resumido,
        decimal precioVenta,
        decimal precioDelivery,
        decimal precioLlevar,
        decimal precioCanal4,
        decimal precioCanal5,
        bool activo,
        string usuario)
    {
        Validar(CodigoProducto, grupo, detallado, usuario, precioVenta, precioDelivery, precioLlevar, precioCanal4, precioCanal5, PrecioVentaRefer);
        Grupo = grupo.Trim();
        SubGrupo = Limpiar(subGrupo);
        Detallado = detallado.Trim();
        Resumido = string.IsNullOrWhiteSpace(resumido) ? detallado.Trim() : resumido.Trim();
        PrecioVenta = precioVenta;
        PrecioDelivery = precioDelivery;
        PrecioLlevar = precioLlevar;
        PrecioCanal4 = precioCanal4;
        PrecioCanal5 = precioCanal5;
        Activo = activo;
        Usuario = usuario.Trim();
    }

    public decimal ObtenerPrecioPorCanal(CanalVenta canalVenta)
        => canalVenta switch
        {
            CanalVenta.Local => PrecioVenta,
            CanalVenta.Delivery => PrecioDelivery > 0 ? PrecioDelivery : PrecioVenta,
            CanalVenta.Llevar => PrecioLlevar > 0 ? PrecioLlevar : PrecioVenta,
            CanalVenta.Canal4 => PrecioCanal4 > 0 ? PrecioCanal4 : PrecioVenta,
            CanalVenta.Canal5 => PrecioCanal5 > 0 ? PrecioCanal5 : PrecioVenta,
            _ => throw new DomainException("Canal de venta inválido para el producto.", "PRODUCTO_CANAL_INVALIDO")
        };

    public bool DisponibleEnCanal(CanalVenta canalVenta)
        => canalVenta switch
        {
            CanalVenta.Local => Local,
            CanalVenta.Delivery => Delivery,
            CanalVenta.Llevar => Llevar,
            CanalVenta.Canal4 => Canal4,
            CanalVenta.Canal5 => Canal5,
            _ => false
        };

    public IReadOnlyList<bool> ObtenerImpuestosActivos() =>
    [
        Impuesto1, Impuesto2, Impuesto3, Impuesto4, Impuesto5,
        Impuesto6, Impuesto7, Impuesto8, Impuesto9, Impuesto10,
        Impuesto11, Impuesto12, Impuesto13, Impuesto14, Impuesto15
    ];

    private void AsignarImpuestos(IEnumerable<bool>? impuestos)
    {
        var valores = (impuestos ?? Enumerable.Repeat(false, 15)).Take(15).ToArray();
        if (valores.Length < 15)
            valores = valores.Concat(Enumerable.Repeat(false, 15 - valores.Length)).ToArray();

        Impuesto1 = valores[0];
        Impuesto2 = valores[1];
        Impuesto3 = valores[2];
        Impuesto4 = valores[3];
        Impuesto5 = valores[4];
        Impuesto6 = valores[5];
        Impuesto7 = valores[6];
        Impuesto8 = valores[7];
        Impuesto9 = valores[8];
        Impuesto10 = valores[9];
        Impuesto11 = valores[10];
        Impuesto12 = valores[11];
        Impuesto13 = valores[12];
        Impuesto14 = valores[13];
        Impuesto15 = valores[14];
    }

    private static void Validar(
        string codigoProducto,
        string grupo,
        string detallado,
        string usuario,
        decimal precioVenta,
        decimal precioDelivery,
        decimal precioLlevar,
        decimal precioCanal4,
        decimal precioCanal5,
        decimal precioVentaRefer)
    {
        if (string.IsNullOrWhiteSpace(codigoProducto))
            throw new DomainException("El código del producto es obligatorio.", "PRODUCTO_MAESTRO_CODIGO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(grupo))
            throw new DomainException("El grupo del producto es obligatorio.", "PRODUCTO_MAESTRO_GRUPO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(detallado))
            throw new DomainException("La descripción detallada del producto es obligatoria.", "PRODUCTO_MAESTRO_DETALLADO_REQUERIDO");

        if (string.IsNullOrWhiteSpace(usuario))
            throw new DomainException("El usuario del producto es obligatorio.", "PRODUCTO_MAESTRO_USUARIO_REQUERIDO");

        if (new[] { precioVenta, precioDelivery, precioLlevar, precioCanal4, precioCanal5, precioVentaRefer }.Any(p => p < 0))
            throw new DomainException("Los precios del producto no pueden ser negativos.", "PRODUCTO_MAESTRO_PRECIO_INVALIDO");
    }

    private static string? Limpiar(string? valor)
        => string.IsNullOrWhiteSpace(valor) ? null : valor.Trim();
}
