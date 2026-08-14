using Inforest.Domain.Entities.Maestros;

namespace Inforest.Application.Maestros;

/// <summary>
/// Contrato base para repositorios de maestros operativos.
/// Legacy: TPRODUCTO, TGRUPOPRODUCTO, TSUBGRUPOPRODUCTO, TCLIENTE, TMESA, TSALON, TMOZO.
/// </summary>
public interface IMaestroRepository<T>
{
    Task<IReadOnlyList<T>> ObtenerTodosAsync(CancellationToken ct = default);
    Task<T?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default);
    Task<bool> InsertarAsync(T entidad, CancellationToken ct = default);
    Task<bool> ActualizarAsync(T entidad, CancellationToken ct = default);
}

public interface IGrupoProductoRepository : IMaestroRepository<GrupoProducto>;
public interface ISubGrupoProductoRepository : IMaestroRepository<SubGrupoProducto>;
public interface IClienteRepository : IMaestroRepository<Cliente>;
public interface IMesaRepository : IMaestroRepository<Mesa>;
public interface ISalonRepository : IMaestroRepository<Salon>;
public interface IMozoRepository : IMaestroRepository<Mozo>;
public interface IProductoMaestroRepository : IMaestroRepository<ProductoMaestro>;

/// <summary>
/// Contrato del repositorio de insumos / platos de stock crítico (TINSUMO).
/// Legacy: frmInsumo.frm, frmInsumoDetalle.frm. BR-INSUMO-001..004.
/// SPs: USP_LISTARINSUMOS, usp_agregarinsumos, USP_MODIFICARINSUMOS, USP_ELIMINARINSUMOS.
/// </summary>
public interface IInsumoRepository
{
    Task<IReadOnlyList<Insumo>> ObtenerTodosAsync(bool moduloInforRest = true, CancellationToken ct = default);
    Task<Insumo?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default);
    Task<bool> ExisteDescripcionAsync(string descripcion, CancellationToken ct = default);
    Task<string> ObtenerProximoCodigoAsync(CancellationToken ct = default);
    Task AgregarAsync(Insumo insumo, string caja, bool moduloInforRest = true, CancellationToken ct = default);
    Task ModificarAsync(Insumo insumo, string caja, bool moduloInforRest = true, CancellationToken ct = default);
    Task EliminarAsync(string codigo, CancellationToken ct = default);
}

/// <summary>
/// Contrato de cuentas corrientes (vCompania / TDELIVERY lClienteCtaCte=1).
/// Legacy: frmCtaCte.frm, frmCuentaCobrar.frm. BR-CTACTE-001/002/003.
/// </summary>
public interface ICuentaCorrienteRepository
{
    Task<IReadOnlyList<CuentaCorriente>> ObtenerTodosAsync(CancellationToken ct = default);
    Task<CuentaCorriente?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default);
    Task<IReadOnlyList<DocumentoPendienteCobro>> ObtenerDocumentosPendientesCobroAsync(
        DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default);
}
