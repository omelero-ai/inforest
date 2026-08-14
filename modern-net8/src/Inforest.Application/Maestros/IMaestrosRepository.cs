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
