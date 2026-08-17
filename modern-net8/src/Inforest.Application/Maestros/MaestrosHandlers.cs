using Inforest.Domain.Common;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Application.Maestros;

/// <summary>
/// Query para obtener GrupoProducto por código.
/// </summary>
public sealed record ObtenerGrupoProductoPorCodigoQuery(string Codigo);

public sealed class ObtenerGrupoProductoPorCodigoHandler
{
    private readonly IGrupoProductoRepository _repository;

    public ObtenerGrupoProductoPorCodigoHandler(IGrupoProductoRepository repository)
        => _repository = repository;

    public async Task<Result<GrupoProducto>> HandleAsync(ObtenerGrupoProductoPorCodigoQuery query, CancellationToken ct = default)
    {
        var entidad = await _repository.ObtenerPorCodigoAsync(query.Codigo, ct);
        return entidad is null
            ? Result.Fail<GrupoProducto>("No se encontró el registro solicitado.", "GRUPOPRODUCTO_NO_ENCONTRADO")
            : Result.Ok(entidad);
    }
}

/// <summary>
/// Query para obtener grupoproductos activos.
/// </summary>
public sealed record ObtenerGrupoProductosActivosQuery;

public sealed class ObtenerGrupoProductosActivosHandler
{
    private readonly IGrupoProductoRepository _repository;

    public ObtenerGrupoProductosActivosHandler(IGrupoProductoRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<GrupoProducto>>> HandleAsync(ObtenerGrupoProductosActivosQuery query, CancellationToken ct = default)
    {
        var entidades = await _repository.ObtenerTodosAsync(ct);
        return Result.Ok<IReadOnlyList<GrupoProducto>>(entidades.Where(static x => x.Activo).ToList());
    }
}

/// <summary>
/// Comando para crear GrupoProducto.
/// </summary>
public sealed record CrearGrupoProductoCommand(GrupoProducto Entidad);

public sealed class CrearGrupoProductoHandler
{
    private readonly IGrupoProductoRepository _repository;

    public CrearGrupoProductoHandler(IGrupoProductoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(CrearGrupoProductoCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoGrupo, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un registro con el código indicado.", "GRUPOPRODUCTO_YA_EXISTE");

        var creado = await _repository.InsertarAsync(command.Entidad, ct);
        return creado
            ? Result.Ok()
            : Result.Fail("No se pudo insertar el registro.", "GRUPOPRODUCTO_INSERT_FALLO");
    }
}

/// <summary>
/// Comando para actualizar GrupoProducto.
/// </summary>
public sealed record ActualizarGrupoProductoCommand(GrupoProducto Entidad);

public sealed class ActualizarGrupoProductoHandler
{
    private readonly IGrupoProductoRepository _repository;

    public ActualizarGrupoProductoHandler(IGrupoProductoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(ActualizarGrupoProductoCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoGrupo, ct);
        if (existente is null)
            return Result.Fail("No se encontró el registro a actualizar.", "GRUPOPRODUCTO_NO_ENCONTRADO");

        var actualizado = await _repository.ActualizarAsync(command.Entidad, ct);
        return actualizado
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar el registro.", "GRUPOPRODUCTO_UPDATE_FALLO");
    }
}

/// <summary>
/// Query para obtener SubGrupoProducto por código.
/// </summary>
public sealed record ObtenerSubGrupoProductoPorCodigoQuery(string Codigo);

public sealed class ObtenerSubGrupoProductoPorCodigoHandler
{
    private readonly ISubGrupoProductoRepository _repository;

    public ObtenerSubGrupoProductoPorCodigoHandler(ISubGrupoProductoRepository repository)
        => _repository = repository;

    public async Task<Result<SubGrupoProducto>> HandleAsync(ObtenerSubGrupoProductoPorCodigoQuery query, CancellationToken ct = default)
    {
        var entidad = await _repository.ObtenerPorCodigoAsync(query.Codigo, ct);
        return entidad is null
            ? Result.Fail<SubGrupoProducto>("No se encontró el registro solicitado.", "SUBGRUPOPRODUCTO_NO_ENCONTRADO")
            : Result.Ok(entidad);
    }
}

/// <summary>
/// Query para obtener subgrupoproductos activos.
/// </summary>
public sealed record ObtenerSubGrupoProductosActivosQuery;

public sealed class ObtenerSubGrupoProductosActivosHandler
{
    private readonly ISubGrupoProductoRepository _repository;

    public ObtenerSubGrupoProductosActivosHandler(ISubGrupoProductoRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<SubGrupoProducto>>> HandleAsync(ObtenerSubGrupoProductosActivosQuery query, CancellationToken ct = default)
    {
        var entidades = await _repository.ObtenerTodosAsync(ct);
        return Result.Ok<IReadOnlyList<SubGrupoProducto>>(entidades.Where(static x => x.Activo).ToList());
    }
}

/// <summary>
/// Comando para crear SubGrupoProducto.
/// </summary>
public sealed record CrearSubGrupoProductoCommand(SubGrupoProducto Entidad);

public sealed class CrearSubGrupoProductoHandler
{
    private readonly ISubGrupoProductoRepository _repository;

    public CrearSubGrupoProductoHandler(ISubGrupoProductoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(CrearSubGrupoProductoCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoSubGrupo, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un registro con el código indicado.", "SUBGRUPOPRODUCTO_YA_EXISTE");

        var creado = await _repository.InsertarAsync(command.Entidad, ct);
        return creado
            ? Result.Ok()
            : Result.Fail("No se pudo insertar el registro.", "SUBGRUPOPRODUCTO_INSERT_FALLO");
    }
}

/// <summary>
/// Comando para actualizar SubGrupoProducto.
/// </summary>
public sealed record ActualizarSubGrupoProductoCommand(SubGrupoProducto Entidad);

public sealed class ActualizarSubGrupoProductoHandler
{
    private readonly ISubGrupoProductoRepository _repository;

    public ActualizarSubGrupoProductoHandler(ISubGrupoProductoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(ActualizarSubGrupoProductoCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoSubGrupo, ct);
        if (existente is null)
            return Result.Fail("No se encontró el registro a actualizar.", "SUBGRUPOPRODUCTO_NO_ENCONTRADO");

        var actualizado = await _repository.ActualizarAsync(command.Entidad, ct);
        return actualizado
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar el registro.", "SUBGRUPOPRODUCTO_UPDATE_FALLO");
    }
}

/// <summary>
/// Query para obtener Cliente por código.
/// </summary>
public sealed record ObtenerClientePorCodigoQuery(string Codigo);

public sealed class ObtenerClientePorCodigoHandler
{
    private readonly IClienteRepository _repository;

    public ObtenerClientePorCodigoHandler(IClienteRepository repository)
        => _repository = repository;

    public async Task<Result<Cliente>> HandleAsync(ObtenerClientePorCodigoQuery query, CancellationToken ct = default)
    {
        var entidad = await _repository.ObtenerPorCodigoAsync(query.Codigo, ct);
        return entidad is null
            ? Result.Fail<Cliente>("No se encontró el registro solicitado.", "CLIENTE_NO_ENCONTRADO")
            : Result.Ok(entidad);
    }
}

/// <summary>
/// Query para obtener clientes activos.
/// </summary>
public sealed record ObtenerClientesActivosQuery;

public sealed class ObtenerClientesActivosHandler
{
    private readonly IClienteRepository _repository;

    public ObtenerClientesActivosHandler(IClienteRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<Cliente>>> HandleAsync(ObtenerClientesActivosQuery query, CancellationToken ct = default)
    {
        var entidades = await _repository.ObtenerTodosAsync(ct);
        return Result.Ok<IReadOnlyList<Cliente>>(entidades.Where(static x => x.Activo).ToList());
    }
}

/// <summary>
/// Comando para crear Cliente.
/// </summary>
public sealed record CrearClienteCommand(Cliente Entidad);

public sealed class CrearClienteHandler
{
    private readonly IClienteRepository _repository;

    public CrearClienteHandler(IClienteRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(CrearClienteCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoCliente, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un registro con el código indicado.", "CLIENTE_YA_EXISTE");

        var creado = await _repository.InsertarAsync(command.Entidad, ct);
        return creado
            ? Result.Ok()
            : Result.Fail("No se pudo insertar el registro.", "CLIENTE_INSERT_FALLO");
    }
}

/// <summary>
/// Comando para actualizar Cliente.
/// </summary>
public sealed record ActualizarClienteCommand(Cliente Entidad);

public sealed class ActualizarClienteHandler
{
    private readonly IClienteRepository _repository;

    public ActualizarClienteHandler(IClienteRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(ActualizarClienteCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoCliente, ct);
        if (existente is null)
            return Result.Fail("No se encontró el registro a actualizar.", "CLIENTE_NO_ENCONTRADO");

        var actualizado = await _repository.ActualizarAsync(command.Entidad, ct);
        return actualizado
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar el registro.", "CLIENTE_UPDATE_FALLO");
    }
}

/// <summary>
/// Query para obtener Mesa por código.
/// </summary>
public sealed record ObtenerMesaPorCodigoQuery(string Codigo);

public sealed class ObtenerMesaPorCodigoHandler
{
    private readonly IMesaRepository _repository;

    public ObtenerMesaPorCodigoHandler(IMesaRepository repository)
        => _repository = repository;

    public async Task<Result<Mesa>> HandleAsync(ObtenerMesaPorCodigoQuery query, CancellationToken ct = default)
    {
        var entidad = await _repository.ObtenerPorCodigoAsync(query.Codigo, ct);
        return entidad is null
            ? Result.Fail<Mesa>("No se encontró el registro solicitado.", "MESA_NO_ENCONTRADO")
            : Result.Ok(entidad);
    }
}

/// <summary>
/// Query para obtener mesas activas.
/// </summary>
public sealed record ObtenerMesasActivosQuery;

public sealed class ObtenerMesasActivosHandler
{
    private readonly IMesaRepository _repository;

    public ObtenerMesasActivosHandler(IMesaRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<Mesa>>> HandleAsync(ObtenerMesasActivosQuery query, CancellationToken ct = default)
    {
        var entidades = await _repository.ObtenerTodosAsync(ct);
        return Result.Ok<IReadOnlyList<Mesa>>(entidades.Where(static x => x.Activo).ToList());
    }
}

/// <summary>
/// Comando para crear Mesa.
/// </summary>
public sealed record CrearMesaCommand(Mesa Entidad);

public sealed class CrearMesaHandler
{
    private readonly IMesaRepository _repository;

    public CrearMesaHandler(IMesaRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(CrearMesaCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoMesa, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un registro con el código indicado.", "MESA_YA_EXISTE");

        var creado = await _repository.InsertarAsync(command.Entidad, ct);
        return creado
            ? Result.Ok()
            : Result.Fail("No se pudo insertar el registro.", "MESA_INSERT_FALLO");
    }
}

/// <summary>
/// Comando para actualizar Mesa.
/// </summary>
public sealed record ActualizarMesaCommand(Mesa Entidad);

public sealed class ActualizarMesaHandler
{
    private readonly IMesaRepository _repository;

    public ActualizarMesaHandler(IMesaRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(ActualizarMesaCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoMesa, ct);
        if (existente is null)
            return Result.Fail("No se encontró el registro a actualizar.", "MESA_NO_ENCONTRADO");

        var actualizado = await _repository.ActualizarAsync(command.Entidad, ct);
        return actualizado
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar el registro.", "MESA_UPDATE_FALLO");
    }
}

/// <summary>
/// Query para obtener Salon por código.
/// </summary>
public sealed record ObtenerSalonPorCodigoQuery(string Codigo);

public sealed class ObtenerSalonPorCodigoHandler
{
    private readonly ISalonRepository _repository;

    public ObtenerSalonPorCodigoHandler(ISalonRepository repository)
        => _repository = repository;

    public async Task<Result<Salon>> HandleAsync(ObtenerSalonPorCodigoQuery query, CancellationToken ct = default)
    {
        var entidad = await _repository.ObtenerPorCodigoAsync(query.Codigo, ct);
        return entidad is null
            ? Result.Fail<Salon>("No se encontró el registro solicitado.", "SALON_NO_ENCONTRADO")
            : Result.Ok(entidad);
    }
}

/// <summary>
/// Query para obtener salones activos.
/// </summary>
public sealed record ObtenerSalonesActivosQuery;

public sealed class ObtenerSalonesActivosHandler
{
    private readonly ISalonRepository _repository;

    public ObtenerSalonesActivosHandler(ISalonRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<Salon>>> HandleAsync(ObtenerSalonesActivosQuery query, CancellationToken ct = default)
    {
        var entidades = await _repository.ObtenerTodosAsync(ct);
        return Result.Ok<IReadOnlyList<Salon>>(entidades.Where(static x => x.Activo).ToList());
    }
}

/// <summary>
/// Comando para crear Salon.
/// </summary>
public sealed record CrearSalonCommand(Salon Entidad);

public sealed class CrearSalonHandler
{
    private readonly ISalonRepository _repository;

    public CrearSalonHandler(ISalonRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(CrearSalonCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoSalon, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un registro con el código indicado.", "SALON_YA_EXISTE");

        var creado = await _repository.InsertarAsync(command.Entidad, ct);
        return creado
            ? Result.Ok()
            : Result.Fail("No se pudo insertar el registro.", "SALON_INSERT_FALLO");
    }
}

/// <summary>
/// Comando para actualizar Salon.
/// </summary>
public sealed record ActualizarSalonCommand(Salon Entidad);

public sealed class ActualizarSalonHandler
{
    private readonly ISalonRepository _repository;

    public ActualizarSalonHandler(ISalonRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(ActualizarSalonCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoSalon, ct);
        if (existente is null)
            return Result.Fail("No se encontró el registro a actualizar.", "SALON_NO_ENCONTRADO");

        var actualizado = await _repository.ActualizarAsync(command.Entidad, ct);
        return actualizado
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar el registro.", "SALON_UPDATE_FALLO");
    }
}

/// <summary>
/// Query para obtener Mozo por código.
/// </summary>
public sealed record ObtenerMozoPorCodigoQuery(string Codigo);

public sealed class ObtenerMozoPorCodigoHandler
{
    private readonly IMozoRepository _repository;

    public ObtenerMozoPorCodigoHandler(IMozoRepository repository)
        => _repository = repository;

    public async Task<Result<Mozo>> HandleAsync(ObtenerMozoPorCodigoQuery query, CancellationToken ct = default)
    {
        var entidad = await _repository.ObtenerPorCodigoAsync(query.Codigo, ct);
        return entidad is null
            ? Result.Fail<Mozo>("No se encontró el registro solicitado.", "MOZO_NO_ENCONTRADO")
            : Result.Ok(entidad);
    }
}

/// <summary>
/// Query para obtener mozos activos.
/// </summary>
public sealed record ObtenerMozosActivosQuery;

public sealed class ObtenerMozosActivosHandler
{
    private readonly IMozoRepository _repository;

    public ObtenerMozosActivosHandler(IMozoRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<Mozo>>> HandleAsync(ObtenerMozosActivosQuery query, CancellationToken ct = default)
    {
        var entidades = await _repository.ObtenerTodosAsync(ct);
        return Result.Ok<IReadOnlyList<Mozo>>(entidades.Where(static x => x.Activo).ToList());
    }
}

/// <summary>
/// Comando para crear Mozo.
/// </summary>
public sealed record CrearMozoCommand(Mozo Entidad);

public sealed class CrearMozoHandler
{
    private readonly IMozoRepository _repository;

    public CrearMozoHandler(IMozoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(CrearMozoCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoMozo, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un registro con el código indicado.", "MOZO_YA_EXISTE");

        var creado = await _repository.InsertarAsync(command.Entidad, ct);
        return creado
            ? Result.Ok()
            : Result.Fail("No se pudo insertar el registro.", "MOZO_INSERT_FALLO");
    }
}

/// <summary>
/// Comando para actualizar Mozo.
/// </summary>
public sealed record ActualizarMozoCommand(Mozo Entidad);

public sealed class ActualizarMozoHandler
{
    private readonly IMozoRepository _repository;

    public ActualizarMozoHandler(IMozoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(ActualizarMozoCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoMozo, ct);
        if (existente is null)
            return Result.Fail("No se encontró el registro a actualizar.", "MOZO_NO_ENCONTRADO");

        var actualizado = await _repository.ActualizarAsync(command.Entidad, ct);
        return actualizado
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar el registro.", "MOZO_UPDATE_FALLO");
    }
}

/// <summary>
/// Query para obtener ProductoMaestro por código.
/// </summary>
public sealed record ObtenerProductoMaestroPorCodigoQuery(string Codigo);

public sealed class ObtenerProductoMaestroPorCodigoHandler
{
    private readonly IProductoMaestroRepository _repository;

    public ObtenerProductoMaestroPorCodigoHandler(IProductoMaestroRepository repository)
        => _repository = repository;

    public async Task<Result<ProductoMaestro>> HandleAsync(ObtenerProductoMaestroPorCodigoQuery query, CancellationToken ct = default)
    {
        var entidad = await _repository.ObtenerPorCodigoAsync(query.Codigo, ct);
        return entidad is null
            ? Result.Fail<ProductoMaestro>("No se encontró el registro solicitado.", "PRODUCTOMAESTRO_NO_ENCONTRADO")
            : Result.Ok(entidad);
    }
}

/// <summary>
/// Query para obtener productosmaestro activos.
/// </summary>
public sealed record ObtenerProductosMaestroActivosQuery;

public sealed class ObtenerProductosMaestroActivosHandler
{
    private readonly IProductoMaestroRepository _repository;

    public ObtenerProductosMaestroActivosHandler(IProductoMaestroRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<ProductoMaestro>>> HandleAsync(ObtenerProductosMaestroActivosQuery query, CancellationToken ct = default)
    {
        var entidades = await _repository.ObtenerTodosAsync(ct);
        return Result.Ok<IReadOnlyList<ProductoMaestro>>(entidades.Where(static x => x.Activo).ToList());
    }
}

/// <summary>
/// Comando para crear ProductoMaestro.
/// </summary>
public sealed record CrearProductoMaestroCommand(ProductoMaestro Entidad);

public sealed class CrearProductoMaestroHandler
{
    private readonly IProductoMaestroRepository _repository;

    public CrearProductoMaestroHandler(IProductoMaestroRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(CrearProductoMaestroCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoProducto, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un registro con el código indicado.", "PRODUCTOMAESTRO_YA_EXISTE");

        var creado = await _repository.InsertarAsync(command.Entidad, ct);
        return creado
            ? Result.Ok()
            : Result.Fail("No se pudo insertar el registro.", "PRODUCTOMAESTRO_INSERT_FALLO");
    }
}

/// <summary>
/// Comando para actualizar ProductoMaestro.
/// </summary>
public sealed record ActualizarProductoMaestroCommand(ProductoMaestro Entidad);

public sealed class ActualizarProductoMaestroHandler
{
    private readonly IProductoMaestroRepository _repository;

    public ActualizarProductoMaestroHandler(IProductoMaestroRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(ActualizarProductoMaestroCommand command, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorCodigoAsync(command.Entidad.CodigoProducto, ct);
        if (existente is null)
            return Result.Fail("No se encontró el registro a actualizar.", "PRODUCTOMAESTRO_NO_ENCONTRADO");

        var actualizado = await _repository.ActualizarAsync(command.Entidad, ct);
        return actualizado
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar el registro.", "PRODUCTOMAESTRO_UPDATE_FALLO");
    }
}

// ── Mesa — Cambiar Estado ─────────────────────────────────────────────────────

/// <summary>
/// Comando para cambiar el estado de una mesa.
/// Legacy: UPDATE TMESA SET tEstadoMesa='XX' WHERE tCodigoMesa='...' en frmMesaConsulta.frm.
/// BR-MESACONSULTA-001.
/// </summary>
public sealed record CambiarEstadoMesaCommand(string CodigoMesa, EstadoMesa NuevoEstado);

/// <summary>Handler de <see cref="CambiarEstadoMesaCommand"/>.</summary>
public sealed class CambiarEstadoMesaHandler
{
    private readonly IMesaRepository _repository;

    public CambiarEstadoMesaHandler(IMesaRepository repository)
        => _repository = repository;

    /// <summary>
    /// Cambia el estado de una mesa activa.
    /// BR-MESACONSULTA-001: no permite cambiar estado de una mesa Ocupada (EstadoMesa.Ocupada).
    /// </summary>
    public async Task<Result> HandleAsync(CambiarEstadoMesaCommand command, CancellationToken ct = default)
    {
        var mesa = await _repository.ObtenerPorCodigoAsync(command.CodigoMesa, ct);
        if (mesa is null)
            return Result.Fail("Mesa no encontrada.", "MESA_NO_ENCONTRADA");

        // BR-MESACONSULTA-001: no cambiar estado a mesas Ocupadas
        if (mesa.Estado == EstadoMesa.Ocupada)
            return Result.Fail("No se puede cambiar el estado de una mesa ocupada.", "MESA_OCUPADA");

        var ok = await _repository.CambiarEstadoAsync(command.CodigoMesa, command.NuevoEstado, ct);
        return ok
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar el estado de la mesa.", "MESA_ESTADO_UPDATE_FALLO");
    }
}

// ── Junta Mesas ───────────────────────────────────────────────────────────────

/// <summary>Query para obtener las mesas juntas a un pedido.</summary>
public sealed record ObtenerMesasJuntadasQuery(string CodigoPedido);

/// <summary>Handler de <see cref="ObtenerMesasJuntadasQuery"/>.</summary>
public sealed class ObtenerMesasJuntadasHandler
{
    private readonly IJuntaMesaRepository _repository;

    public ObtenerMesasJuntadasHandler(IJuntaMesaRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<string>>> HandleAsync(ObtenerMesasJuntadasQuery query, CancellationToken ct = default)
    {
        var mesas = await _repository.ObtenerMesasJuntadasAsync(query.CodigoPedido, ct);
        return Result.Ok<IReadOnlyList<string>>(mesas);
    }
}

/// <summary>
/// Comando para actualizar las mesas asignadas a un pedido.
/// Legacy: frmJuntaMesa.frm — DELETE/INSERT TPEDIDOMESA + UPDATE TMESA estados.
/// BR-JUNTA-001.
/// </summary>
public sealed record ActualizarJuntaMesasCommand(string CodigoPedido, IEnumerable<string> CodigosMesas);

/// <summary>Handler de <see cref="ActualizarJuntaMesasCommand"/>.</summary>
public sealed class ActualizarJuntaMesasHandler
{
    private readonly IJuntaMesaRepository _repository;

    public ActualizarJuntaMesasHandler(IJuntaMesaRepository repository)
        => _repository = repository;

    /// <summary>
    /// Actualiza la asignación de mesas del pedido.
    /// BR-JUNTA-001: limpia mesas previas antes de insertar las nuevas.
    /// </summary>
    public async Task<Result> HandleAsync(ActualizarJuntaMesasCommand command, CancellationToken ct = default)
    {
        var ok = await _repository.ActualizarJuntaMesasAsync(command.CodigoPedido, command.CodigosMesas, ct);
        return ok
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar la junta de mesas.", "JUNTA_MESA_UPDATE_FALLO");
    }
}
