using Inforest.Domain.Common;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Application.Maestros;

// ── Queries ──────────────────────────────────────────────────────────────────

/// <summary>
/// Lista todos los insumos/platos de stock crítico.
/// Legacy: USP_LISTARINSUMOS. BR-INSUMO-001.
/// </summary>
public sealed record ListarInsumosQuery(bool ModuloInfoRest = true);

public sealed class ListarInsumosHandler
{
    private readonly IInsumoRepository _repository;

    public ListarInsumosHandler(IInsumoRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<Insumo>>> HandleAsync(
        ListarInsumosQuery query, CancellationToken ct = default)
    {
        var items = await _repository.ObtenerTodosAsync(query.ModuloInfoRest, ct);
        return Result.Ok<IReadOnlyList<Insumo>>(items);
    }
}

/// <summary>
/// Obtiene un insumo por código.
/// </summary>
public sealed record ObtenerInsumoPorCodigoQuery(string Codigo);

public sealed class ObtenerInsumoPorCodigoHandler
{
    private readonly IInsumoRepository _repository;

    public ObtenerInsumoPorCodigoHandler(IInsumoRepository repository)
        => _repository = repository;

    public async Task<Result<Insumo>> HandleAsync(
        ObtenerInsumoPorCodigoQuery query, CancellationToken ct = default)
    {
        var item = await _repository.ObtenerPorCodigoAsync(query.Codigo, ct);
        return item is null
            ? Result.Fail<Insumo>("No se encontró el insumo solicitado.", "INSUMO_NO_ENCONTRADO")
            : Result.Ok(item);
    }
}

// ── Commands ─────────────────────────────────────────────────────────────────

/// <summary>
/// Agrega un nuevo insumo/plato de stock crítico.
/// Legacy: usp_agregarinsumos. BR-INSUMO-001, BR-INSUMO-002.
/// </summary>
public sealed record AgregarInsumoCommand(
    string Descripcion,
    string Usuario,
    double Stock,
    bool EsInsumo,
    string Caja,
    bool ModuloInfoRest = true);

public sealed class AgregarInsumoHandler
{
    private readonly IInsumoRepository _repository;

    public AgregarInsumoHandler(IInsumoRepository repository)
        => _repository = repository;

    public async Task<Result<string>> HandleAsync(
        AgregarInsumoCommand command, CancellationToken ct = default)
    {
        // BR-INSUMO-001: no duplicados por descripción
        if (await _repository.ExisteDescripcionAsync(command.Descripcion.Trim().ToUpperInvariant(), ct))
            return Result.Fail<string>(
                $"Ya existe un insumo/plato con el nombre: {command.Descripcion}",
                "INSUMO_DESCRIPCION_DUPLICADA");

        var codigo = await _repository.ObtenerProximoCodigoAsync(ct);

        // BR-INSUMO-002: descripción en mayúsculas — la entidad normaliza
        var insumo = Insumo.Crear(
            codigo,
            command.Descripcion,
            command.Usuario,
            command.Stock,
            command.EsInsumo);

        await _repository.AgregarAsync(insumo, command.Caja, command.ModuloInfoRest, ct);
        return Result.Ok(codigo);
    }
}

/// <summary>
/// Modifica un insumo/plato existente.
/// Legacy: USP_MODIFICARINSUMOS. BR-INSUMO-003.
/// </summary>
public sealed record ModificarInsumoCommand(
    string Codigo,
    string Descripcion,
    string Usuario,
    double Stock,
    bool EsInsumo,
    bool Activo,
    string Caja,
    bool ModuloInfoRest = true);

public sealed class ModificarInsumoHandler
{
    private readonly IInsumoRepository _repository;

    public ModificarInsumoHandler(IInsumoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(
        ModificarInsumoCommand command, CancellationToken ct = default)
    {
        var insumo = await _repository.ObtenerPorCodigoAsync(command.Codigo, ct);
        if (insumo is null)
            return Result.Fail("No se encontró el insumo a modificar.", "INSUMO_NO_ENCONTRADO");

        insumo.Actualizar(
            command.Descripcion,
            command.Usuario,
            command.Stock,
            command.EsInsumo,
            command.Activo);

        await _repository.ModificarAsync(insumo, command.Caja, command.ModuloInfoRest, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Elimina un insumo.
/// Legacy: USP_ELIMINARINSUMOS. BR-INSUMO-004.
/// </summary>
public sealed record EliminarInsumoCommand(string Codigo);

public sealed class EliminarInsumoHandler
{
    private readonly IInsumoRepository _repository;

    public EliminarInsumoHandler(IInsumoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(
        EliminarInsumoCommand command, CancellationToken ct = default)
    {
        var insumo = await _repository.ObtenerPorCodigoAsync(command.Codigo, ct);
        if (insumo is null)
            return Result.Fail("No se encontró el insumo a eliminar.", "INSUMO_NO_ENCONTRADO");

        await _repository.EliminarAsync(command.Codigo, ct);
        return Result.Ok();
    }
}
