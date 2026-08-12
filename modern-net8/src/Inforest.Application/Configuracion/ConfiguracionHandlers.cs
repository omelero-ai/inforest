using Inforest.Domain.Common;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Application.Configuracion;

/// <summary>
/// Legacy: modPuntoVenta.bas carga TPARAMETRO al iniciar. BR-SQL-003
/// </summary>
public sealed record ObtenerConfiguracionSistemaQuery;

/// <summary>Handler de <see cref="ObtenerConfiguracionSistemaQuery"/>.</summary>
public sealed class ObtenerConfiguracionSistemaHandler
{
    private readonly IParametroRepository _repository;

    public ObtenerConfiguracionSistemaHandler(IParametroRepository repository)
        => _repository = repository;

    public async Task<Result<ConfiguracionSistema?>> HandleAsync(ObtenerConfiguracionSistemaQuery _, CancellationToken ct = default)
        => Result.Ok(await _repository.ObtenerConfiguracionAsync(ct));
}

/// <summary>
/// Legacy: modPuntoVenta.bas y módulos operativos cargan TCAJA según la caja activa. BR-SQL-004
/// </summary>
public sealed record ObtenerConfiguracionCajaQuery(string CodigoCaja);

/// <summary>Handler de <see cref="ObtenerConfiguracionCajaQuery"/>.</summary>
public sealed class ObtenerConfiguracionCajaHandler
{
    private readonly IParametroRepository _repository;

    public ObtenerConfiguracionCajaHandler(IParametroRepository repository)
        => _repository = repository;

    public async Task<Result<ConfiguracionCaja?>> HandleAsync(ObtenerConfiguracionCajaQuery query, CancellationToken ct = default)
    {
        var configuracion = await _repository.ObtenerConfiguracionCajaAsync(query.CodigoCaja, ct);
        return configuracion is null
            ? Result.Fail<ConfiguracionCaja?>("No se encontró la configuración de la caja solicitada.", "CONFIGURACION_CAJA_NO_ENCONTRADA")
            : Result.Ok<ConfiguracionCaja?>(configuracion);
    }
}

/// <summary>
/// Legacy: TCAJA define perfiles operativos por terminal/caja. BR-SQL-004
/// </summary>
public sealed record ObtenerTodasCajasQuery;

/// <summary>Handler de <see cref="ObtenerTodasCajasQuery"/>.</summary>
public sealed class ObtenerTodasCajasHandler
{
    private readonly IParametroRepository _repository;

    public ObtenerTodasCajasHandler(IParametroRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<ConfiguracionCaja>>> HandleAsync(ObtenerTodasCajasQuery _, CancellationToken ct = default)
        => Result.Ok(await _repository.ObtenerTodasCajasAsync(ct));
}
