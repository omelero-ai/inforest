using Inforest.Domain.Common;
using Inforest.Domain.Entities.Configuracion;
using Inforest.Application.Interfaces;

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

/// <summary>
/// Legacy: modPuntoVenta.bas valida TPARAMETRO y TCAJA antes de abrir el POS. BR-POS-002/BR-POS-005/BR-POS-006
/// </summary>
public sealed record ValidarInicioPosQuery(string CodigoCaja);

/// <summary>
/// Resultado de validación de arranque POS.
/// </summary>
public sealed record InicioPosValidado(string CodigoCaja, bool RequiereLogin, ConfiguracionCaja ConfiguracionCaja);

/// <summary>Handler de <see cref="ValidarInicioPosQuery"/>.</summary>
public sealed class ValidarInicioPosHandler
{
    private readonly IParametroService _parametroService;
    private readonly IParametroRepository _repository;

    public ValidarInicioPosHandler(IParametroService parametroService, IParametroRepository repository)
    {
        _parametroService = parametroService;
        _repository = repository;
    }

    public async Task<Result<InicioPosValidado>> HandleAsync(ValidarInicioPosQuery query, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoCaja))
            return Result.Fail<InicioPosValidado>("Debe indicar el código de caja para iniciar POS.", "POS_CODIGO_CAJA_REQUERIDO");

        if (!await _parametroService.TieneConfiguracionAsync(ct))
            return Result.Fail<InicioPosValidado>("No se encontró configuración global en TPARAMETRO.", "CONFIGURACION_SISTEMA_NO_ENCONTRADA");

        var caja = await _repository.ObtenerConfiguracionCajaAsync(query.CodigoCaja, ct);
        if (caja is null)
            return Result.Fail<InicioPosValidado>("No se encontró la configuración de la caja solicitada.", "CONFIGURACION_CAJA_NO_ENCONTRADA");

        return Result.Ok(new InicioPosValidado(caja.tCaja!, RequiereLogin: !caja.lMCPV, caja));
    }
}
