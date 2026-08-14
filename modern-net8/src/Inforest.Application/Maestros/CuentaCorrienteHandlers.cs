using Inforest.Domain.Common;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Application.Maestros;

// ──────────────────────────────────────────────────────────────────────────────
// POS-FUNC-010  Cliente y cuentas corrientes
// Legacy: frmNuevoCliente.frm, frmCtaCte.frm, frmCuentaCobrar.frm
// BR-CLIENTE-001..004, BR-CTACTE-001..003
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Obtiene todas las cuentas corrientes activas (vCompania / TDELIVERY lClienteCtaCte=1).
/// BR-CTACTE-003: Sólo se muestran los registros con lActivo=1 y lClienteCtaCte=1.
/// </summary>
public sealed record ObtenerCuentasCorrientesQuery;

public sealed class ObtenerCuentasCorrientesHandler(ICuentaCorrienteRepository repository)
{
    public async Task<Result<IReadOnlyList<CuentaCorriente>>> HandleAsync(
        ObtenerCuentasCorrientesQuery _,
        CancellationToken ct = default)
    {
        var lista = await repository.ObtenerTodosAsync(ct);
        return Result.Ok(lista);
    }
}

/// <summary>
/// Obtiene una cuenta corriente por código.
/// </summary>
public sealed record ObtenerCuentaCorrientePorCodigoQuery(string Codigo);

public sealed class ObtenerCuentaCorrientePorCodigoHandler(ICuentaCorrienteRepository repository)
{
    public async Task<Result<CuentaCorriente>> HandleAsync(
        ObtenerCuentaCorrientePorCodigoQuery query,
        CancellationToken ct = default)
    {
        var entidad = await repository.ObtenerPorCodigoAsync(query.Codigo, ct);
        return entidad is null
            ? Result.Fail<CuentaCorriente>("No se encontró la cuenta corriente.", "CTACTE_NO_ENCONTRADA")
            : Result.Ok(entidad);
    }
}

/// <summary>
/// Obtiene los documentos pendientes de cobro en un rango de fechas.
/// BR-CTACTE-002: Estado de documento '03' = CtaCte pendiente de cobro.
/// Legacy: frmCuentaCobrar.frm — usa vDocumentoGrilla WHERE tEstadoDocumento='03'.
/// </summary>
public sealed record ObtenerDocumentosPendientesCobroQuery(DateTime FechaInicio, DateTime FechaFin);

public sealed class ObtenerDocumentosPendientesCobroHandler(ICuentaCorrienteRepository repository)
{
    public async Task<Result<IReadOnlyList<DocumentoPendienteCobro>>> HandleAsync(
        ObtenerDocumentosPendientesCobroQuery query,
        CancellationToken ct = default)
    {
        if (query.FechaFin < query.FechaInicio)
            return Result.Fail<IReadOnlyList<DocumentoPendienteCobro>>(
                "La fecha fin no puede ser anterior a la fecha inicio.", "CTACTE_FECHA_INVALIDA");

        var lista = await repository.ObtenerDocumentosPendientesCobroAsync(
            query.FechaInicio, query.FechaFin, ct);
        return Result.Ok(lista);
    }
}

/// <summary>
/// Registra un nuevo cliente en TCLIENTE desde el flujo POS (frmNuevoCliente.frm).
/// BR-CLIENTE-001: El código de cliente debe ser único.
/// BR-CLIENTE-002: La identidad (RUC/DNI) debe tener formato válido según tipo.
/// BR-CLIENTE-003: El campo razón social es obligatorio.
/// </summary>
public sealed record RegistrarNuevoClientePosCommand(Cliente Entidad);

public sealed class RegistrarNuevoClientePosHandler(IClienteRepository repository)
{
    public async Task<Result> HandleAsync(
        RegistrarNuevoClientePosCommand command,
        CancellationToken ct = default)
    {
        var existente = await repository.ObtenerPorCodigoAsync(command.Entidad.CodigoCliente, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un cliente con ese código.", "CLIENTE_YA_EXISTE");

        var ok = await repository.InsertarAsync(command.Entidad, ct);
        return ok
            ? Result.Ok()
            : Result.Fail("No se pudo registrar el cliente.", "CLIENTE_INSERT_FALLIDO");
    }
}

/// <summary>
/// Actualiza datos de un cliente existente desde el flujo POS.
/// BR-CLIENTE-004: Sólo se actualizan datos de un cliente que existe en TCLIENTE.
/// </summary>
public sealed record ActualizarClientePosCommand(Cliente Entidad);

public sealed class ActualizarClientePosHandler(IClienteRepository repository)
{
    public async Task<Result> HandleAsync(
        ActualizarClientePosCommand command,
        CancellationToken ct = default)
    {
        var existente = await repository.ObtenerPorCodigoAsync(command.Entidad.CodigoCliente, ct);
        if (existente is null)
            return Result.Fail("No se encontró el cliente a actualizar.", "CLIENTE_NO_ENCONTRADO");

        var ok = await repository.ActualizarAsync(command.Entidad, ct);
        return ok
            ? Result.Ok()
            : Result.Fail("No se pudo actualizar el cliente.", "CLIENTE_UPDATE_FALLIDO");
    }
}
