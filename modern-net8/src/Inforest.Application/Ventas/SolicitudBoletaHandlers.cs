using Inforest.Application.Maestros;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Application.Ventas;

/// <summary>
/// Legacy: frmSolicitudBoleta.frm + modPuntoVenta.bas.
/// BR-BOLETA-001/002/003.
/// </summary>
public sealed record ObtenerClienteGeneralBoletaQuery;

/// <summary>
/// Resuelve el cliente general usado cuando la boleta se emite sin datos.
/// Legacy: sCodigoClienteGeneral / CLIENTEGENERAL / TCLIENTE.
/// </summary>
public sealed class ObtenerClienteGeneralBoletaHandler
{
    public const string IdentidadClienteGeneral = "00000000";

    private readonly IClienteRepository _repository;

    public ObtenerClienteGeneralBoletaHandler(IClienteRepository repository)
    {
        _repository = repository;
    }

    public async Task<Result<Cliente>> HandleAsync(ObtenerClienteGeneralBoletaQuery _, CancellationToken ct = default)
    {
        var existente = await _repository.ObtenerPorIdentidadAsync(IdentidadClienteGeneral, ct);
        if (existente is not null)
            return Result.Ok(existente);

        var codigo = await _repository.ObtenerProximoCodigoAsync(ct);
        var cliente = Cliente.Crear(
            codigo,
            "Cliente General",
            IdentidadClienteGeneral,
            "Lima",
            "ADMIN",
            correo: "cliente@cliente.com",
            tipoIdentidad: "01",
            activo: true);

        var creado = await _repository.InsertarAsync(cliente, ct);
        return creado
            ? Result.Ok(cliente)
            : Result.Fail<Cliente>("No se pudo crear el cliente general para boleta sin datos.", "BOLETA_CLIENTE_GENERAL_NO_CREADO");
    }
}
