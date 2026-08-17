using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Application.Ventas;

/// <summary>
/// Query para obtener motivos de anulación activos.
/// </summary>
public sealed record ObtenerMotivosAnulacionQuery();

/// <summary>
/// Resultado de la consulta de motivos de anulación.
/// </summary>
public sealed record ObtenerMotivosAnulacionResult(IReadOnlyList<MotivoAnulacion> Motivos);

/// <summary>
/// Obtiene la lista de motivos de anulación activos para selección en el POS.
/// Legacy: frmListaMotivos.frm / Form_Load: vMotivoANULACION lActivo=1.
/// BR-MOTIVO-001: motivos obtenidos de vMotivoAnulacion con lActivo = 1.
/// BR-MOTIVO-002: si no hay motivos configurados, retorna falla.
/// </summary>
public sealed class ObtenerMotivosAnulacionHandler(IMotivoAnulacionRepository repository)
{
    /// <summary>
    /// Ejecuta la consulta de motivos activos.
    /// </summary>
    public async Task<Result<ObtenerMotivosAnulacionResult>> HandleAsync(
        ObtenerMotivosAnulacionQuery query,
        CancellationToken ct = default)
    {
        var motivos = await repository.ObtenerActivosAsync(ct);

        // BR-MOTIVO-002: sin motivos configurados la operación no es posible
        if (motivos.Count == 0)
            return Result.Fail<ObtenerMotivosAnulacionResult>(
                "No existe Motivos configuradas para esta caja",
                "MOTIVO_VACIO");

        return Result.Ok(new ObtenerMotivosAnulacionResult(motivos));
    }
}
