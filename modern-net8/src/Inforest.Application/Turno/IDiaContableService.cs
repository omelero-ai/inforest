namespace Inforest.Application.Turno;

public interface IDiaContableService
{
    Task<DateTime> ObtenerDiaContableAsync(bool esAutomatico, string horaCierre, string usuario, CancellationToken ct = default);
    Task<DateTime> ObtenerFechaServidorAsync(CancellationToken ct = default);
    Task InsertarDiaContableAsync(DateTime fecha, string usuario, CancellationToken ct = default);

    /// <summary>
    /// Cierra el día contable registrado para la fecha indicada.
    /// Legacy: frmDiaContable.frm cmdCerrar_Click → UPDATE tdiacontable SET lcierre=1, tusuariocierre, fregistrocierre.
    /// </summary>
    Task CerrarDiaContableAsync(DateTime fecha, string usuario, CancellationToken ct = default);

    /// <summary>
    /// Obtiene la fecha máxima registrada en TDIACONTABLE para validar apertura.
    /// Legacy: clsDiaContable.validaFechaIngreso → SELECT max(fdiacontable) FROM tdiacontable.
    /// </summary>
    Task<DateTime?> ObtenerFechaMaximaDiaContableAsync(CancellationToken ct = default);
}
