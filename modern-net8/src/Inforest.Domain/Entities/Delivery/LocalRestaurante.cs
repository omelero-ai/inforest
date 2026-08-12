using Inforest.Domain.Common;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Delivery;

/// <summary>
/// Local (sucursal) del restaurante para operación multi-local.
/// <para>
/// Legacy: tabla <c>TLOCAL</c> — controla IP de conexión, base de datos INFOREST por local
/// y flag de réplica para sincronización con CENTRALDELIVERY.
/// </para>
/// Regla BR-DEL-007.
/// </summary>
public class LocalRestaurante : Entity
{
    public string CodigoLocal { get; private set; } = string.Empty;
    public string? Detallado { get; private set; }
    public string? Resumido { get; private set; }
    public string? CodigoSector { get; private set; }
    public string? DireccionIp { get; private set; }

    /// <summary>Nombre de la base de datos INFOREST del local. Legacy: <c>tBaseDatosINF</c>.</summary>
    public string? BaseDatosInf { get; private set; }

    /// <summary>Nombre de la base de datos ALMACEN del local. Legacy: <c>tBaseDatosALM</c>.</summary>
    public string? BaseDatosAlm { get; private set; }

    public decimal Meta { get; private set; }
    public bool Activo { get; private set; }

    /// <summary>
    /// Si el local participa en réplica multi-local.
    /// Legacy: <c>lReplica</c> en <c>TLOCAL</c>.
    /// Regla BR-DEL-007.
    /// </summary>
    public bool ParticipanEnReplica { get; private set; }

    // Dapper constructor
    private LocalRestaurante() { }

    public static LocalRestaurante Crear(
        string codigoLocal,
        string? detallado,
        string? direccionIp,
        string? baseDatosInf,
        bool activo,
        bool participaEnReplica)
    {
        if (string.IsNullOrWhiteSpace(codigoLocal))
            throw new DomainException("El código de local es obligatorio.", "LOCAL_CODIGO_REQUERIDO");

        return new LocalRestaurante
        {
            CodigoLocal = codigoLocal,
            Detallado = detallado,
            DireccionIp = direccionIp,
            BaseDatosInf = baseDatosInf,
            Activo = activo,
            ParticipanEnReplica = participaEnReplica
        };
    }
}
