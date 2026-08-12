using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Entities.Cocina;

/// <summary>
/// Configuración de producción/cocina para comandas y KDS.
/// Legacy: TPARAMETRO.tOrderInfo/tOrderInfox/tBump + lKDS/lKDSInforest/lKDSxArea/lKDSTodos.
/// Reglas BR-010 y BR-011.
/// </summary>
public sealed class ConfiguracionProduccionCocina
{
    public bool KdsHabilitado { get; }
    public bool KdsInforestHabilitado { get; }
    public bool KdsPorAreaHabilitado { get; }
    public bool KdsTodosHabilitado { get; }
    public string? DirectorioOrderInfoPrimario { get; }
    public string? DirectorioOrderInfoSecundario { get; }
    public string? DirectorioBumpPrimario { get; }
    public string? DirectorioBumpSecundario { get; }

    public bool TieneModeloSecundario => !string.IsNullOrWhiteSpace(DirectorioOrderInfoSecundario);

    public bool UsaKdsXmlExterno => KdsHabilitado && !KdsInforestHabilitado;

    private ConfiguracionProduccionCocina(
        bool kdsHabilitado,
        bool kdsInforestHabilitado,
        bool kdsPorAreaHabilitado,
        bool kdsTodosHabilitado,
        string? directorioOrderInfoPrimario,
        string? directorioOrderInfoSecundario,
        string? directorioBumpPrimario,
        string? directorioBumpSecundario)
    {
        KdsHabilitado = kdsHabilitado;
        KdsInforestHabilitado = kdsInforestHabilitado;
        KdsPorAreaHabilitado = kdsPorAreaHabilitado;
        KdsTodosHabilitado = kdsTodosHabilitado;
        DirectorioOrderInfoPrimario = directorioOrderInfoPrimario;
        DirectorioOrderInfoSecundario = directorioOrderInfoSecundario;
        DirectorioBumpPrimario = directorioBumpPrimario;
        DirectorioBumpSecundario = directorioBumpSecundario;
    }

    public static ConfiguracionProduccionCocina Crear(
        bool kdsHabilitado,
        bool kdsInforestHabilitado,
        bool kdsPorAreaHabilitado,
        bool kdsTodosHabilitado,
        string? directorioOrderInfoPrimario,
        string? directorioOrderInfoSecundario,
        string? directorioBumpPrimario,
        string? directorioBumpSecundario)
    {
        if (kdsPorAreaHabilitado && kdsTodosHabilitado)
            throw new DomainException(
                "La configuración KDS no puede habilitar simultáneamente KDS por área y KDS todos.",
                "COCINA_KDS_CONFIG_INVALIDA");

        return new ConfiguracionProduccionCocina(
            kdsHabilitado,
            kdsInforestHabilitado,
            kdsPorAreaHabilitado,
            kdsTodosHabilitado,
            Limpiar(directorioOrderInfoPrimario),
            Limpiar(directorioOrderInfoSecundario),
            Limpiar(directorioBumpPrimario),
            Limpiar(directorioBumpSecundario));
    }

    private static string? Limpiar(string? valor)
        => string.IsNullOrWhiteSpace(valor) ? null : valor.Trim();
}
