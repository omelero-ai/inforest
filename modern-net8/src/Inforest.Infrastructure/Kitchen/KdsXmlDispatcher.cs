using System.Xml.Linq;
using Inforest.Domain.Entities.Cocina;
using Inforest.Domain.Exceptions;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Kitchen;

/// <summary>
/// Adaptador KDS por XML/directorio.
/// Legacy: modKDS.bas + DPEDIDOKDS.
/// Regla BR-010.
/// </summary>
internal sealed class KdsXmlDispatcher : IKdsDispatcher
{
    private readonly IKdsLegacyGateway _legacyGateway;
    private readonly ILogger<KdsXmlDispatcher> _logger;

    public KdsXmlDispatcher(
        IKdsLegacyGateway legacyGateway,
        ILogger<KdsXmlDispatcher> logger)
    {
        _legacyGateway = legacyGateway;
        _logger = logger;
    }

    public async Task EnviarNuevaOrdenAsync(
        OrdenProduccion orden,
        PlanProduccion plan,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default)
    {
        foreach (var grupo in ObtenerDestinosKds(plan, configuracion))
        {
            var document = BuildNuevaOrden(orden, grupo.Items);
            await GuardarDocumentoAsync(document, grupo.Directorio, $"{orden.ObtenerIdNumerico()}.xml", cancellationToken);
        }
    }

    public async Task EliminarOrdenAsync(
        OrdenProduccion orden,
        PlanProduccion plan,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default)
    {
        foreach (var grupo in ObtenerDestinosKds(plan, configuracion, fallbackATodosLosDirectorios: true))
        {
            var document = BuildEliminarOrden(orden);
            await GuardarDocumentoAsync(document, grupo.Directorio, $"{orden.ObtenerIdNumerico()}.xml", cancellationToken);
        }
    }

    public async Task EliminarProductoAsync(
        OrdenProduccion orden,
        ItemProduccion item,
        PlanProduccion plan,
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default)
    {
        foreach (var grupo in ObtenerDestinosKds(plan, configuracion, item))
        {
            var document = BuildEliminarProducto(orden, item);
            await GuardarDocumentoAsync(document, grupo.Directorio, $"{orden.ObtenerIdNumerico()}.xml", cancellationToken);
            await GuardarDocumentoAsync(
                BuildMensajeEstacion(item.ObtenerEstacionKds(), $"({orden.ObtenerIdNumerico()})ELIMINADO:{item.NombreSalida}"),
                grupo.Directorio,
                "message.xml",
                cancellationToken);
        }
    }

    public async Task<int> ProcesarBumpNotificationsAsync(
        ConfiguracionProduccionCocina configuracion,
        CancellationToken cancellationToken = default)
    {
        var directorios = new[]
        {
            configuracion.DirectorioBumpPrimario,
            configuracion.DirectorioBumpSecundario
        }
        .Where(path => !string.IsNullOrWhiteSpace(path))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToArray();

        var procesados = 0;

        foreach (var directorio in directorios)
        {
            if (string.IsNullOrWhiteSpace(directorio) || !Directory.Exists(directorio))
                continue;

            var historial = Path.Combine(directorio, "Historial");
            Directory.CreateDirectory(historial);

            foreach (var archivo in Directory.EnumerateFiles(directorio, "*.xml"))
            {
                var xml = await File.ReadAllTextAsync(archivo, cancellationToken);
                var document = XDocument.Parse(xml);
                var orderNode = document.Root?.Element("Order");
                var codigoPedido = orderNode?.Element("ID")?.Value;
                var item = orderNode?.Elements("Item").FirstOrDefault()?.Element("ID")?.Value;

                if (!string.IsNullOrWhiteSpace(codigoPedido) && !string.IsNullOrWhiteSpace(item))
                {
                    await _legacyGateway.RegistrarTiempoSalidaAsync(
                        codigoPedido,
                        item,
                        File.GetCreationTimeUtc(archivo),
                        cancellationToken);
                    procesados++;
                }

                var destino = Path.Combine(historial, Path.GetFileName(archivo));
                if (File.Exists(destino))
                    File.Delete(destino);

                File.Move(archivo, destino);
            }
        }

        return procesados;
    }

    private static XDocument BuildNuevaOrden(OrdenProduccion orden, IReadOnlyList<ItemProduccion> items)
    {
        return new XDocument(
            new XElement("Transaction",
                new XElement("Order",
                    new XElement("ID", orden.ObtenerIdNumerico()),
                    new XElement("PosTerminal", ExtraerNumero(orden.CodigoCaja)),
                    new XElement("TransType", "1"),
                    new XElement("OrderStatus", "1"),
                    new XElement("OrderType", string.Empty),
                    new XElement("ServerName", orden.NombreServidor ?? string.Empty),
                    new XElement("Destination", orden.Destino ?? string.Empty),
                    new XElement("GuestTable", orden.ObtenerNombreMesa()),
                    new XElement("UserInfo"),
                    items.Select(BuildNuevoItem))));
    }

    private static XDocument BuildEliminarOrden(OrdenProduccion orden)
    {
        return new XDocument(
            new XElement("Transaction",
                new XElement("Order",
                    new XElement("ID", orden.ObtenerIdNumerico()),
                    new XElement("PosTerminal", ExtraerNumero(orden.CodigoCaja)),
                    new XElement("TransType", "2"),
                    new XElement("ServerName", orden.UsuarioAccion ?? orden.NombreServidor ?? string.Empty),
                    new XElement("Destination", orden.Destino ?? string.Empty),
                    new XElement("GuestTable"),
                    new XElement("UserInfo"))));
    }

    private static XDocument BuildEliminarProducto(OrdenProduccion orden, ItemProduccion item)
    {
        return new XDocument(
            new XElement("Transaction",
                new XElement("Order",
                    new XElement("ID", orden.ObtenerIdNumerico()),
                    new XElement("PosTerminal", ExtraerNumero(orden.CodigoCaja)),
                    new XElement("TransType", "3"),
                    new XElement("OrderStatus", "1"),
                    new XElement("OrderType", string.Empty),
                    new XElement("ServerName", orden.UsuarioAccion ?? orden.NombreServidor ?? string.Empty),
                    new XElement("Destination", orden.Destino ?? string.Empty),
                    new XElement("GuestTable"),
                    new XElement("UserInfo"),
                    new XElement("Item",
                        new XElement("ID", ObtenerIdItem(item)),
                        new XElement("TransType", "2")))));
    }

    private static XElement BuildNuevoItem(ItemProduccion item)
    {
        return new XElement("Item",
            new XElement("ID", ObtenerIdItem(item)),
            new XElement("TransType", "1"),
            new XElement("Name", item.NombreSalida),
            new XElement("Category", "Monitor1"),
            new XElement("Quantity", item.Cantidad),
            new XElement("Color"),
            new XElement("KDSStation", item.ObtenerEstacionKds()),
            item.Condimentos.Select(condimento =>
                new XElement("Condiment",
                    new XElement("ID", ExtraerNumero(condimento.Codigo)),
                    new XElement("TransType", "1"),
                    new XElement("Name", condimento.Nombre),
                    new XElement("Color"),
                    new XElement("Action"))));
    }

    private static XDocument BuildMensajeEstacion(string stationId, string info)
    {
        return new XDocument(
            new XElement("StationInfo",
                new XElement("StationID", ExtraerNumero(stationId)),
                new XElement("User", "0"),
                new XElement("Info", info)));
    }

    private static string ObtenerIdItem(ItemProduccion item)
    {
        var baseItem = ExtraerNumero(item.Item);
        if (string.IsNullOrWhiteSpace(item.ItemCombo))
            return baseItem;

        var combo = ExtraerNumero(item.ItemCombo);
        return $"{baseItem}{combo.PadLeft(2, '0')}";
    }

    private static string ExtraerNumero(string? valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
            return "0";

        var digits = new string(valor.Where(char.IsDigit).ToArray());
        return string.IsNullOrWhiteSpace(digits) ? "0" : digits;
    }

    private async Task GuardarDocumentoAsync(
        XDocument document,
        string directorio,
        string archivo,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(directorio))
            throw new DomainException("No existe directorio configurado para la salida KDS.", "COCINA_KDS_PATH_REQUERIDO");

        Directory.CreateDirectory(directorio);

        var ruta = Path.Combine(directorio, archivo);
        await using var stream = File.Create(ruta);
        await document.SaveAsync(stream, SaveOptions.None, cancellationToken);
        _logger.LogInformation("Archivo KDS generado en {Ruta}", ruta);
    }

    private static IReadOnlyList<KdsDestino> ObtenerDestinosKds(
        PlanProduccion plan,
        ConfiguracionProduccionCocina configuracion,
        ItemProduccion? itemFiltrado = null,
        bool fallbackATodosLosDirectorios = false)
    {
        if (!configuracion.UsaKdsXmlExterno)
            return [];

        var canales = itemFiltrado is null
            ? plan.CanalesKds
            : plan.CanalesKds.Where(canal => canal.Items.Contains(itemFiltrado)).ToArray();

        if (canales.Count == 0)
        {
            return fallbackATodosLosDirectorios
                ? ObtenerTodosLosDirectorios(configuracion)
                : [];
        }

        return canales
            .Select(canal => new KdsDestino(
                canal.Canal,
                ResolverDirectorio(configuracion, canal.Canal),
                canal.Items))
            .ToArray();
    }

    private static IReadOnlyList<KdsDestino> ObtenerTodosLosDirectorios(ConfiguracionProduccionCocina configuracion)
    {
        var destinos = new List<KdsDestino>();

        if (!string.IsNullOrWhiteSpace(configuracion.DirectorioOrderInfoPrimario))
            destinos.Add(new KdsDestino(0, configuracion.DirectorioOrderInfoPrimario!, []));

        if (!string.IsNullOrWhiteSpace(configuracion.DirectorioOrderInfoSecundario))
            destinos.Add(new KdsDestino(1, configuracion.DirectorioOrderInfoSecundario!, []));

        return destinos;
    }

    private static string ResolverDirectorio(ConfiguracionProduccionCocina configuracion, int canal)
    {
        if (canal == 1 && !string.IsNullOrWhiteSpace(configuracion.DirectorioOrderInfoSecundario))
            return configuracion.DirectorioOrderInfoSecundario!;

        return configuracion.DirectorioOrderInfoPrimario
            ?? throw new DomainException("No existe directorio KDS primario configurado.", "COCINA_KDS_PATH_REQUERIDO");
    }

    private sealed record KdsDestino(int Canal, string Directorio, IReadOnlyList<ItemProduccion> Items);
}
