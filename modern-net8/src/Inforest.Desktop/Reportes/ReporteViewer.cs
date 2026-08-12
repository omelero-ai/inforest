using System.Data;
using Microsoft.Extensions.Logging;

namespace Inforest.Desktop.Reportes;

/// <summary>
/// Envoltorio centralizado sobre FastReport .NET para renderizar reportes del sistema.
/// <para>
/// ADR-007: FastReport .NET es el motor objetivo de reportes (reemplaza Crystal Reports).
/// Convención de plantillas: <c>Reports/Templates/Rep{Nombre}.frx</c>
/// </para>
/// Uso típico:
/// <code>
/// var viewer = new ReporteViewer(logger, "Reports/Templates");
/// viewer.CargarPlantilla("RepPropina.frx");
/// viewer.AgregarDataSource("Propinas", dataTable);
/// viewer.Mostrar();
/// </code>
/// </summary>
public sealed class ReporteViewer : IDisposable
{
    private readonly ILogger<ReporteViewer> _logger;
    private readonly string _rutaPlantillas;
    private bool _disposed;

    // FastReport.Report cargado en tiempo de ejecución para evitar dependencia
    // en tiempo de compilación cuando FastReport no está instalado (ej. CI).
    private dynamic? _report;
    private bool _reportCargado;

    /// <summary>
    /// Inicializa el viewer con la ruta base de plantillas FastReport.
    /// </summary>
    /// <param name="logger">Logger de diagnóstico.</param>
    /// <param name="rutaPlantillas">Ruta absoluta o relativa al directorio de plantillas <c>.frx</c>.</param>
    public ReporteViewer(ILogger<ReporteViewer> logger, string rutaPlantillas)
    {
        _logger = logger;
        _rutaPlantillas = ResolverRutaPlantillas(rutaPlantillas);
    }

    /// <summary>
    /// Carga una plantilla FastReport <c>.frx</c> desde el directorio de plantillas configurado.
    /// </summary>
    /// <param name="nombrePlantilla">Nombre del archivo, ej. <c>RepPropina.frx</c>.</param>
    /// <exception cref="FileNotFoundException">Si la plantilla no existe.</exception>
    public void CargarPlantilla(string nombrePlantilla)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var rutaCompleta = ResolverRutaPlantilla(nombrePlantilla);
        if (!File.Exists(rutaCompleta))
        {
            _logger.LogWarning("Plantilla FastReport no encontrada: {Ruta}", rutaCompleta);
            throw new FileNotFoundException($"Plantilla de reporte no encontrada: {rutaCompleta}");
        }

        _report = CrearReporte();
        _report!.Load(rutaCompleta);
        _reportCargado = true;
        _logger.LogDebug("Plantilla cargada: {NombrePlantilla}", nombrePlantilla);
    }

    /// <summary>
    /// Agrega un <see cref="DataTable"/> como fuente de datos con nombre.
    /// Equivale a registrar el dataset en el reporte antes de procesar.
    /// </summary>
    /// <param name="nombreDataSource">Nombre del datasource dentro de la plantilla <c>.frx</c>.</param>
    /// <param name="datos">Tabla de datos.</param>
    public void AgregarDataSource(string nombreDataSource, DataTable datos)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        EnsureReportCargado();
        _report!.RegisterData(datos, nombreDataSource);
        if (!string.Equals(nombreDataSource, "Result", StringComparison.OrdinalIgnoreCase))
            _report!.RegisterData(datos, "Result");

        _logger.LogDebug("DataSource '{Nombre}' registrado ({Filas} filas)", nombreDataSource, datos.Rows.Count);
    }

    /// <summary>
    /// Agrega una lista de objetos como fuente de datos con nombre.
    /// </summary>
    /// <typeparam name="T">Tipo de los elementos.</typeparam>
    /// <param name="nombreDataSource">Nombre del datasource.</param>
    /// <param name="datos">Lista de datos.</param>
    public void AgregarDataSource<T>(string nombreDataSource, IEnumerable<T> datos)
    {
        var dt = ConvertirADataTable(datos);
        AgregarDataSource(nombreDataSource, dt);
    }

    /// <summary>
    /// Muestra el reporte en el visor FastReport.
    /// </summary>
    public void Mostrar()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        EnsureReportCargado();
        try
        {
            _report!.Show();
            _logger.LogInformation("Reporte mostrado correctamente");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al mostrar reporte");
            throw;
        }
    }

    /// <summary>
    /// Envía el reporte a la impresora predeterminada.
    /// </summary>
    public void Imprimir()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        EnsureReportCargado();
        try
        {
            _report!.Print();
            _logger.LogInformation("Reporte enviado a impresora");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al imprimir reporte");
            throw;
        }
    }

    /// <summary>
    /// Exporta el reporte a PDF en la ruta indicada.
    /// </summary>
    /// <param name="rutaDestino">Ruta completa del archivo PDF de destino.</param>
    public void ExportarPdf(string rutaDestino)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        EnsureReportCargado();
        try
        {
            // Exportar a PDF vía FastReport PDFExport
            var exportType = Type.GetType("FastReport.Export.Pdf.PDFExport, FastReport");
            if (exportType is null)
            {
                _logger.LogWarning("PDFExport de FastReport no disponible");
                throw new InvalidOperationException("El módulo de exportación PDF de FastReport no está instalado.");
            }

            dynamic pdfExport = Activator.CreateInstance(exportType)!;
            _report!.Export(pdfExport, rutaDestino);
            _logger.LogInformation("Reporte exportado a PDF: {Ruta}", rutaDestino);
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al exportar reporte a PDF");
            throw;
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (!_disposed)
        {
            (_report as IDisposable)?.Dispose();
            _report = null;
            _disposed = true;
        }
    }

    // ── Helpers privados ────────────────────────────────────────────────────

    private void EnsureReportCargado()
    {
        if (!_reportCargado || _report is null)
            throw new InvalidOperationException("Debe llamar a CargarPlantilla() antes de usar el reporte.");
    }


    private string ResolverRutaPlantilla(string nombrePlantilla)
    {
        var rutaDirecta = Path.Combine(_rutaPlantillas, nombrePlantilla);
        if (File.Exists(rutaDirecta))
            return rutaDirecta;

        var rutaCanonica = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reports", "Templates", nombrePlantilla);
        return rutaCanonica;
    }

    private static string ResolverRutaPlantillas(string rutaPlantillas)
    {
        var rutaCanonica = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reports", "Templates");
        if (string.IsNullOrWhiteSpace(rutaPlantillas))
            return rutaCanonica;

        if (Path.IsPathRooted(rutaPlantillas))
            return rutaPlantillas;

        var normalizada = rutaPlantillas.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
        var relativa = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, normalizada);
        return Directory.Exists(relativa) ? relativa : rutaCanonica;
    }

    /// <summary>
    /// Crea una instancia de <c>FastReport.Report</c> mediante reflexión para
    /// evitar dependencia en tiempo de compilación cuando el paquete no está disponible.
    /// </summary>
    private static dynamic? CrearReporte()
    {
        var reportType = Type.GetType("FastReport.Report, FastReport");
        if (reportType is null)
            throw new InvalidOperationException(
                "FastReport.Report no está disponible. Instale el paquete NuGet FastReport.Net.");
        return Activator.CreateInstance(reportType);
    }

    /// <summary>
    /// Convierte una colección tipada a <see cref="DataTable"/> mediante reflexión de propiedades.
    /// </summary>
    private static DataTable ConvertirADataTable<T>(IEnumerable<T> datos)
    {
        var dt = new DataTable(typeof(T).Name);
        var props = typeof(T).GetProperties();
        foreach (var p in props)
            dt.Columns.Add(p.Name, Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType);

        foreach (var item in datos)
        {
            var row = dt.NewRow();
            foreach (var p in props)
                row[p.Name] = p.GetValue(item) ?? DBNull.Value;
            dt.Rows.Add(row);
        }
        return dt;
    }
}
