using Inforest.Application.Reportes;

namespace Inforest.Desktop.Consultas;

/// <summary>
/// Legacy: <c>mdiConsulta.frm</c>, <c>Consulta.exe</c>, <c>modConsulta.bas</c>.
/// </summary>
public class FrmConsultas : Form
{
    private readonly Func<FrmConsultaVentas>? _ventasFactory;
    private readonly Func<FrmConsultaDocumentos>? _documentosFactory;
    private readonly TreeView _tree;
    private readonly Panel _panelDerecho;
    private Form? _formularioActivo;

    public FrmConsultas(Func<FrmConsultaVentas>? ventasFactory = null, Func<FrmConsultaDocumentos>? documentosFactory = null)
    {
        _ventasFactory = ventasFactory;
        _documentosFactory = documentosFactory;

        Text = "Consultas";
        WindowState = FormWindowState.Maximized;

        var toolbar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 44, Padding = new Padding(8) };
        var btnEjecutar = new Button { Text = "Ejecutar", Width = 100 };
        var btnExportar = new Button { Text = "Exportar", Width = 100 };
        var btnImprimir = new Button { Text = "Imprimir", Width = 100 };
        toolbar.Controls.AddRange([btnEjecutar, btnExportar, btnImprimir]);

        _tree = new TreeView { Dock = DockStyle.Left, Width = 260 };
        _tree.Nodes.Add("Ventas", "Consulta de Ventas");
        _tree.Nodes.Add("Documentos", "Consulta de Documentos");
        _tree.AfterSelect += (_, e) => MostrarConsulta(e.Node.Name);

        _panelDerecho = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8), BackColor = Color.WhiteSmoke };

        btnEjecutar.Click += async (_, _) => await EjecutarAccionAsync(static c => c.EjecutarConsultaAsync());
        btnExportar.Click += async (_, _) => await EjecutarAccionAsync(static c => c.ExportarAsync());
        btnImprimir.Click += async (_, _) => await EjecutarAccionAsync(static c => c.ImprimirAsync());

        Controls.Add(_panelDerecho);
        Controls.Add(_tree);
        Controls.Add(toolbar);
        Load += (_, _) =>
        {
            _tree.SelectedNode = _tree.Nodes[0];
            MostrarConsulta("Ventas");
        };
    }

    private void MostrarConsulta(string clave)
    {
        _formularioActivo?.Close();
        _panelDerecho.Controls.Clear();

        _formularioActivo = clave switch
        {
            "Documentos" => _documentosFactory?.Invoke() ?? new FrmConsultaDocumentos(new DummyReporteRepository()),
            _ => _ventasFactory?.Invoke() ?? new FrmConsultaVentas(new ObtenerReporteVentaMensualIntegradoHandler(new DummyReporteRepository()))
        };

        _formularioActivo.TopLevel = false;
        _formularioActivo.FormBorderStyle = FormBorderStyle.None;
        _formularioActivo.Dock = DockStyle.Fill;
        _panelDerecho.Controls.Add(_formularioActivo);
        _formularioActivo.Show();
    }

    private async Task EjecutarAccionAsync(Func<IConsultaOperable, Task> accion)
    {
        if (_formularioActivo is IConsultaOperable consulta)
            await accion(consulta);
    }
}

public interface IConsultaOperable
{
    Task EjecutarConsultaAsync();
    Task ExportarAsync();
    Task ImprimirAsync();
}
