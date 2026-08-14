using Inforest.Application.Reportes;

namespace Inforest.Desktop.Consultas;

/// <summary>
/// Legacy: <c>frmRepRegVenta.frm</c> + <c>spRep_RegVenta</c>.
/// </summary>
public class FrmConsultaVentas : Form, IConsultaOperable
{
    private readonly ObtenerReporteVentaMensualIntegradoHandler _handler;
    private readonly DataGridView _grid;
    private readonly DateTimePicker _dtpDesde;
    private readonly DateTimePicker _dtpHasta;
    private readonly ComboBox _cmbCaja;
    private readonly ComboBox _cmbMozo;
    private readonly ComboBox _cmbCanal;

    public FrmConsultaVentas(ObtenerReporteVentaMensualIntegradoHandler handler)
    {
        _handler = handler;
        Text = "Consulta de Ventas";

        var filtros = new TableLayoutPanel { Dock = DockStyle.Top, Height = 88, ColumnCount = 8, Padding = new Padding(8) };
        filtros.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
        filtros.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        filtros.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
        filtros.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        filtros.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
        filtros.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        filtros.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
        filtros.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _dtpDesde = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddDays(-7) };
        _dtpHasta = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _cmbCaja = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        _cmbMozo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        _cmbCanal = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        _cmbCaja.Items.AddRange(["Todas", "001", "002"]);
        _cmbMozo.Items.AddRange(["Todos", "Mesero", "Delivery"]);
        _cmbCanal.Items.AddRange(["Todos", "Local", "Delivery", "Llevar"]);
        _cmbCaja.SelectedIndex = _cmbMozo.SelectedIndex = _cmbCanal.SelectedIndex = 0;

        filtros.Controls.Add(new Label { Text = "Desde", AutoSize = true }, 0, 0);
        filtros.Controls.Add(_dtpDesde, 1, 0);
        filtros.Controls.Add(new Label { Text = "Hasta", AutoSize = true }, 2, 0);
        filtros.Controls.Add(_dtpHasta, 3, 0);
        filtros.Controls.Add(new Label { Text = "Caja", AutoSize = true }, 4, 0);
        filtros.Controls.Add(_cmbCaja, 5, 0);
        filtros.Controls.Add(new Label { Text = "Mozo", AutoSize = true }, 6, 0);
        filtros.Controls.Add(_cmbMozo, 7, 0);
        filtros.Controls.Add(new Label { Text = "Canal", AutoSize = true }, 0, 1);
        filtros.Controls.Add(_cmbCanal, 1, 1);

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(8) };
        var btnConsultar = new Button { Text = "Consultar", Width = 110 };
        var btnExportar = new Button { Text = "Exportar", Width = 110 };
        acciones.Controls.AddRange([btnConsultar, btnExportar]);
        btnConsultar.Click += async (_, _) => await EjecutarConsultaAsync();
        btnExportar.Click += async (_, _) => await ExportarAsync();

        _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = true, AllowUserToAddRows = false, ReadOnly = true };

        Controls.Add(_grid);
        Controls.Add(acciones);
        Controls.Add(filtros);
        Load += async (_, _) => await EjecutarConsultaAsync();
    }

    public async Task EjecutarConsultaAsync()
    {
        var desde = _dtpDesde.Value.Date;
        var hasta = _dtpHasta.Value.Date;
        var parametros = new VentaMensualIntegradoParametros
        {
            Ano = desde.Year.ToString(),
            Mes = desde.Month.ToString("00"),
            Tipo = "D",
            TipoPrecio = ExpresionPrecio.Venta,
            HoraCorte = 6,
            MetaMensual = 0,
            DiasDelMes = DateTime.DaysInMonth(desde.Year, desde.Month)
        };

        var resultado = await _handler.HandleAsync(new ObtenerReporteVentaMensualIntegradoQuery(parametros));
        var filas = resultado.Filas
            .Where(row => row.Dia >= desde.Day && row.Dia <= Math.Min(hasta.Day, DateTime.DaysInMonth(desde.Year, desde.Month)))
            .Select(row => new ConsultaVentaRow(row.Dia, row.Salon, row.Delivery, row.Llevar, row.Venta, row.Cantidad, _cmbCaja.Text, _cmbCanal.Text))
            .ToList();
        _grid.DataSource = filas;
    }

    public async Task ExportarAsync()
    {
        await EjecutarConsultaAsync();
        using var dialog = new SaveFileDialog { Filter = "CSV|*.csv", FileName = "consulta-ventas.csv" };
        if (dialog.ShowDialog(this) != DialogResult.OK || _grid.DataSource is not List<ConsultaVentaRow> filas)
            return;

        var lineas = new List<string> { "Dia,Salon,Delivery,Llevar,Venta,Cantidad,Caja,Canal" };
        lineas.AddRange(filas.Select(f => string.Join(',', f.Dia, f.Salon, f.Delivery, f.Llevar, f.Venta, f.Cantidad, f.Caja, f.Canal)));
        await File.WriteAllLinesAsync(dialog.FileName, lineas);
    }

    public Task ImprimirAsync()
    {
        MessageBox.Show("Use el módulo de reportes para impresión formal de ventas.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        return Task.CompletedTask;
    }

    private sealed record ConsultaVentaRow(int Dia, double Salon, double Delivery, double Llevar, double Venta, int Cantidad, string Caja, string Canal);
}

