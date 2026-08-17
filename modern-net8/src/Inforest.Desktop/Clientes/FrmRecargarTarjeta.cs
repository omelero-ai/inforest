using Inforest.Application.Delivery;
using Inforest.Application.Interfaces;

namespace Inforest.Desktop.Clientes;

/// <summary>
/// Listado de recargas de tarjetas RFID.
/// Legacy: FrmRecargarTarjeta.frm.
/// Reglas: BR-RFID-005.
/// </summary>
public sealed class FrmRecargarTarjeta : Form
{
    private readonly ObtenerRecargasTarjetaHandler _obtenerHandler;
    private readonly RegistrarRecargaTarjetaHandler _registrarHandler;
    private readonly ObtenerTarjetasProximidadHandler _obtenerTarjetasHandler;
    private readonly ISessionService _sessionService;

    private readonly DateTimePicker _dtpDesde = new() { Format = DateTimePickerFormat.Short, Width = 120 };
    private readonly DateTimePicker _dtpHasta = new() { Format = DateTimePickerFormat.Short, Width = 120 };
    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AutoGenerateColumns = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false
    };
    private readonly Label _lblEstado = new() { AutoSize = true };
    private List<RecargaTarjetaMovimiento> _movimientos = [];

    public FrmRecargarTarjeta(
        ObtenerRecargasTarjetaHandler obtenerHandler,
        RegistrarRecargaTarjetaHandler registrarHandler,
        ObtenerTarjetasProximidadHandler obtenerTarjetasHandler,
        ISessionService sessionService)
    {
        _obtenerHandler = obtenerHandler;
        _registrarHandler = registrarHandler;
        _obtenerTarjetasHandler = obtenerTarjetasHandler;
        _sessionService = sessionService;

        Text = "Lista de Recargas";
        Width = 1200;
        Height = 680;
        StartPosition = FormStartPosition.CenterScreen;

        _dtpDesde.Value = DateTime.Today.AddDays(-7);
        _dtpHasta.Value = DateTime.Today;

        ConfigurarColumnas();
        Controls.Add(BuildLayout());

        Load += async (_, _) => await CargarAsync();
    }

    private void ConfigurarColumnas()
    {
        _grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Id", HeaderText = "Correlativo", Width = 95 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "CodigoRfid", HeaderText = "Tarjeta", Width = 170 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Tipo", HeaderText = "Tipo", Width = 70 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "DocumentoReferencia", HeaderText = "Doc. Referencia", Width = 170 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "FechaRegistro", HeaderText = "Fecha", Width = 110 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "MontoIngreso", HeaderText = "Monto Ingreso", Width = 120 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "MontoSalida", HeaderText = "Monto Salida", Width = 120 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "MontoAnterior", HeaderText = "Monto Anterior", Width = 120 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "MontoFinal", HeaderText = "Monto Final", Width = 120 });
    }

    private Control BuildLayout()
    {
        var tool = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 42,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(6)
        };

        var btnBuscar = new Button { Text = "Procesar", Width = 90 };
        var btnNuevo = new Button { Text = "Nueva recarga", Width = 120 };
        var btnCerrar = new Button { Text = "Cerrar", Width = 90 };

        btnBuscar.Click += async (_, _) => await CargarAsync();
        btnNuevo.Click += async (_, _) => await NuevaRecargaAsync();
        btnCerrar.Click += (_, _) => Close();

        tool.Controls.Add(new Label { Text = "Desde:", AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
        tool.Controls.Add(_dtpDesde);
        tool.Controls.Add(new Label { Text = "Hasta:", AutoSize = true, Margin = new Padding(8, 8, 0, 0) });
        tool.Controls.Add(_dtpHasta);
        tool.Controls.Add(btnBuscar);
        tool.Controls.Add(new Label { Width = 20 });
        tool.Controls.Add(btnNuevo);
        tool.Controls.Add(btnCerrar);

        var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 28, Padding = new Padding(6, 4, 6, 2) };
        bottom.Controls.Add(_lblEstado);

        var root = new Panel { Dock = DockStyle.Fill };
        root.Controls.Add(_grid);
        root.Controls.Add(bottom);
        root.Controls.Add(tool);
        return root;
    }

    private async Task CargarAsync()
    {
        var desde = _dtpDesde.Value.Date;
        var hasta = _dtpHasta.Value.Date.AddDays(1).AddSeconds(-1);
        var result = await _obtenerHandler.HandleAsync(new ObtenerRecargasTarjetaQuery(desde, hasta));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _movimientos = result.Valor!.ToList();
        _grid.DataSource = _movimientos.Select(x => new
        {
            x.Id,
            x.CodigoRfid,
            x.Tipo,
            x.DocumentoReferencia,
            FechaRegistro = x.FechaRegistro == DateTime.MinValue ? string.Empty : x.FechaRegistro.ToString("dd/MM/yyyy"),
            x.MontoIngreso,
            x.MontoSalida,
            x.MontoAnterior,
            x.MontoFinal
        }).ToList();

        _lblEstado.Text = $"Registros: {_movimientos.Count}";
    }

    private async Task NuevaRecargaAsync()
    {
        var sesion = _sessionService.SesionActual;
        var usuario = sesion?.CodigoUsuario ?? string.Empty;
        if (string.IsNullOrWhiteSpace(usuario))
        {
            MessageBox.Show("No existe una sesión activa para registrar recargas.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var modal = new FrmRecargarTarjetaDetalle(
            _obtenerTarjetasHandler,
            _registrarHandler,
            usuario);

        if (modal.ShowDialog(this) == DialogResult.OK)
            await CargarAsync();
    }
}
