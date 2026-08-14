using Inforest.Application.Motorizado;
using MotorizadoEntity = Inforest.Domain.Entities.Motorizado.Motorizado;

namespace Inforest.Desktop.Motorizado;

/// <summary>
/// Formulario de control de llegada y salida del motorizado.
/// <para>
/// Legacy: <c>frmLlegadaSalida.frm</c> — es el formulario de startup del exe Motorizado.
/// Muestra grilla de pedidos asignados al motorizado y permite registrar salida/llegada.
/// Se actualiza automáticamente cada 10 segundos (Timer Interval=10000 en Legacy).
/// </para>
/// Regla BR-DEL-006.
/// </summary>
public partial class LlegadaSalidaForm : Form
{
    private readonly ObtenerMotorizadosActivosHandler _motorizadosHandler;
    private readonly RegistrarLlegadaSalidaHandler _llegadaSalidaHandler;
    private readonly System.Windows.Forms.Timer _timer;

    public LlegadaSalidaForm(
        ObtenerMotorizadosActivosHandler motorizadosHandler,
        RegistrarLlegadaSalidaHandler llegadaSalidaHandler)
    {
        _motorizadosHandler = motorizadosHandler;
        _llegadaSalidaHandler = llegadaSalidaHandler;
        _timer = new System.Windows.Forms.Timer { Interval = 10_000 };
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Llegada / Salida Motorizado";
        Size = new Size(1100, 650);
        WindowState = FormWindowState.Maximized;

        var toolbar = new ToolStrip();
        var btnSalida = new ToolStripButton("Salida");
        var btnLlegada = new ToolStripButton("Llegada");
        var btnRefrescar = new ToolStripButton("Refrescar");
        var btnPedidosEntregados = new ToolStripButton("Pedidos Entregados");
        toolbar.Items.AddRange([btnSalida, btnLlegada, new ToolStripSeparator(), btnRefrescar, btnPedidosEntregados]);

        var cmbMotorizado = new ComboBox { Left = 10, Top = 35, Width = 200, Name = "cmbMotorizado", DropDownStyle = ComboBoxStyle.DropDownList };
        var lblMotorizado = new Label { Text = "Motorizado:", Left = 10, Top = 15, Width = 90 };

        var grid = new DataGridView
        {
            Name = "grdPedidos",
            Left = 0, Top = 65,
            Width = ClientSize.Width,
            Height = ClientSize.Height - 65,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        Controls.AddRange([toolbar, lblMotorizado, cmbMotorizado, grid]);

        Load += async (_, _) =>
        {
            await CargarMotorizadosAsync(cmbMotorizado);
            _timer.Start();
        };

        FormClosed += (_, _) => _timer.Stop();

        _timer.Tick += async (_, _) => await RefrescarAsync(cmbMotorizado, grid);

        btnRefrescar.Click += async (_, _) => await RefrescarAsync(cmbMotorizado, grid);
        btnSalida.Click += async (_, _) => await RegistrarMovimientoAsync(cmbMotorizado, grid, esLlegada: false);
        btnLlegada.Click += async (_, _) => await RegistrarMovimientoAsync(cmbMotorizado, grid, esLlegada: true);
    }

    private async Task CargarMotorizadosAsync(ComboBox cmb)
    {
        var result = await _motorizadosHandler.HandleAsync(new ObtenerMotorizadosActivosQuery());
        if (!result.EsExitoso) return;
        cmb.DataSource = result.Valor?.ToList();
        cmb.DisplayMember = "Codigo";
        cmb.ValueMember = "Codigo";
    }

    private Task RefrescarAsync(ComboBox cmb, DataGridView grid)
    {
        // En la implementación completa, recarga los pedidos del motorizado seleccionado
        return Task.CompletedTask;
    }

    private async Task RegistrarMovimientoAsync(ComboBox cmb, DataGridView grid, bool esLlegada)
    {
        if (cmb.SelectedValue is not string codigoMotorizado) return;
        if (grid.SelectedRows.Count == 0) return;

        // Obtiene codigoPedido de la fila seleccionada
        var codigoPedido = grid.SelectedRows[0].Cells["CodigoPedido"].Value?.ToString() ?? string.Empty;

        var result = await _llegadaSalidaHandler.HandleAsync(
            new RegistrarLlegadaSalidaCommand(codigoPedido, codigoMotorizado, esLlegada));

        if (!result.EsExitoso)
            MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        else
            await RefrescarAsync(cmb, grid);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) _timer.Dispose();
        base.Dispose(disposing);
    }
}
