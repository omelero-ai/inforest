using Inforest.Application.Ventas;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Formulario para actualizar cortesía y canal de venta de un pedido/documento.
/// Legacy: FrmActualizarPedidos (frmUpdateDatosPedido.frm). BR-ACTPED-001/002/003.
/// SP: usp_ActualizarCabPeDoc (@opcion='1' cortesía, @opcion='2' canal de venta).
/// </summary>
public sealed class FrmActualizarDatosPedido : Form
{
    private readonly string _codigoDocumento;
    private readonly ObtenerDatosPedidoActualizarHandler? _obtenerHandler;
    private readonly ActualizarCortesiaPedidoHandler? _cortesiaHandler;
    private readonly ActualizarCanalVentaPedidoHandler? _canalHandler;

    private DatosPedidoActualizar? _datosActuales;
    private string _codigoClienteDelivery = string.Empty;

    // Cortesía
    private readonly Label _lblCortesiaActual;
    private readonly ComboBox _cboCortesia;
    private readonly Button _btnModificarCortesia;

    // Canal de venta
    private readonly Label _lblCanalActual;
    private readonly ComboBox _cboCanal;
    private readonly Button _btnBuscarClienteDelivery;
    private readonly Label _lblClienteDelivery;
    private readonly Button _btnModificarCanal;
    private readonly Button _btnCerrar;

    public FrmActualizarDatosPedido(
        string codigoDocumento,
        ObtenerDatosPedidoActualizarHandler? obtenerHandler = null,
        ActualizarCortesiaPedidoHandler? cortesiaHandler = null,
        ActualizarCanalVentaPedidoHandler? canalHandler = null)
    {
        _codigoDocumento = codigoDocumento;
        _obtenerHandler = obtenerHandler;
        _cortesiaHandler = cortesiaHandler;
        _canalHandler = canalHandler;

        Text = "Actualizar Cortesía / Canal de Venta";
        Width = 620;
        Height = 380;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        _lblCortesiaActual = new Label { AutoSize = true };
        _cboCortesia = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
        _btnModificarCortesia = new Button { Text = "Modificar", Width = 100, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White };
        _btnModificarCortesia.FlatAppearance.BorderSize = 0;

        _lblCanalActual = new Label { AutoSize = true };
        _cboCanal = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
        _lblClienteDelivery = new Label { AutoSize = true, Text = "(sin cliente delivery)", ForeColor = Color.Gray };
        _btnBuscarClienteDelivery = new Button { Text = "...", Width = 32, Height = 26 };
        _btnModificarCanal = new Button { Text = "Modificar", Width = 100, Height = 28, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White };
        _btnModificarCanal.FlatAppearance.BorderSize = 0;

        _btnCerrar = new Button { Text = "Cerrar", Width = 90, Height = 30, FlatStyle = FlatStyle.Flat };

        _btnModificarCortesia.Click += async (_, _) => await ModificarCortesiaAsync();
        _btnModificarCanal.Click += async (_, _) => await ModificarCanalAsync();
        _btnBuscarClienteDelivery.Click += (_, _) => BuscarClienteDelivery();
        _btnCerrar.Click += (_, _) => Close();
        _cboCanal.SelectedIndexChanged += (_, _) =>
        {
            bool esDelivery = _cboCanal.SelectedItem is OpcionCatalogo op && op.Codigo == "02";
            _btnBuscarClienteDelivery.Visible = esDelivery;
            _lblClienteDelivery.Visible = esDelivery;
        };

        var main = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(16),
            ColumnCount = 1,
            RowCount = 4,
            AutoSize = true
        };

        // --- Bloque Cortesía ---
        var groupCortesia = new GroupBox { Text = "Motivo Cortesía", Dock = DockStyle.Fill, Height = 110, Padding = new Padding(8) };
        var panelCort = new TableLayoutPanel { ColumnCount = 3, RowCount = 2, AutoSize = true, Dock = DockStyle.Fill };
        panelCort.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
        panelCort.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panelCort.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));

        panelCort.Controls.Add(new Label { Text = "Cortesía actual:", AutoSize = true }, 0, 0);
        panelCort.Controls.Add(_lblCortesiaActual, 1, 0);
        panelCort.Controls.Add(new Label { Text = "Cambiar por:", AutoSize = true }, 0, 1);
        panelCort.Controls.Add(_cboCortesia, 1, 1);
        panelCort.Controls.Add(_btnModificarCortesia, 2, 1);

        groupCortesia.Controls.Add(panelCort);
        main.Controls.Add(groupCortesia, 0, 0);

        // --- Bloque Canal de Venta ---
        var groupCanal = new GroupBox { Text = "Canal de Venta", Dock = DockStyle.Fill, Height = 130, Padding = new Padding(8) };
        var panelCanal = new TableLayoutPanel { ColumnCount = 3, RowCount = 3, AutoSize = true, Dock = DockStyle.Fill };
        panelCanal.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
        panelCanal.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panelCanal.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));

        panelCanal.Controls.Add(new Label { Text = "Canal actual:", AutoSize = true }, 0, 0);
        panelCanal.Controls.Add(_lblCanalActual, 1, 0);
        panelCanal.Controls.Add(new Label { Text = "Cambiar por:", AutoSize = true }, 0, 1);
        panelCanal.Controls.Add(_cboCanal, 1, 1);
        panelCanal.Controls.Add(_btnModificarCanal, 2, 1);

        var deliveryRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
        deliveryRow.Controls.Add(_lblClienteDelivery);
        deliveryRow.Controls.Add(_btnBuscarClienteDelivery);
        panelCanal.Controls.Add(new Label { Text = "Cliente delivery:", AutoSize = true }, 0, 2);
        panelCanal.Controls.Add(deliveryRow, 1, 2);

        groupCanal.Controls.Add(panelCanal);
        main.Controls.Add(groupCanal, 0, 1);

        // --- Botones ---
        var btnRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom };
        btnRow.Controls.Add(_btnCerrar);
        main.Controls.Add(btnRow, 0, 3);

        Controls.Add(main);

        Load += async (_, _) => await CargarAsync();
    }

    private async Task CargarAsync()
    {
        if (_obtenerHandler is null)
        {
            MessageBox.Show("Servicio no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _obtenerHandler.HandleAsync(new ObtenerDatosPedidoActualizarQuery(_codigoDocumento));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var (datos, cortesias, canales) = result.Valor!;
        _datosActuales = datos;

        _lblCortesiaActual.Text = string.IsNullOrWhiteSpace(datos.DescripcionCortesia)
            ? "(sin cortesía)"
            : $"{datos.CodigoCortesia} — {datos.DescripcionCortesia}";

        _lblCanalActual.Text = string.IsNullOrWhiteSpace(datos.DescripcionCanalVenta)
            ? "(sin canal)"
            : $"{datos.CodigoCanalVenta} — {datos.DescripcionCanalVenta}";

        _cboCortesia.Items.Clear();
        foreach (var c in cortesias)
            _cboCortesia.Items.Add(c);
        _cboCortesia.DisplayMember = nameof(OpcionCatalogo.Descripcion);
        _cboCortesia.ValueMember = nameof(OpcionCatalogo.Codigo);

        _cboCanal.Items.Clear();
        foreach (var c in canales)
            _cboCanal.Items.Add(c);
        _cboCanal.DisplayMember = nameof(OpcionCatalogo.Descripcion);
        _cboCanal.ValueMember = nameof(OpcionCatalogo.Codigo);

        _btnBuscarClienteDelivery.Visible = false;
        _lblClienteDelivery.Visible = false;
    }

    private async Task ModificarCortesiaAsync()
    {
        if (_cortesiaHandler is null)
        {
            MessageBox.Show("Servicio no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_datosActuales is null) return;

        if (_cboCortesia.SelectedItem is not OpcionCatalogo seleccionada)
        {
            MessageBox.Show("Seleccione la cortesía de reemplazo.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _cortesiaHandler.HandleAsync(new ActualizarCortesiaPedidoCommand(
            CodigoDocumento: _codigoDocumento,
            CodigoCortesiaAnterior: _datosActuales.CodigoCortesia,
            CodigoCortesiaNueva: seleccionada.Codigo,
            CodigoUsuario: Environment.UserName));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show("Cortesía modificada correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        await CargarAsync();
    }

    private async Task ModificarCanalAsync()
    {
        if (_canalHandler is null)
        {
            MessageBox.Show("Servicio no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_datosActuales is null) return;

        if (_cboCanal.SelectedItem is not OpcionCatalogo seleccionado)
        {
            MessageBox.Show("Seleccione el canal de venta.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _canalHandler.HandleAsync(new ActualizarCanalVentaPedidoCommand(
            CodigoDocumento: _codigoDocumento,
            CodigoCanalAnterior: _datosActuales.CodigoCanalVenta,
            CodigoCanalNuevo: seleccionado.Codigo,
            CodigoClienteDelivery: _codigoClienteDelivery,
            CodigoUsuario: Environment.UserName));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show("Canal de venta modificado correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        await CargarAsync();
    }

    private void BuscarClienteDelivery()
    {
        using var dlg = new Form
        {
            Text = "Buscar Cliente Delivery",
            Width = 340,
            Height = 130,
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false
        };
        var lbl = new Label { Text = "Código de cliente:", Left = 12, Top = 16, AutoSize = true };
        var txt = new TextBox { Left = 12, Top = 38, Width = 280 };
        var btnOk = new Button { Text = "Aceptar", Left = 120, Top = 70, Width = 80, DialogResult = DialogResult.OK };
        var btnCan = new Button { Text = "Cancelar", Left = 210, Top = 70, Width = 80, DialogResult = DialogResult.Cancel };
        dlg.Controls.AddRange([lbl, txt, btnOk, btnCan]);
        dlg.AcceptButton = btnOk;
        dlg.CancelButton = btnCan;

        if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(txt.Text))
        {
            _codigoClienteDelivery = txt.Text.Trim();
            _lblClienteDelivery.Text = _codigoClienteDelivery;
        }
    }
}
