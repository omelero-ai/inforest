using Inforest.Application.Delivery;

namespace Inforest.Desktop.Clientes;

/// <summary>
/// Detalle de recarga RFID.
/// Legacy: FrmRecargarTarjetaDetalle.frm.
/// Reglas: BR-RFID-006, BR-RFID-007.
/// </summary>
public sealed class FrmRecargarTarjetaDetalle : Form
{
    private readonly ObtenerTarjetasProximidadHandler _obtenerTarjetasHandler;
    private readonly RegistrarRecargaTarjetaHandler _registrarHandler;
    private readonly string _usuario;

    private readonly ComboBox _cbTarjetas = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 340 };
    private readonly Label _lblCliente = new() { AutoSize = true };
    private readonly Label _lblSaldo = new() { AutoSize = true };
    private readonly NumericUpDown _nudMonto = new()
    {
        DecimalPlaces = 2,
        Maximum = 99999999,
        ThousandsSeparator = true,
        Width = 140
    };
    private readonly ComboBox _cbTipoPago = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
    private readonly TextBox _txtReferencia = new() { Width = 220, MaxLength = 40 };
    private readonly TextBox _txtDocumento = new() { Width = 220, MaxLength = 40 };

    private List<TarjetaOption> _tarjetas = [];

    public FrmRecargarTarjetaDetalle(
        ObtenerTarjetasProximidadHandler obtenerTarjetasHandler,
        RegistrarRecargaTarjetaHandler registrarHandler,
        string usuario)
    {
        _obtenerTarjetasHandler = obtenerTarjetasHandler;
        _registrarHandler = registrarHandler;
        _usuario = usuario;

        Text = "Recargar Tarjeta";
        Width = 560;
        Height = 360;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;

        _cbTipoPago.Items.AddRange(["Efectivo MN (01)", "Tarjeta (02)", "Cheque (03)"]);
        _cbTipoPago.SelectedIndex = 0;
        _txtDocumento.Text = $"REC-{DateTime.Now:yyyyMMddHHmmss}";

        Controls.Add(BuildLayout());
        Load += async (_, _) => await CargarTarjetasAsync();
        _cbTarjetas.SelectedIndexChanged += (_, _) => MostrarTarjetaSeleccionada();
    }

    private Control BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(14),
            ColumnCount = 2,
            RowCount = 8
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        AddField(root, 0, "Tarjeta:", _cbTarjetas);
        AddField(root, 1, "Cliente:", _lblCliente);
        AddField(root, 2, "Saldo actual:", _lblSaldo);
        AddField(root, 3, "Monto recarga:", _nudMonto);
        AddField(root, 4, "Tipo pago:", _cbTipoPago);
        AddField(root, 5, "Referencia:", _txtReferencia);
        AddField(root, 6, "Doc referencia:", _txtDocumento);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft
        };
        var btnCancelar = new Button { Text = "Cancelar", Width = 90 };
        var btnGuardar = new Button { Text = "Guardar", Width = 90 };
        btnCancelar.Click += (_, _) => DialogResult = DialogResult.Cancel;
        btnGuardar.Click += async (_, _) => await GuardarAsync();
        buttons.Controls.Add(btnCancelar);
        buttons.Controls.Add(btnGuardar);
        root.Controls.Add(new Label(), 0, 7);
        root.Controls.Add(buttons, 1, 7);
        return root;
    }

    private static void AddField(TableLayoutPanel root, int row, string label, Control control)
    {
        root.Controls.Add(new Label
        {
            Text = label,
            AutoSize = true,
            Margin = new Padding(0, 7, 0, 0)
        }, 0, row);
        control.Margin = new Padding(0, 3, 0, 6);
        root.Controls.Add(control, 1, row);
    }

    private async Task CargarTarjetasAsync()
    {
        var result = await _obtenerTarjetasHandler.HandleAsync(new ObtenerTarjetasProximidadQuery());
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _tarjetas = result.Valor!
            .Select(t => new TarjetaOption(t.CodigoRfid, t.NombreCliente, t.MontoDisponible, t.Estado))
            .OrderBy(t => t.CodigoRfid)
            .ToList();

        _cbTarjetas.DataSource = _tarjetas;
        _cbTarjetas.DisplayMember = nameof(TarjetaOption.Display);
        _cbTarjetas.ValueMember = nameof(TarjetaOption.CodigoRfid);

        MostrarTarjetaSeleccionada();
    }

    private void MostrarTarjetaSeleccionada()
    {
        var item = _cbTarjetas.SelectedItem as TarjetaOption;
        _lblCliente.Text = item is null
            ? "*** Cliente ***"
            : $"{item.NombreCliente} ({item.Estado})";
        _lblSaldo.Text = item is null
            ? "0.00"
            : item.Saldo.ToString("N2");
    }

    private async Task GuardarAsync()
    {
        var item = _cbTarjetas.SelectedItem as TarjetaOption;
        if (item is null)
        {
            MessageBox.Show("Debe seleccionar una tarjeta.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_nudMonto.Value <= 0)
        {
            MessageBox.Show("Ingrese un monto mayor a cero.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _registrarHandler.HandleAsync(new RegistrarRecargaTarjetaCommand(
            item.CodigoRfid,
            _nudMonto.Value,
            _usuario,
            _txtDocumento.Text.Trim()));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        MessageBox.Show(
            $"Recarga registrada.\nSaldo anterior: {result.Valor!.MontoAnterior:N2}\nSaldo final: {result.Valor!.MontoFinal:N2}",
            Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        DialogResult = DialogResult.OK;
    }

    private sealed record TarjetaOption(
        string CodigoRfid,
        string NombreCliente,
        decimal Saldo,
        string Estado)
    {
        public string Display => $"{CodigoRfid} - {NombreCliente}";
    }
}
