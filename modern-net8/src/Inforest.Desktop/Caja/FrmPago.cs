using System.ComponentModel;
using Inforest.Application.Caja;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Entities.Hardware;

namespace Inforest.Desktop.Caja;

/// <summary>
/// Pantalla de cobranza de documentos.
/// Legacy: frmPago.frm. BR-007, BR-013.
/// </summary>
public class FrmPago : Form
{
    private readonly string _codigoDocumento;
    private readonly decimal _totalDocumento;
    private readonly string? _codigoCaja;
    private readonly ObtenerMediosPagoHandler? _mediosPagoHandler;
    private readonly PagarDocumentoHandler? _pagarHandler;
    private readonly RegistrarPagosMultiplesHandler? _registrarPagosMultiplesHandler;
    private readonly ProcesarPagoPinPadHandler? _procesarPagoPinPadHandler;
    private readonly ObtenerTerminalesPinPadHandler? _obtenerTerminalesPinPadHandler;

    private readonly BindingList<PagoGridRow> _pagos = [];
    private readonly Label _lblVuelto;
    private readonly Label _lblTotal;
    private readonly Label _lblSaldoPendiente;
    private readonly Label _lblTotalPagado;
    private readonly NumericUpDown _nudMonto;
    private readonly ComboBox _cboMedios;
    private IReadOnlyList<MedioPago> _mediosPago = [];

    public FrmPago(
        string codigoDocumento,
        decimal totalDocumento,
        string? codigoCaja = null,
        ObtenerMediosPagoHandler? mediosPagoHandler = null,
        PagarDocumentoHandler? pagarHandler = null,
        RegistrarPagosMultiplesHandler? registrarPagosMultiplesHandler = null,
        ProcesarPagoPinPadHandler? procesarPagoPinPadHandler = null,
        ObtenerTerminalesPinPadHandler? obtenerTerminalesPinPadHandler = null)
    {
        _codigoDocumento = codigoDocumento;
        _totalDocumento = totalDocumento;
        _codigoCaja = codigoCaja;
        _mediosPagoHandler = mediosPagoHandler;
        _pagarHandler = pagarHandler;
        _registrarPagosMultiplesHandler = registrarPagosMultiplesHandler;
        _procesarPagoPinPadHandler = procesarPagoPinPadHandler;
        _obtenerTerminalesPinPadHandler = obtenerTerminalesPinPadHandler;

        Text = "Cobro de Documento";
        Width = 760;
        Height = 580;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        _lblTotal = new Label { AutoSize = true, Font = new Font("Segoe UI", 12, FontStyle.Bold), Text = $"{_totalDocumento:0.00}" };
        _lblVuelto = new Label { AutoSize = true };
        _lblSaldoPendiente = new Label { AutoSize = true };
        _lblTotalPagado = new Label { AutoSize = true };
        _nudMonto = new NumericUpDown { DecimalPlaces = 2, Maximum = 999999, Width = 160, Value = _totalDocumento };
        _cboMedios = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        _nudMonto.ValueChanged += (_, _) => ActualizarResumen();

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 220,
            Padding = new Padding(12),
            ColumnCount = 2
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));

        panel.Controls.Add(new Label { Text = "Documento:", AutoSize = true }, 0, 0);
        panel.Controls.Add(new Label { Text = _codigoDocumento, AutoSize = true }, 1, 0);
        panel.Controls.Add(new Label { Text = "Total a pagar:", AutoSize = true }, 0, 1);
        panel.Controls.Add(_lblTotal, 1, 1);
        panel.Controls.Add(new Label { Text = "Medio de pago:", AutoSize = true }, 0, 2);
        panel.Controls.Add(_cboMedios, 1, 2);
        panel.Controls.Add(new Label { Text = "Monto pago:", AutoSize = true }, 0, 3);
        panel.Controls.Add(_nudMonto, 1, 3);
        panel.Controls.Add(new Label { Text = "Total pagado:", AutoSize = true }, 0, 4);
        panel.Controls.Add(_lblTotalPagado, 1, 4);
        panel.Controls.Add(new Label { Text = "Saldo pendiente:", AutoSize = true }, 0, 5);
        panel.Controls.Add(_lblSaldoPendiente, 1, 5);
        panel.Controls.Add(new Label { Text = "Vuelto:", AutoSize = true }, 0, 6);
        panel.Controls.Add(_lblVuelto, 1, 6);

        var btnAgregar = new Button { Text = "Agregar", Width = 90 };
        var btnCobrar = new Button { Text = "Cobrar", Width = 90 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 90 };
        btnCancelar.Click += (_, _) => Close();
        btnAgregar.Click += async (_, _) => await AgregarPagoAsync();
        btnCobrar.Click += async (_, _) => await CobrarAsync();

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(12, 0, 12, 0) };
        acciones.Controls.Add(btnAgregar);
        acciones.Controls.Add(btnCobrar);
        acciones.Controls.Add(btnCancelar);

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            ReadOnly = true,
            DataSource = _pagos,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        Controls.Add(grid);
        Controls.Add(acciones);
        Controls.Add(panel);

        Load += async (_, _) => await CargarMediosPagoAsync();
        ActualizarResumen();
    }

    private async Task CargarMediosPagoAsync()
    {
        if (_mediosPagoHandler is null)
        {
            _cboMedios.Items.AddRange(["Efectivo", "Tarjeta", "Cheque", "Vale", "Otro"]);
            if (_cboMedios.Items.Count > 0) _cboMedios.SelectedIndex = 0;
            return;
        }

        try
        {
            var result = await _mediosPagoHandler.HandleAsync(new ObtenerMediosPagoQuery());
            if (result.EsExitoso && result.Valor?.Count > 0)
            {
                _mediosPago = result.Valor;
                _cboMedios.DataSource = _mediosPago.ToList();
                _cboMedios.DisplayMember = nameof(MedioPago.Descripcion);
                _cboMedios.ValueMember = nameof(MedioPago.Codigo);
                return;
            }
        }
        catch { /* fallback */ }

        _cboMedios.Items.AddRange(["Efectivo", "Tarjeta"]);
        if (_cboMedios.Items.Count > 0) _cboMedios.SelectedIndex = 0;
    }

    private async Task AgregarPagoAsync()
    {
        var medioCodigo = _cboMedios.SelectedValue?.ToString() ?? _cboMedios.Text;
        var medioDesc = _cboMedios.Text;

        if (EsMedioTarjetaSeleccionado())
        {
            await AgregarPagoPinPadAsync(medioCodigo, medioDesc);
            return;
        }

        var saldoPendiente = ObtenerSaldoPendiente();
        var vuelto = Math.Max(0, _nudMonto.Value - saldoPendiente);
        _pagos.Add(new PagoGridRow(medioCodigo, medioDesc, _nudMonto.Value, vuelto, null));
        ActualizarResumen();
    }

    private async Task CobrarAsync()
    {
        if (_pagos.Count > 0 && _registrarPagosMultiplesHandler is not null)
        {
            var registroResult = await _registrarPagosMultiplesHandler.HandleAsync(new RegistrarPagosMultiplesCommand(
                _codigoDocumento,
                _pagos.Select(p => new LineaPago(p.CodigoMedio, p.Monto, p.Vuelto, p.Referencia)).ToList()));

            if (!registroResult.EsExitoso)
            {
                MessageBox.Show(registroResult.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var vueltoMulti = Math.Max(0, _pagos.Sum(p => p.Monto) - _totalDocumento);
            MessageBox.Show($"Cobro registrado. Vuelto: {vueltoMulti:0.00}", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult = DialogResult.OK;
            Close();
            return;
        }

        if (_pagarHandler is null)
        {
            MessageBox.Show("Handler de cobro no configurado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var totalPagado = _pagos.Count > 0 ? _pagos.Sum(p => p.Monto) : _nudMonto.Value;
        if (totalPagado < _totalDocumento)
        {
            MessageBox.Show($"El monto ingresado ({totalPagado:0.00}) no cubre el total ({_totalDocumento:0.00}).", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var medioCodigo = _cboMedios.SelectedValue?.ToString() ?? _cboMedios.Text;
        if (string.IsNullOrWhiteSpace(medioCodigo))
        {
            MessageBox.Show("Seleccione un medio de pago.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var vuelto = Math.Max(0, totalPagado - _totalDocumento);
        var command = new PagarDocumentoCommand(_codigoDocumento, medioCodigo, totalPagado, vuelto, null);
        var result = await _pagarHandler.HandleAsync(command);

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show($"Cobro registrado. Vuelto: {vuelto:0.00}", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        DialogResult = DialogResult.OK;
        Close();
    }

    private async Task AgregarPagoPinPadAsync(string medioCodigo, string medioDesc)
    {
        if (_procesarPagoPinPadHandler is null)
        {
            MessageBox.Show("La integración de PinPad no está configurada.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var saldoPendiente = ObtenerSaldoPendiente();
        if (_nudMonto.Value > saldoPendiente)
        {
            MessageBox.Show("El pago con tarjeta no puede exceder el saldo pendiente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var terminalSeleccionado = await SeleccionarTerminalPinPadAsync();
        if (terminalSeleccionado.Cancelled)
            return;

        var result = await _procesarPagoPinPadHandler.HandleAsync(new ProcesarPagoPinPadCommand(
            _nudMonto.Value,
            1,
            terminalSeleccionado.Terminal?.CodigoTerminal));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var descripcion = terminalSeleccionado.Terminal is null
            ? medioDesc
            : $"{medioDesc} ({terminalSeleccionado.Terminal.Descripcion})";

        _pagos.Add(new PagoGridRow(
            medioCodigo,
            descripcion,
            _nudMonto.Value,
            0m,
            result.Valor!.Referencia));

        ActualizarResumen();
    }

    private bool EsMedioTarjetaSeleccionado()
    {
        if (_mediosPago.Count > 0 && _cboMedios.SelectedValue is string codigo)
            return _mediosPago.FirstOrDefault(m => m.Codigo == codigo)?.Tipo == TipoMedioPago.Tarjeta;

        return _cboMedios.Text.Contains("tarjeta", StringComparison.OrdinalIgnoreCase);
    }

    private decimal ObtenerSaldoPendiente()
        => Math.Max(0, _totalDocumento - _pagos.Sum(p => p.Monto));

    private void ActualizarResumen()
    {
        var totalPagado = _pagos.Sum(p => p.Monto);
        var saldoPendiente = Math.Max(0, _totalDocumento - totalPagado);
        var vuelto = _pagos.Count > 0
            ? Math.Max(0, totalPagado - _totalDocumento)
            : Math.Max(0, _nudMonto.Value - _totalDocumento);

        _lblTotalPagado.Text = $"{totalPagado:0.00}";
        _lblSaldoPendiente.Text = $"{saldoPendiente:0.00}";
        _lblVuelto.Text = $"{vuelto:0.00}";

        if (saldoPendiente > 0 && _nudMonto.Value > saldoPendiente && !EsMedioTarjetaSeleccionado())
            return;

        if (saldoPendiente > 0 && _nudMonto.Value != saldoPendiente)
            _nudMonto.Value = saldoPendiente;
    }

    private async Task<TerminalPinPadSelection> SeleccionarTerminalPinPadAsync()
    {
        if (_obtenerTerminalesPinPadHandler is null || string.IsNullOrWhiteSpace(_codigoCaja))
            return TerminalPinPadSelection.None();

        var queryResult = await _obtenerTerminalesPinPadHandler.HandleAsync(new ObtenerTerminalesPinPadQuery(_codigoCaja));
        if (!queryResult.EsExitoso || queryResult.Valor is null || queryResult.Valor.Count == 0)
            return TerminalPinPadSelection.None();

        if (queryResult.Valor.Count == 1)
            return TerminalPinPadSelection.From(queryResult.Valor[0]);

        using var selector = new FrmPagoPinPad(queryResult.Valor);
        return selector.ShowDialog(this) == DialogResult.OK && selector.TerminalSeleccionado is not null
            ? TerminalPinPadSelection.From(selector.TerminalSeleccionado)
            : TerminalPinPadSelection.Cancel();
    }

    private sealed record PagoGridRow(string CodigoMedio, string MedioPago, decimal Monto, decimal Vuelto, string? Referencia);

    private sealed record TerminalPinPadSelection(TerminalPinPad? Terminal, bool Cancelled)
    {
        public static TerminalPinPadSelection None() => new(null, false);
        public static TerminalPinPadSelection From(TerminalPinPad terminal) => new(terminal, false);
        public static TerminalPinPadSelection Cancel() => new(null, true);
    }
}
