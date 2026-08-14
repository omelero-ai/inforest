using System.ComponentModel;
using Inforest.Application.Caja;
using Inforest.Domain.Entities.Caja;

namespace Inforest.Desktop.Caja;

/// <summary>
/// Pantalla de cobranza de documentos.
/// Legacy: frmPago.frm. BR-007, BR-013.
/// </summary>
public class FrmPago : Form
{
    private readonly string _codigoDocumento;
    private readonly decimal _totalDocumento;
    private readonly ObtenerMediosPagoHandler? _mediosPagoHandler;
    private readonly PagarDocumentoHandler? _pagarHandler;

    private readonly BindingList<PagoGridRow> _pagos = [];
    private readonly Label _lblVuelto;
    private readonly Label _lblTotal;
    private readonly NumericUpDown _nudMonto;
    private readonly ComboBox _cboMedios;
    private IReadOnlyList<MedioPago> _mediosPago = [];

    public FrmPago(
        string codigoDocumento,
        decimal totalDocumento,
        ObtenerMediosPagoHandler? mediosPagoHandler = null,
        PagarDocumentoHandler? pagarHandler = null)
    {
        _codigoDocumento = codigoDocumento;
        _totalDocumento = totalDocumento;
        _mediosPagoHandler = mediosPagoHandler;
        _pagarHandler = pagarHandler;

        Text = "Cobro de Documento";
        Width = 760;
        Height = 540;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        _lblTotal = new Label { AutoSize = true, Font = new Font("Segoe UI", 12, FontStyle.Bold), Text = $"{_totalDocumento:0.00}" };
        _lblVuelto = new Label { AutoSize = true };
        _nudMonto = new NumericUpDown { DecimalPlaces = 2, Maximum = 999999, Width = 160, Value = _totalDocumento };
        _cboMedios = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        _nudMonto.ValueChanged += (_, _) => ActualizarVuelto();

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 160,
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
        panel.Controls.Add(new Label { Text = "Vuelto:", AutoSize = true }, 0, 4);
        panel.Controls.Add(_lblVuelto, 1, 4);

        var btnAgregar = new Button { Text = "Agregar", Width = 90 };
        var btnCobrar = new Button { Text = "Cobrar", Width = 90 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 90 };
        btnCancelar.Click += (_, _) => Close();
        btnAgregar.Click += (_, _) => AgregarPago();
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
        ActualizarVuelto();
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

    private void AgregarPago()
    {
        var medioCodigo = _cboMedios.SelectedValue?.ToString() ?? _cboMedios.Text;
        var medioDesc = _cboMedios.Text;
        var vuelto = Math.Max(0, _nudMonto.Value - _totalDocumento);
        _pagos.Add(new PagoGridRow(medioCodigo, medioDesc, _nudMonto.Value, vuelto));
        ActualizarVuelto();
    }

    private async Task CobrarAsync()
    {
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

    private void ActualizarVuelto()
    {
        var vuelto = Math.Max(0, _nudMonto.Value - _totalDocumento);
        _lblVuelto.Text = $"{vuelto:0.00}";
    }

    private sealed record PagoGridRow(string CodigoMedio, string MedioPago, decimal Monto, decimal Vuelto);
}
