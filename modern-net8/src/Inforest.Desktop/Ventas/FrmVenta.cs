using System.ComponentModel;
using Inforest.Application.Caja;
using Inforest.Application.Ventas;
using Inforest.Desktop.Caja;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Pantalla de emisión de documentos de venta.
/// Legacy: frmVenta.frm. BR-002, BR-013.
/// </summary>
public class FrmVenta : Form
{
    private readonly Pedido _pedido;
    private readonly EmitirDocumentoHandler? _emitirHandler;
    private readonly ObtenerMediosPagoHandler? _mediosPagoHandler;
    private readonly PagarDocumentoHandler? _pagarHandler;
    private readonly RegistrarPagosMultiplesHandler? _registrarPagosMultiplesHandler;
    private readonly ProcesarPagoPinPadHandler? _procesarPagoPinPadHandler;
    private readonly ObtenerTerminalesPinPadHandler? _obtenerTerminalesPinPadHandler;
    private readonly AnularDocumentoHandler? _anularHandler;

    private readonly BindingList<VentaItemRow> _items = [];
    private readonly Label _lblTotal;
    private readonly Label _lblNeto;
    private readonly Label _lblImpuesto;
    private readonly NumericUpDown _nudPropina;
    private readonly NumericUpDown _nudDescuento;
    private readonly ComboBox _cboTipoDocumento;
    private readonly TextBox _txtCliente;

    public FrmVenta(
        Pedido pedido,
        EmitirDocumentoHandler? emitirHandler = null,
        ObtenerMediosPagoHandler? mediosPagoHandler = null,
        PagarDocumentoHandler? pagarHandler = null,
        RegistrarPagosMultiplesHandler? registrarPagosMultiplesHandler = null,
        ProcesarPagoPinPadHandler? procesarPagoPinPadHandler = null,
        ObtenerTerminalesPinPadHandler? obtenerTerminalesPinPadHandler = null,
        AnularDocumentoHandler? anularHandler = null)
    {
        _pedido = pedido;
        _emitirHandler = emitirHandler;
        _mediosPagoHandler = mediosPagoHandler;
        _pagarHandler = pagarHandler;
        _registrarPagosMultiplesHandler = registrarPagosMultiplesHandler;
        _procesarPagoPinPadHandler = procesarPagoPinPadHandler;
        _obtenerTerminalesPinPadHandler = obtenerTerminalesPinPadHandler;
        _anularHandler = anularHandler;

        Text = "Venta / Emisión de Documento";
        Width = 1000;
        Height = 640;
        StartPosition = FormStartPosition.CenterScreen;

        _lblTotal = new Label { AutoSize = true, Font = new Font("Segoe UI", 11, FontStyle.Bold) };
        _lblNeto = new Label { AutoSize = true };
        _lblImpuesto = new Label { AutoSize = true };
        _nudPropina = new NumericUpDown { DecimalPlaces = 2, Maximum = 999999, Width = 140 };
        _nudDescuento = new NumericUpDown { DecimalPlaces = 2, Maximum = 999999, Width = 140 };
        _cboTipoDocumento = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
        _cboTipoDocumento.Items.AddRange(["01 - Boleta", "03 - Factura"]);
        _cboTipoDocumento.SelectedIndex = 0;
        _txtCliente = new TextBox { Width = 200 };

        _nudPropina.ValueChanged += (_, _) => ActualizarTotales();
        _nudDescuento.ValueChanged += (_, _) => ActualizarTotales();

        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 620 };

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            ReadOnly = true,
            DataSource = _items,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };
        split.Panel1.Controls.Add(grid);

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 10,
            Padding = new Padding(12),
            AutoSize = true
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));

        panel.Controls.Add(new Label { Text = "Pedido:", AutoSize = true }, 0, 0);
        panel.Controls.Add(new Label { Text = _pedido.CodigoPedido, AutoSize = true }, 1, 0);
        panel.Controls.Add(new Label { Text = "Mesa:", AutoSize = true }, 0, 1);
        panel.Controls.Add(new Label { Text = _pedido.CodigoMesa ?? string.Empty, AutoSize = true }, 1, 1);
        panel.Controls.Add(new Label { Text = "Tipo documento:", AutoSize = true }, 0, 2);
        panel.Controls.Add(_cboTipoDocumento, 1, 2);
        panel.Controls.Add(new Label { Text = "Cliente (RUC/DNI):", AutoSize = true }, 0, 3);
        panel.Controls.Add(_txtCliente, 1, 3);
        panel.Controls.Add(new Label { Text = "Propina:", AutoSize = true }, 0, 4);
        panel.Controls.Add(_nudPropina, 1, 4);
        panel.Controls.Add(new Label { Text = "Descuento:", AutoSize = true }, 0, 5);
        panel.Controls.Add(_nudDescuento, 1, 5);
        panel.Controls.Add(new Label { Text = "Neto:", AutoSize = true }, 0, 6);
        panel.Controls.Add(_lblNeto, 1, 6);
        panel.Controls.Add(new Label { Text = "Impuestos:", AutoSize = true }, 0, 7);
        panel.Controls.Add(_lblImpuesto, 1, 7);
        panel.Controls.Add(new Label { Text = "Total:", AutoSize = true }, 0, 8);
        panel.Controls.Add(_lblTotal, 1, 8);

        var btnEmitir = new Button { Text = "Emitir", Width = 100 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 100 };
        btnCancelar.Click += (_, _) => Close();
        btnEmitir.Click += async (_, _) => await EmitirAsync();

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
        acciones.Controls.Add(btnEmitir);
        acciones.Controls.Add(btnCancelar);
        panel.Controls.Add(acciones, 1, 9);

        split.Panel2.Controls.Add(panel);
        Controls.Add(split);

        CargarItems();
        ActualizarTotales();
    }

    private void CargarItems()
    {
        foreach (var det in _pedido.Detalles)
        {
            _items.Add(new VentaItemRow(
                det.Item,
                det.CodigoProducto,
                det.CodigoProducto,
                det.Cantidad,
                det.PrecioNeto,
                det.PrecioVenta,
                det.Impuesto1 + det.Impuesto2 + det.Impuesto3));
        }
    }

    private void ActualizarTotales()
    {
        var neto = _items.Sum(x => x.Cantidad * x.PrecioNeto);
        var impuesto = _items.Sum(x => x.ImpuestoTotal);
        var total = neto + impuesto + _nudPropina.Value - _nudDescuento.Value;

        _lblNeto.Text = $"{neto:0.00}";
        _lblImpuesto.Text = $"{impuesto:0.00}";
        _lblTotal.Text = $"{total:0.00}";
    }

    private async Task EmitirAsync()
    {
        if (_emitirHandler is null)
        {
            MessageBox.Show("Handler de emisión no configurado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var tipoCodigo = _cboTipoDocumento.SelectedIndex == 0 ? "01" : "03";
        var command = new EmitirDocumentoCommand(
            _pedido.CodigoPedido,
            tipoCodigo,
            string.IsNullOrWhiteSpace(_txtCliente.Text) ? null : _txtCliente.Text.Trim(),
            _nudPropina.Value,
            _nudDescuento.Value);

        var result = await _emitirHandler.HandleAsync(command);
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var documento = result.Valor!;
        using var frmPago = new FrmPago(
            documento.CodigoDocumento,
            documento.Total,
            documento.CodigoCaja,
            _mediosPagoHandler,
            _pagarHandler,
            _registrarPagosMultiplesHandler,
            _procesarPagoPinPadHandler,
            _obtenerTerminalesPinPadHandler);
        if (frmPago.ShowDialog(this) == DialogResult.OK)
        {
            using var frmDocumento = new FrmDocumento(documento, _anularHandler);
            frmDocumento.ShowDialog(this);
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    private sealed record VentaItemRow(
        string Item,
        string CodigoProducto,
        string Descripcion,
        decimal Cantidad,
        decimal PrecioNeto,
        decimal PrecioVenta,
        decimal ImpuestoTotal);
}
