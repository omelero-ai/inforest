using Inforest.Application.Caja;
using Inforest.Domain.Entities.Caja;

namespace Inforest.Desktop.Caja.Recibos;

/// <summary>
/// Detalle/ingreso de recibo de ingreso de caja.
/// Legacy: frmReciboIngresoDetalle.frm.
/// BR-RECIBO-001, BR-RECIBO-003, BR-RECIBO-004, BR-RECIBO-005, BR-RECIBO-011.
/// </summary>
public class FrmReciboIngresoDetalle : Form
{
    private readonly RegistrarIngresoHandler? _registrarHandler;
    private readonly ReciboIngreso? _reciboExistente;

    private readonly TextBox txtRecibo      = new() { ReadOnly = true, BackColor = SystemColors.Control, MaxLength = 10 };
    private readonly TextBox txtFecha       = new() { ReadOnly = true, BackColor = SystemColors.Control };
    private readonly TextBox txtMonto       = new() { MaxLength = 12 };
    private readonly TextBox txtDescripcion = new() { MaxLength = 100 };
    private readonly TextBox txtAutorizacion = new() { MaxLength = 50 };
    private readonly TextBox txtReferencia  = new() { MaxLength = 30 };
    private readonly TextBox txtTurno       = new() { ReadOnly = true, BackColor = SystemColors.Control };
    private readonly TextBox txtPedido      = new() { MaxLength = 10 };
    private readonly ComboBox cboTipoPago   = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox cboMoneda     = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly CheckBox chkAnticipo   = new() { Text = "Es Anticipo", Checked = true, AutoSize = true };
    private readonly Label lblEstado        = new() { AutoSize = true, ForeColor = Color.DarkRed };
    private readonly Button btnGuardar      = new() { Text = "Guardar", Width = 90 };
    private readonly Button btnCancelar     = new() { Text = "Cancelar", Width = 90 };

    public string CodigoTurno { get; set; } = string.Empty;
    public string CodigoCaja  { get; set; } = string.Empty;
    public string Usuario     { get; set; } = string.Empty;

    /// <param name="reciboExistente">Si null, modo NUEVO; si no null, modo LECTURA/ANULADO.</param>
    public FrmReciboIngresoDetalle(ReciboIngreso? reciboExistente = null, RegistrarIngresoHandler? registrarHandler = null)
    {
        _reciboExistente = reciboExistente;
        _registrarHandler = registrarHandler;

        Text = "Recibo de Ingreso";
        Width = 500;
        Height = 430;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;

        InicializarCombos();
        Controls.Add(BuildLayout());

        Load += (_, _) => CargarDatos();
        btnGuardar.Click += async (_, _) => await GuardarAsync();
        btnCancelar.Click += (_, _) => Close();
    }

    private void InicializarCombos()
    {
        cboTipoPago.Items.AddRange(["Efectivo MN (01)", "Tarjeta (02)", "Cheque (03)"]);
        cboTipoPago.SelectedIndex = 0;
        cboMoneda.Items.AddRange(["Moneda Nacional (01)", "Moneda Extranjera (02)"]);
        cboMoneda.SelectedIndex = 0;
    }

    private void CargarDatos()
    {
        if (_reciboExistente is null)
        {
            // Modo nuevo
            txtFecha.Text = DateTime.Now.ToString("dd/MM/yyyy");
            txtTurno.Text = CodigoTurno;
            txtRecibo.Text = "(auto)";
            lblEstado.Text = string.Empty;
        }
        else
        {
            // Modo lectura
            txtRecibo.Text = _reciboExistente.CodigoRecibo;
            txtFecha.Text = _reciboExistente.Fecha.ToString("dd/MM/yyyy");
            txtMonto.Text = _reciboExistente.Monto.ToString("N2");
            txtDescripcion.Text = _reciboExistente.Descripcion;
            txtAutorizacion.Text = _reciboExistente.Autorizacion ?? string.Empty;
            txtReferencia.Text = _reciboExistente.Referencia ?? string.Empty;
            txtTurno.Text = _reciboExistente.CodigoTurno ?? string.Empty;
            txtPedido.Text = _reciboExistente.CodigoPedido ?? string.Empty;
            chkAnticipo.Checked = _reciboExistente.EsAnticipo;
            lblEstado.Text = _reciboExistente.DescripcionEstado;
            lblEstado.ForeColor = _reciboExistente.EstaAnulado ? Color.Red : Color.DarkGreen;

            // Seleccionar tipo pago y moneda
            cboTipoPago.SelectedIndex = _reciboExistente.TipoPago switch
            {
                "02" => 1,
                "03" => 2,
                _ => 0
            };
            cboMoneda.SelectedIndex = _reciboExistente.Moneda == "02" ? 1 : 0;

            // BR-RECIBO-003: bloquear si anulado o pagado
            bool bloqueado = _reciboExistente.EstaAnulado || _reciboExistente.EstaPagado;
            txtMonto.ReadOnly = bloqueado;
            txtDescripcion.ReadOnly = bloqueado;
            txtAutorizacion.ReadOnly = bloqueado;
            txtReferencia.ReadOnly = bloqueado;
            txtPedido.ReadOnly = bloqueado;
            cboTipoPago.Enabled = !bloqueado;
            cboMoneda.Enabled = !bloqueado;
            chkAnticipo.Enabled = !bloqueado;
            btnGuardar.Enabled = false; // no permitir re-guardar en modo lectura
        }
    }

    private async Task GuardarAsync()
    {
        if (!decimal.TryParse(txtMonto.Text.Replace(",", ""), out var monto) || monto <= 0)
        {
            MessageBox.Show("Ingrese el monto del recibo.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(txtDescripcion.Text))
        {
            MessageBox.Show("Ingrese la descripción del recibo.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_registrarHandler is null)
        {
            MessageBox.Show("Servicio de registro no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var tipoPago = cboTipoPago.SelectedIndex switch { 1 => "02", 2 => "03", _ => "01" };
        var moneda   = cboMoneda.SelectedIndex == 1 ? "02" : "01";

        var result = await _registrarHandler.HandleAsync(new RegistrarIngresoCommand(
            Fecha: DateTime.Now,
            Monto: monto,
            Descripcion: txtDescripcion.Text.Trim(),
            Moneda: moneda,
            TipoPago: tipoPago,
            Referencia: txtReferencia.Text.Trim(),
            Autorizacion: txtAutorizacion.Text.Trim(),
            EsAnticipo: chkAnticipo.Checked,
            Turno: CodigoTurno,
            Caja: CodigoCaja,
            Usuario: Usuario,
            CodigoPedido: txtPedido.Text.Trim()));

        if (result.EsExitoso)
        {
            MessageBox.Show("Registro guardado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }
        else
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private Control BuildLayout()
    {
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            ColumnCount = 2,
            RowCount = 11
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        void Row(int r, string label, Control ctrl)
        {
            table.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Padding = new Padding(0, 4, 0, 0) }, 0, r);
            ctrl.Dock = DockStyle.Fill;
            table.Controls.Add(ctrl, 1, r);
        }

        Row(0, "Recibo:", txtRecibo);
        Row(1, "Fecha:", txtFecha);
        Row(2, "Monto:", txtMonto);
        Row(3, "Moneda:", cboMoneda);
        Row(4, "Tipo Pago:", cboTipoPago);
        Row(5, "Descripción:", txtDescripcion);
        Row(6, "Referencia:", txtReferencia);
        Row(7, "Autoriza:", txtAutorizacion);
        Row(8, "Turno:", txtTurno);
        Row(9, "Pedido Ref.:", txtPedido);

        var panelCheck = new FlowLayoutPanel { Dock = DockStyle.Fill };
        panelCheck.Controls.Add(chkAnticipo);
        panelCheck.Controls.Add(lblEstado);
        table.Controls.Add(panelCheck, 1, 10);
        table.Controls.Add(new Label { Text = string.Empty }, 0, 10);

        var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 38, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        btnPanel.Controls.Add(btnCancelar);
        btnPanel.Controls.Add(btnGuardar);

        var container = new Panel { Dock = DockStyle.Fill };
        container.Controls.Add(table);
        container.Controls.Add(btnPanel);
        return container;
    }
}
