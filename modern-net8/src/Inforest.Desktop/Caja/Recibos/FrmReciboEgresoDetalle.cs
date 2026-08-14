using Inforest.Application.Caja;
using Inforest.Domain.Entities.Caja;

namespace Inforest.Desktop.Caja.Recibos;

/// <summary>
/// Detalle/ingreso de recibo de egreso de caja.
/// Legacy: frmReciboEgresoDetalle.frm.
/// BR-RECIBO-007, BR-RECIBO-008, BR-RECIBO-009, BR-RECIBO-010, BR-RECIBO-011, BR-RECIBO-012.
/// </summary>
public class FrmReciboEgresoDetalle : Form
{
    private readonly RegistrarEgresoHandler _registrarHandler;
    private readonly ReciboEgreso? _reciboExistente;
    private readonly bool _solicitaClaveEgreso;

    private readonly TextBox txtRecibo       = new() { ReadOnly = true, BackColor = SystemColors.Control };
    private readonly TextBox txtFecha        = new() { ReadOnly = true, BackColor = SystemColors.Control };
    private readonly TextBox txtMonto        = new() { MaxLength = 14 };
    private readonly TextBox txtDescripcion  = new() { MaxLength = 100 };
    private readonly TextBox txtAutorizacion = new() { MaxLength = 50 };
    private readonly TextBox txtObservacion  = new() { MaxLength = 100 };
    private readonly TextBox txtNotaCredito  = new() { MaxLength = 20 };
    private readonly TextBox txtTurno        = new() { ReadOnly = true, BackColor = SystemColors.Control };
    private readonly ComboBox cboMoneda      = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox cboModoEgreso  = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly TextBox txtTipoEgreso   = new() { MaxLength = 3 };
    private readonly Label lblEstado         = new() { AutoSize = true, ForeColor = Color.DarkRed };
    private readonly Button btnGuardar       = new() { Text = "Guardar", Width = 90 };
    private readonly Button btnCancelar      = new() { Text = "Cancelar", Width = 90 };

    public string CodigoTurno { get; set; } = string.Empty;
    public string CodigoCaja  { get; set; } = string.Empty;
    public string Usuario     { get; set; } = string.Empty;

    public FrmReciboEgresoDetalle(
        RegistrarEgresoHandler registrarHandler,
        ReciboEgreso? reciboExistente = null,
        bool solicitaClaveEgreso = false)
    {
        _registrarHandler = registrarHandler;
        _reciboExistente = reciboExistente;
        _solicitaClaveEgreso = solicitaClaveEgreso;

        Text = "Recibo de Egreso";
        Width = 520;
        Height = 500;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;

        InicializarCombos();
        Controls.Add(BuildLayout());

        Load += (_, _) => CargarDatos();
        btnGuardar.Click += async (_, _) => await GuardarAsync();
        btnCancelar.Click += (_, _) => Close();

        // BR-RECIBO-012: txtNotaCredito habilitado sólo en modo NOTACREDITO
        cboModoEgreso.SelectedIndexChanged += (_, _) =>
            txtNotaCredito.Enabled = cboModoEgreso.SelectedItem?.ToString() == "NOTACREDITO";
    }

    private void InicializarCombos()
    {
        cboMoneda.Items.AddRange(["Moneda Nacional (01)", "Moneda Extranjera (02)"]);
        cboMoneda.SelectedIndex = 0;
        // BR-RECIBO-010: modos de egreso disponibles
        cboModoEgreso.Items.AddRange(["VARIOS", "NOTACREDITO", "MERCADERIA"]);
        cboModoEgreso.SelectedIndex = 0;
    }

    private void CargarDatos()
    {
        if (_reciboExistente is null)
        {
            txtFecha.Text = DateTime.Now.ToString("dd/MM/yyyy");
            txtTurno.Text = CodigoTurno;
            txtRecibo.Text = "(auto)";
            txtMonto.Text = "0.00";
            lblEstado.Text = string.Empty;
        }
        else
        {
            txtRecibo.Text       = _reciboExistente.CodigoRecibo;
            txtFecha.Text        = _reciboExistente.Fecha.ToString("dd/MM/yyyy");
            txtMonto.Text        = _reciboExistente.Monto.ToString("N2");
            txtDescripcion.Text  = _reciboExistente.Descripcion;
            txtAutorizacion.Text = _reciboExistente.Autorizacion ?? string.Empty;
            txtObservacion.Text  = _reciboExistente.Observacion ?? string.Empty;
            txtNotaCredito.Text  = _reciboExistente.Referencia ?? string.Empty;
            txtTurno.Text        = _reciboExistente.CodigoTurno ?? string.Empty;
            txtTipoEgreso.Text   = _reciboExistente.TipoEgreso ?? string.Empty;
            lblEstado.Text       = _reciboExistente.DescripcionEstado;
            lblEstado.ForeColor  = _reciboExistente.EstaAnulado ? Color.Red : Color.DarkGreen;

            cboMoneda.SelectedIndex     = _reciboExistente.Moneda == "02" ? 1 : 0;
            cboModoEgreso.SelectedIndex = _reciboExistente.ModoEgreso switch
            {
                "NOTACREDITO" => 1,
                "MERCADERIA"  => 2,
                _ => 0
            };

            // BR-RECIBO-008: bloquear si anulado
            bool bloqueado = _reciboExistente.EstaAnulado;
            txtMonto.ReadOnly        = bloqueado;
            txtDescripcion.ReadOnly  = bloqueado;
            txtAutorizacion.ReadOnly = bloqueado;
            txtObservacion.ReadOnly  = bloqueado;
            txtNotaCredito.ReadOnly  = bloqueado;
            txtTipoEgreso.ReadOnly   = bloqueado;
            cboMoneda.Enabled        = !bloqueado;
            cboModoEgreso.Enabled    = !bloqueado;
            btnGuardar.Enabled       = false; // no re-guardar en modo lectura
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

        var modoEgreso = cboModoEgreso.SelectedItem?.ToString() ?? "VARIOS";

        // BR-RECIBO-012: NOTACREDITO requiere referencia
        if (modoEgreso == "NOTACREDITO" && string.IsNullOrWhiteSpace(txtNotaCredito.Text))
        {
            MessageBox.Show("Seleccione una nota de crédito.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var moneda = cboMoneda.SelectedIndex == 1 ? "02" : "01";

        var result = await _registrarHandler.HandleAsync(new RegistrarEgresoCommand(
            Fecha: DateTime.Now,
            Monto: monto,
            Descripcion: txtDescripcion.Text.Trim(),
            ModoEgreso: modoEgreso,
            Moneda: moneda,
            Autorizacion: txtAutorizacion.Text.Trim(),
            TipoEgreso: txtTipoEgreso.Text.Trim(),
            Referencia: modoEgreso == "NOTACREDITO" ? txtNotaCredito.Text.Trim() : null,
            Turno: CodigoTurno,
            Caja: CodigoCaja,
            Usuario: Usuario,
            Observacion: txtObservacion.Text.Trim()));

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
            RowCount = 12
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
        Row(4, "Modo Egreso:", cboModoEgreso);
        Row(5, "Tipo Egreso:", txtTipoEgreso);
        Row(6, "Descripción:", txtDescripcion);
        Row(7, "Nota de Crédito:", txtNotaCredito);
        Row(8, "Autoriza:", txtAutorizacion);
        Row(9, "Observación:", txtObservacion);
        Row(10, "Turno:", txtTurno);

        var panelEstado = new FlowLayoutPanel { Dock = DockStyle.Fill };
        panelEstado.Controls.Add(lblEstado);
        table.Controls.Add(new Label { Text = string.Empty }, 0, 11);
        table.Controls.Add(panelEstado, 1, 11);

        var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 38, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        btnPanel.Controls.Add(btnCancelar);
        btnPanel.Controls.Add(btnGuardar);

        var container = new Panel { Dock = DockStyle.Fill };
        container.Controls.Add(table);
        container.Controls.Add(btnPanel);
        return container;
    }
}
