using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Creación y visualización de una nota de crédito.
/// Legacy: frmNotaCreditoDetalle.frm — cmdOpcion_Click (Grabar / Anular).
/// BR-NC-001: Total > 0.
/// BR-NC-002: Impuestos >= 0.
/// BR-NC-003: Total NC <= saldo disponible del documento.
/// BR-NC-004: Observación obligatoria.
/// BR-NC-006: No reanular NC anulada.
/// POS-FUNC-006.
/// </summary>
public class FrmNotaCreditoDetalle : Form
{
    private readonly NotaCredito? _nc;
    private readonly AnularNotaCreditoHandler _anularHandler;
    private readonly EmitirNotaCreditoHandler? _emitirHandler;
    private readonly string _usuarioActual;
    private readonly Func<NotaCredito, Task>? _onSaved;

    private readonly TextBox _txtCodigo;
    private readonly DateTimePicker _dtpFecha;
    private readonly TextBox _txtDocumento;
    private readonly TextBox _txtNeto;
    private readonly TextBox _txtImpuesto1;
    private readonly TextBox _txtImpuesto2;
    private readonly TextBox _txtImpuesto3;
    private readonly TextBox _txtTotal;
    private readonly TextBox _txtObservacion;
    private readonly Label _lblEstado;
    private readonly Button _btnAnular;

    public FrmNotaCreditoDetalle(
        NotaCredito? nc,
        AnularNotaCreditoHandler anularHandler,
        string usuarioActual,
        Func<NotaCredito, Task>? onSaved,
        EmitirNotaCreditoHandler? emitirHandler = null)
    {
        _nc = nc;
        _anularHandler = anularHandler;
        _emitirHandler = emitirHandler;
        _usuarioActual = usuarioActual;
        _onSaved = onSaved;

        bool modoVer = nc is not null;

        Text = modoVer ? "Detalle Nota de Crédito" : "Nueva Nota de Crédito";
        Width = 520;
        Height = 440;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(16),
            ColumnCount = 2,
            RowCount = 10
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));

        _txtCodigo = new TextBox { ReadOnly = modoVer, Width = 200 };
        _dtpFecha = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 200, Enabled = !modoVer };
        _txtDocumento = new TextBox { ReadOnly = modoVer, Width = 200 };
        _txtNeto = new TextBox { ReadOnly = modoVer, Width = 200, Text = "0.00" };
        _txtImpuesto1 = new TextBox { ReadOnly = modoVer, Width = 200, Text = "0.00" };
        _txtImpuesto2 = new TextBox { ReadOnly = modoVer, Width = 200, Text = "0.00" };
        _txtImpuesto3 = new TextBox { ReadOnly = modoVer, Width = 200, Text = "0.00" };
        _txtTotal = new TextBox { ReadOnly = true, Width = 200, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
        _txtObservacion = new TextBox { ReadOnly = modoVer, Width = 200, MaxLength = 50 };
        _lblEstado = new Label { AutoSize = true };

        layout.Controls.Add(new Label { Text = "Código NC:", AutoSize = true }, 0, 0);
        layout.Controls.Add(_txtCodigo, 1, 0);
        layout.Controls.Add(new Label { Text = "Fecha:", AutoSize = true }, 0, 1);
        layout.Controls.Add(_dtpFecha, 1, 1);
        layout.Controls.Add(new Label { Text = "Documento:", AutoSize = true }, 0, 2);
        layout.Controls.Add(_txtDocumento, 1, 2);
        layout.Controls.Add(new Label { Text = "Neto:", AutoSize = true }, 0, 3);
        layout.Controls.Add(_txtNeto, 1, 3);
        layout.Controls.Add(new Label { Text = "Impuesto 1:", AutoSize = true }, 0, 4);
        layout.Controls.Add(_txtImpuesto1, 1, 4);
        layout.Controls.Add(new Label { Text = "Impuesto 2:", AutoSize = true }, 0, 5);
        layout.Controls.Add(_txtImpuesto2, 1, 5);
        layout.Controls.Add(new Label { Text = "Impuesto 3:", AutoSize = true }, 0, 6);
        layout.Controls.Add(_txtImpuesto3, 1, 6);
        layout.Controls.Add(new Label { Text = "Total:", AutoSize = true }, 0, 7);
        layout.Controls.Add(_txtTotal, 1, 7);
        layout.Controls.Add(new Label { Text = "Observación:", AutoSize = true }, 0, 8);
        layout.Controls.Add(_txtObservacion, 1, 8);
        layout.Controls.Add(new Label { Text = "Estado:", AutoSize = true }, 0, 9);
        layout.Controls.Add(_lblEstado, 1, 9);

        if (!modoVer)
        {
            _txtNeto.TextChanged += (_, _) => ActualizarTotal();
            _txtImpuesto1.TextChanged += (_, _) => ActualizarTotal();
            _txtImpuesto2.TextChanged += (_, _) => ActualizarTotal();
            _txtImpuesto3.TextChanged += (_, _) => ActualizarTotal();
        }

        var botones = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(8, 8, 0, 0) };

        if (modoVer)
        {
            _btnAnular = new Button { Text = "Anular", Width = 90 };
            _btnAnular.Click += async (_, _) => await AnularAsync();
            _btnAnular.Enabled = nc!.Estado != "AN";
            botones.Controls.Add(_btnAnular);
        }
        else
        {
            _btnAnular = new Button { Visible = false };
            var btnGrabar = new Button { Text = "Grabar", Width = 90 };
            btnGrabar.Click += async (_, _) => await GrabarAsync();
            botones.Controls.Add(btnGrabar);
        }

        var btnCerrar = new Button { Text = "Cerrar", Width = 90 };
        btnCerrar.Click += (_, _) => Close();
        botones.Controls.Add(btnCerrar);

        Controls.Add(layout);
        Controls.Add(botones);

        if (modoVer) CargarDesdeDominio(nc!);
    }

    private void CargarDesdeDominio(NotaCredito nc)
    {
        _txtCodigo.Text = nc.CodigoNotaCredito;
        _dtpFecha.Value = nc.Fecha == default ? DateTime.Today : nc.Fecha;
        _txtDocumento.Text = nc.CodigoDocumento;
        _txtNeto.Text = nc.Neto.ToString("N2");
        _txtImpuesto1.Text = nc.Impuesto1.ToString("N2");
        _txtImpuesto2.Text = nc.Impuesto2.ToString("N2");
        _txtImpuesto3.Text = nc.Impuesto3.ToString("N2");
        _txtTotal.Text = nc.Total.ToString("N2");
        _txtObservacion.Text = nc.Observacion;
        _lblEstado.Text = nc.Estado switch { "PE" => "Pendiente", "PR" => "Procesado", "AN" => "Anulado", _ => nc.Estado };
    }

    private void ActualizarTotal()
    {
        if (decimal.TryParse(_txtNeto.Text, out var neto)
            && decimal.TryParse(_txtImpuesto1.Text, out var imp1)
            && decimal.TryParse(_txtImpuesto2.Text, out var imp2)
            && decimal.TryParse(_txtImpuesto3.Text, out var imp3))
        {
            _txtTotal.Text = (neto + imp1 + imp2 + imp3).ToString("N2");
        }
    }

    private async Task AnularAsync()
    {
        if (_nc is null) return;
        if (MessageBox.Show("¿Desea anular esta nota de crédito?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        var result = await _anularHandler.HandleAsync(new AnularNotaCreditoCommand(_nc.CodigoNotaCredito, _usuarioActual));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _lblEstado.Text = "Anulado";
        _btnAnular.Enabled = false;
        MessageBox.Show("Nota de crédito anulada.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private async Task GrabarAsync()
    {
        if (!decimal.TryParse(_txtNeto.Text, out var neto)
            || !decimal.TryParse(_txtImpuesto1.Text, out var imp1)
            || !decimal.TryParse(_txtImpuesto2.Text, out var imp2)
            || !decimal.TryParse(_txtImpuesto3.Text, out var imp3))
        {
            MessageBox.Show("Ingrese valores numéricos válidos.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_emitirHandler is null)
        {
            MessageBox.Show("Handler de emisión no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var command = new EmitirNotaCreditoCommand(
            _txtCodigo.Text.Trim(),
            _dtpFecha.Value.Date,
            _txtDocumento.Text.Trim(),
            neto,
            imp1,
            imp2,
            imp3,
            _txtObservacion.Text.Trim());

        var result = await _emitirHandler.HandleAsync(command);
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show("Nota de crédito registrada correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

        if (_onSaved is not null)
            await _onSaved(result.Valor!);

        Close();
    }
}
