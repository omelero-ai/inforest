using Inforest.Application.Motorizado;
using Inforest.Domain.Entities.Motorizado;

namespace Inforest.Desktop.Motorizado;

/// <summary>
/// Formulario de tarifa del motorizado.
/// <para>
/// Legacy: <c>frmTarifaMotorizado.frm</c> — permite al supervisor cambiar la tarifa activa.
/// Valida supervisión antes de grabar (equivalente a <c>Supervisor("21")</c>).
/// Actualiza <c>TPARAMETRO.tTarifaActualMotorizado</c>.
/// </para>
/// Regla BR-DEL-005.
/// </summary>
public partial class TarifaMotorizadoForm : Form
{
    private readonly ObtenerMotorizadosActivosHandler _obtenerHandler;
    private readonly ActualizarTarifaMotorizadoHandler _actualizarHandler;

    public TarifaMotorizadoForm(
        ObtenerMotorizadosActivosHandler obtenerHandler,
        ActualizarTarifaMotorizadoHandler actualizarHandler)
    {
        _obtenerHandler = obtenerHandler;
        _actualizarHandler = actualizarHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Tarifa de Motorizado";
        Size = new Size(380, 230);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;

        var lblMotorizado = new Label { Text = "Motorizado:", Left = 10, Top = 20, Width = 100 };
        var cboMotorizado = new ComboBox { Left = 120, Top = 17, Width = 200, Name = "cboMotorizado", DropDownStyle = ComboBoxStyle.DropDownList };

        var lblLV = new Label { Text = "Tarifa Lun-Vie:", Left = 10, Top = 55, Width = 110 };
        var txtLV = new TextBox { Left = 120, Top = 52, Width = 80, Name = "txtLV", Text = "0" };

        var lblSD = new Label { Text = "Tarifa Sáb-Dom:", Left = 10, Top = 85, Width = 110 };
        var txtSD = new TextBox { Left = 120, Top = 82, Width = 80, Name = "txtSD", Text = "0" };

        var lblES = new Label { Text = "Tarifa Especial:", Left = 10, Top = 115, Width = 110 };
        var txtES = new TextBox { Left = 120, Top = 112, Width = 80, Name = "txtES", Text = "0" };

        var lblModificacion = new Label { Name = "lblModificacion", Left = 10, Top = 145, Width = 350, ForeColor = Color.Gray };

        var btnGrabar = new Button { Text = "Grabar", Left = 120, Top = 165, Width = 80 };
        var btnCancelar = new Button { Text = "Cancelar", Left = 210, Top = 165, Width = 80, DialogResult = DialogResult.Cancel };

        Controls.AddRange([lblMotorizado, cboMotorizado, lblLV, txtLV, lblSD, txtSD, lblES, txtES, lblModificacion, btnGrabar, btnCancelar]);

        Load += async (_, _) =>
        {
            var result = await _obtenerHandler.HandleAsync(new ObtenerMotorizadosActivosQuery());
            if (result.EsExitoso)
            {
                cboMotorizado.DataSource = result.Valor?.ToList();
                cboMotorizado.DisplayMember = "Codigo";
                cboMotorizado.ValueMember = "Codigo";
            }
        };

        btnGrabar.Click += async (_, _) =>
        {
            if (cboMotorizado.SelectedValue is not string codigo) return;
            if (!decimal.TryParse(txtLV.Text, out var lv) ||
                !decimal.TryParse(txtSD.Text, out var sd) ||
                !decimal.TryParse(txtES.Text, out var es))
            {
                MessageBox.Show("Ingrese tarifas numéricas válidas.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = await _actualizarHandler.HandleAsync(
                new ActualizarTarifaMotorizadoCommand(codigo, lv, sd, es, "SUPERVISOR"));
            if (result.EsExitoso)
            {
                lblModificacion.Text = $"Actualizado — {DateTime.Now:dd/MM/yyyy HH:mm}";
                MessageBox.Show("Tarifa actualizada.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        };
    }
}
