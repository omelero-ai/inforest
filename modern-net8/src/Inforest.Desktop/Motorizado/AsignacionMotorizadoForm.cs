using Inforest.Application.Configuracion;

namespace Inforest.Desktop.Motorizado;

/// <summary>
/// Formulario informativo del monto máximo de asignación por motorizado.
/// <para>
/// Legacy: <c>frmAsignacionMotorizado.frm</c> — solo consulta
/// <c>TPARAMETRO.nAsignacionMotorizado</c> y lo muestra en pantalla.
/// </para>
/// Regla BR-DEL-005.
/// </summary>
public partial class AsignacionMotorizadoForm : Form
{
    private readonly ObtenerConfiguracionSistemaHandler _configuracionHandler;

    public AsignacionMotorizadoForm(ObtenerConfiguracionSistemaHandler configuracionHandler)
    {
        _configuracionHandler = configuracionHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Asignación Motorizado";
        Size = new Size(420, 150);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;

        var lblTitulo = new Label
        {
            Text = "Asignación Máxima a Motorizados S/.",
            Left = 15,
            Top = 20,
            Width = 260,
            Font = new Font(Font, FontStyle.Bold)
        };

        var lblAsignacion = new Label
        {
            Name = "lblAsignacion",
            Text = "0.00",
            Left = 285,
            Top = 20,
            Width = 100,
            Font = new Font(Font, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleRight
        };

        var btnCerrar = new Button
        {
            Text = "Cerrar",
            Left = 155,
            Top = 60,
            Width = 100,
            DialogResult = DialogResult.OK
        };

        Controls.AddRange([lblTitulo, lblAsignacion, btnCerrar]);

        Load += async (_, _) =>
        {
            var result = await _configuracionHandler.HandleAsync(new ObtenerConfiguracionSistemaQuery());
            if (result.EsExitoso && result.Valor is not null)
            {
                lblAsignacion.Text = result.Valor.nAsignacionMotorizado.ToString("###,##0.00");
            }
        };
    }
}
