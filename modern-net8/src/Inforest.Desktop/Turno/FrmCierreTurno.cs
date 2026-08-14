using Inforest.Application.Turno;

namespace Inforest.Desktop.Turno;

    public class FrmCierreTurno : Form
    {
        private readonly CerrarTurnoHandler? _cerrarTurnoHandler;
        private readonly string _codigoTurno;
        private readonly string _codigoCaja;

        private Label lblResumen = null!;
        private TextBox txtMontoFinal = null!;
        private Button btnCerrar = null!;
        private Button btnCancelar = null!;

        public FrmCierreTurno(string codigoTurno, string codigoCaja, CerrarTurnoHandler? cerrarTurnoHandler = null)
        {
            _codigoTurno = codigoTurno;
            _codigoCaja = codigoCaja;
            _cerrarTurnoHandler = cerrarTurnoHandler;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "Cierre de Turno";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(420, 210);

            lblResumen = new Label
            {
                Left = 20,
                Top = 20,
                Width = 360,
                Height = 60,
                Text = "Resumen del turno: " + _codigoTurno + Environment.NewLine + "Ingrese el monto final contado para el cierre."
            };
            var lblMonto = new Label { Left = 20, Top = 100, Width = 100, Text = "Monto Final:" };
            txtMontoFinal = new TextBox { Left = 130, Top = 98, Width = 240, Name = "txtMontoFinal", Text = "0.00" };
            btnCerrar = new Button { Left = 130, Top = 145, Width = 110, Text = "Cerrar", Name = "btnCerrar" };
            btnCancelar = new Button { Left = 260, Top = 145, Width = 110, Text = "Cancelar", Name = "btnCancelar" };

            btnCancelar.Click += (_, _) => Close();
            btnCerrar.Click += async (_, _) => await CerrarAsync();

            Controls.AddRange([lblResumen, lblMonto, txtMontoFinal, btnCerrar, btnCancelar]);
        }

        private async Task CerrarAsync()
        {
            if (_cerrarTurnoHandler is null)
            {
                MessageBox.Show("Handler de cierre no configurado.", "Turno", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(txtMontoFinal.Text, out var montoFinal))
            {
                MessageBox.Show("Ingrese un monto final válido.", "Turno", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = await _cerrarTurnoHandler.HandleAsync(new CerrarTurnoCommand(_codigoTurno, _codigoCaja, montoFinal, CodigoUsuario: Environment.UserName));
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Turno", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
