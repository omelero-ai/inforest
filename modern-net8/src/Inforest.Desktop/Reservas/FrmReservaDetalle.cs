using Inforest.Application.Reservas;
using Inforest.Domain.Entities.Reservas;

namespace Inforest.Desktop.Reservas;

/// <summary>
/// Formulario de detalle / edición de una reserva.
/// Legacy: frmReservaDetalle.frm
/// Permite crear, modificar y (a través de FrmReserva) anular reservas.
/// BR-RESERVA-001/002/003/004
/// </summary>
public class FrmReservaDetalle : Form
{
    private readonly CrearReservaHandler    _crearHandler;
    private readonly ModificarReservaHandler _modificarHandler;
    private readonly Reserva?               _reserva;  // null = nuevo

    public event Func<Task>? ReservaCreadaOModificada;

    // ── Controls ─────────────────────────────────────────────────────────────
    private readonly TextBox txtCodigo      = new() { Width = 120, ReadOnly = true };
    private readonly TextBox txtCliente     = new() { Width = 280 };
    private readonly TextBox txtTelefono    = new() { Width = 140 };
    private readonly TextBox txtMesero      = new() { Width = 200 };
    private readonly TextBox txtMesa        = new() { Width = 80  };
    private readonly TextBox txtPax         = new() { Width = 80  };
    private readonly TextBox txtObservacion = new() { Width = 380, Multiline = true, Height = 56 };
    private readonly DateTimePicker dtpFecha = new() { Width = 140, Format = DateTimePickerFormat.Short };
    private readonly DateTimePicker dtpHora  = new() { Width = 140, Format = DateTimePickerFormat.Time, ShowUpDown = true };
    private readonly Label lblEstado        = new() { AutoSize = true, ForeColor = System.Drawing.Color.DarkBlue };

    private readonly Button btnGrabar = new() { Text = "Grabar",  Width = 90 };
    private readonly Button btnSalir  = new() { Text = "Salir",   Width = 80 };

    public FrmReservaDetalle(
        CrearReservaHandler crearHandler,
        ModificarReservaHandler modificarHandler,
        Reserva? reserva)
    {
        _crearHandler     = crearHandler;
        _modificarHandler = modificarHandler;
        _reserva          = reserva;

        Text            = reserva is null ? "Nueva Reserva" : $"Reserva {reserva.CodigoReserva}";
        Width           = 540;
        Height          = 440;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox     = false;
        MinimizeBox     = false;
        StartPosition   = FormStartPosition.CenterParent;

        BuildLayout();

        btnGrabar.Click += async (_, _) => await OnGrabarAsync();
        btnSalir.Click  += (_, _) => Close();

        if (reserva is not null)
            CargarDatos(reserva);
    }

    private void BuildLayout()
    {
        int y = 14;
        const int lblW = 100;
        const int lx = 12;
        const int cx = 120;

        void AddRow(string label, Control ctrl)
        {
            Controls.Add(new Label { Text = label, Left = lx, Top = y + 3, Width = lblW, AutoSize = false, TextAlign = ContentAlignment.MiddleRight });
            ctrl.Left = cx; ctrl.Top = y;
            Controls.Add(ctrl);
            y += ctrl.Height + 8;
        }

        AddRow("Código:",      txtCodigo);
        AddRow("Cliente:",     txtCliente);
        AddRow("Teléfono:",    txtTelefono);
        AddRow("Mesero:",      txtMesero);
        AddRow("Mesa:",        txtMesa);
        AddRow("PAX:",         txtPax);
        AddRow("Fecha:",       dtpFecha);
        AddRow("Hora:",        dtpHora);
        AddRow("Observación:", txtObservacion);
        AddRow("Estado:",      lblEstado);

        y += 10;
        btnGrabar.Top = y; btnGrabar.Left = cx;
        btnSalir.Top  = y; btnSalir.Left  = cx + btnGrabar.Width + 10;
        Controls.Add(btnGrabar);
        Controls.Add(btnSalir);

        txtPax.Text  = "1";
        dtpFecha.Value = DateTime.Today;
        dtpHora.Value  = DateTime.Now;
        lblEstado.Text = "Pendiente";
    }

    private void CargarDatos(Reserva r)
    {
        txtCodigo.Text       = r.CodigoReserva;
        txtCliente.Text      = r.NombreCliente ?? r.CodigoCliente ?? string.Empty;
        txtTelefono.Text     = r.Telefono ?? string.Empty;
        txtMesero.Text       = r.CodigoMesero ?? string.Empty;
        txtMesa.Text         = r.CodigoMesa ?? string.Empty;
        txtPax.Text          = r.Pax.ToString();
        dtpFecha.Value       = r.FechaReserva == default ? DateTime.Today : r.FechaReserva;
        dtpHora.Value        = DateTime.Today;   // hora como string en Legacy
        txtObservacion.Text  = r.Observacion ?? string.Empty;
        lblEstado.Text       = r.Estado switch
        {
            EstadoReserva.Atendida => "Atendida",
            EstadoReserva.Anulada  => "Anulada",
            _                      => "Pendiente"
        };

        // Bloquear edición si ya está atendida/anulada
        bool editable = r.Estado == EstadoReserva.Pendiente;
        txtCliente.ReadOnly     = !editable;
        txtTelefono.ReadOnly    = !editable;
        txtMesero.ReadOnly      = !editable;
        txtMesa.ReadOnly        = !editable;
        txtPax.ReadOnly         = !editable;
        txtObservacion.ReadOnly = !editable;
        dtpFecha.Enabled        = editable;
        dtpHora.Enabled         = editable;
        btnGrabar.Enabled       = editable;
    }

    private async Task OnGrabarAsync()
    {
        if (string.IsNullOrWhiteSpace(txtCliente.Text))
        {
            MessageBox.Show("Ingrese el Cliente.", "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            txtCliente.Focus();
            return;
        }

        if (!int.TryParse(txtPax.Text, out int pax) || pax <= 0)
        {
            MessageBox.Show("Ingrese el N° de PAX (mayor a 0).", "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            txtPax.Focus();
            return;
        }

        if (_reserva is null)
        {
            // Nueva reserva
            var command = new CrearReservaCommand(
                FechaReserva  : dtpFecha.Value.Date,
                HoraReserva   : dtpHora.Value.ToString("HH:mm"),
                CodigoCliente : null,
                NombreCliente : txtCliente.Text.Trim(),
                Telefono      : txtTelefono.Text.Trim(),
                CodigoMotivo  : null,
                CodigoMesero  : string.IsNullOrWhiteSpace(txtMesero.Text) ? null : txtMesero.Text.Trim(),
                Pax           : pax,
                CodigoMesa    : string.IsNullOrWhiteSpace(txtMesa.Text) ? null : txtMesa.Text.Trim(),
                Observacion   : txtObservacion.Text.Trim());

            var result = await _crearHandler.HandleAsync(command);
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("Registro Guardado.", "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (ReservaCreadaOModificada is not null) await ReservaCreadaOModificada();
            Close();
        }
        else
        {
            // Modificar
            var command = new ModificarReservaCommand(
                CodigoReserva : _reserva.CodigoReserva,
                FechaReserva  : dtpFecha.Value.Date,
                HoraReserva   : dtpHora.Value.ToString("HH:mm"),
                CodigoCliente : _reserva.CodigoCliente,
                NombreCliente : txtCliente.Text.Trim(),
                Telefono      : txtTelefono.Text.Trim(),
                CodigoMotivo  : null,
                CodigoMesero  : string.IsNullOrWhiteSpace(txtMesero.Text) ? null : txtMesero.Text.Trim(),
                Pax           : pax,
                CodigoMesa    : string.IsNullOrWhiteSpace(txtMesa.Text) ? null : txtMesa.Text.Trim(),
                Observacion   : txtObservacion.Text.Trim());

            var result = await _modificarHandler.HandleAsync(command);
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (ReservaCreadaOModificada is not null) await ReservaCreadaOModificada();
            Close();
        }
    }
}
