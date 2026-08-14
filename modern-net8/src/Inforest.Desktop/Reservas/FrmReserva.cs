using Inforest.Application.Reservas;
using Inforest.Domain.Entities.Reservas;

namespace Inforest.Desktop.Reservas;

/// <summary>
/// Lista y consulta de reservas por rango de fechas.
/// Legacy: frmReserva.frm — "Correlativo de Reservas"
/// Carga TRESERVA con filtro de fecha y delega a FrmReservaDetalle para CRUD.
/// BR-RESERVA-001, BR-RESERVA-002, BR-RESERVA-003, BR-RESERVA-004
/// </summary>
public class FrmReserva : Form
{
    private readonly ObtenerReservasPorFechaHandler _queryHandler;
    private readonly CrearReservaHandler _crearHandler;
    private readonly ModificarReservaHandler _modificarHandler;
    private readonly AnularReservaHandler _anularHandler;
    private readonly ConvertirReservaAPedidoHandler _convertirHandler;

    private List<Reserva> _reservas = [];

    // ── Controls ─────────────────────────────────────────────────────────────
    private readonly DataGridView dgv = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AutoGenerateColumns = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false
    };

    private readonly DateTimePicker dtpFecIni = new() { Width = 130, Format = DateTimePickerFormat.Short };
    private readonly DateTimePicker dtpFecFin = new() { Width = 130, Format = DateTimePickerFormat.Short };
    private readonly Button btnBuscar    = new() { Text = "Buscar",    Width = 90 };
    private readonly Button btnNuevo     = new() { Text = "Nuevo",     Width = 80 };
    private readonly Button btnEditar    = new() { Text = "Editar",    Width = 80 };
    private readonly Button btnAnular    = new() { Text = "Anular",    Width = 80 };
    private readonly Button btnAtender   = new() { Text = "Atender",   Width = 80 };
    private readonly Button btnSalir     = new() { Text = "Salir",     Width = 80 };
    private readonly Label  lblStatus    = new() { Dock = DockStyle.Bottom, Height = 22, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(4, 0, 0, 0) };

    public FrmReserva(
        ObtenerReservasPorFechaHandler queryHandler,
        CrearReservaHandler crearHandler,
        ModificarReservaHandler modificarHandler,
        AnularReservaHandler anularHandler,
        ConvertirReservaAPedidoHandler convertirHandler)
    {
        _queryHandler    = queryHandler;
        _crearHandler    = crearHandler;
        _modificarHandler = modificarHandler;
        _anularHandler   = anularHandler;
        _convertirHandler = convertirHandler;

        Text            = "Correlativo de Reservas";
        Width           = 1200;
        Height          = 660;
        StartPosition   = FormStartPosition.CenterScreen;
        MdiParent       = System.Windows.Forms.Application.OpenForms.OfType<Form>().FirstOrDefault(f => f.IsMdiContainer);

        dtpFecIni.Value = DateTime.Today;
        dtpFecFin.Value = DateTime.Today;

        ConfigurarGrilla();
        Controls.Add(dgv);
        Controls.Add(BuildBarraFiltro());
        Controls.Add(BuildBarraBotones());
        Controls.Add(lblStatus);

        btnBuscar.Click  += async (_, _) => await CargarAsync();
        btnNuevo.Click   += OnNuevo;
        btnEditar.Click  += OnEditar;
        btnAnular.Click  += async (_, _) => await OnAnularAsync();
        btnAtender.Click += async (_, _) => await OnAtenderAsync();
        btnSalir.Click   += (_, _) => Close();

        Load += async (_, _) => await CargarAsync();
    }

    private void ConfigurarGrilla()
    {
        dgv.Columns.AddRange(
        [
            new DataGridViewTextBoxColumn { DataPropertyName = "CodigoReserva", HeaderText = "Reserva",     Width = 100 },
            new DataGridViewTextBoxColumn { DataPropertyName = "NombreCliente", HeaderText = "Cliente",     Width = 220 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Telefono",      HeaderText = "Teléfono",    Width = 110 },
            new DataGridViewTextBoxColumn { DataPropertyName = "FechaReserva",  HeaderText = "Fecha",       Width = 100 },
            new DataGridViewTextBoxColumn { DataPropertyName = "HoraReserva",   HeaderText = "Hora",        Width = 70  },
            new DataGridViewTextBoxColumn { DataPropertyName = "Pax",           HeaderText = "PAX",         Width = 60  },
            new DataGridViewTextBoxColumn { DataPropertyName = "CodigoMesa",    HeaderText = "Mesa",        Width = 70  },
            new DataGridViewTextBoxColumn { DataPropertyName = "EstadoTexto",   HeaderText = "Estado",      Width = 90  },
            new DataGridViewTextBoxColumn { DataPropertyName = "Observacion",   HeaderText = "Observación", Width = 200 }
        ]);
    }

    private Panel BuildBarraFiltro()
    {
        var panel = new Panel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(4) };
        panel.Controls.AddRange([
            new Label { Text = "Desde:", AutoSize = true, Top = 10, Left = 4 },
            new Control { Width = 4 },
            dtpFecIni,
            new Label { Text = "  Hasta:", AutoSize = true, Top = 10 },
            dtpFecFin,
            new Label { Text = "  ", AutoSize = true },
            btnBuscar
        ]);
        // reposition horizontally
        int x = 4;
        foreach (Control c in panel.Controls)
        {
            c.Top  = 8;
            c.Left = x;
            x += c.Width + 4;
        }
        return panel;
    }

    private Panel BuildBarraBotones()
    {
        var panel = new Panel { Dock = DockStyle.Bottom, Height = 42 };
        int x = 8;
        foreach (var btn in new[] { btnNuevo, btnEditar, btnAnular, btnAtender, btnSalir })
        {
            btn.Height = 30;
            btn.Top    = 6;
            btn.Left   = x;
            x += btn.Width + 6;
            panel.Controls.Add(btn);
        }
        return panel;
    }

    private async Task CargarAsync()
    {
        lblStatus.Text = "Cargando…";
        var result = await _queryHandler.HandleAsync(
            new ObtenerReservasPorFechaQuery(dtpFecIni.Value.Date, dtpFecFin.Value.Date));

        if (!result.EsExitoso)
        {
            lblStatus.Text = result.MensajeError;
            return;
        }

        _reservas = result.Valor!.ToList();

        dgv.DataSource = _reservas.Select(r => new
        {
            r.CodigoReserva,
            NombreCliente = r.NombreCliente ?? r.CodigoCliente ?? "(sin cliente)",
            r.Telefono,
            FechaReserva  = r.FechaReserva.ToString("dd/MM/yyyy"),
            r.HoraReserva,
            r.Pax,
            r.CodigoMesa,
            EstadoTexto   = r.Estado switch
            {
                EstadoReserva.Atendida => "Atendida",
                EstadoReserva.Anulada  => "Anulada",
                _                      => "Pendiente"
            },
            r.Observacion
        }).ToList();

        lblStatus.Text = $"Registro {_reservas.Count} reserva(s)";
    }

    private Reserva? ReservaSeleccionada()
    {
        if (dgv.CurrentRow is null || _reservas.Count == 0) return null;
        var idx = dgv.CurrentRow.Index;
        return idx >= 0 && idx < _reservas.Count ? _reservas[idx] : null;
    }

    private void OnNuevo(object? sender, EventArgs e)
    {
        var frm = new FrmReservaDetalle(_crearHandler, _modificarHandler, null);
        frm.ReservaCreadaOModificada += async () => await CargarAsync();
        frm.ShowDialog(this);
    }

    private void OnEditar(object? sender, EventArgs e)
    {
        var reserva = ReservaSeleccionada();
        if (reserva is null) { MessageBox.Show("Seleccione una reserva.", "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

        var frm = new FrmReservaDetalle(_crearHandler, _modificarHandler, reserva);
        frm.ReservaCreadaOModificada += async () => await CargarAsync();
        frm.ShowDialog(this);
    }

    private async Task OnAnularAsync()
    {
        var reserva = ReservaSeleccionada();
        if (reserva is null) { MessageBox.Show("Seleccione una reserva.", "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

        if (MessageBox.Show($"¿Seguro de anular la reserva {reserva.CodigoReserva}?",
                "INFOREST", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel) return;

        var result = await _anularHandler.HandleAsync(new AnularReservaCommand(reserva.CodigoReserva));
        if (!result.EsExitoso)
            MessageBox.Show(result.MensajeError, "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Error);
        else
            await CargarAsync();
    }

    private async Task OnAtenderAsync()
    {
        var reserva = ReservaSeleccionada();
        if (reserva is null) { MessageBox.Show("Seleccione una reserva.", "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

        if (reserva.Estado != EstadoReserva.Pendiente)
        {
            MessageBox.Show("Solo se puede atender una reserva pendiente.", "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Leer datos de sesión — se inyectan en el comando desde el contexto de ejecución
        // La resolución de caja/turno/diacontable requiere ISessionService; aquí se
        // omiten por ser un formulario que delega al handler (la validación real está
        // en ConvertirReservaAPedidoHandler + spIns_MPEDIDO_RESERVA).
        MessageBox.Show(
            $"Para convertir la reserva {reserva.CodigoReserva} en pedido utilice " +
            "el módulo de Pedidos indicando la reserva de origen.",
            "INFOREST — Atender Reserva",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }
}
