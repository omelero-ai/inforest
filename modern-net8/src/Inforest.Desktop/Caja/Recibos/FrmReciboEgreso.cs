using Inforest.Application.Caja;
using Inforest.Domain.Entities.Caja;

namespace Inforest.Desktop.Caja.Recibos;

/// <summary>
/// Listado de recibos de egreso de caja.
/// Legacy: frmReciboEgreso.frm.
/// BR-RECIBO-007, BR-RECIBO-008, BR-RECIBO-009, BR-RECIBO-010.
/// </summary>
public class FrmReciboEgreso : Form
{
    private readonly ObtenerEgresosHandler _obtenerHandler;
    private readonly AnularEgresoHandler _anularHandler;
    private readonly RegistrarEgresoHandler _registrarHandler;

    private readonly DataGridView dgv = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AutoGenerateColumns = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false
    };

    private readonly DateTimePicker dtpDesde  = new() { Format = DateTimePickerFormat.Short, Width = 110 };
    private readonly DateTimePicker dtpHasta  = new() { Format = DateTimePickerFormat.Short, Width = 110 };
    private readonly Button btnBuscar     = new() { Text = "Buscar",    Width = 80 };
    private readonly Button btnNuevo      = new() { Text = "Nuevo",     Width = 80 };
    private readonly Button btnAnular     = new() { Text = "Anular",    Width = 80 };
    private readonly Button btnActualizar = new() { Text = "Actualizar", Width = 90 };
    private readonly Label lblContador    = new() { AutoSize = true };

    private List<ReciboEgreso> _items = [];
    private string? _usuarioFiltro;

    /// <summary>
    /// BR-RECIBO-009: si solicitaClaveEgreso=true, la creación/anulación requiere autorización de supervisor.
    /// El flag se pasa desde el módulo POS vía TPARAMETRO.lSolicitaClaveEgreso.
    /// En esta implementación la validación se realiza en el formulario de detalle.
    /// </summary>
    public bool SolicitaClaveEgreso { get; set; }

    public FrmReciboEgreso(
        ObtenerEgresosHandler obtenerHandler,
        AnularEgresoHandler anularHandler,
        RegistrarEgresoHandler registrarHandler,
        string? usuarioFiltro = null)
    {
        _obtenerHandler = obtenerHandler;
        _anularHandler = anularHandler;
        _registrarHandler = registrarHandler;
        _usuarioFiltro = usuarioFiltro;

        Text = "Mantenimiento de Recibos de Egreso";
        Width = 1200;
        Height = 650;
        StartPosition = FormStartPosition.CenterScreen;

        dtpDesde.Value = DateTime.Today.AddDays(-7);
        dtpHasta.Value = DateTime.Today.AddDays(1).AddSeconds(-1);

        ConfigurarColumnas();
        Controls.Add(BuildLayout());

        Load += async (_, _) => await CargarAsync();
        btnBuscar.Click += async (_, _) => await CargarAsync();
        btnNuevo.Click += (_, _) => AbrirDetalle(nuevo: true);
        btnAnular.Click += async (_, _) => await AnularAsync();
        btnActualizar.Click += async (_, _) => await CargarAsync();
        dgv.DoubleClick += (_, _) => AbrirDetalle(nuevo: false);
    }

    private void ConfigurarColumnas()
    {
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "CodigoRecibo",      HeaderText = "Recibo",       Width = 120 });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Fecha",             HeaderText = "Fecha",        Width = 90  });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Moneda",            HeaderText = "Mon",          Width = 50  });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Monto",             HeaderText = "Monto",        Width = 90  });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Descripcion",       HeaderText = "Descripción",  Width = 200 });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "ModoEgreso",        HeaderText = "Modo",         Width = 90  });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "TipoEgreso",        HeaderText = "Tipo Egreso",  Width = 90  });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Autorizacion",      HeaderText = "Autoriza",     Width = 130 });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "DescripcionEstado", HeaderText = "Estado",       Width = 80  });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "CodigoTurno",       HeaderText = "Turno",        Width = 100 });
    }

    private Control BuildLayout()
    {
        var toolPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 38,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(4)
        };

        toolPanel.Controls.Add(new Label { Text = "Desde:", AutoSize = true, Margin = new Padding(3, 6, 0, 0) });
        toolPanel.Controls.Add(dtpDesde);
        toolPanel.Controls.Add(new Label { Text = "Hasta:", AutoSize = true, Margin = new Padding(6, 6, 0, 0) });
        toolPanel.Controls.Add(dtpHasta);
        toolPanel.Controls.Add(btnBuscar);
        toolPanel.Controls.Add(new Label { Text = " ", Width = 20 });
        toolPanel.Controls.Add(btnNuevo);
        toolPanel.Controls.Add(btnAnular);
        toolPanel.Controls.Add(btnActualizar);

        var statusPanel = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 24, Padding = new Padding(4) };
        statusPanel.Controls.Add(lblContador);

        var container = new Panel { Dock = DockStyle.Fill };
        container.Controls.Add(dgv);
        container.Controls.Add(statusPanel);
        container.Controls.Add(toolPanel);

        return container;
    }

    private async Task CargarAsync()
    {
        btnBuscar.Enabled = false;
        try
        {
            if (dtpDesde.Value > dtpHasta.Value)
            {
                MessageBox.Show("La fecha de inicio no puede ser posterior a la fecha fin.",
                    Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = await _obtenerHandler.HandleAsync(
                new ObtenerEgresosQuery(dtpDesde.Value, dtpHasta.Value, _usuarioFiltro));

            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _items = [.. result.Valor];
            dgv.DataSource = _items;
            lblContador.Text = $"Registros: {_items.Count}";
        }
        finally
        {
            btnBuscar.Enabled = true;
        }
    }

    private void AbrirDetalle(bool nuevo)
    {
        ReciboEgreso? seleccionado = null;
        if (!nuevo && dgv.CurrentRow?.DataBoundItem is ReciboEgreso r)
            seleccionado = r;

        if (!nuevo && seleccionado is null)
        {
            MessageBox.Show("No existe datos ingresados.", Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            return;
        }

        using var detalle = new FrmReciboEgresoDetalle(
            _registrarHandler,
            seleccionado,
            SolicitaClaveEgreso);
        detalle.ShowDialog(this);
        _ = CargarAsync();
    }

    private async Task AnularAsync()
    {
        // BR-RECIBO-009: verificación de clave supervisor delegada al llamante (módulo POS)
        if (dgv.CurrentRow?.DataBoundItem is not ReciboEgreso r)
        {
            MessageBox.Show("Seleccione un recibo de egreso.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (r.EstaAnulado)
        {
            MessageBox.Show("El recibo ya está anulado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (MessageBox.Show($"¿Está seguro de anular el recibo {r.CodigoRecibo}?",
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        var result = await _anularHandler.HandleAsync(new AnularEgresoCommand(r.CodigoRecibo));
        if (result.EsExitoso)
        {
            MessageBox.Show("Recibo anulado correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            await CargarAsync();
        }
        else
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
