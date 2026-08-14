using Inforest.Application.Almacen;
using Inforest.Domain.Entities.Almacen;

namespace Inforest.Desktop.Almacen;

/// <summary>
/// Lista de requerimientos de almacén pendientes de importar al POS.
/// Legacy: frmImportacionRequerimientos.frm.
/// BR-IMPORT-001: Solo requerimientos en estado '02' y área habilitada.
/// BR-IMPORT-002: Validación de duplicados antes de importar.
/// BR-IMPORT-003: Verificación de enlace de producto con INFOREST.
/// BR-IMPORT-004: Marcado del requerimiento como procesado.
/// </summary>
public class FrmImportacionRequerimientos : Form
{
    private readonly ObtenerRequerimientosPendientesHandler _obtenerHandler;
    private readonly ImportarRequerimientoHandler _importarHandler;
    private readonly string _codigoCaja;
    private readonly string _codigoTurno;
    private readonly string _codigoSalon;
    private readonly string _codigoUsuario;
    private readonly string _baseDatosInforest;
    private readonly DateTime _fechaDiaContable;

    private readonly DateTimePicker dtpFechaInicio = new() { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
    private readonly DateTimePicker dtpFechaFin    = new() { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
    private readonly ComboBox cboTipoPedido = new() { Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly DataGridView dgv = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AutoGenerateColumns = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false
    };
    private readonly Button btnBuscar   = new() { Text = "Buscar",     Width = 90 };
    private readonly Button btnDetalle  = new() { Text = "Ver Detalle", Width = 100 };
    private readonly Button btnImportar = new() { Text = "Importar",   Width = 90 };
    private readonly Button btnSalir    = new() { Text = "Salir",      Width = 80 };
    private readonly Label lblEstado = new() { AutoSize = true, ForeColor = Color.DarkBlue };

    private List<RequerimientoAlmacen> _requerimientos = [];
    private ObtenerDetalleRequerimientoHandler? _detalleHandler;

    public FrmImportacionRequerimientos(
        ObtenerRequerimientosPendientesHandler obtenerHandler,
        ImportarRequerimientoHandler importarHandler,
        string codigoCaja,
        string codigoTurno,
        string codigoSalon,
        string codigoUsuario,
        DateTime fechaDiaContable,
        string baseDatosInforest = "INFOREST")
    {
        _obtenerHandler = obtenerHandler;
        _importarHandler = importarHandler;
        _codigoCaja = codigoCaja;
        _codigoTurno = codigoTurno;
        _codigoSalon = codigoSalon;
        _codigoUsuario = codigoUsuario;
        _fechaDiaContable = fechaDiaContable;
        _baseDatosInforest = baseDatosInforest;

        Text = "Importacion de Pedidos";
        Width = 1000;
        Height = 650;
        StartPosition = FormStartPosition.CenterScreen;

        InicializarCombos();
        ConfigurarColumnas();
        Controls.Add(BuildLayout());

        btnBuscar.Click += async (_, _) => await BuscarAsync();
        btnDetalle.Click += (_, _) => VerDetalle();
        btnImportar.Click += async (_, _) => await ImportarAsync();
        btnSalir.Click += (_, _) => Close();
    }

    /// <summary>
    /// Permite inyectar el handler de detalle para abrir la ventana de detalle.
    /// </summary>
    public void AsignarHandlerDetalle(ObtenerDetalleRequerimientoHandler detalleHandler)
        => _detalleHandler = detalleHandler;

    // ── Search ───────────────────────────────────────────────────────────────

    private async Task BuscarAsync()
    {
        lblEstado.Text = "Buscando...";
        var result = await _obtenerHandler.HandleAsync(
            new ObtenerRequerimientosPendientesQuery(dtpFechaInicio.Value, dtpFechaFin.Value));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            lblEstado.Text = string.Empty;
            return;
        }

        _requerimientos = [.. result.Valor!];
        dgv.DataSource = _requerimientos.Select(r => new
        {
            r.Rq,
            Fecha = r.Fecha.ToString("dd/MM/yyyy"),
            r.Area,
            r.Responsable,
            r.CodEstado,
            Pedido = r.CodigoPedido ?? string.Empty
        }).ToList();

        lblEstado.Text = $"Registros: {_requerimientos.Count}";
    }

    // ── Detail ───────────────────────────────────────────────────────────────

    private void VerDetalle()
    {
        if (_requerimientos.Count == 0)
        {
            MessageBox.Show("No existen datos ingresados.", "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (dgv.CurrentRow?.Index is null or < 0) return;
        var requerimiento = _requerimientos[dgv.CurrentRow.Index];

        if (_detalleHandler is null)
        {
            MessageBox.Show("El visualizador de detalle no está disponible.", "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var frmDetalle = new FrmImportacionRequerimientoDetalle(
            _detalleHandler, requerimiento.Rq, _baseDatosInforest);
        frmDetalle.ShowDialog(this);
    }

    // ── Import ───────────────────────────────────────────────────────────────

    private async Task ImportarAsync()
    {
        if (cboTipoPedido.SelectedValue is not string tipoPedido || string.IsNullOrEmpty(tipoPedido))
        {
            MessageBox.Show("Seleccionar un Canal de Venta.", "Validación",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_requerimientos.Count == 0)
        {
            MessageBox.Show("No existe datos a importar.", "Información",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (dgv.CurrentRow?.Index is null or < 0) return;
        var requerimiento = _requerimientos[dgv.CurrentRow.Index];

        // BR-IMPORT-001: verificación previa
        if (!requerimiento.PuedeImportarse())
        {
            MessageBox.Show($"El requerimiento '{requerimiento.Rq}' no puede importarse (ya importado o estado incorrecto).",
                "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        btnImportar.Enabled = false;
        lblEstado.Text = "Importando...";

        var cmd = new ImportarRequerimientoCommand(
            Rq: requerimiento.Rq,
            CodigoCaja: _codigoCaja,
            CodigoTurno: _codigoTurno,
            CodigoSalon: _codigoSalon,
            CodigoUsuario: _codigoUsuario,
            TipoPedido: tipoPedido,
            FechaDiaContable: _fechaDiaContable,
            BaseDatosInforest: _baseDatosInforest);

        var result = await _importarHandler.HandleAsync(cmd);

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "Error al importar",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        else
        {
            MessageBox.Show($"Pedido '{result.Valor!.CodigoPedido}' Generado Correctamente",
                "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        btnImportar.Enabled = true;
        lblEstado.Text = string.Empty;
        await BuscarAsync();
    }

    // ── Layout helpers ───────────────────────────────────────────────────────

    private void InicializarCombos()
    {
        cboTipoPedido.Items.Add(new { Codigo = "01", Descripcion = "Salón" });
        cboTipoPedido.Items.Add(new { Codigo = "02", Descripcion = "Delivery" });
        cboTipoPedido.Items.Add(new { Codigo = "03", Descripcion = "Para llevar" });
        cboTipoPedido.DisplayMember = "Descripcion";
        cboTipoPedido.ValueMember = "Codigo";
        if (cboTipoPedido.Items.Count > 0) cboTipoPedido.SelectedIndex = 0;
    }

    private void ConfigurarColumnas()
    {
        dgv.Columns.AddRange(
            new DataGridViewTextBoxColumn { DataPropertyName = "Rq",          HeaderText = "Requerimiento", Width = 120 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Fecha",        HeaderText = "Fecha",         Width = 90  },
            new DataGridViewTextBoxColumn { DataPropertyName = "Area",         HeaderText = "Área",          Width = 150 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Responsable",  HeaderText = "Responsable",   Width = 150 },
            new DataGridViewTextBoxColumn { DataPropertyName = "CodEstado",    HeaderText = "Estado",        Width = 70  },
            new DataGridViewTextBoxColumn { DataPropertyName = "Pedido",       HeaderText = "Pedido",        Width = 100 }
        );
    }

    private Control BuildLayout()
    {
        var pnlTop = new Panel { Dock = DockStyle.Top, Height = 70, Padding = new Padding(5) };

        pnlTop.Controls.Add(new Label { Text = "Del:", AutoSize = true, Location = new Point(5, 8) });
        dtpFechaInicio.Location = new Point(40, 5);
        pnlTop.Controls.Add(dtpFechaInicio);

        pnlTop.Controls.Add(new Label { Text = "Al:", AutoSize = true, Location = new Point(225, 8) });
        dtpFechaFin.Location = new Point(250, 5);
        pnlTop.Controls.Add(dtpFechaFin);

        pnlTop.Controls.Add(new Label { Text = "Canal:", AutoSize = true, Location = new Point(5, 40) });
        cboTipoPedido.Location = new Point(55, 37);
        pnlTop.Controls.Add(cboTipoPedido);

        btnBuscar.Location = new Point(270, 5);
        pnlTop.Controls.Add(btnBuscar);

        var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(5) };
        btnDetalle.Location  = new Point(5,  5);
        btnImportar.Location = new Point(115, 5);
        btnSalir.Location    = new Point(215, 5);
        lblEstado.Location   = new Point(310, 10);
        pnlBottom.Controls.AddRange([btnDetalle, btnImportar, btnSalir, lblEstado]);

        var pnlGrid = new Panel { Dock = DockStyle.Fill };
        pnlGrid.Controls.Add(dgv);

        var container = new Panel { Dock = DockStyle.Fill };
        container.Controls.Add(pnlGrid);
        container.Controls.Add(pnlBottom);
        container.Controls.Add(pnlTop);
        return container;
    }
}
