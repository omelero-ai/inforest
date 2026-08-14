using Inforest.Application.Motorizado;
using MotorizadoEntity = Inforest.Domain.Entities.Motorizado.Motorizado;
using PedidoReasignacionMotorizado = Inforest.Domain.Entities.Delivery.PedidoReasignacionMotorizado;

namespace Inforest.Desktop.Motorizado;

/// <summary>
/// Formulario de reasignación de motorizado para pedidos delivery.
/// <para>
/// Legacy: <c>frmReasignacionMotorizado.frm</c> — busca pedidos delivery
/// activos, permite asignar/desasignar <c>tMotorizadoN</c> y controla el
/// monto máximo por motorizado vía <c>TPARAMETRO.nAsignacionMotorizado</c>.
/// </para>
/// Reglas BR-DEL-002 y BR-DEL-005.
/// </summary>
public partial class ReasignacionMotorizadoForm : Form
{
    private readonly ObtenerPedidosReasignacionHandler _pedidosHandler;
    private readonly ObtenerMotorizadosActivosHandler _motorizadosHandler;
    private readonly AsignarReasignacionMotorizadoHandler _asignarHandler;
    private readonly DesasignarReasignacionMotorizadoHandler _desasignarHandler;
    private DateTimePicker? _dtpFecIni;
    private DateTimePicker? _dtpHorIni;
    private DateTimePicker? _dtpFecFin;
    private DateTimePicker? _dtpHorFin;
    private TextBox? _txtCriterio;
    private DataGridView? _grid;

    public ReasignacionMotorizadoForm(
        ObtenerPedidosReasignacionHandler pedidosHandler,
        ObtenerMotorizadosActivosHandler motorizadosHandler,
        AsignarReasignacionMotorizadoHandler asignarHandler,
        DesasignarReasignacionMotorizadoHandler desasignarHandler)
    {
        _pedidosHandler = pedidosHandler;
        _motorizadosHandler = motorizadosHandler;
        _asignarHandler = asignarHandler;
        _desasignarHandler = desasignarHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Reasignación Motorizado";
        Size = new Size(900, 620);
        StartPosition = FormStartPosition.CenterScreen;

        var grpFiltros = new GroupBox
        {
            Text = "Opciones",
            Left = 10,
            Top = 10,
            Width = 860,
            Height = 140
        };

        var lblFechaIni = new Label { Text = "Fecha Inicial:", Left = 15, Top = 30, Width = 100 };
        _dtpFecIni = new DateTimePicker { Left = 120, Top = 25, Width = 130, Value = DateTime.Today, Name = "dtpFecIni" };
        _dtpHorIni = new DateTimePicker
        {
            Left = 260,
            Top = 25,
            Width = 90,
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "HH:mm",
            ShowUpDown = true,
            Value = DateTime.Today
        };

        var lblFechaFin = new Label { Text = "Fecha Final:", Left = 15, Top = 65, Width = 100 };
        _dtpFecFin = new DateTimePicker { Left = 120, Top = 60, Width = 130, Value = DateTime.Today, Name = "dtpFecFin" };
        _dtpHorFin = new DateTimePicker
        {
            Left = 260,
            Top = 60,
            Width = 90,
            Format = DateTimePickerFormat.Custom,
            CustomFormat = "HH:mm",
            ShowUpDown = true,
            Value = DateTime.Today.AddHours(23).AddMinutes(59)
        };

        var lblPedido = new Label { Text = "Pedido:", Left = 380, Top = 30, Width = 60 };
        _txtCriterio = new TextBox { Left = 445, Top = 25, Width = 180, Name = "txtCriterio" };

        var lblMotorizado = new Label { Text = "Motorizado:", Left = 350, Top = 65, Width = 85 };
        var cboMotorizado = new ComboBox
        {
            Left = 445,
            Top = 60,
            Width = 250,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Name = "cboMotorizado"
        };

        var btnBuscar = new Button { Text = "Buscar", Left = 715, Top = 23, Width = 110 };
        var btnAsignar = new Button { Text = "Asignar Motorizado", Left = 715, Top = 58, Width = 110, Height = 30 };

        grpFiltros.Controls.AddRange([
            lblFechaIni, _dtpFecIni, _dtpHorIni,
            lblFechaFin, _dtpFecFin, _dtpHorFin,
            lblPedido, _txtCriterio, lblMotorizado, cboMotorizado,
            btnBuscar, btnAsignar
        ]);

        _grid = new DataGridView
        {
            Name = "grdPedidos",
            Left = 10,
            Top = 160,
            Width = 700,
            Height = 400,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoGenerateColumns = false
        };
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(PedidoReasignacionMotorizado.CodigoPedido),
            HeaderText = "Pedido",
            Width = 150
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(PedidoReasignacionMotorizado.MotorizadoDescripcion),
            HeaderText = "Motorizado",
            Width = 350
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(PedidoReasignacionMotorizado.TarifaMotorizado),
            HeaderText = "Tarifa",
            Width = 120,
            DefaultCellStyle = new DataGridViewCellStyle { Format = "N2" }
        });

        var pnlAcciones = new Panel
        {
            Left = 720,
            Top = 160,
            Width = 150,
            Height = 400,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        var btnDesasignar = new Button { Text = "Desasignar", Left = 5, Top = 10, Width = 135, Height = 50 };
        var btnSalir = new Button { Text = "Salir", Left = 5, Top = 320, Width = 135, Height = 40 };
        pnlAcciones.Controls.AddRange([btnDesasignar, btnSalir]);

        Controls.AddRange([grpFiltros, _grid, pnlAcciones]);

        Load += async (_, _) =>
        {
            await CargarMotorizadosAsync(cboMotorizado);
            await RecargarGridAsync();
        };

        btnBuscar.Click += async (_, _) => await RecargarGridAsync();
        btnAsignar.Click += async (_, _) => await AsignarAsync(cboMotorizado, _grid);
        btnDesasignar.Click += async (_, _) => await DesasignarAsync(_grid);
        btnSalir.Click += (_, _) => Close();
    }

    private async Task CargarMotorizadosAsync(ComboBox combo)
    {
        var result = await _motorizadosHandler.HandleAsync(new ObtenerMotorizadosActivosQuery());
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "Motorizados", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        combo.DataSource = result.Valor?.ToList();
        combo.DisplayMember = nameof(MotorizadoEntity.Codigo);
        combo.ValueMember = nameof(MotorizadoEntity.Codigo);
    }

    private async Task BuscarAsync(
        DateTimePicker fechaInicio,
        DateTimePicker horaInicio,
        DateTimePicker fechaFin,
        DateTimePicker horaFin,
        TextBox criterio,
        DataGridView grid)
    {
        var inicio = fechaInicio.Value.Date.Add(horaInicio.Value.TimeOfDay);
        var fin = fechaFin.Value.Date.Add(horaFin.Value.TimeOfDay);

        var result = await _pedidosHandler.HandleAsync(
            new ObtenerPedidosReasignacionQuery(inicio, fin, criterio.Text.Trim()));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "Reasignación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        grid.DataSource = result.Valor?.ToList();
    }

    private async Task AsignarAsync(ComboBox combo, DataGridView grid)
    {
        if (grid.SelectedRows.Count == 0)
            return;

        if (combo.SelectedValue is not string codigoMotorizado)
            return;

        var pedido = (PedidoReasignacionMotorizado)grid.SelectedRows[0].DataBoundItem;
        var result = await _asignarHandler.HandleAsync(
            new AsignarReasignacionMotorizadoCommand(pedido.CodigoPedido, codigoMotorizado, "SISTEMA"));

        if (!result.EsExitoso && result.CodigoError == "REASIGNACION_SUPERA_MONTO_MAXIMO")
        {
            var confirmar = MessageBox.Show(
                "Con la asignación de este pedido se supera el monto máximo por motorizado. ¿Desea continuar?",
                "Reasignación",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirmar == DialogResult.Yes)
            {
                result = await _asignarHandler.HandleAsync(
                    new AsignarReasignacionMotorizadoCommand(
                        pedido.CodigoPedido,
                        codigoMotorizado,
                        "SISTEMA",
                        AutorizarTarifaExtra: true));
            }
        }

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "Reasignación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        await RecargarGridAsync();
    }

    private async Task DesasignarAsync(DataGridView grid)
    {
        if (grid.SelectedRows.Count == 0)
            return;

        var pedido = (PedidoReasignacionMotorizado)grid.SelectedRows[0].DataBoundItem;
        var result = await _desasignarHandler.HandleAsync(
            new DesasignarReasignacionMotorizadoCommand(pedido.CodigoPedido));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "Reasignación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        await RecargarGridAsync();
    }

    private async Task RecargarGridAsync()
    {
        if (_dtpFecIni is null || _dtpHorIni is null || _dtpFecFin is null || _dtpHorFin is null || _txtCriterio is null || _grid is null)
            return;

        await BuscarAsync(_dtpFecIni, _dtpHorIni, _dtpFecFin, _dtpHorFin, _txtCriterio, _grid);
    }
}
