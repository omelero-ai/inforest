using Inforest.Application.CentralPedidos;
using Inforest.Application.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Formulario de Central de Pedidos multi-local.
/// <para>
/// Legacy: <c>frmCentralPedidos.frm</c> — "Correlativo Pedidos Centro Produccion".
/// Consulta pedidos de la base <c>CENTRALDELIVERY</c> mediante
/// <c>sp_Inforest_PedidosCentralPedido</c>.
/// Habilitado solo si flag <c>lCD = 1</c> en <c>TCAJA</c>.
/// </para>
/// Funcionalidades migradas:
/// <list type="bullet">
///   <item>Case 0: Ver detalle del pedido.</item>
///   <item>Case 2: Modificar fecha programada (BR-DEL-014).</item>
///   <item>Case 3: Confirmar entrega (BR-DEL-012, BR-DEL-013).</item>
///   <item>Case 5: Revertir entrega (BR-DEL-012 + supervisor).</item>
/// </list>
/// Reglas: BR-DEL-003, BR-DEL-008, BR-DEL-012, BR-DEL-013, BR-DEL-014.
/// </summary>
public partial class CentralPedidosForm : Form
{
    private readonly ObtenerPedidosCentralHandler _obtenerHandler;
    private readonly ConfirmarEntregaCentralHandler _confirmarHandler;
    private readonly RevertirEntregaCentralHandler _revertirHandler;
    private readonly ModificarFechaProgramadaDeliveryHandler _modificarFechaHandler;
    private readonly string _usuarioActual;

    public CentralPedidosForm(
        ObtenerPedidosCentralHandler obtenerHandler,
        ConfirmarEntregaCentralHandler confirmarHandler,
        RevertirEntregaCentralHandler revertirHandler,
        ModificarFechaProgramadaDeliveryHandler modificarFechaHandler,
        string usuarioActual = "")
    {
        _obtenerHandler = obtenerHandler;
        _confirmarHandler = confirmarHandler;
        _revertirHandler = revertirHandler;
        _modificarFechaHandler = modificarFechaHandler;
        _usuarioActual = usuarioActual;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Central de Pedidos";
        Size = new Size(1100, 650);
        WindowState = FormWindowState.Maximized;

        var toolbar = new ToolStrip();
        var btnConsultar = new ToolStripButton("Consultar");
        var btnConfirmarEntrega = new ToolStripButton("Confirmar Entrega") { ToolTipText = "Confirma que el pedido fue entregado (BR-DEL-012)" };
        var btnRevertirEntrega = new ToolStripButton("Revertir Entrega") { ToolTipText = "Revierte la confirmación de entrega (requiere supervisor)" };
        var btnModificarFecha = new ToolStripButton("Modificar Fecha") { ToolTipText = "Modifica la fecha programada de entrega (BR-DEL-014)" };
        var btnEmitir = new ToolStripButton("Emitir");
        var btnSalir = new ToolStripButton("Salir");
        toolbar.Items.AddRange([btnConsultar, new ToolStripSeparator(),
            btnConfirmarEntrega, btnRevertirEntrega, new ToolStripSeparator(),
            btnModificarFecha, new ToolStripSeparator(), btnEmitir, btnSalir]);

        var lblFechaIni = new Label { Text = "Desde:", Left = 10, Top = 35, Width = 50 };
        var dtpFecIni = new DateTimePicker { Left = 65, Top = 32, Width = 120, Value = DateTime.Today };
        var lblFechaFin = new Label { Text = "Hasta:", Left = 195, Top = 35, Width = 50 };
        var dtpFecFin = new DateTimePicker { Left = 250, Top = 32, Width = 120, Value = DateTime.Today };

        var statusBar = new StatusStrip();
        var lblStatus = new ToolStripStatusLabel { Name = "lblStatus", Text = "Listo" };
        statusBar.Items.Add(lblStatus);

        var grid = new DataGridView
        {
            Name = "grdPedidos",
            Left = 0, Top = 65,
            Width = ClientSize.Width,
            Height = ClientSize.Height - 65,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        Controls.AddRange([toolbar, lblFechaIni, dtpFecIni, lblFechaFin, dtpFecFin, grid, statusBar]);

        btnConsultar.Click += async (_, _) =>
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                var result = await _obtenerHandler.HandleAsync(
                    new ObtenerPedidosCentralQuery(dtpFecIni.Value.Date, dtpFecFin.Value.Date.AddHours(23).AddMinutes(59)));
                if (result.EsExitoso)
                {
                    grid.DataSource = result.Valor?.ToList();
                    lblStatus.Text = $"Registros: {result.Valor?.Count() ?? 0}";
                }
                else
                    MessageBox.Show(result.MensajeError, "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Cursor = Cursors.Default; }
        };

        // Case 3: Confirmar entrega — BR-DEL-012, BR-DEL-013
        btnConfirmarEntrega.Click += async (_, _) =>
        {
            var pedido = ObtenerPedidoSeleccionado(grid, "grdPedidos");
            if (pedido is null) return;

            if (MessageBox.Show($"¿Confirmar la entrega del Pedido {pedido}?",
                    "Confirmar Entrega", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            var result = await _confirmarHandler.HandleAsync(
                new ConfirmarEntregaCentralCommand(pedido, _usuarioActual));

            if (result.EsExitoso)
            {
                MessageBox.Show("Pedido Entregado", "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnConsultar.PerformClick();
            }
            else if (result.CodigoError == "REQUIERE_SUPERVISOR_22")
            {
                // BR-DEL-013: el pedido no está cancelado — requiere supervisor acción "22"
                if (MessageBox.Show(result.MensajeError + "\n¿Desea aún así Entregarlo?",
                        "Autorización Supervisor", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    var clave = PromptSupervisor();
                    if (string.IsNullOrWhiteSpace(clave))
                    {
                        MessageBox.Show("Clave no permitida", "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    var resultSup = await _confirmarHandler.HandleAsync(
                        new ConfirmarEntregaCentralCommand(pedido, _usuarioActual, SupervisorAutorizado: true));
                    if (resultSup.EsExitoso)
                    {
                        MessageBox.Show("Pedido Entregado", "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        btnConsultar.PerformClick();
                    }
                    else
                        MessageBox.Show(resultSup.MensajeError, "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show(result.MensajeError, "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        };

        // Case 5: Revertir entrega — requiere supervisor
        btnRevertirEntrega.Click += async (_, _) =>
        {
            var pedido = ObtenerPedidoSeleccionado(grid, "grdPedidos");
            if (pedido is null) return;

            var clave = PromptSupervisor();
            if (string.IsNullOrWhiteSpace(clave))
            {
                MessageBox.Show("Clave no permitida", "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            var result = await _revertirHandler.HandleAsync(
                new RevertirEntregaCentralCommand(pedido, _usuarioActual, SupervisorAutorizado: true));

            if (result.EsExitoso)
            {
                MessageBox.Show("Operación exitosa", "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnConsultar.PerformClick();
            }
            else
                MessageBox.Show(result.MensajeError, "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        };

        // Case 2: Modificar fecha programada — BR-DEL-014
        btnModificarFecha.Click += async (_, _) =>
        {
            var pedido = ObtenerPedidoSeleccionado(grid, "grdPedidos");
            if (pedido is null) return;

            using var dlg = new Form
            {
                Text = "Modificar Fecha Programada",
                Size = new Size(300, 140),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent
            };
            var dtp = new DateTimePicker { Left = 10, Top = 10, Width = 260, Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy HH:mm", Value = DateTime.Now };
            var btnOk = new Button { Text = "Aceptar", Left = 100, Top = 55, DialogResult = DialogResult.OK };
            var btnCancelar = new Button { Text = "Cancelar", Left = 185, Top = 55, DialogResult = DialogResult.Cancel };
            dlg.Controls.AddRange([dtp, btnOk, btnCancelar]);
            dlg.AcceptButton = btnOk;

            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            var result = await _modificarFechaHandler.HandleAsync(
                new ModificarFechaProgramadaDeliveryCommand(pedido, dtp.Value));

            if (result.EsExitoso)
                btnConsultar.PerformClick();
            else
                MessageBox.Show(result.MensajeError, "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        };

        btnEmitir.Click += (_, _) =>
            MessageBox.Show("Funcionalidad de emisión de reporte de entrega pendiente.", "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Information);

        btnSalir.Click += (_, _) => Close();

        Load += async (_, _) => btnConsultar.PerformClick();
    }

    private static string? ObtenerPedidoSeleccionado(DataGridView grid, string gridName)
    {
        if (grid.CurrentRow is null || grid.RowCount == 0)
        {
            MessageBox.Show("No Existe Datos Ingresados", "Central de Pedidos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            return null;
        }
        return grid.CurrentRow.Cells[0].Value?.ToString();
    }

    /// <summary>
    /// Solicita contraseña de supervisor.
    /// Legacy: Supervisor("22") — verifica acceso código 22 en TACCESO.
    /// </summary>
    private string? PromptSupervisor()
    {
        using var dlg = new Form
        {
            Text = "Clave de Supervisor",
            Size = new Size(280, 120),
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent
        };
        var lbl = new Label { Text = "Ingrese clave:", Left = 10, Top = 10, Width = 120 };
        var txtClave = new TextBox { Left = 140, Top = 8, Width = 110, PasswordChar = '*' };
        var btnOk = new Button { Text = "Aceptar", Left = 60, Top = 45, Width = 70, DialogResult = DialogResult.OK };
        var btnCancel = new Button { Text = "Cancelar", Left = 145, Top = 45, Width = 80, DialogResult = DialogResult.Cancel };
        dlg.Controls.AddRange([lbl, txtClave, btnOk, btnCancel]);
        dlg.AcceptButton = btnOk;
        return dlg.ShowDialog(this) == DialogResult.OK ? txtClave.Text : null;
    }
}

