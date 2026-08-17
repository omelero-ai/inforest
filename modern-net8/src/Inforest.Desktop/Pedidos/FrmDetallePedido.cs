using Inforest.Application.Pedidos;
using Inforest.Domain.Common;

namespace Inforest.Desktop.Pedidos;

/// <summary>
/// Formulario de consulta de detalle de un pedido (solo lectura).
/// Legacy: frmDetallePedido.frm — Form_Load, SubAsignar, Asignar.
/// Reglas: BR-PEDIDO-005 (solo ítems activos), BR-PEDIDO-006 (vista combos).
/// </summary>
public sealed class FrmDetallePedido : Form
{
    private readonly ObtenerDetallePedidoExtendidoHandler _handler;
    private readonly string _codigoPedido;

    // ── Controles ──────────────────────────────────────────────────────────────
    private Label _lblPedido = null!;
    private Label _lblCabeceraInfo = null!;
    private Label _lblTotal = null!;
    private Label _lblProductCount = null!;
    private DataGridView _dgvItems = null!;
    private DataGridView _dgvCombos = null!;

    public FrmDetallePedido(string codigoPedido, ObtenerDetallePedidoExtendidoHandler handler)
    {
        _codigoPedido = codigoPedido;
        _handler = handler;

        Text = $"Detalle de pedido — {codigoPedido}";
        Size = new Size(1100, 680);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        BuildLayout();
        Load += async (_, _) => await CargarAsync();
    }

    // ── Layout ──────────────────────────────────────────────────────────────────

    private void BuildLayout()
    {
        // ── Header azul ──
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 36,
            BackColor = Color.FromArgb(0, 84, 166)
        };
        _lblPedido = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Left,
            Width = 300,
            Height = 36,
            Padding = new Padding(10, 8, 0, 0),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.White,
            Text = $"Pedido: {_codigoPedido}"
        };
        _lblCabeceraInfo = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Fill,
            Height = 36,
            Padding = new Padding(8, 8, 0, 0),
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.White,
            Text = ""
        };
        pnlHeader.Controls.Add(_lblCabeceraInfo);
        pnlHeader.Controls.Add(_lblPedido);

        // ── Botones de cierre ──
        var pnlFooter = new Panel { Dock = DockStyle.Bottom, Height = 44 };
        var btnCerrar = new Button
        {
            Text = "Cerrar",
            Width = 100,
            Height = 30,
            Left = pnlFooter.Width - 120,
            Top = 7,
            Anchor = AnchorStyles.Right | AnchorStyles.Top
        };
        btnCerrar.Click += (_, _) => Close();

        _lblTotal = new Label
        {
            AutoSize = false,
            Left = 10,
            Top = 12,
            Width = 280,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Text = "Total: —"
        };
        _lblProductCount = new Label
        {
            AutoSize = false,
            Left = 300,
            Top = 12,
            Width = 200,
            Font = new Font("Segoe UI", 9),
            Text = ""
        };
        pnlFooter.Controls.AddRange([_lblTotal, _lblProductCount, btnCerrar]);

        // ── Split: items (top) + combos (bottom) ──
        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = 380
        };

        split.Panel1.Controls.Add(BuildItemsPanel());
        split.Panel2.Controls.Add(BuildCombosPanel());

        Controls.Add(split);
        Controls.Add(pnlFooter);
        Controls.Add(pnlHeader);
    }

    private Panel BuildItemsPanel()
    {
        var lbl = new Label
        {
            Dock = DockStyle.Top,
            Height = 22,
            Padding = new Padding(4, 4, 0, 0),
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            BackColor = Color.FromArgb(230, 237, 247),
            Text = "Ítems del pedido"
        };

        _dgvItems = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            RowHeadersVisible = false,
            Font = new Font("Segoe UI", 8.5f)
        };
        ConfigurarColumnasItems();

        var pnl = new Panel { Dock = DockStyle.Fill };
        pnl.Controls.Add(_dgvItems);
        pnl.Controls.Add(lbl);
        return pnl;
    }

    private Panel BuildCombosPanel()
    {
        var lbl = new Label
        {
            Dock = DockStyle.Top,
            Height = 22,
            Padding = new Padding(4, 4, 0, 0),
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            BackColor = Color.FromArgb(230, 237, 247),
            Text = "Combos del pedido"
        };

        _dgvCombos = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            RowHeadersVisible = false,
            Font = new Font("Segoe UI", 8.5f)
        };
        ConfigurarColumnasCombos();

        var pnl = new Panel { Dock = DockStyle.Fill };
        pnl.Controls.Add(_dgvCombos);
        pnl.Controls.Add(lbl);
        return pnl;
    }

    private void ConfigurarColumnasItems()
    {
        _dgvItems.AutoGenerateColumns = false;
        _dgvItems.Columns.AddRange(
            new DataGridViewTextBoxColumn { Name = "Item",         HeaderText = "Ítem",       DataPropertyName = "Item",         FillWeight = 5 },
            new DataGridViewTextBoxColumn { Name = "Producto",     HeaderText = "Producto",   DataPropertyName = "Producto",     FillWeight = 30 },
            new DataGridViewTextBoxColumn { Name = "PrecioOficial",HeaderText = "P.Oficial",  DataPropertyName = "PrecioOficial",FillWeight = 10, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { Name = "Descuento",    HeaderText = "Dcto",       DataPropertyName = "Descuento",    FillWeight = 7,  DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { Name = "PrecioVenta",  HeaderText = "P.Venta",    DataPropertyName = "PrecioVenta",  FillWeight = 10, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { Name = "Cantidad",     HeaderText = "Cant.",      DataPropertyName = "Cantidad",     FillWeight = 7,  DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { Name = "SubTotal",     HeaderText = "SubTotal",   DataPropertyName = "SubTotal",     FillWeight = 12, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { Name = "Facturado",    HeaderText = "F",          DataPropertyName = "Facturado",    FillWeight = 4 },
            new DataGridViewTextBoxColumn { Name = "Documento",    HeaderText = "Documento",  DataPropertyName = "Documento",    FillWeight = 15 },
            new DataGridViewTextBoxColumn { Name = "Usuario",      HeaderText = "Usuario",    DataPropertyName = "Usuario",      FillWeight = 10 });
    }

    private void ConfigurarColumnasCombos()
    {
        _dgvCombos.AutoGenerateColumns = false;
        _dgvCombos.Columns.AddRange(
            new DataGridViewTextBoxColumn { Name = "Item",      HeaderText = "Ítem",    DataPropertyName = "Item",      FillWeight = 8 },
            new DataGridViewTextBoxColumn { Name = "ItemCombo", HeaderText = "Combo",   DataPropertyName = "ItemCombo", FillWeight = 8 },
            new DataGridViewTextBoxColumn { Name = "Producto",  HeaderText = "Producto",DataPropertyName = "Producto",  FillWeight = 50 },
            new DataGridViewTextBoxColumn { Name = "Cantidad",  HeaderText = "Cant.",   DataPropertyName = "Cantidad",  FillWeight = 10, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } });
    }

    // ── Lógica ──────────────────────────────────────────────────────────────────

    private async Task CargarAsync()
    {
        var result = await _handler.HandleAsync(
            new ObtenerDetallePedidoExtendidoQuery(_codigoPedido));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError ?? "Error al cargar el pedido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            Close();
            return;
        }

        var dato = result.Valor!;

        // Cabecera
        _lblCabeceraInfo.Text = $"Adultos: {dato.NumeroAdultos}  |  Fecha: {dato.FechaRegistro:dd/MM/yyyy HH:mm}" +
            (string.IsNullOrWhiteSpace(dato.Observacion) ? "" : $"  |  Obs: {dato.Observacion}");

        // Items grid
        var bsItems = new BindingSource();
        bsItems.DataSource = dato.Items.ToList();
        _dgvItems.DataSource = bsItems;

        // Combos grid
        var bsCombos = new BindingSource();
        bsCombos.DataSource = dato.Combos.ToList();
        _dgvCombos.DataSource = bsCombos;

        // Totales pie
        _lblTotal.Text = $"Total: S/ {dato.Total:N2}";
        _lblProductCount.Text = $"{dato.Items.Count} producto(s)";
    }
}
