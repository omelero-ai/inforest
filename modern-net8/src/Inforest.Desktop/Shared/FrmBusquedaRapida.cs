namespace Inforest.Desktop.Shared;

/// <summary>
/// Diálogo genérico de búsqueda/selección rápida.
/// Legacy: <c>frmBusquedaRapida.frm</c>.
/// <para>
/// El formulario muestra una lista de ítems <see cref="BusquedaItem"/> en un DataGridView
/// con filtrado por texto en tiempo real sobre la columna <c>Descripcion</c> (columna predeterminada).
/// El resultado se devuelve mediante <see cref="Resultado"/>, o <c>null</c> si el usuario canceló.
/// </para>
/// <para>
/// Uso:
/// <code>
/// using var dlg = new FrmBusquedaRapida("Buscar producto", items, "Código", "Descripción");
/// if (dlg.ShowDialog(this) == DialogResult.OK &amp;&amp; dlg.Resultado is not null)
///     procesoSeleccion(dlg.Resultado.Codigo, dlg.Resultado.Descripcion);
/// </code>
/// </para>
/// Reglas de negocio migradas: filtrado en tiempo real (BR-BUSQ-001), selección con Enter
/// o doble clic (BR-BUSQ-002), cancelación con Escape (BR-BUSQ-003).
/// </summary>
public sealed class FrmBusquedaRapida : Form
{
    private static readonly Color SurfaceColor = Color.FromArgb(245, 246, 247);
    private static readonly Color HeaderColor = Color.FromArgb(235, 236, 240);
    private static readonly Color AccentColor = Color.FromArgb(0, 122, 204);

    private readonly string _titulo;
    private readonly IReadOnlyList<BusquedaItem> _items;
    private readonly string _headerCodigo;
    private readonly string _headerDescripcion;
    private readonly string _headerExtra;
    private readonly bool _mostrarExtra;

    private List<BusquedaItem> _filtrados;
    private readonly DataGridView _dgv;
    private readonly TextBox _txtFiltro;
    private readonly Label _lblContador;

    /// <summary>
    /// Ítem seleccionado por el usuario. Null si el diálogo fue cancelado.
    /// Equivale a comprobar <c>wEnter = True</c> + leer <c>sCodigo / sDescrip</c> en VB6.
    /// </summary>
    public BusquedaResultado? Resultado { get; private set; }

    /// <param name="titulo">Texto del título del formulario.</param>
    /// <param name="items">Lista de ítems a mostrar.</param>
    /// <param name="headerCodigo">Encabezado de la columna Código.</param>
    /// <param name="headerDescripcion">Encabezado de la columna Descripción.</param>
    /// <param name="headerExtra">Encabezado de la columna Extra (si aplica).</param>
    public FrmBusquedaRapida(
        string titulo,
        IReadOnlyList<BusquedaItem> items,
        string headerCodigo = "Código",
        string headerDescripcion = "Descripción",
        string headerExtra = "")
    {
        _titulo = titulo;
        _items = items;
        _filtrados = [.. items];
        _headerCodigo = headerCodigo;
        _headerDescripcion = headerDescripcion;
        _headerExtra = headerExtra;
        _mostrarExtra = !string.IsNullOrWhiteSpace(headerExtra);

        Text = titulo;
        Size = new Size(700, 520);
        MinimumSize = new Size(520, 380);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = SurfaceColor;
        FormBorderStyle = FormBorderStyle.Sizable;
        KeyPreview = true;

        // Barra de filtro
        var panelFiltro = new Panel
        {
            Dock = DockStyle.Top,
            Height = 50,
            BackColor = HeaderColor,
            Padding = new Padding(10, 8, 10, 8)
        };

        var lblFiltro = new Label
        {
            Text = "Buscar:",
            AutoSize = true,
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(35, 35, 35),
            Left = 10,
            Top = 15
        };

        _txtFiltro = new TextBox
        {
            Width = 360,
            Height = 28,
            Left = 65,
            Top = 11,
            Font = new Font("Segoe UI", 10),
            BorderStyle = BorderStyle.FixedSingle
        };
        _txtFiltro.TextChanged += OnFiltroChanged;

        _lblContador = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 8),
            ForeColor = Color.FromArgb(100, 100, 100),
            Left = 435,
            Top = 16
        };

        panelFiltro.Controls.Add(lblFiltro);
        panelFiltro.Controls.Add(_txtFiltro);
        panelFiltro.Controls.Add(_lblContador);

        // DataGridView
        _dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            BackgroundColor = SurfaceColor,
            BorderStyle = BorderStyle.None,
            GridColor = Color.FromArgb(220, 222, 225),
            RowHeadersVisible = false,
            Font = new Font("Segoe UI", 9),
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = HeaderColor,
                ForeColor = Color.FromArgb(35, 35, 35),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleLeft
            },
            DefaultCellStyle = new DataGridViewCellStyle
            {
                SelectionBackColor = Color.FromArgb(210, 230, 255),
                SelectionForeColor = Color.FromArgb(35, 35, 35)
            }
        };
        _dgv.CellDoubleClick += OnCellDoubleClick;

        // Pie de página
        var footer = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 48,
            BackColor = SurfaceColor,
            Padding = new Padding(10, 6, 10, 6)
        };

        var btnSeleccionar = new Button
        {
            Text = "Seleccionar",
            Width = 120,
            Height = 34,
            Left = footer.Width - 252,
            Top = 7,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
            FlatStyle = FlatStyle.Flat,
            BackColor = AccentColor,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnSeleccionar.FlatAppearance.BorderSize = 0;
        btnSeleccionar.Click += (_, _) => ConfirmarSeleccion();

        var btnCancelar = new Button
        {
            Text = "Cancelar",
            Width = 110,
            Height = 34,
            Left = footer.Width - 130,
            Top = 7,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnCancelar.FlatAppearance.BorderSize = 0;
        btnCancelar.Click += (_, _) =>
        {
            Resultado = null;
            DialogResult = DialogResult.Cancel;
            Close();
        };

        footer.Controls.Add(btnSeleccionar);
        footer.Controls.Add(btnCancelar);

        Controls.Add(_dgv);
        Controls.Add(panelFiltro);
        Controls.Add(footer);

        Load += (_, _) =>
        {
            ConfigurarColumnas();
            ActualizarGrid();
            _txtFiltro.Focus();
        };

        KeyDown += OnFormKeyDown;
    }

    private void ConfigurarColumnas()
    {
        _dgv.Columns.Clear();

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Codigo",
            HeaderText = _headerCodigo,
            FillWeight = 25,
            SortMode = DataGridViewColumnSortMode.Automatic
        });

        _dgv.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Descripcion",
            HeaderText = _headerDescripcion,
            FillWeight = _mostrarExtra ? 50 : 75,
            SortMode = DataGridViewColumnSortMode.Automatic
        });

        if (_mostrarExtra)
        {
            _dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Extra",
                HeaderText = _headerExtra,
                FillWeight = 25,
                SortMode = DataGridViewColumnSortMode.Automatic
            });
        }
    }

    private void ActualizarGrid()
    {
        _dgv.Rows.Clear();

        foreach (var item in _filtrados)
        {
            if (_mostrarExtra)
                _dgv.Rows.Add(item.Codigo, item.Descripcion, item.Extra);
            else
                _dgv.Rows.Add(item.Codigo, item.Descripcion);
        }

        _lblContador.Text = $"{_filtrados.Count} registro(s)";

        if (_dgv.Rows.Count > 0)
            _dgv.Rows[0].Selected = true;
    }

    private void OnFiltroChanged(object? sender, EventArgs e)
    {
        var texto = _txtFiltro.Text.Trim();

        _filtrados = string.IsNullOrEmpty(texto)
            ? [.. _items]
            : [.. _items.Where(i =>
                i.Descripcion.Contains(texto, StringComparison.OrdinalIgnoreCase) ||
                i.Codigo.Contains(texto, StringComparison.OrdinalIgnoreCase))];

        ActualizarGrid();
    }

    private void OnCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex >= 0)
            ConfirmarSeleccion();
    }

    private void OnFormKeyDown(object? sender, KeyEventArgs e)
    {
        switch (e.KeyCode)
        {
            case Keys.Escape:
                // BR-BUSQ-003: cancelar con Escape
                Resultado = null;
                DialogResult = DialogResult.Cancel;
                Close();
                e.Handled = true;
                break;

            case Keys.Enter:
                // BR-BUSQ-002: confirmar con Enter
                ConfirmarSeleccion();
                e.Handled = true;
                break;

            case Keys.Down when _txtFiltro.Focused && _dgv.Rows.Count > 0:
                _dgv.Focus();
                if (_dgv.SelectedRows.Count == 0)
                    _dgv.Rows[0].Selected = true;
                e.Handled = true;
                break;

            case Keys.Up when _dgv.Focused && _dgv.CurrentRow?.Index == 0:
                _txtFiltro.Focus();
                e.Handled = true;
                break;
        }
    }

    private void ConfirmarSeleccion()
    {
        if (_dgv.SelectedRows.Count == 0 || _dgv.Rows.Count == 0)
            return;

        var row = _dgv.SelectedRows[0];
        var codigo = row.Cells["Codigo"].Value?.ToString() ?? string.Empty;
        var descripcion = row.Cells["Descripcion"].Value?.ToString() ?? string.Empty;

        Resultado = new BusquedaResultado(codigo, descripcion);
        DialogResult = DialogResult.OK;
        Close();
    }
}
