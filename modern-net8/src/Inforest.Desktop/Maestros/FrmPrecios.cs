using Inforest.Application.Precios;

namespace Inforest.Desktop.Maestros;

/// <summary>
/// Listado de precios de productos (solo consulta).
/// Legacy: frmPrecios.frm — "Listado de Precios"
/// SELECT Codigo, Grupo, SubGrupo, Descripcion, nPrecioVenta, nPrecioDelivery, nPrecioLlevar, lActivo FROM vProducto
/// BR-PRECIO-001 (solo lectura)
/// </summary>
public sealed class FrmPrecios : Form
{
    private readonly ObtenerListadoPreciosHandler _handler;

    // ── Controls ──────────────────────────────────────────────────────────────
    private readonly DataGridView   _dgv          = new() { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AllowUserToAddRows = false };
    private readonly TextBox        _txtFiltro    = new() { Width = 260, PlaceholderText = "Filtrar por código, grupo o descripción…" };
    private readonly Button         _btnSalir     = new() { Text = "Salir", Width = 80 };
    private readonly Label          _lblStatus    = new() { Dock = DockStyle.Bottom, Height = 22, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(4, 0, 0, 0) };

    private List<PrecioProductoFila> _todos = [];

    public FrmPrecios(ObtenerListadoPreciosHandler handler)
    {
        _handler = handler;

        Text            = "Listado de Precios";
        Width           = 920;
        Height          = 600;
        StartPosition   = FormStartPosition.CenterScreen;
        MdiParent       = System.Windows.Forms.Application.OpenForms.OfType<Form>().FirstOrDefault(f => f.IsMdiContainer);

        ConfigurarGrilla();
        Controls.Add(_dgv);
        Controls.Add(BuildToolbar());
        Controls.Add(_lblStatus);

        _txtFiltro.TextChanged += (_, _) => AplicarFiltro();
        _btnSalir.Click        += (_, _) => Close();

        Load += async (_, _) => await CargarAsync();
    }

    private void ConfigurarGrilla()
    {
        _dgv.Columns.AddRange(
        [
            new DataGridViewTextBoxColumn { DataPropertyName = "Codigo",        HeaderText = "Código",   Width = 80  },
            new DataGridViewTextBoxColumn { DataPropertyName = "Grupo",         HeaderText = "Grupo",    Width = 130 },
            new DataGridViewTextBoxColumn { DataPropertyName = "SubGrupo",      HeaderText = "SubGrupo", Width = 150 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Descripcion",   HeaderText = "Producto", Width = 250 },
            new DataGridViewTextBoxColumn { DataPropertyName = "PrecioVenta",   HeaderText = "Salón",    Width = 90,  DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { DataPropertyName = "PrecioDelivery",HeaderText = "Delivery", Width = 90,  DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewTextBoxColumn { DataPropertyName = "PrecioLlevar",  HeaderText = "Llevar",   Width = 90,  DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight } },
            new DataGridViewCheckBoxColumn { DataPropertyName = "Activo",       HeaderText = "Activo",   Width = 60  }
        ]);
    }

    private Panel BuildToolbar()
    {
        var panel = new Panel { Dock = DockStyle.Top, Height = 42 };
        int x = 8;

        void Add(Control c) { c.Top = 6; c.Left = x; x += c.Width + 6; panel.Controls.Add(c); }

        Add(new Label { Text = "Filtro:", AutoSize = true, Top = 12 });
        Add(_txtFiltro);
        Add(new Label { Text = "  ", AutoSize = true });
        Add(_btnSalir);
        return panel;
    }

    private async Task CargarAsync()
    {
        _lblStatus.Text = "Cargando…";
        var result = await _handler.HandleAsync(new ObtenerListadoPreciosQuery());
        if (!result.EsExitoso)
        {
            _lblStatus.Text = result.MensajeError;
            return;
        }

        _todos = result.Valor!.ToList();
        AplicarFiltro();
        _lblStatus.Text = $"{_todos.Count} producto(s) cargados.";
    }

    private void AplicarFiltro()
    {
        var texto = _txtFiltro.Text.Trim();
        var lista = string.IsNullOrEmpty(texto)
            ? _todos
            : _todos.Where(r =>
                r.Codigo.Contains(texto, StringComparison.OrdinalIgnoreCase) ||
                r.Grupo.Contains(texto, StringComparison.OrdinalIgnoreCase) ||
                r.SubGrupo.Contains(texto, StringComparison.OrdinalIgnoreCase) ||
                r.Descripcion.Contains(texto, StringComparison.OrdinalIgnoreCase))
              .ToList();

        _dgv.DataSource = lista;
        _lblStatus.Text = $"{lista.Count} / {_todos.Count} producto(s).";
    }
}
