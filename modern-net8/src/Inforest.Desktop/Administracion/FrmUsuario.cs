using System.ComponentModel;
using Inforest.Application.Interfaces;

namespace Inforest.Desktop.Administracion;

/// <summary>
/// Legacy: <c>frmUsuario.frm</c>.
/// </summary>
public class FrmUsuario : Form
{
    private readonly IAuthService _authService;
    private readonly IRbacService _rbacService;
    private readonly BindingList<UsuarioAdminRow> _usuarios = [];
    private readonly DataGridView _grid;
    private readonly TextBox _txtCodigo;
    private readonly TextBox _txtNombre;
    private readonly TextBox _txtGrupo;
    private readonly CheckBox _chkActivo;
    private readonly Button _btnGuardar;
    private readonly Button _btnCancelar;
    private UsuarioAdminRow? _usuarioEnEdicion;

    public FrmUsuario(IAuthService authService, IRbacService rbacService)
    {
        _authService = authService;
        _rbacService = rbacService;

        Text = "Usuarios";
        Dock = DockStyle.Fill;

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            DataSource = _usuarios
        };
        _grid.SelectionChanged += (_, _) => CargarSeleccion();

        _txtCodigo = new TextBox { Width = 120 };
        _txtNombre = new TextBox { Width = 240 };
        _txtGrupo = new TextBox { Width = 140 };
        _chkActivo = new CheckBox { Text = "Activo", AutoSize = true };

        var panelEditor = new TableLayoutPanel { Dock = DockStyle.Top, Height = 110, ColumnCount = 4, Padding = new Padding(8) };
        panelEditor.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
        panelEditor.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        panelEditor.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
        panelEditor.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panelEditor.Controls.Add(new Label { Text = "Código", AutoSize = true }, 0, 0);
        panelEditor.Controls.Add(_txtCodigo, 1, 0);
        panelEditor.Controls.Add(new Label { Text = "Grupo", AutoSize = true }, 2, 0);
        panelEditor.Controls.Add(_txtGrupo, 3, 0);
        panelEditor.Controls.Add(new Label { Text = "Nombre", AutoSize = true }, 0, 1);
        panelEditor.Controls.Add(_txtNombre, 1, 1);
        panelEditor.Controls.Add(_chkActivo, 3, 1);

        var panelBotones = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 42, Padding = new Padding(8) };
        var btnNuevo = new Button { Text = "Nuevo", Width = 90 };
        var btnEditar = new Button { Text = "Editar", Width = 90 };
        _btnGuardar = new Button { Text = "Guardar", Width = 90 };
        _btnCancelar = new Button { Text = "Cancelar", Width = 90 };
        panelBotones.Controls.AddRange([btnNuevo, btnEditar, _btnGuardar, _btnCancelar]);

        btnNuevo.Click += (_, _) => PrepararNuevo();
        btnEditar.Click += (_, _) => PrepararEdicion();
        _btnGuardar.Click += async (_, _) => await GuardarAsync();
        _btnCancelar.Click += (_, _) => LimpiarEditor();

        Controls.Add(_grid);
        Controls.Add(panelBotones);
        Controls.Add(panelEditor);
        Load += FrmUsuario_Load;
    }

    private async void FrmUsuario_Load(object? sender, EventArgs e)
    {
        await _rbacService.ObtenerPermisosAsync("ADMIN", "02");
        _usuarios.Clear();
        _usuarios.Add(new UsuarioAdminRow("ADMIN", "Administrador General", "ADM", true));
        _usuarios.Add(new UsuarioAdminRow("CAJERO1", "Cajero Principal", "CAJ", true));
        _usuarios.Add(new UsuarioAdminRow("AUDITOR", "Consulta y reportes", "CON", false));
        LimpiarEditor();
    }

    private void PrepararNuevo()
    {
        _usuarioEnEdicion = null;
        _txtCodigo.Enabled = true;
        _txtCodigo.Clear();
        _txtNombre.Clear();
        _txtGrupo.Clear();
        _chkActivo.Checked = true;
        _txtCodigo.Focus();
    }

    private void PrepararEdicion()
    {
        if (_grid.CurrentRow?.DataBoundItem is UsuarioAdminRow row)
        {
            _usuarioEnEdicion = row;
            _txtCodigo.Enabled = false;
            _txtCodigo.Text = row.Codigo;
            _txtNombre.Text = row.Nombre;
            _txtGrupo.Text = row.Grupo;
            _chkActivo.Checked = row.Activo;
        }
    }

    private async Task GuardarAsync()
    {
        var codigo = _txtCodigo.Text.Trim().ToUpperInvariant();
        var nombre = _txtNombre.Text.Trim();
        var grupo = _txtGrupo.Text.Trim().ToUpperInvariant();

        if (string.IsNullOrWhiteSpace(codigo) || string.IsNullOrWhiteSpace(nombre))
        {
            MessageBox.Show("Código y nombre son obligatorios.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _btnGuardar.Enabled = false;
        try
        {
            await _authService.TieneAccesoAsync(codigo, "02", "USR_EDI");
            if (_usuarioEnEdicion is null)
            {
                _usuarios.Add(new UsuarioAdminRow(codigo, nombre, grupo, _chkActivo.Checked));
            }
            else
            {
                var index = _usuarios.IndexOf(_usuarioEnEdicion);
                _usuarios[index] = new UsuarioAdminRow(codigo, nombre, grupo, _chkActivo.Checked);
            }

            _grid.Refresh();
            LimpiarEditor();
        }
        finally
        {
            _btnGuardar.Enabled = true;
        }
    }

    private void CargarSeleccion()
    {
        if (_grid.CurrentRow?.DataBoundItem is not UsuarioAdminRow row)
            return;

        _txtCodigo.Text = row.Codigo;
        _txtNombre.Text = row.Nombre;
        _txtGrupo.Text = row.Grupo;
        _chkActivo.Checked = row.Activo;
    }

    private void LimpiarEditor()
    {
        _usuarioEnEdicion = null;
        _txtCodigo.Enabled = true;
        _txtCodigo.Clear();
        _txtNombre.Clear();
        _txtGrupo.Clear();
        _chkActivo.Checked = true;
    }

    private sealed record UsuarioAdminRow(string Codigo, string Nombre, string Grupo, bool Activo);
}
