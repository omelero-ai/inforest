using Inforest.Application.Kitchen;
using Inforest.Domain.Entities.Cocina;

namespace Inforest.Desktop.Kitchen;

/// <summary>
/// Mantenimiento/listado de mensajes operativos de cocina.
/// Legacy: frmMensajeCocina.frm.
/// Reglas BR-MSGCOC-004 y BR-MSGCOC-005.
/// </summary>
public sealed class FrmMensajeCocina : Form
{
    private readonly ObtenerMensajesCocinaHandler _obtenerHandler;
    private readonly AgregarMensajeCocinaHandler _agregarHandler;
    private readonly ModificarMensajeCocinaHandler _modificarHandler;
    private readonly EliminarMensajeCocinaHandler _eliminarHandler;
    private readonly string _codigoCajaActual;
    private readonly string _usuarioActual;

    private readonly DateTimePicker _dtpDesde = new() { Format = DateTimePickerFormat.Short, Width = 110, Value = DateTime.Today };
    private readonly DateTimePicker _dtpHasta = new() { Format = DateTimePickerFormat.Short, Width = 110, Value = DateTime.Today };
    private readonly TextBox _txtCaja = new() { Width = 60, MaxLength = 3 };
    private readonly Button _btnBuscar = new() { Text = "Buscar", Width = 90 };
    private readonly Button _btnNuevo = new() { Text = "Nuevo", Width = 90 };
    private readonly Button _btnEditar = new() { Text = "Editar", Width = 90 };
    private readonly Button _btnSalir = new() { Text = "Salir", Width = 90 };
    private readonly Label _lblEstado = new() { AutoSize = true, ForeColor = Color.DarkBlue };
    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AutoGenerateColumns = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        MultiSelect = false,
        AllowUserToAddRows = false
    };

    private List<MensajeCocina> _mensajes = [];

    public FrmMensajeCocina(
        ObtenerMensajesCocinaHandler obtenerHandler,
        AgregarMensajeCocinaHandler agregarHandler,
        ModificarMensajeCocinaHandler modificarHandler,
        EliminarMensajeCocinaHandler eliminarHandler,
        string codigoCajaActual,
        string usuarioActual)
    {
        _obtenerHandler = obtenerHandler;
        _agregarHandler = agregarHandler;
        _modificarHandler = modificarHandler;
        _eliminarHandler = eliminarHandler;
        _codigoCajaActual = codigoCajaActual;
        _usuarioActual = usuarioActual;

        Text = "Mantenimiento de Mensajes en Adición";
        Width = 1180;
        Height = 640;
        StartPosition = FormStartPosition.CenterScreen;
        MdiParent = System.Windows.Forms.Application.OpenForms.OfType<Form>().FirstOrDefault(f => f.IsMdiContainer);

        _txtCaja.Text = codigoCajaActual;

        ConfigurarColumnas();
        Controls.Add(BuildLayout());

        Load += async (_, _) => await CargarAsync();
        _btnBuscar.Click += async (_, _) => await CargarAsync();
        _btnNuevo.Click += (_, _) => AbrirDetalle(null);
        _btnEditar.Click += (_, _) => AbrirDetalle(MensajeSeleccionado());
        _btnSalir.Click += (_, _) => Close();
        _grid.DoubleClick += (_, _) => AbrirDetalle(MensajeSeleccionado());
    }

    private void ConfigurarColumnas()
    {
        _grid.Columns.AddRange(
        [
            new DataGridViewTextBoxColumn { DataPropertyName = nameof(MensajeCocina.Codigo), HeaderText = "Código", Width = 90 },
            new DataGridViewTextBoxColumn { DataPropertyName = nameof(MensajeCocina.UsuarioRegistro), HeaderText = "Usuario", Width = 110 },
            new DataGridViewTextBoxColumn { DataPropertyName = nameof(MensajeCocina.Mensaje), HeaderText = "Mensaje", Width = 360 },
            new DataGridViewTextBoxColumn { DataPropertyName = nameof(MensajeCocina.FechaRegistro), HeaderText = "Fecha Registro", Width = 140, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm:ss" } },
            new DataGridViewTextBoxColumn { DataPropertyName = nameof(MensajeCocina.FechaModificacion), HeaderText = "Fecha Modificación", Width = 140, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm:ss" } },
            new DataGridViewTextBoxColumn { DataPropertyName = nameof(MensajeCocina.UsuarioModificacion), HeaderText = "Modificó", Width = 110 },
            new DataGridViewTextBoxColumn { DataPropertyName = nameof(MensajeCocina.Caja), HeaderText = "Caja", Width = 70 },
            new DataGridViewCheckBoxColumn { DataPropertyName = nameof(MensajeCocina.Activo), HeaderText = "Activo", Width = 60 },
            new DataGridViewCheckBoxColumn { DataPropertyName = nameof(MensajeCocina.ActivarAlerta), HeaderText = "Alerta", Width = 60 }
        ]);
    }

    private Control BuildLayout()
    {
        var top = new Panel { Dock = DockStyle.Top, Height = 44, Padding = new Padding(6) };
        top.Controls.AddRange(
        [
            new Label { Text = "Desde:", AutoSize = true, Left = 6, Top = 12 },
            _dtpDesde,
            new Label { Text = "Hasta:", AutoSize = true, Top = 12 },
            _dtpHasta,
            new Label { Text = "Caja:", AutoSize = true, Top = 12 },
            _txtCaja,
            _btnBuscar,
            _btnNuevo,
            _btnEditar,
            _btnSalir,
            _lblEstado
        ]);

        int x = 6;
        foreach (Control control in top.Controls)
        {
            control.Left = x;
            control.Top = control switch
            {
                Label => 12,
                _ => 8
            };
            x += control.Width + 6;
        }

        var panel = new Panel { Dock = DockStyle.Fill };
        panel.Controls.Add(_grid);
        panel.Controls.Add(top);
        return panel;
    }

    internal async Task CargarAsync()
    {
        _lblEstado.Text = "Cargando...";
        var result = await _obtenerHandler.HandleAsync(new ObtenerMensajesCocinaQuery(_dtpDesde.Value.Date, _dtpHasta.Value.Date, _txtCaja.Text.Trim()));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            _lblEstado.Text = string.Empty;
            return;
        }

        _mensajes = result.Valor!.OrderBy(m => m.Codigo).ToList();
        _grid.DataSource = null;
        _grid.DataSource = _mensajes;
        _lblEstado.Text = $"Registros: {_mensajes.Count}";
    }

    private MensajeCocina? MensajeSeleccionado()
    {
        if (_grid.CurrentRow?.Index is null or < 0 || _grid.CurrentRow.Index >= _mensajes.Count)
        {
            MessageBox.Show("No existe datos ingresados.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return null;
        }

        return _mensajes[_grid.CurrentRow.Index];
    }

    private void AbrirDetalle(MensajeCocina? mensaje)
    {
        if (mensaje is null && _btnEditar.Focused)
            return;

        using var frm = new FrmMensajeCocinaDetalle(
            _agregarHandler,
            _modificarHandler,
            _eliminarHandler,
            mensaje,
            _usuarioActual,
            string.IsNullOrWhiteSpace(_txtCaja.Text) ? _codigoCajaActual : _txtCaja.Text.Trim(),
            async () => await CargarAsync());
        frm.ShowDialog(this);
    }
}
