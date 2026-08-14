using Inforest.Application.Kitchen;
using Inforest.Domain.Entities.Cocina;

namespace Inforest.Desktop.Kitchen;

/// <summary>
/// Alta/edición puntual de mensajes de cocina.
/// Legacy: frmMensajeCocinaDetalle.frm.
/// Reglas BR-MSGCOC-001, BR-MSGCOC-002, BR-MSGCOC-003.
/// </summary>
public sealed class FrmMensajeCocinaDetalle : Form
{
    private readonly AgregarMensajeCocinaHandler _agregarHandler;
    private readonly ModificarMensajeCocinaHandler _modificarHandler;
    private readonly EliminarMensajeCocinaHandler _eliminarHandler;
    private readonly MensajeCocina? _mensajeExistente;
    private readonly string _usuarioActual;
    private readonly string _codigoCajaActual;
    private readonly Func<Task>? _onSaved;

    private readonly TextBox _txtCodigo = new() { ReadOnly = true, BackColor = SystemColors.Control, Width = 120 };
    private readonly TextBox _txtMensaje = new() { MaxLength = MensajeCocina.LongitudMaximaMensaje, Width = 420 };
    private readonly CheckBox _chkActivo = new() { Text = "Activo", AutoSize = true, Checked = true };
    private readonly CheckBox _chkAlerta = new() { Text = "Activar alerta", AutoSize = true };
    private readonly Button _btnAgregar = new() { Text = "Agregar", Width = 90 };
    private readonly Button _btnGuardar = new() { Text = "Guardar", Width = 90 };
    private readonly Button _btnEliminar = new() { Text = "Eliminar", Width = 90 };
    private readonly Button _btnCancelar = new() { Text = "Cancelar", Width = 90 };

    public FrmMensajeCocinaDetalle(
        AgregarMensajeCocinaHandler agregarHandler,
        ModificarMensajeCocinaHandler modificarHandler,
        EliminarMensajeCocinaHandler eliminarHandler,
        MensajeCocina? mensajeExistente,
        string usuarioActual,
        string codigoCajaActual,
        Func<Task>? onSaved)
    {
        _agregarHandler = agregarHandler;
        _modificarHandler = modificarHandler;
        _eliminarHandler = eliminarHandler;
        _mensajeExistente = mensajeExistente;
        _usuarioActual = usuarioActual;
        _codigoCajaActual = codigoCajaActual;
        _onSaved = onSaved;

        Text = "Mantenimiento de Mensajes";
        Width = 620;
        Height = 230;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        Controls.Add(BuildLayout());
        Load += (_, _) => CargarDatos();
        _btnAgregar.Click += async (_, _) => await AgregarAsync();
        _btnGuardar.Click += async (_, _) => await GuardarAsync();
        _btnEliminar.Click += async (_, _) => await EliminarAsync();
        _btnCancelar.Click += (_, _) => Close();
    }

    private Control BuildLayout()
    {
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            ColumnCount = 2,
            RowCount = 4
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        table.Controls.Add(new Label { Text = "Código:", AutoSize = true }, 0, 0);
        table.Controls.Add(_txtCodigo, 1, 0);
        table.Controls.Add(new Label { Text = "Mensaje:", AutoSize = true }, 0, 1);
        table.Controls.Add(_txtMensaje, 1, 1);

        var opciones = new FlowLayoutPanel { AutoSize = true };
        opciones.Controls.AddRange([_chkActivo, _chkAlerta]);
        table.Controls.Add(opciones, 1, 2);

        var botones = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(6) };
        botones.Controls.AddRange([_btnCancelar, _btnEliminar, _btnGuardar, _btnAgregar]);

        var panel = new Panel { Dock = DockStyle.Fill };
        panel.Controls.Add(table);
        panel.Controls.Add(botones);
        return panel;
    }

    private void CargarDatos()
    {
        if (_mensajeExistente is null)
        {
            _txtCodigo.Text = "(auto)";
            _chkActivo.Checked = true;
            _chkAlerta.Checked = false;
            _btnGuardar.Enabled = false;
            _btnEliminar.Enabled = false;
            _txtMensaje.Focus();
            return;
        }

        _txtCodigo.Text = _mensajeExistente.Codigo;
        _txtMensaje.Text = _mensajeExistente.Mensaje;
        _chkActivo.Checked = _mensajeExistente.Activo;
        _chkAlerta.Checked = _mensajeExistente.ActivarAlerta;
        _btnAgregar.Enabled = false;
    }

    private async Task AgregarAsync()
    {
        var result = await _agregarHandler.HandleAsync(new AgregarMensajeCocinaCommand(
            _txtMensaje.Text,
            _usuarioActual,
            _codigoCajaActual,
            _chkActivo.Checked,
            _chkAlerta.Checked));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show("Registro guardado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        if (_onSaved is not null) await _onSaved();
        DialogResult = DialogResult.OK;
        Close();
    }

    private async Task GuardarAsync()
    {
        if (_mensajeExistente is null)
            return;

        var result = await _modificarHandler.HandleAsync(new ModificarMensajeCocinaCommand(
            _mensajeExistente.Codigo,
            _txtMensaje.Text,
            _usuarioActual,
            _codigoCajaActual,
            _chkActivo.Checked,
            _chkAlerta.Checked));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show("Registro modificado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        if (_onSaved is not null) await _onSaved();
        DialogResult = DialogResult.OK;
        Close();
    }

    private async Task EliminarAsync()
    {
        if (_mensajeExistente is null)
            return;

        if (MessageBox.Show("¿Seguro de eliminar el mensaje?", Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
            return;

        var result = await _eliminarHandler.HandleAsync(new EliminarMensajeCocinaCommand(_mensajeExistente.Codigo));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_onSaved is not null) await _onSaved();
        DialogResult = DialogResult.OK;
        Close();
    }
}
