using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Maestros;

/// <summary>
/// Maestro WinForms de clientes.
/// Legacy: frmCliente.frm.
/// </summary>
public class FrmCliente : Form
{
    private readonly IClienteRepository _repository;
    private readonly DataGridView dgv = new() { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
    private readonly TextBox txtCodigo = new();
    private readonly TextBox txtEmpresa = new();
    private readonly TextBox txtIdentidad = new();
    private readonly TextBox txtDireccion = new();
    private readonly TextBox txtCorreo = new();
    private readonly ComboBox cboTipo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly CheckBox chkActivo = new() { Text = "Activo", Checked = true, AutoSize = true };
    private readonly Button btnNuevo = new() { Text = "Nuevo" };
    private readonly Button btnEditar = new() { Text = "Editar" };
    private readonly Button btnEliminar = new() { Text = "Eliminar" };
    private readonly Button btnGuardar = new() { Text = "Guardar" };
    private List<Cliente> _items = [];

    public FrmCliente(IClienteRepository repository)
    {
        _repository = repository;
        Text = "Clientes";
        Width = 1100;
        Height = 650;
        StartPosition = FormStartPosition.CenterScreen;
        Controls.Add(BuildLayout());
        cboTipo.DataSource = new[] { "RUC", "DNI", "OTRO" };
        Load += async (_, _) => await CargarAsync();
        dgv.SelectionChanged += (_, _) => CargarSeleccion();
        btnNuevo.Click += (_, _) => Limpiar();
        btnEditar.Click += (_, _) => CargarSeleccion();
        btnEliminar.Click += async (_, _) => await EliminarAsync();
        btnGuardar.Click += async (_, _) => await GuardarAsync();
    }

    private Control BuildLayout()
    {
        var datos = new TableLayoutPanel { Dock = DockStyle.Top, Height = 160, ColumnCount = 4, Padding = new Padding(8) };
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        datos.Controls.Add(new Label { Text = "Código", AutoSize = true }, 0, 0);
        datos.Controls.Add(txtCodigo, 1, 0);
        datos.Controls.Add(new Label { Text = "Empresa", AutoSize = true }, 2, 0);
        datos.Controls.Add(txtEmpresa, 3, 0);
        datos.Controls.Add(new Label { Text = "Identidad", AutoSize = true }, 0, 1);
        datos.Controls.Add(txtIdentidad, 1, 1);
        datos.Controls.Add(new Label { Text = "Dirección", AutoSize = true }, 2, 1);
        datos.Controls.Add(txtDireccion, 3, 1);
        datos.Controls.Add(new Label { Text = "Correo", AutoSize = true }, 0, 2);
        datos.Controls.Add(txtCorreo, 1, 2);
        datos.Controls.Add(new Label { Text = "Tipo", AutoSize = true }, 2, 2);
        datos.Controls.Add(cboTipo, 3, 2);
        datos.Controls.Add(chkActivo, 3, 3);

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(8) };
        acciones.Controls.AddRange([btnNuevo, btnEditar, btnEliminar, btnGuardar]);

        var panel = new Panel { Dock = DockStyle.Fill };
        panel.Controls.Add(dgv);
        panel.Controls.Add(acciones);
        panel.Controls.Add(datos);
        return panel;
    }

    private async Task CargarAsync()
    {
        _items = (await _repository.ObtenerTodosAsync()).ToList();
        dgv.DataSource = null;
        dgv.DataSource = _items;
    }

    private void CargarSeleccion()
    {
        if (dgv.CurrentRow?.DataBoundItem is not Cliente item) return;
        txtCodigo.Text = item.CodigoCliente;
        txtCodigo.Enabled = false;
        txtEmpresa.Text = item.Empresa;
        txtIdentidad.Text = item.Identidad;
        txtDireccion.Text = item.Direccion;
        txtCorreo.Text = item.Correo;
        cboTipo.SelectedItem = item.TipoIdentidad ?? "OTRO";
        chkActivo.Checked = item.Activo;
    }

    private void Limpiar()
    {
        txtCodigo.Enabled = true;
        txtCodigo.Clear();
        txtEmpresa.Clear();
        txtIdentidad.Clear();
        txtDireccion.Clear();
        txtCorreo.Clear();
        cboTipo.SelectedIndex = 0;
        chkActivo.Checked = true;
    }

    private async Task GuardarAsync()
    {
        try
        {
            var tipo = cboTipo.SelectedItem?.ToString();
            var existente = await _repository.ObtenerPorCodigoAsync(txtCodigo.Text.Trim());
            if (existente is null)
            {
                var nuevo = Cliente.Crear(txtCodigo.Text, txtEmpresa.Text, txtIdentidad.Text, txtDireccion.Text, Environment.UserName, txtCorreo.Text, tipo);
                await _repository.InsertarAsync(nuevo);
            }
            else
            {
                existente.Actualizar(txtEmpresa.Text, txtIdentidad.Text, txtDireccion.Text, Environment.UserName, txtCorreo.Text, tipo, null, null, null, null, null, chkActivo.Checked);
                await _repository.ActualizarAsync(existente);
            }

            await CargarAsync();
            Limpiar();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private async Task EliminarAsync()
    {
        if (dgv.CurrentRow?.DataBoundItem is not Cliente item) return;
        item.Inactivar();
        await _repository.ActualizarAsync(item);
        await CargarAsync();
    }
}
