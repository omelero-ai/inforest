using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Maestros;

/// <summary>
/// Maestro WinForms de grupos de producto.
/// Legacy: frmGrupo.frm.
/// </summary>
public class FrmGrupoProducto : Form
{
    private readonly IGrupoProductoRepository _repository;
    private readonly DataGridView dgv = new() { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
    private readonly TextBox txtCodigo = new();
    private readonly TextBox txtDescripcion = new();
    private readonly TextBox txtResumido = new();
    private readonly CheckBox chkActivo = new() { Text = "Activo", Checked = true, AutoSize = true };
    private readonly Button btnNuevo = new() { Text = "Nuevo" };
    private readonly Button btnEditar = new() { Text = "Editar" };
    private readonly Button btnEliminar = new() { Text = "Eliminar" };
    private readonly Button btnGuardar = new() { Text = "Guardar" };
    private List<GrupoProducto> _items = [];

    public FrmGrupoProducto(IGrupoProductoRepository repository)
    {
        _repository = repository;
        Text = "Grupos de producto";
        Width = 1000;
        Height = 600;
        StartPosition = FormStartPosition.CenterScreen;
        Controls.Add(BuildLayout());
        Load += async (_, _) => await CargarAsync();
        dgv.SelectionChanged += (_, _) => CargarSeleccion();
        btnNuevo.Click += (_, _) => Limpiar();
        btnEditar.Click += (_, _) => CargarSeleccion();
        btnEliminar.Click += async (_, _) => await EliminarAsync();
        btnGuardar.Click += async (_, _) => await GuardarAsync();
    }

    private Control BuildLayout()
    {
        var datos = new TableLayoutPanel { Dock = DockStyle.Top, Height = 100, ColumnCount = 4, Padding = new Padding(8) };
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        datos.Controls.Add(new Label { Text = "Código", AutoSize = true }, 0, 0);
        datos.Controls.Add(txtCodigo, 1, 0);
        datos.Controls.Add(new Label { Text = "Descripción", AutoSize = true }, 2, 0);
        datos.Controls.Add(txtDescripcion, 3, 0);
        datos.Controls.Add(new Label { Text = "Resumido", AutoSize = true }, 0, 1);
        datos.Controls.Add(txtResumido, 1, 1);
        datos.Controls.Add(chkActivo, 3, 1);

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
        if (dgv.CurrentRow?.DataBoundItem is not GrupoProducto item) return;
        txtCodigo.Text = item.CodigoGrupo;
        txtCodigo.Enabled = false;
        txtDescripcion.Text = item.Detallado;
        txtResumido.Text = item.Resumido;
        chkActivo.Checked = item.Activo;
    }

    private void Limpiar()
    {
        txtCodigo.Enabled = true;
        txtCodigo.Clear();
        txtDescripcion.Clear();
        txtResumido.Clear();
        chkActivo.Checked = true;
    }

    private async Task GuardarAsync()
    {
        try
        {
            var existente = await _repository.ObtenerPorCodigoAsync(txtCodigo.Text.Trim());
            if (existente is null)
            {
                var nuevo = GrupoProducto.Crear(txtCodigo.Text, txtDescripcion.Text, txtResumido.Text, Environment.UserName, chkActivo.Checked);
                await _repository.InsertarAsync(nuevo);
            }
            else
            {
                existente.Actualizar(txtDescripcion.Text, txtResumido.Text, Environment.UserName, chkActivo.Checked);
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
        if (dgv.CurrentRow?.DataBoundItem is not GrupoProducto item) return;
        item.Inactivar();
        await _repository.ActualizarAsync(item);
        await CargarAsync();
    }
}
