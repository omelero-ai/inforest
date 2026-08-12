using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Maestros;

/// <summary>
/// Maestro WinForms de productos.
/// Legacy: frmProducto.frm.
/// </summary>
public class FrmProducto : Form
{
    private readonly IProductoMaestroRepository _repository;
    private readonly IGrupoProductoRepository _grupoRepository;
    private readonly ISubGrupoProductoRepository _subGrupoRepository;
    private readonly DataGridView dgv = new() { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
    private readonly TextBox txtCodigo = new();
    private readonly TextBox txtDescripcion = new();
    private readonly NumericUpDown nudPrecioVenta = new() { DecimalPlaces = 2, Maximum = 1000000, Increment = 0.50m };
    private readonly ComboBox cboGrupo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox cboSubGrupo = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly CheckBox chkActivo = new() { Text = "Activo", Checked = true, AutoSize = true };
    private readonly Button btnNuevo = new() { Text = "Nuevo" };
    private readonly Button btnEditar = new() { Text = "Editar" };
    private readonly Button btnEliminar = new() { Text = "Eliminar" };
    private readonly Button btnGuardar = new() { Text = "Guardar" };
    private List<ProductoMaestro> _items = [];
    private List<SubGrupoProducto> _subGrupos = [];

    public FrmProducto(IProductoMaestroRepository repository, IGrupoProductoRepository grupoRepository, ISubGrupoProductoRepository subGrupoRepository)
    {
        _repository = repository;
        _grupoRepository = grupoRepository;
        _subGrupoRepository = subGrupoRepository;
        Text = "Productos";
        Width = 1200;
        Height = 650;
        StartPosition = FormStartPosition.CenterScreen;
        Controls.Add(BuildLayout());
        Load += async (_, _) => await InicializarAsync();
        dgv.SelectionChanged += (_, _) => CargarSeleccion();
        cboGrupo.SelectedValueChanged += (_, _) => RefrescarSubGrupos();
        btnNuevo.Click += (_, _) => Limpiar();
        btnEditar.Click += (_, _) => CargarSeleccion();
        btnEliminar.Click += async (_, _) => await EliminarAsync();
        btnGuardar.Click += async (_, _) => await GuardarAsync();
    }

    private Control BuildLayout()
    {
        var datos = new TableLayoutPanel { Dock = DockStyle.Top, Height = 140, ColumnCount = 4, Padding = new Padding(8) };
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        datos.Controls.Add(new Label { Text = "Código", AutoSize = true }, 0, 0);
        datos.Controls.Add(txtCodigo, 1, 0);
        datos.Controls.Add(new Label { Text = "Descripción", AutoSize = true }, 2, 0);
        datos.Controls.Add(txtDescripcion, 3, 0);
        datos.Controls.Add(new Label { Text = "Grupo", AutoSize = true }, 0, 1);
        datos.Controls.Add(cboGrupo, 1, 1);
        datos.Controls.Add(new Label { Text = "Subgrupo", AutoSize = true }, 2, 1);
        datos.Controls.Add(cboSubGrupo, 3, 1);
        datos.Controls.Add(new Label { Text = "Precio venta", AutoSize = true }, 0, 2);
        datos.Controls.Add(nudPrecioVenta, 1, 2);
        datos.Controls.Add(chkActivo, 3, 2);

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(8) };
        acciones.Controls.AddRange([btnNuevo, btnEditar, btnEliminar, btnGuardar]);

        var panel = new Panel { Dock = DockStyle.Fill };
        panel.Controls.Add(dgv);
        panel.Controls.Add(acciones);
        panel.Controls.Add(datos);
        return panel;
    }

    private async Task InicializarAsync()
    {
        cboGrupo.DataSource = (await _grupoRepository.ObtenerTodosAsync()).Where(g => g.Activo).ToList();
        cboGrupo.DisplayMember = nameof(GrupoProducto.Detallado);
        cboGrupo.ValueMember = nameof(GrupoProducto.CodigoGrupo);
        _subGrupos = (await _subGrupoRepository.ObtenerTodosAsync()).Where(s => s.Activo).ToList();
        RefrescarSubGrupos();
        await CargarAsync();
    }

    private void RefrescarSubGrupos()
    {
        var codigoGrupo = cboGrupo.SelectedValue?.ToString();
        cboSubGrupo.DataSource = _subGrupos.Where(s => s.CodigoGrupo == codigoGrupo).ToList();
        cboSubGrupo.DisplayMember = nameof(SubGrupoProducto.Detallado);
        cboSubGrupo.ValueMember = nameof(SubGrupoProducto.CodigoSubGrupo);
    }

    private async Task CargarAsync()
    {
        _items = (await _repository.ObtenerTodosAsync()).ToList();
        dgv.DataSource = null;
        dgv.DataSource = _items;
    }

    private void CargarSeleccion()
    {
        if (dgv.CurrentRow?.DataBoundItem is not ProductoMaestro item) return;
        txtCodigo.Text = item.CodigoProducto;
        txtCodigo.Enabled = false;
        txtDescripcion.Text = item.Detallado;
        cboGrupo.SelectedValue = item.Grupo;
        RefrescarSubGrupos();
        cboSubGrupo.SelectedValue = item.SubGrupo;
        nudPrecioVenta.Value = item.PrecioVenta;
        chkActivo.Checked = item.Activo;
    }

    private void Limpiar()
    {
        txtCodigo.Enabled = true;
        txtCodigo.Clear();
        txtDescripcion.Clear();
        nudPrecioVenta.Value = 0;
        chkActivo.Checked = true;
    }

    private async Task GuardarAsync()
    {
        try
        {
            var grupo = cboGrupo.SelectedValue?.ToString() ?? string.Empty;
            var subGrupo = cboSubGrupo.SelectedValue?.ToString();
            var existente = await _repository.ObtenerPorCodigoAsync(txtCodigo.Text.Trim());
            if (existente is null)
            {
                var nuevo = ProductoMaestro.Crear(txtCodigo.Text, grupo, txtDescripcion.Text, Environment.UserName, nudPrecioVenta.Value, subGrupo, resumido: txtDescripcion.Text, activo: chkActivo.Checked);
                await _repository.InsertarAsync(nuevo);
            }
            else
            {
                existente.ActualizarBasico(grupo, subGrupo, txtDescripcion.Text, txtDescripcion.Text, nudPrecioVenta.Value, existente.PrecioDelivery, existente.PrecioLlevar, existente.PrecioCanal4, existente.PrecioCanal5, chkActivo.Checked, Environment.UserName);
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
        if (dgv.CurrentRow?.DataBoundItem is not ProductoMaestro item) return;
        item.ActualizarBasico(item.Grupo, item.SubGrupo, item.Detallado, item.Resumido, item.PrecioVenta, item.PrecioDelivery, item.PrecioLlevar, item.PrecioCanal4, item.PrecioCanal5, false, Environment.UserName);
        await _repository.ActualizarAsync(item);
        await CargarAsync();
    }
}
