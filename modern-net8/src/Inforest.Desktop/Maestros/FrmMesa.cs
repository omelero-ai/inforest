using System.Drawing;
using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Maestros;

/// <summary>
/// Maestro WinForms de mesas por salón.
/// Legacy: frmMesa.frm / frmMesas.frm.
/// </summary>
public class FrmMesa : Form
{
    private readonly IMesaRepository _repository;
    private readonly ISalonRepository _salonRepository;
    private readonly DataGridView dgv = new() { Dock = DockStyle.Fill, ReadOnly = true, AutoGenerateColumns = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
    private readonly TextBox txtCodigo = new();
    private readonly TextBox txtDescripcion = new();
    private readonly TextBox txtResumido = new();
    private readonly NumericUpDown nudPersonas = new() { Minimum = 1, Maximum = 50, Value = 4 };
    private readonly ComboBox cboSalon = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox cboEstado = new() { DropDownStyle = ComboBoxStyle.DropDownList, DataSource = Enum.GetValues<EstadoMesa>() };
    private readonly CheckBox chkFumador = new() { Text = "Fumador", AutoSize = true };
    private readonly CheckBox chkActivo = new() { Text = "Activo", Checked = true, AutoSize = true };
    private readonly Button btnNuevo = new() { Text = "Nuevo" };
    private readonly Button btnEditar = new() { Text = "Editar" };
    private readonly Button btnEliminar = new() { Text = "Eliminar" };
    private readonly Button btnGuardar = new() { Text = "Guardar" };
    private List<Mesa> _items = [];

    public FrmMesa(IMesaRepository repository, ISalonRepository salonRepository)
    {
        _repository = repository;
        _salonRepository = salonRepository;
        Text = "Mesas";
        Width = 1200;
        Height = 650;
        StartPosition = FormStartPosition.CenterScreen;
        Controls.Add(BuildLayout());
        Load += async (_, _) => await InicializarAsync();
        dgv.SelectionChanged += (_, _) => CargarSeleccion();
        dgv.CellFormatting += Dgv_CellFormatting;
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
        datos.Controls.Add(new Label { Text = "Descripción", AutoSize = true }, 2, 0);
        datos.Controls.Add(txtDescripcion, 3, 0);
        datos.Controls.Add(new Label { Text = "Resumido", AutoSize = true }, 0, 1);
        datos.Controls.Add(txtResumido, 1, 1);
        datos.Controls.Add(new Label { Text = "Salón", AutoSize = true }, 2, 1);
        datos.Controls.Add(cboSalon, 3, 1);
        datos.Controls.Add(new Label { Text = "Personas", AutoSize = true }, 0, 2);
        datos.Controls.Add(nudPersonas, 1, 2);
        datos.Controls.Add(new Label { Text = "Estado", AutoSize = true }, 2, 2);
        datos.Controls.Add(cboEstado, 3, 2);
        datos.Controls.Add(chkFumador, 1, 3);
        datos.Controls.Add(chkActivo, 3, 3);

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
        var salones = (await _salonRepository.ObtenerTodosAsync()).Where(s => s.Activo).ToList();
        cboSalon.DataSource = salones;
        cboSalon.DisplayMember = nameof(Salon.Detallado);
        cboSalon.ValueMember = nameof(Salon.CodigoSalon);
        await CargarAsync();
    }

    private async Task CargarAsync()
    {
        _items = (await _repository.ObtenerTodosAsync()).ToList();
        dgv.DataSource = null;
        dgv.DataSource = _items;
    }

    private void CargarSeleccion()
    {
        if (dgv.CurrentRow?.DataBoundItem is not Mesa item) return;
        txtCodigo.Text = item.CodigoMesa;
        txtCodigo.Enabled = false;
        txtDescripcion.Text = item.Detallado;
        txtResumido.Text = item.Resumido;
        cboSalon.SelectedValue = item.CodigoSalon;
        nudPersonas.Value = item.NumeroPersonas;
        cboEstado.SelectedItem = item.Estado;
        chkFumador.Checked = item.Fumador;
        chkActivo.Checked = item.Activo;
    }

    private void Limpiar()
    {
        txtCodigo.Enabled = true;
        txtCodigo.Clear();
        txtDescripcion.Clear();
        txtResumido.Clear();
        nudPersonas.Value = 4;
        cboEstado.SelectedItem = EstadoMesa.Libre;
        chkFumador.Checked = false;
        chkActivo.Checked = true;
    }

    private async Task GuardarAsync()
    {
        try
        {
            var codigoSalon = cboSalon.SelectedValue?.ToString() ?? string.Empty;
            var estado = cboEstado.SelectedItem is EstadoMesa valor ? valor : EstadoMesa.Libre;
            var existente = await _repository.ObtenerPorCodigoAsync(txtCodigo.Text.Trim());
            if (existente is null)
            {
                var nueva = Mesa.Crear(txtCodigo.Text, txtDescripcion.Text, txtResumido.Text, codigoSalon, chkFumador.Checked, null, null, (int)nudPersonas.Value, Environment.UserName, estado, chkActivo.Checked);
                await _repository.InsertarAsync(nueva);
            }
            else
            {
                existente.Actualizar(txtDescripcion.Text, txtResumido.Text, codigoSalon, chkFumador.Checked, null, null, (int)nudPersonas.Value, Environment.UserName, chkActivo.Checked);
                switch (estado)
                {
                    case EstadoMesa.Libre: existente.Liberar(); break;
                    case EstadoMesa.Ocupada: if (existente.Estado != EstadoMesa.Ocupada) existente.Ocupar(); break;
                    case EstadoMesa.Reservada: existente.Reservar(); break;
                    case EstadoMesa.EnCuenta: existente.MarcarEnCuenta(); break;
                }
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
        if (dgv.CurrentRow?.DataBoundItem is not Mesa item) return;
        item.Inactivar();
        await _repository.ActualizarAsync(item);
        await CargarAsync();
    }

    private void Dgv_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (dgv.Rows[e.RowIndex].DataBoundItem is not Mesa mesa) return;
        dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = mesa.Estado switch
        {
            EstadoMesa.Ocupada => Color.MistyRose,
            EstadoMesa.Reservada => Color.LightGoldenrodYellow,
            EstadoMesa.EnCuenta => Color.LightBlue,
            _ => Color.Honeydew
        };
    }
}
