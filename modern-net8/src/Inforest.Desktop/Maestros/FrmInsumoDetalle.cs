using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Maestros;

/// <summary>
/// Detalle/edición de insumo — equivalente a frmInsumoDetalle.frm.
/// Legacy: frmInsumoDetalle.frm / usp_agregarinsumos / USP_MODIFICARINSUMOS.
/// BR-INSUMO-001, BR-INSUMO-002, BR-INSUMO-003.
/// </summary>
public class FrmInsumoDetalle : Form
{
    private readonly IInsumoRepository _repository;
    private readonly Insumo? _insumoExistente;
    private readonly bool _moduloInfoRest;

    private readonly RadioButton optInsumo   = new() { Text = "Insumo",  Checked = true, AutoSize = true };
    private readonly RadioButton optProducto = new() { Text = "Plato/Producto", AutoSize = true };
    private readonly TextBox txtCodigo       = new() { ReadOnly = true, BackColor = SystemColors.Control };
    private readonly TextBox txtDescripcion  = new() { MaxLength = 50 };
    private readonly TextBox txtStock        = new() { Text = "0", MaxLength = 8 };
    private readonly CheckBox chkActivo      = new() { Text = "Activo", Checked = true, AutoSize = true };
    private readonly Button btnGuardar       = new() { Text = "Guardar", Width = 90 };
    private readonly Button btnCancelar      = new() { Text = "Cancelar", Width = 90 };

    /// <summary>Código de caja del contexto (capturado de sesión o vacío).</summary>
    public string Caja { get; set; } = string.Empty;
    /// <summary>Usuario del contexto (capturado de sesión).</summary>
    public string Usuario { get; set; } = string.Empty;

    public FrmInsumoDetalle(
        IInsumoRepository repository,
        Insumo? insumoExistente,
        bool moduloInfoRest = true)
    {
        _repository = repository;
        _insumoExistente = insumoExistente;
        _moduloInfoRest = moduloInfoRest;

        Text = "Mantenimiento de Insumos/Platos de Stock Crítico";
        Width = 500;
        Height = 280;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;

        Controls.Add(BuildLayout());

        Load += (_, _) => CargarDatos();
        btnGuardar.Click += async (_, _) => await GuardarAsync();
        btnCancelar.Click += (_, _) => Close();

        // BR-INSUMO-004: si es Plato, el botón de búsqueda estaría visible
        optInsumo.CheckedChanged += (_, _) => txtDescripcion.Enabled = optInsumo.Checked;
        optProducto.CheckedChanged += (_, _) => txtDescripcion.Enabled = !optProducto.Checked;
    }

    private Control BuildLayout()
    {
        var datos = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 6,
            Padding = new Padding(12),
            AutoSize = true
        };
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        datos.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        datos.Controls.Add(new Label { Text = "Código:", AutoSize = true }, 0, 0);
        datos.Controls.Add(txtCodigo, 1, 0);
        datos.Controls.Add(new Label { Text = "Tipo:", AutoSize = true }, 0, 1);

        var tipoPanel = new FlowLayoutPanel { AutoSize = true };
        tipoPanel.Controls.AddRange([optInsumo, optProducto]);
        datos.Controls.Add(tipoPanel, 1, 1);

        datos.Controls.Add(new Label { Text = "Descripción:", AutoSize = true }, 0, 2);
        datos.Controls.Add(txtDescripcion, 1, 2);

        // Stock solo visible en módulo INFOREST
        var lblStock = new Label { Text = "Cant. Disponible:", AutoSize = true, Visible = _moduloInfoRest };
        txtStock.Visible = _moduloInfoRest;
        datos.Controls.Add(lblStock, 0, 3);
        datos.Controls.Add(txtStock, 1, 3);

        datos.Controls.Add(chkActivo, 1, 4);

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(6) };
        acciones.Controls.AddRange([btnCancelar, btnGuardar]);

        var panel = new Panel { Dock = DockStyle.Fill };
        panel.Controls.Add(datos);
        panel.Controls.Add(acciones);
        return panel;
    }

    private void CargarDatos()
    {
        if (_insumoExistente is null)
        {
            // Nuevo
            txtCodigo.Text = "(auto)";
            chkActivo.Checked = true;
        }
        else
        {
            txtCodigo.Text      = _insumoExistente.Codigo;
            txtDescripcion.Text = _insumoExistente.Descripcion;
            txtStock.Text       = _insumoExistente.Stock.ToString("F2");
            chkActivo.Checked   = _insumoExistente.Activo;
            optInsumo.Checked   = _insumoExistente.EsInsumo;
            optProducto.Checked = !_insumoExistente.EsInsumo;
        }
    }

    private async Task GuardarAsync()
    {
        // BR-INSUMO-001: descripción requerida
        if (string.IsNullOrWhiteSpace(txtDescripcion.Text))
        {
            MessageBox.Show("Ingrese el Insumo", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            txtDescripcion.Focus();
            return;
        }

        if (!double.TryParse(txtStock.Text, out double stock)) stock = 0;

        if (_insumoExistente is null)
        {
            // Agregar
            var agregarHandler = new AgregarInsumoHandler(_repository);
            var cmd = new AgregarInsumoCommand(
                txtDescripcion.Text,
                Usuario,
                stock,
                optInsumo.Checked,
                Caja,
                _moduloInfoRest);
            var result = await agregarHandler.HandleAsync(cmd);
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            MessageBox.Show("Registro agregado.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            // Modificar
            var modificarHandler = new ModificarInsumoHandler(_repository);
            var cmd = new ModificarInsumoCommand(
                _insumoExistente.Codigo,
                txtDescripcion.Text,
                Usuario,
                stock,
                optInsumo.Checked,
                chkActivo.Checked,
                Caja,
                _moduloInfoRest);
            var result = await modificarHandler.HandleAsync(cmd);
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            MessageBox.Show("Registro modificado.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        DialogResult = DialogResult.OK;
        Close();
    }
}
