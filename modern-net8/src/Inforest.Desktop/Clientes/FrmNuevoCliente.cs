using Inforest.Application.Maestros;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Clientes;

/// <summary>
/// Formulario para agregar o editar un cliente desde el flujo POS.
/// Legacy: frmNuevoCliente.frm — "Agregar Cliente"
/// BR-CLIENTE-001: Código único. BR-CLIENTE-002: Identidad válida (RUC 11 dígitos / DNI 8 dígitos).
/// BR-CLIENTE-003: Razón social obligatoria. BR-CLIENTE-004: Permite editar cliente existente.
/// </summary>
public class FrmNuevoCliente : Form
{
    private readonly IClienteRepository _repository;

    private readonly TextBox txtCodigo     = new() { MaxLength = 5 };
    private readonly TextBox txtEmpresa    = new() { MaxLength = 150 };
    private readonly TextBox txtIdentidad  = new() { MaxLength = 15 };
    private readonly TextBox txtDireccion  = new() { MaxLength = 200 };
    private readonly TextBox txtCorreo     = new() { MaxLength = 400 };
    private readonly ComboBox cboTipoId    = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly CheckBox chkActivo    = new() { Text = "Activo", Checked = true, AutoSize = true };
    private readonly Button btnGuardar     = new() { Text = "Guardar" };
    private readonly Button btnCancelar    = new() { Text = "Cancelar" };

    /// <summary>Código del cliente guardado; disponible tras cierre exitoso.</summary>
    public string? CodigoGuardado { get; private set; }

    /// <summary>
    /// Modo edición: si se provee un cliente existente, se carga en el formulario.
    /// </summary>
    public Cliente? ClienteExistente { get; set; }

    public FrmNuevoCliente(IClienteRepository repository)
    {
        _repository = repository;
        Text = "Agregar / Editar Cliente";
        Width = 520;
        Height = 420;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;

        Controls.Add(BuildLayout());
        cboTipoId.Items.AddRange(["RUC", "DNI", "CE", "OTRO"]);
        cboTipoId.SelectedIndex = 0;

        Load += OnLoad;
        btnGuardar.Click  += async (_, _) => await OnGuardarAsync();
        btnCancelar.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
    }

    private void OnLoad(object? sender, EventArgs e)
    {
        if (ClienteExistente is null) return;
        // BR-CLIENTE-004: carga datos del cliente existente para edición
        txtCodigo.Text    = ClienteExistente.CodigoCliente;
        txtCodigo.Enabled = false;
        txtEmpresa.Text   = ClienteExistente.Empresa;
        txtIdentidad.Text = ClienteExistente.Identidad;
        txtDireccion.Text = ClienteExistente.Direccion;
        txtCorreo.Text    = ClienteExistente.Correo ?? string.Empty;
        chkActivo.Checked = ClienteExistente.Activo;
        if (ClienteExistente.TipoIdentidad is not null)
            cboTipoId.SelectedItem = ClienteExistente.TipoIdentidad;
    }

    private async Task OnGuardarAsync()
    {
        if (string.IsNullOrWhiteSpace(txtCodigo.Text))
        { MessageBox.Show("El código del cliente es obligatorio.", "Validación"); return; }
        if (string.IsNullOrWhiteSpace(txtEmpresa.Text))
        { MessageBox.Show("La razón social es obligatoria.", "Validación"); return; }
        if (string.IsNullOrWhiteSpace(txtIdentidad.Text))
        { MessageBox.Show("La identidad es obligatoria.", "Validación"); return; }
        if (string.IsNullOrWhiteSpace(txtDireccion.Text))
        { MessageBox.Show("La dirección es obligatoria.", "Validación"); return; }

        try
        {
            var tipoId = cboTipoId.SelectedItem?.ToString();
            Result resultado;

            if (ClienteExistente is null)
            {
                // BR-CLIENTE-001: verificar unicidad antes de crear
                var cliente = Cliente.Crear(
                    txtCodigo.Text.Trim(), txtEmpresa.Text.Trim(),
                    txtIdentidad.Text.Trim(), txtDireccion.Text.Trim(),
                    Environment.UserName,
                    correo: string.IsNullOrWhiteSpace(txtCorreo.Text) ? null : txtCorreo.Text.Trim(),
                    tipoIdentidad: tipoId,
                    activo: chkActivo.Checked);

                var handler = new RegistrarNuevoClientePosHandler(_repository);
                resultado = await handler.HandleAsync(new RegistrarNuevoClientePosCommand(cliente));
            }
            else
            {
                ClienteExistente.Actualizar(
                    txtEmpresa.Text.Trim(), txtIdentidad.Text.Trim(),
                    txtDireccion.Text.Trim(), Environment.UserName,
                    correo: string.IsNullOrWhiteSpace(txtCorreo.Text) ? null : txtCorreo.Text.Trim(),
                    tipoIdentidad: tipoId, enlace: null, tipoCliente: null,
                    ubigeo: null, urbanizacion: null, tipoResponsable: null,
                    activo: chkActivo.Checked);

                var handler = new ActualizarClientePosHandler(_repository);
                resultado = await handler.HandleAsync(new ActualizarClientePosCommand(ClienteExistente));
            }

            if (!resultado.EsExitoso)
            { MessageBox.Show(resultado.MensajeError, "Error"); return; }

            CodigoGuardado = txtCodigo.Text.Trim();
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Domain.Exceptions.DomainException ex)
        {
            MessageBox.Show(ex.Message, "Validación");
        }
    }

    private Control BuildLayout()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 8,
            Padding = new Padding(12)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

        void AddRow(int row, string label, Control ctrl)
        {
            panel.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 0, row);
            ctrl.Dock = DockStyle.Fill;
            panel.Controls.Add(ctrl, 1, row);
        }

        AddRow(0, "Código:",          txtCodigo);
        AddRow(1, "Razón social:",     txtEmpresa);
        AddRow(2, "Tipo identidad:",   cboTipoId);
        AddRow(3, "Identidad:",        txtIdentidad);
        AddRow(4, "Dirección:",        txtDireccion);
        AddRow(5, "Correo:",           txtCorreo);
        panel.Controls.Add(chkActivo, 1, 6);

        var btnPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(4)
        };
        btnCancelar.Width = btnGuardar.Width = 90;
        btnPanel.Controls.AddRange([btnCancelar, btnGuardar]);
        panel.Controls.Add(btnPanel, 0, 7);
        panel.SetColumnSpan(btnPanel, 2);

        return panel;
    }
}
