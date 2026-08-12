using Inforest.Application.Delivery;
using Inforest.Domain.Entities.Delivery;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Formulario para crear/editar cliente delivery frecuente.
/// <para>
/// Legacy: <c>frmNuevoDelivery.frm</c> — "Agregar Cliente Frecuente".
/// Campos: teléfono, apellido, nombre, dirección, zona, distrito, tarjeta, tipo cliente, identidad.
/// Permite búsqueda en Central (CENTRALDELIVERY) por número de identidad.
/// </para>
/// Reglas: BR-DEL-001, BR-DEL-008.
/// </summary>
public partial class NuevoDeliveryForm : Form
{
    private readonly CrearClienteDeliveryHandler _crearHandler;
    private readonly ActualizarClienteDeliveryHandler _actualizarHandler;
    private readonly BuscarClienteDeliveryHandler _buscarHandler;

    public ClienteDelivery? ClienteResultado { get; private set; }

    public NuevoDeliveryForm(
        CrearClienteDeliveryHandler crearHandler,
        ActualizarClienteDeliveryHandler actualizarHandler,
        BuscarClienteDeliveryHandler buscarHandler)
    {
        _crearHandler = crearHandler;
        _actualizarHandler = actualizarHandler;
        _buscarHandler = buscarHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = "Agregar Cliente Frecuente";
        Size = new Size(550, 500);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var yPos = 15;
        const int labelWidth = 110;
        const int inputLeft = 125;
        const int inputWidth = 300;

        Label MkLabel(string text) => new() { Text = text, Left = 10, Top = yPos, Width = labelWidth };
        TextBox MkTextBox(string name) { var tb = new TextBox { Name = name, Left = inputLeft, Top = yPos, Width = inputWidth }; yPos += 30; return tb; }

        var lblTelefono = MkLabel("Teléfono:"); var txtTelefono = MkTextBox("txtTelefono");
        var lblApellido = MkLabel("Apellido:"); var txtApellido = MkTextBox("txtApellido");
        var lblNombre = MkLabel("Nombre:"); var txtNombre = MkTextBox("txtNombre");
        var lblDireccion = MkLabel("Dirección:"); var txtDireccion = MkTextBox("txtDireccion");
        var lblReferencia = MkLabel("Referencia:"); var txtReferencia = MkTextBox("txtReferencia");
        var lblDistrito = MkLabel("Distrito:"); var txtDistrito = MkTextBox("txtDistrito");
        var lblZona = MkLabel("Zona:"); var txtZona = MkTextBox("txtZona");
        var lblObservacion = MkLabel("Observación:"); var txtObservacion = MkTextBox("txtObservacion");
        var lblEmail = MkLabel("Email:"); var txtEmail = MkTextBox("txtEmail");

        var btnGrabar = new Button { Text = "Grabar", Left = 130, Top = yPos + 10, Width = 90, DialogResult = DialogResult.None };
        var btnCancelar = new Button { Text = "Cancelar", Left = 230, Top = yPos + 10, Width = 90, DialogResult = DialogResult.Cancel };

        Controls.AddRange([
            lblTelefono, txtTelefono, lblApellido, txtApellido,
            lblNombre, txtNombre, lblDireccion, txtDireccion,
            lblReferencia, txtReferencia, lblDistrito, txtDistrito,
            lblZona, txtZona, lblObservacion, txtObservacion,
            lblEmail, txtEmail, btnGrabar, btnCancelar
        ]);

        btnGrabar.Click += async (_, _) =>
        {
            var cmd = new CrearClienteDeliveryCommand(
                Guid.NewGuid().ToString("N")[..7].ToUpper(),
                null,
                txtApellido.Text.Trim(),
                txtNombre.Text.Trim(),
                txtTelefono.Text.Trim(),
                txtDireccion.Text.Trim(),
                txtReferencia.Text.Trim(),
                txtZona.Text.Trim(),
                txtDistrito.Text.Trim(),
                null, null, null, null,
                txtObservacion.Text.Trim(),
                txtEmail.Text.Trim());

            var result = await _crearHandler.HandleAsync(cmd);
            if (result.EsExitoso)
                DialogResult = DialogResult.OK;
            else
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        };
    }
}
