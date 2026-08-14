using Inforest.Application.Delivery;
using Inforest.Domain.Entities.Delivery;
using System.IO;

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
    private readonly ActualizarFotoClienteDeliveryHandler _actualizarFotoHandler;
    private readonly BuscarClienteDeliveryHandler _buscarHandler;

    public ClienteDelivery? ClienteResultado { get; private set; }

    public NuevoDeliveryForm(
        CrearClienteDeliveryHandler crearHandler,
        ActualizarClienteDeliveryHandler actualizarHandler,
        ActualizarFotoClienteDeliveryHandler actualizarFotoHandler,
        BuscarClienteDeliveryHandler buscarHandler)
    {
        _crearHandler = crearHandler;
        _actualizarHandler = actualizarHandler;
        _actualizarFotoHandler = actualizarFotoHandler;
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
        var lblFoto = MkLabel("Foto:"); yPos += 30;
        string? fotoSeleccionada = null;
        var txtFoto = new TextBox { Name = "txtFoto", Left = inputLeft, Top = yPos - 30, Width = 220, ReadOnly = true };
        var btnFoto = new Button { Text = "Seleccionar", Left = inputLeft + 230, Top = yPos - 31, Width = 90 };

        var btnGrabar = new Button { Text = "Grabar", Left = 130, Top = yPos + 10, Width = 90, DialogResult = DialogResult.None };
        var btnCancelar = new Button { Text = "Cancelar", Left = 230, Top = yPos + 10, Width = 90, DialogResult = DialogResult.Cancel };

        Controls.AddRange([
            lblTelefono, txtTelefono, lblApellido, txtApellido,
            lblNombre, txtNombre, lblDireccion, txtDireccion,
            lblReferencia, txtReferencia, lblDistrito, txtDistrito,
            lblZona, txtZona, lblObservacion, txtObservacion,
            lblEmail, txtEmail, lblFoto, txtFoto, btnFoto, btnGrabar, btnCancelar
        ]);

        btnFoto.Click += (_, _) =>
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Imágenes|*.jpg;*.jpeg;*.png;*.bmp|Todos|*.*",
                Multiselect = false,
                CheckFileExists = true
            };
            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            fotoSeleccionada = dialog.FileName;
            txtFoto.Text = Path.GetFileName(fotoSeleccionada);
        };

        btnGrabar.Click += async (_, _) =>
        {
            var codigoDelivery = Guid.NewGuid().ToString("N")[..7].ToUpper();
            var cmd = new CrearClienteDeliveryCommand(
                codigoDelivery,
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
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!string.IsNullOrWhiteSpace(fotoSeleccionada))
            {
                var fotoBytes = await System.IO.File.ReadAllBytesAsync(fotoSeleccionada);
                var fotoResult = await _actualizarFotoHandler.HandleAsync(
                    new ActualizarFotoClienteDeliveryCommand(codigoDelivery, fotoBytes));
                if (!fotoResult.EsExitoso)
                {
                    MessageBox.Show(fotoResult.MensajeError, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            DialogResult = DialogResult.OK;
        };
    }
}
