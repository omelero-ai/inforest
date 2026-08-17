using Inforest.Application.Delivery;
using Inforest.Domain.Entities.Delivery;
using System.IO;

namespace Inforest.Desktop.Delivery;

/// <summary>
/// Formulario para crear/editar cliente delivery frecuente.
/// <para>
/// Legacy: <c>frmNuevoDelivery.frm</c>.
/// </para>
/// Reglas: BR-DEL-001, BR-DEL-011.
/// </summary>
public partial class NuevoDeliveryForm : Form
{
    private readonly string? _codigoDeliveryEdicion;
    private readonly ObtenerSiguienteCodigoClienteDeliveryHandler _codigoHandler;
    private readonly ObtenerClienteDeliveryPorCodigoHandler _obtenerClienteHandler;
    private readonly CrearClienteDeliveryHandler _crearHandler;
    private readonly ActualizarClienteDeliveryHandler _actualizarHandler;
    private readonly ActualizarFotoClienteDeliveryHandler _actualizarFotoHandler;
    private TextBox _txtTelefono = null!;
    private TextBox _txtApellido = null!;
    private TextBox _txtNombre = null!;
    private TextBox _txtDireccion = null!;
    private TextBox _txtReferencia = null!;
    private TextBox _txtDistrito = null!;
    private TextBox _txtZona = null!;
    private TextBox _txtObservacion = null!;
    private TextBox _txtEmail = null!;
    private TextBox _txtTipoCliente = null!;
    private TextBox _txtTipoIdentidad = null!;
    private TextBox _txtNumeroIdentidad = null!;
    private TextBox _txtFoto = null!;
    private string? _fotoSeleccionada;

    public ClienteDelivery? ClienteResultado { get; private set; }

    public NuevoDeliveryForm(
        ObtenerSiguienteCodigoClienteDeliveryHandler codigoHandler,
        ObtenerClienteDeliveryPorCodigoHandler obtenerClienteHandler,
        CrearClienteDeliveryHandler crearHandler,
        ActualizarClienteDeliveryHandler actualizarHandler,
        ActualizarFotoClienteDeliveryHandler actualizarFotoHandler,
        string? codigoDeliveryEdicion = null)
    {
        _codigoHandler = codigoHandler;
        _obtenerClienteHandler = obtenerClienteHandler;
        _crearHandler = crearHandler;
        _actualizarHandler = actualizarHandler;
        _actualizarFotoHandler = actualizarFotoHandler;
        _codigoDeliveryEdicion = codigoDeliveryEdicion;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        var esEdicion = !string.IsNullOrWhiteSpace(_codigoDeliveryEdicion);

        Text = esEdicion ? "Modificar Cliente Delivery" : "Agregar Cliente Frecuente";
        Size = new Size(620, 700);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.FromArgb(245, 246, 247);

        var header = new Label
        {
            Dock = DockStyle.Top,
            Height = 46,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(14, 0, 0, 0),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            BackColor = Color.FromArgb(235, 236, 240),
            Text = esEdicion ? "Editar cliente delivery frecuente" : "Registrar cliente delivery frecuente"
        };

        var body = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16, 14, 16, 14) };
        var yPos = 8;
        const int labelWidth = 120;
        const int inputLeft = 130;
        const int inputWidth = 430;

        Label MkLabel(string text) => new()
        {
            Text = text,
            Left = 0,
            Top = yPos + 4,
            Width = labelWidth,
            TextAlign = ContentAlignment.MiddleRight
        };
        TextBox MkTextBox(string name)
        {
            var tb = new TextBox
            {
                Name = name,
                Left = inputLeft,
                Top = yPos,
                Width = inputWidth,
                Font = new Font("Segoe UI", 9F)
            };
            yPos += 34;
            return tb;
        }

        var lblTelefono = MkLabel("Teléfono:");
        _txtTelefono = MkTextBox("txtTelefono");
        var lblApellido = MkLabel("Apellido:");
        _txtApellido = MkTextBox("txtApellido");
        var lblNombre = MkLabel("Nombre:");
        _txtNombre = MkTextBox("txtNombre");
        var lblDireccion = MkLabel("Dirección:");
        _txtDireccion = MkTextBox("txtDireccion");
        var lblReferencia = MkLabel("Referencia:");
        _txtReferencia = MkTextBox("txtReferencia");
        var lblDistrito = MkLabel("Distrito:");
        _txtDistrito = MkTextBox("txtDistrito");
        var lblZona = MkLabel("Zona:");
        _txtZona = MkTextBox("txtZona");
        var lblTipoCliente = MkLabel("Tipo cliente:");
        _txtTipoCliente = MkTextBox("txtTipoCliente");
        var lblTipoIdentidad = MkLabel("Tipo identidad:");
        _txtTipoIdentidad = MkTextBox("txtTipoIdentidad");
        var lblNumeroIdentidad = MkLabel("Nro identidad:");
        _txtNumeroIdentidad = MkTextBox("txtNumeroIdentidad");
        var lblObservacion = MkLabel("Observación:");
        _txtObservacion = MkTextBox("txtObservacion");
        var lblEmail = MkLabel("Email:");
        _txtEmail = MkTextBox("txtEmail");
        var lblFoto = MkLabel("Foto:");
        _txtFoto = new TextBox { Name = "txtFoto", Left = inputLeft, Top = yPos, Width = 320, ReadOnly = true, Font = new Font("Segoe UI", 9F) };
        var btnFoto = new Button
        {
            Text = "Seleccionar",
            Left = inputLeft + 330,
            Top = yPos - 1,
            Width = 100,
            Height = 28
        };
        yPos += 40;

        var footer = new Panel { Dock = DockStyle.Bottom, Height = 54 };
        var btnGrabar = new Button
        {
            Text = "Grabar",
            Width = 110,
            Height = 34,
            Top = 10,
            Left = Width - 270,
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            BackColor = Color.FromArgb(37, 99, 235),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        btnGrabar.FlatAppearance.BorderSize = 0;
        var btnCancelar = new Button
        {
            Text = "Cancelar",
            Width = 110,
            Height = 34,
            Top = 10,
            Left = Width - 150,
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        btnCancelar.FlatAppearance.BorderSize = 0;
        btnCancelar.Click += (_, _) =>
        {
            DialogResult = DialogResult.Cancel;
            Close();
        };
        footer.Controls.AddRange([btnGrabar, btnCancelar]);

        body.Controls.AddRange([
            lblTelefono, _txtTelefono, lblApellido, _txtApellido,
            lblNombre, _txtNombre, lblDireccion, _txtDireccion,
            lblReferencia, _txtReferencia, lblDistrito, _txtDistrito,
            lblZona, _txtZona, lblTipoCliente, _txtTipoCliente,
            lblTipoIdentidad, _txtTipoIdentidad, lblNumeroIdentidad, _txtNumeroIdentidad,
            lblObservacion, _txtObservacion, lblEmail, _txtEmail,
            lblFoto, _txtFoto, btnFoto
        ]);
        Controls.AddRange([body, footer, header]);

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

            _fotoSeleccionada = dialog.FileName;
            _txtFoto.Text = Path.GetFileName(_fotoSeleccionada);
        };

        Load += async (_, _) =>
        {
            if (esEdicion)
                await CargarClienteEdicionAsync();
        };

        btnGrabar.Click += async (_, _) => await GuardarAsync();
    }

    private async Task CargarClienteEdicionAsync()
    {
        var result = await _obtenerClienteHandler.HandleAsync(new ObtenerClienteDeliveryPorCodigoQuery(_codigoDeliveryEdicion!));
        if (!result.EsExitoso || result.Valor is null)
        {
            MessageBox.Show("No se encontró el cliente delivery a editar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            DialogResult = DialogResult.Cancel;
            Close();
            return;
        }

        var cliente = result.Valor;
        _txtTelefono.Text = cliente.Telefono ?? string.Empty;
        _txtApellido.Text = cliente.Apellido ?? string.Empty;
        _txtNombre.Text = cliente.Nombre ?? string.Empty;
        _txtDireccion.Text = cliente.Direccion ?? string.Empty;
        _txtReferencia.Text = cliente.Referencia ?? string.Empty;
        _txtDistrito.Text = cliente.CodigoDistrito ?? string.Empty;
        _txtZona.Text = cliente.CodigoZona ?? string.Empty;
        _txtTipoCliente.Text = cliente.TipoCliente ?? string.Empty;
        _txtTipoIdentidad.Text = cliente.TipoIdentidad ?? string.Empty;
        _txtNumeroIdentidad.Text = cliente.NumeroIdentidad ?? string.Empty;
        _txtObservacion.Text = cliente.Observacion ?? string.Empty;
        _txtEmail.Text = cliente.Email ?? string.Empty;
    }

    private async Task GuardarAsync()
    {
        NormalizarTexto();

        if (string.IsNullOrWhiteSpace(_txtTipoCliente.Text))
        {
            MessageBox.Show("Ingrese Tipo de Cliente", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(_txtTelefono.Text) && string.IsNullOrWhiteSpace(_txtNumeroIdentidad.Text))
        {
            MessageBox.Show("Ingrese el Teléfono o Id", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var codigoDelivery = _codigoDeliveryEdicion;

        if (string.IsNullOrWhiteSpace(codigoDelivery))
        {
            var codigoResult = await _codigoHandler.HandleAsync(new ObtenerSiguienteCodigoClienteDeliveryQuery());
            if (!codigoResult.EsExitoso || string.IsNullOrWhiteSpace(codigoResult.Valor))
            {
                MessageBox.Show("No se pudo generar el código de cliente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            codigoDelivery = codigoResult.Valor;
            var cmd = new CrearClienteDeliveryCommand(
                codigoDelivery,
                _txtTipoCliente.Text.Trim(),
                _txtApellido.Text.Trim(),
                _txtNombre.Text.Trim(),
                _txtTelefono.Text.Trim(),
                _txtDireccion.Text.Trim(),
                _txtReferencia.Text.Trim(),
                _txtZona.Text.Trim(),
                _txtDistrito.Text.Trim(),
                _txtTipoIdentidad.Text.Trim(),
                _txtNumeroIdentidad.Text.Trim(),
                null,
                null,
                _txtObservacion.Text.Trim(),
                _txtEmail.Text.Trim());

            var result = await _crearHandler.HandleAsync(cmd);
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        else
        {
            var cmd = new ActualizarClienteDeliveryCommand(
                codigoDelivery,
                _txtApellido.Text.Trim(),
                _txtNombre.Text.Trim(),
                _txtTelefono.Text.Trim(),
                _txtDireccion.Text.Trim(),
                _txtReferencia.Text.Trim(),
                _txtZona.Text.Trim(),
                _txtDistrito.Text.Trim(),
                _txtTipoIdentidad.Text.Trim(),
                _txtNumeroIdentidad.Text.Trim(),
                _txtObservacion.Text.Trim(),
                _txtEmail.Text.Trim());

            var result = await _actualizarHandler.HandleAsync(cmd);
            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        if (!string.IsNullOrWhiteSpace(_fotoSeleccionada))
        {
            var fotoBytes = await File.ReadAllBytesAsync(_fotoSeleccionada);
            var fotoResult = await _actualizarFotoHandler.HandleAsync(
                new ActualizarFotoClienteDeliveryCommand(codigoDelivery, fotoBytes));
            if (!fotoResult.EsExitoso)
            {
                MessageBox.Show(fotoResult.MensajeError, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        var clienteResult = await _obtenerClienteHandler.HandleAsync(new ObtenerClienteDeliveryPorCodigoQuery(codigoDelivery));
        ClienteResultado = clienteResult.Valor;
        DialogResult = DialogResult.OK;
        Close();
    }

    private void NormalizarTexto()
    {
        _txtApellido.Text = _txtApellido.Text.Trim().ToUpperInvariant();
        _txtNombre.Text = _txtNombre.Text.Trim().ToUpperInvariant();
        _txtDireccion.Text = _txtDireccion.Text.Trim().ToUpperInvariant();
        _txtReferencia.Text = _txtReferencia.Text.Trim().ToUpperInvariant();
        _txtObservacion.Text = _txtObservacion.Text.Trim().ToUpperInvariant();
    }
}
