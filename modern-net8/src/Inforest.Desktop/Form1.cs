using Inforest.Application.Interfaces;
using Inforest.Application.Configuracion;
using Inforest.Desktop.POS;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Inforest.Desktop;

/// <summary>
/// Formulario inicial de acceso.
/// Legacy: frmAcceso.frm.
/// </summary>
public partial class Form1 : Form
{
    private readonly IAuthService _authService;
    private readonly ILicenseService _licenseService;
    private readonly ValidarInicioPosHandler _validarInicioPosHandler;
    private readonly IConfiguration _configuration;
    private readonly IServiceProvider _serviceProvider;
    private InicioPosValidado? _inicioPos;

    public Form1(
        IAuthService authService,
        ILicenseService licenseService,
        ValidarInicioPosHandler validarInicioPosHandler,
        IConfiguration configuration,
        IServiceProvider serviceProvider)
    {
        _authService = authService;
        _licenseService = licenseService;
        _validarInicioPosHandler = validarInicioPosHandler;
        _configuration = configuration;
        _serviceProvider = serviceProvider;
        InitializeComponent();
    }

    private async void Form1_Load(object sender, EventArgs e)
    {
        txtUsuario.Focus();
        await ValidarLicenciaAsync();
        await ValidarInicioPosAsync();
    }

    private async void btnIngresar_Click(object sender, EventArgs e)
    {
        btnIngresar.Enabled = false;
        lblEstado.Text = "Autenticando...";

        try
        {
            var request = new AuthRequest(
                txtUsuario.Text,
                txtPassword.Text,
                _configuration["Inforest:Modulo"] ?? "INFOREST",
                _inicioPos?.CodigoCaja ?? _configuration["Inforest:CodigoCaja"] ?? "01",
                _configuration["Inforest:CodigoTerminal"] ?? Environment.MachineName,
                ResolveDatabaseName());

            var result = await _authService.AutenticarAsync(request);
            if (!result.Exitoso || result.Sesion is null)
            {
                lblEstado.Text = result.MensajeError ?? "No se pudo iniciar sesión.";
                txtPassword.Clear();
                txtPassword.Focus();
                return;
            }

            lblEstado.Text = $"Sesión activa: {result.Sesion.CodigoUsuario} / Caja {result.Sesion.CodigoCaja}";
            lblSesion.Text = $"Módulo {result.Sesion.Modulo} · {result.Sesion.Permisos.Count} permisos cargados";
            txtPassword.Clear();

            // Navegar al módulo principal tras login exitoso
            AbrirShellPrincipal();
        }
        finally
        {
            btnIngresar.Enabled = true;
        }
    }

    private void AbrirShellPrincipal()
    {
        try
        {
            var frmPOS = _serviceProvider.GetRequiredService<FrmPuntoVenta>();
            frmPOS.FormClosed += (_, _) => Show();
            Hide();
            frmPOS.Show();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"No se pudo abrir el módulo POS: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void btnCancelar_Click(object sender, EventArgs e)
    {
        Close();
    }

    protected override async void OnFormClosing(FormClosingEventArgs e)
    {
        if (_authService is not null)
            await _authService.CerrarSesionAsync();

        base.OnFormClosing(e);
    }

    private async Task ValidarLicenciaAsync()
    {
        var request = new LicenseValidationRequest(
            Habilitado: bool.TryParse(_configuration["Inforest:Licensing:Enabled"], out var enabled) && enabled,
            ModuloEjecutable: _configuration["Inforest:Modulo"] ?? "INFOREST",
            BaseDatos: ResolveDatabaseName(),
            RucEmpresa: _configuration["Inforest:Licensing:RucEmpresa"],
            OmitirHardwareLegacy: true);

        var result = await _licenseService.ValidarLicenciaAsync(request);
        if (!result.EsValida)
        {
            lblEstado.Text = result.Mensaje ?? "Licencia inválida.";
            btnIngresar.Enabled = false;
            return;
        }

        lblEstado.Text = result.Advertencia ?? "Licencia validada.";
    }

    private async Task ValidarInicioPosAsync()
    {
        var codigoCaja = _configuration["Inforest:CodigoCaja"] ?? "01";
        var inicio = await _validarInicioPosHandler.HandleAsync(new ValidarInicioPosQuery(codigoCaja));
        if (!inicio.EsExitoso || inicio.Valor is null)
        {
            lblEstado.Text = inicio.MensajeError ?? "No se pudo cargar la configuración del POS.";
            btnIngresar.Enabled = false;
            return;
        }

        _inicioPos = inicio.Valor;
        lblSesion.Text = $"Caja {_inicioPos.CodigoCaja}: {(_inicioPos.RequiereLogin ? "requiere login" : "acceso directo MCPV")}";

        if (btnIngresar.Enabled && !_inicioPos.RequiereLogin)
            AbrirShellPrincipal();
    }

    private string ResolveDatabaseName()
    {
        var connectionString = _configuration.GetConnectionString("Inforest")
            ?? "Server=localhost;Database=INFOREST;Trusted_Connection=true;";

        var builder = new SqlConnectionStringBuilder(connectionString);
        return string.IsNullOrWhiteSpace(builder.InitialCatalog)
            ? "INFOREST"
            : builder.InitialCatalog;
    }
}
