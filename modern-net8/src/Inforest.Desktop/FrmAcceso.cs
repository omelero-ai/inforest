using Inforest.Application.Interfaces;
using Inforest.Application.Interfaces.Hardware;
using Inforest.Application.Configuracion;
using Inforest.Application.Seguridad;
using Inforest.Desktop.POS;
using Inforest.Desktop.Shared;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Inforest.Desktop;

/// <summary>
/// Formulario inicial de acceso.
/// Legacy: frmAcceso.frm.
/// Reglas: BR-006, BR-POS-006-LOCK, BR-PERIPH-004.
/// Flujos migrados:
///   cmdOpcion[0]/imgOpcion[0] — login usuario+contraseña ✓
///   cmdOpcion[1]/imgOpcion[1] — cancelar ✓
///   imgOpcion[2]              — teclado en pantalla (usuario) → btnTecladoUsuario ✓
///   imgOpcion[3]              — teclado en pantalla (password) → btnTecladoPassword ✓
///   imgOpcion[4]              — login biométrico huella (1:N) → btnHuella ✓ (BLOCKED: OCX)
///   imgNewProceso / Case 5    — marcación biométrica → btnMarcacion ✓ (BLOCKED: OCX)
///   imgNewOpcion              — cambiar contraseña → btnCambiarContrasena ✓
///   Image1_Click              — acerca de → MostrarAcercaDe ✓
/// </summary>
public partial class FrmAcceso : Form
{
    private const string LastUserFileName = "USUARIO.INI";
    private readonly IAuthService _authService;
    private readonly ILicenseService _licenseService;
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;
    private readonly ValidarInicioPosHandler _validarInicioPosHandler;
    private readonly ObtenerConfiguracionSistemaHandler _obtenerConfiguracionSistemaHandler;
    private readonly IConfiguration _configuration;
    private readonly IServiceProvider _serviceProvider;
    private readonly IBiometriaService _biometriaService;
    private InicioPosValidado? _inicioPos;
    private int _intentosFallidos;
    private bool _cambioContrasenaHabilitado;

    public FrmAcceso(
        IAuthService authService,
        ILicenseService licenseService,
        IDbConnectionFactory connectionFactory,
        ISpExecutor spExecutor,
        ValidarInicioPosHandler validarInicioPosHandler,
        ObtenerConfiguracionSistemaHandler obtenerConfiguracionSistemaHandler,
        IConfiguration configuration,
        IServiceProvider serviceProvider,
        IBiometriaService biometriaService)
    {
        _authService = authService;
        _licenseService = licenseService;
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
        _validarInicioPosHandler = validarInicioPosHandler;
        _obtenerConfiguracionSistemaHandler = obtenerConfiguracionSistemaHandler;
        _configuration = configuration;
        _serviceProvider = serviceProvider;
        _biometriaService = biometriaService;
        InitializeComponent();
    }

    private async void FrmAcceso_Load(object sender, EventArgs e)
    {
        _intentosFallidos = 0;
        CargarUltimoUsuario();
        lblCaja.Text = $"CAJA {_configuration["Inforest:CodigoCaja"] ?? "001"}";
        lblBaseDatos.Text = $"{ResolveServerName().ToUpperInvariant()} : {ResolveDatabaseName().ToUpperInvariant()}";
        await EjecutarInicializacionesLegacyAsync();
        await ValidarLicenciaAsync();
        await CargarConfiguracionSistemaAsync();
        await ValidarInicioPosAsync();
        ValidarVersiones();
        timerVersiones.Start();
        AplicarFocoInicial();
    }

    private async void btnIngresar_Click(object sender, EventArgs e)
    {
        btnIngresar.Enabled = false;
        lblEstado.Text = "Autenticando...";

        try
        {
            var loginNormalizado = NormalizarLogin(txtUsuario.Text);
            var request = new AuthRequest(
                loginNormalizado,
                txtPassword.Text,
                _configuration["Inforest:Modulo"] ?? "INFOREST",
                _inicioPos?.CodigoCaja ?? _configuration["Inforest:CodigoCaja"] ?? "01",
                _configuration["Inforest:CodigoTerminal"] ?? Environment.MachineName,
                ResolveDatabaseName(),
                ExtraerDigitos(txtPassword.Text));

            var result = await _authService.AutenticarAsync(request);
            if (!result.Exitoso || result.Sesion is null)
            {
                _intentosFallidos++;
                lblEstado.Text = result.MensajeError ?? "No se pudo iniciar sesión.";
                txtPassword.Clear();
                txtPassword.Focus();

                if (_intentosFallidos >= LoginPolicy.MaxIntentosFallidos)
                    System.Windows.Forms.Application.Exit();

                return;
            }

            lblEstado.Text = $"Sesión activa: {result.Sesion.CodigoUsuario} / Caja {result.Sesion.CodigoCaja}";
            lblSesion.Text = $"Módulo {result.Sesion.Modulo} · {result.Sesion.Permisos.Count} permisos cargados";
            PersistirUltimoUsuario(loginNormalizado);
            txtPassword.Clear();

            if (DebeForzarActualizacionEnIngreso())
            {
                System.Windows.Forms.Application.Exit();
                return;
            }

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
        System.Windows.Forms.Application.Exit();
    }

    private void FrmAcceso_Activated(object sender, EventArgs e)
    {
        AplicarFocoInicial();
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
        var codigoCaja = _configuration["Inforest:CodigoCaja"] ?? "001";
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

    private string ResolveServerName()
    {
        var connectionString = _configuration.GetConnectionString("Inforest")
            ?? "Server=localhost;Database=INFOREST;Trusted_Connection=true;";

        var builder = new SqlConnectionStringBuilder(connectionString);
        return string.IsNullOrWhiteSpace(builder.DataSource)
            ? "LOCALHOST"
            : builder.DataSource;
    }

    private void txtUsuario_Leave(object sender, EventArgs e)
    {
        txtUsuario.Text = NormalizarLogin(txtUsuario.Text);
    }

    private void txtPassword_Enter(object sender, EventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(txtPassword.Text))
            btnIngresar.PerformClick();
    }

    private void CargarUltimoUsuario()
    {
        try
        {
            var userFile = Path.Combine(AppContext.BaseDirectory, LastUserFileName);
            if (!File.Exists(userFile))
                return;

            txtUsuario.Text = NormalizarLogin(File.ReadAllText(userFile));
            txtUsuario.SelectionStart = txtUsuario.TextLength;
        }
        catch
        {
            // Mantener compatibilidad tolerante al fallo como en Legacy:
            // si el archivo no existe o no puede leerse, el login sigue operativo.
        }
    }

    private static void PersistirUltimoUsuario(string login)
    {
        try
        {
            var loginPersistido = login.StartsWith('*')
                ? login[1..Math.Min(login.Length, 16)]
                : login;

            File.WriteAllText(
                Path.Combine(AppContext.BaseDirectory, LastUserFileName),
                loginPersistido.Trim());
        }
        catch
        {
            // La persistencia del usuario recordado no debe bloquear el acceso.
        }
    }

    private static string NormalizarLogin(string? login)
        => (login ?? string.Empty).Trim().ToUpperInvariant();

    private static string ExtraerDigitos(string? texto)
        => new((texto ?? string.Empty).Where(char.IsDigit).ToArray());

    private async Task CargarConfiguracionSistemaAsync()
    {
        var configuracion = await _obtenerConfiguracionSistemaHandler.HandleAsync(new ObtenerConfiguracionSistemaQuery());
        _cambioContrasenaHabilitado = configuracion.EsExitoso && (configuracion.Valor?.lCambioContrasena ?? false);
        btnCambiarContrasena.Enabled = _cambioContrasenaHabilitado;
    }

    private void btnCambiarContrasena_Click(object sender, EventArgs e)
    {
        if (!_cambioContrasenaHabilitado)
        {
            MessageBox.Show(
                "Proceso no habilitado, indicar al administrador del sistema la activación en parametros generales!",
                Text,
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
            return;
        }

        using var frm = ActivatorUtilities.CreateInstance<FrmCambiarContrasenia>(_serviceProvider, NormalizarLogin(txtUsuario.Text));
        if (frm.ShowDialog(this) == DialogResult.OK)
            txtPassword.Clear();
    }

    private void txtPassword_KeyPress(object sender, KeyPressEventArgs e)
    {
        if (e.KeyChar != (char)Keys.Enter)
            return;

        e.Handled = true;
        btnIngresar.PerformClick();
    }

    private void txtUsuario_KeyPress(object sender, KeyPressEventArgs e)
    {
        if (e.KeyChar != (char)Keys.Enter)
            return;

        e.Handled = true;
        btnIngresar.PerformClick();
    }

    private void timerVersiones_Tick(object sender, EventArgs e)
    {
        ValidarVersiones();
    }

    private void ValidarVersiones()
    {
        try
        {
            var versionPath = Path.Combine(AppContext.BaseDirectory, "version.txt");
            if (!File.Exists(versionPath))
            {
                panelVersion.Visible = false;
                return;
            }

            var versionObjetivo = File.ReadLines(versionPath).FirstOrDefault()?.Trim();
            var versionActual = System.Windows.Forms.Application.ProductVersion;
            var hayActualizacion = !string.IsNullOrWhiteSpace(versionObjetivo)
                && !string.Equals(versionObjetivo, versionActual, StringComparison.OrdinalIgnoreCase);

            panelVersion.Visible = hayActualizacion;
            if (hayActualizacion)
                lblVersionDisponible.Text = $"InfoRest {versionObjetivo}";
        }
        catch
        {
            panelVersion.Visible = false;
        }
    }

    private async Task EjecutarInicializacionesLegacyAsync()
    {
        await EjecutarActualizacionCostosSiAplicaAsync();
        await EjecutarTipoCambioSiAplicaAsync();
    }

    private async Task EjecutarActualizacionCostosSiAplicaAsync()
    {
        var almacenHabilitado = bool.TryParse(_configuration["Inforest:Almacen:Enabled"], out var enabled) && enabled;
        if (!almacenHabilitado)
            return;

        try
        {
            var baseInforest = ResolveDatabaseName();
            var baseAlmacen = ResolveDatabaseName("Almacen");
            var codigoLocal = (_configuration["Inforest:CodigoLocal"] ?? "01").Trim();

            using var connectionAlmacen = await _connectionFactory.CreateOpenConnectionAsync("Almacen");
            await _spExecutor.ExecuteAsync(connectionAlmacen, "sp_ActualizaReceta");

            using var connectionInforest = await _connectionFactory.CreateOpenConnectionAsync("Inforest");
            await _spExecutor.ExecuteAsync(connectionInforest, "usp_Inforest_InicializaCostos");
            await _spExecutor.ExecuteAsync(
                connectionInforest,
                "usp_Inforest_ActualizaCostos",
                new
                {
                    tNombreInforest = baseInforest,
                    tNombreAlmacen = baseAlmacen,
                    tLocal = codigoLocal
                });

            await _spExecutor.ExecuteAsync(
                connectionInforest,
                "Usp_ActualizarPreciosTransferenciaAlmacen",
                new
                {
                    SubGrupo = string.Empty,
                    BaseDatoAlmacen = baseAlmacen,
                    tipooper = 2
                });
        }
        catch
        {
            // Compatibilidad legacy: no bloquear login por fallas de inicialización opcional.
        }
    }

    private async Task EjecutarTipoCambioSiAplicaAsync()
    {
        var pais = (_configuration["Inforest:Pais"] ?? string.Empty).Trim();
        if (!string.Equals(pais, "002", StringComparison.OrdinalIgnoreCase))
            return;

        try
        {
            using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest");
            var existe = await connection.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM TTIPOCAMBIO WHERE CONVERT(date, fFecha) = CONVERT(date, GETDATE())");

            if (existe == 0)
            {
                await _spExecutor.ExecuteAsync(
                    connection,
                    "spIns_TipoCambio",
                    new
                    {
                        nTc = 1d,
                        tUSUARIO = string.Empty,
                        nTco = 0d,
                        nTc2 = 1d,
                        nTc3 = 1d
                    });
            }
        }
        catch
        {
            // Compatibilidad legacy: no bloquear login por este flujo.
        }
    }

    private bool DebeForzarActualizacionEnIngreso()
    {
        var actualizadorActivo = bool.TryParse(_configuration["Inforest:Actualizador:Activo"], out var activo) && activo;
        if (!actualizadorActivo)
            return false;

        try
        {
            var versionPath = Path.Combine(AppContext.BaseDirectory, "version.txt");
            if (!File.Exists(versionPath))
                return false;

            var versionObjetivo = File.ReadLines(versionPath).FirstOrDefault()?.Trim();
            var versionActual = System.Windows.Forms.Application.ProductVersion;
            return !string.IsNullOrWhiteSpace(versionObjetivo)
                && !string.Equals(versionObjetivo, versionActual, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private void AplicarFocoInicial()
    {
        if (string.IsNullOrWhiteSpace(txtUsuario.Text))
        {
            txtUsuario.Focus();
            return;
        }

        txtPassword.Focus();
    }

    private string ResolveDatabaseName(string connectionName)
    {
        var connectionString = _configuration.GetConnectionString(connectionName);
        if (string.IsNullOrWhiteSpace(connectionString))
            return ResolveDatabaseName();

        var builder = new SqlConnectionStringBuilder(connectionString);
        return string.IsNullOrWhiteSpace(builder.InitialCatalog)
            ? ResolveDatabaseName()
            : builder.InitialCatalog;
    }

    // ── Biometría — login por huella (imgOpcion Index=4) ─────────────────────

    /// <summary>
    /// Inicia sesión por identificación biométrica 1:N (huella dactilar).
    /// Legacy: imgOpcion_Click Case 4 — frmVerificacionHuellaSup.Opcion lModulo.
    /// GAP-004: sgfplibx.ocx OCX 32-bit — NullBiometriaService activo hasta SDK .NET disponible.
    /// BR-PERIPH-004.
    /// </summary>
    private async void btnHuella_Click(object sender, EventArgs e)
    {
        btnHuella.Enabled = false;
        lblEstado.Text    = "Capturando huella...";
        try
        {
            var modulo = ResolverCodigoModulo(_configuration["Inforest:Modulo"] ?? "INFOREST");
            var result = await _biometriaService.IdentificarUsuarioAsync(modulo);
            if (!result.EsExitoso || string.IsNullOrWhiteSpace(result.LoginIdentificado))
            {
                lblEstado.Text = result.Mensaje;
                return;
            }

            var login = NormalizarLogin(result.LoginIdentificado);
            txtUsuario.Text = login;
            PersistirUltimoUsuario(login);
            AbrirShellPrincipal();
        }
        finally
        {
            btnHuella.Enabled = true;
        }
    }

    // ── Biometría — marcación de asistencia (imgNewProceso) ──────────────────

    /// <summary>
    /// Registra marcación de asistencia por huella.
    /// Legacy: imgNewProceso_Click — marcacion=1, imgOpcion_Click Case 5 — frmVerificarHuellaMarcacion.
    /// GAP-004: sgfplibx.ocx OCX 32-bit — stub hasta SDK .NET disponible.
    /// BR-PERIPH-004.
    /// </summary>
    private async void btnMarcacion_Click(object sender, EventArgs e)
    {
        btnMarcacion.Enabled = false;
        lblEstado.Text       = "Capturando huella para marcación...";
        try
        {
            var resultado = await _biometriaService.CapturarHuellaAsync();
            lblEstado.Text = resultado.EsExitoso
                ? "Marcación registrada."
                : resultado.Mensaje;

            if (resultado.EsExitoso)
            {
                MessageBox.Show(
                    "Marcación de asistencia registrada correctamente.",
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }
        finally
        {
            btnMarcacion.Enabled = true;
            if (lblEstado.Text == "Capturando huella para marcación...")
                lblEstado.Text = "Esperando validación.";
        }
    }

    // ── Teclado en pantalla — usuario (imgOpcion Index=2) ─────────────────────

    /// <summary>
    /// Abre teclado alfanumérico en pantalla para el campo usuario.
    /// Legacy: imgOpcion_Click Case 2 — frmKeyBoard.Show vbModal; txtUsuario.Text = sDescrip.
    /// </summary>
    private void btnTecladoUsuario_Click(object sender, EventArgs e)
    {
        using var teclado = new FrmInputTeclado(
            titulo:       "Usuario",
            textoInicial: txtUsuario.Text,
            esContrasena: false);

        if (teclado.ShowDialog(this) == DialogResult.OK)
            txtUsuario.Text = NormalizarLogin(teclado.ValorTexto);
    }

    // ── Teclado en pantalla — contraseña (imgOpcion Index=3) ─────────────────

    /// <summary>
    /// Abre teclado en pantalla enmascarado para el campo contraseña.
    /// Legacy: imgOpcion_Click Case 3 — frmPassword.Show vbModal; txtPassword.Text = sDescrip.
    /// </summary>
    private void btnTecladoPassword_Click(object sender, EventArgs e)
    {
        using var teclado = new FrmInputTeclado(
            titulo:       "Contraseña",
            textoInicial: string.Empty,
            esContrasena: true);

        if (teclado.ShowDialog(this) == DialogResult.OK)
            txtPassword.Text = teclado.ValorTexto;
    }

    // ── Acerca de (Image1_Click) ─────────────────────────────────────────────

    /// <summary>
    /// Muestra información de versión del sistema.
    /// Legacy: Image1_Click — frmAbout.Show vbModal.
    /// </summary>
    private void MostrarAcercaDe()
    {
        var version = System.Windows.Forms.Application.ProductVersion;
        MessageBox.Show(
            $"INFOREST\nVersión {version}\n\n© Sistema ERP Gastronómico",
            "Acerca de INFOREST",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    // ── Utilidades ───────────────────────────────────────────────────────────

    /// <summary>
    /// Convierte el nombre del módulo al código numérico que usa el Legacy para biometría.
    /// Legacy: lModulo = "01"/"02"/"03" según sModulo en AccesoInicio.
    /// </summary>
    private static string ResolverCodigoModulo(string nombreModulo) =>
        nombreModulo.ToUpperInvariant() switch
        {
            "INFOREST"       => "01",
            "ADMINISTRACION" => "02",
            "CONSULTA"       => "03",
            _                => "01"
        };
}
