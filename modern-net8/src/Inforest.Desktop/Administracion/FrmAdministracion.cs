using Inforest.Application.Interfaces;

namespace Inforest.Desktop.Administracion;

/// <summary>
/// Legacy: <c>mdiAdministracion.frm</c>, <c>Administracion.exe</c>, <c>modAdministracion.bas</c>.
/// </summary>
public class FrmAdministracion : Form
{
    private readonly IAuthService _authService;
    private readonly IRbacService _rbacService;
    private readonly Func<FrmUsuario>? _usuarioFactory;
    private readonly Func<FrmGrupoUsuario>? _grupoUsuarioFactory;
    private readonly Func<FrmAcceso>? _accesoFactory;
    private readonly Func<FrmParametro>? _parametroFactory;
    private readonly Func<FrmConfiguracionCaja>? _configuracionCajaFactory;
    private readonly Panel _panelCentral;
    private readonly Label _lblEstado;

    public FrmAdministracion(
        IAuthService authService,
        IRbacService rbacService,
        Func<FrmUsuario>? usuarioFactory = null,
        Func<FrmGrupoUsuario>? grupoUsuarioFactory = null,
        Func<FrmAcceso>? accesoFactory = null,
        Func<FrmParametro>? parametroFactory = null,
        Func<FrmConfiguracionCaja>? configuracionCajaFactory = null)
    {
        _authService = authService;
        _rbacService = rbacService;
        _usuarioFactory = usuarioFactory;
        _grupoUsuarioFactory = grupoUsuarioFactory;
        _accesoFactory = accesoFactory;
        _parametroFactory = parametroFactory;
        _configuracionCajaFactory = configuracionCajaFactory;

        Text = "Administración";
        WindowState = FormWindowState.Maximized;
        MinimumSize = new Size(1100, 700);

        var menu = new MenuStrip();
        var maestrosItem = new ToolStripMenuItem("Maestros");
        maestrosItem.DropDownItems.Add("Grupos de Usuario", null, (_, _) => MostrarSubFormulario(_grupoUsuarioFactory?.Invoke() ?? CrearFallback("Gestión de grupos")));
        maestrosItem.DropDownItems.Add("Accesos", null, (_, _) => MostrarSubFormulario(_accesoFactory?.Invoke() ?? CrearFallback("Gestión de accesos")));

        var usuariosItem = new ToolStripMenuItem("Usuarios");
        usuariosItem.Click += (_, _) => MostrarSubFormulario(_usuarioFactory?.Invoke() ?? CrearFallback("Gestión de usuarios"));

        var parametrosItem = new ToolStripMenuItem("Parámetros");
        parametrosItem.Click += (_, _) => MostrarSubFormulario(_parametroFactory?.Invoke() ?? CrearFallback("Parámetros TPARAMETRO"));

        var cajasItem = new ToolStripMenuItem("Cajas");
        cajasItem.Click += (_, _) => MostrarSubFormulario(_configuracionCajaFactory?.Invoke() ?? CrearFallback("Configuración de cajas"));

        var salirItem = new ToolStripMenuItem("Salir");
        salirItem.Click += (_, _) => Close();

        menu.Items.AddRange([maestrosItem, usuariosItem, parametrosItem, cajasItem, salirItem]);

        _lblEstado = new Label
        {
            Dock = DockStyle.Bottom,
            Height = 28,
            Padding = new Padding(8, 6, 8, 0),
            ForeColor = Color.DimGray,
            Text = "Inicializando módulo Administración..."
        };

        _panelCentral = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.WhiteSmoke,
            Padding = new Padding(8)
        };

        Controls.Add(_panelCentral);
        Controls.Add(_lblEstado);
        Controls.Add(menu);
        MainMenuStrip = menu;
        Load += FrmAdministracion_Load;
    }

    private async void FrmAdministracion_Load(object? sender, EventArgs e)
    {
        await _rbacService.ObtenerPermisosAsync("ADMIN", "02");
        MostrarSubFormulario(_parametroFactory?.Invoke() ?? CrearFallback("Seleccione una opción del menú"));
        _lblEstado.Text = "Administración lista. Use el menú superior para navegar entre maestros, usuarios y parámetros.";
    }

    private void MostrarSubFormulario(Form form)
    {
        foreach (Control control in _panelCentral.Controls)
        {
            control.Dispose();
        }

        _panelCentral.Controls.Clear();
        form.TopLevel = false;
        form.FormBorderStyle = FormBorderStyle.None;
        form.Dock = DockStyle.Fill;
        _panelCentral.Controls.Add(form);
        form.Show();
        _lblEstado.Text = $"Pantalla activa: {form.Text}";
    }

    private static Form CrearFallback(string titulo)
    {
        var form = new Form { Text = titulo, BackColor = Color.White };
        form.Controls.Add(new Label
        {
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Text = titulo,
            Font = new Font("Segoe UI", 16, FontStyle.Bold)
        });
        return form;
    }
}
