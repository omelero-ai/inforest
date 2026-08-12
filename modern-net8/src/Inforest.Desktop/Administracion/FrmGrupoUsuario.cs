using System.ComponentModel;
using Inforest.Application.Interfaces;

namespace Inforest.Desktop.Administracion;

/// <summary>
/// Legacy: <c>frmGrupoAcceso.frm</c>.
/// </summary>
public class FrmGrupoUsuario : Form
{
    private readonly IRbacService _rbacService;
    private readonly BindingList<GrupoUsuarioRow> _grupos = [];
    private readonly BindingList<AccesoGrupoRow> _accesos = [];
    private readonly DataGridView _gridGrupos;
    private readonly DataGridView _gridAccesos;

    public FrmGrupoUsuario(IRbacService rbacService)
    {
        _rbacService = rbacService;
        Text = "Grupos de Usuario";

        _gridGrupos = new DataGridView
        {
            Dock = DockStyle.Left,
            Width = 300,
            AutoGenerateColumns = true,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            DataSource = _grupos
        };
        _gridGrupos.SelectionChanged += (_, _) => CargarAccesosGrupo();

        _gridAccesos = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            AllowUserToAddRows = false,
            DataSource = _accesos
        };

        Controls.Add(_gridAccesos);
        Controls.Add(_gridGrupos);
        Load += FrmGrupoUsuario_Load;
    }

    private async void FrmGrupoUsuario_Load(object? sender, EventArgs e)
    {
        await _rbacService.ObtenerPermisosAsync("ADMIN", "02");
        _grupos.Clear();
        _grupos.Add(new GrupoUsuarioRow("ADM", "Administradores"));
        _grupos.Add(new GrupoUsuarioRow("CAJ", "Cajeros"));
        _grupos.Add(new GrupoUsuarioRow("CON", "Consultas"));
        CargarAccesosGrupo();
    }

    private void CargarAccesosGrupo()
    {
        _accesos.Clear();
        var codigo = (_gridGrupos.CurrentRow?.DataBoundItem as GrupoUsuarioRow)?.Codigo ?? "ADM";
        foreach (var acceso in ObtenerAccesos(codigo))
            _accesos.Add(acceso);
    }

    private static IEnumerable<AccesoGrupoRow> ObtenerAccesos(string grupo)
    {
        return grupo switch
        {
            "CAJ" =>
            [
                new AccesoGrupoRow(true, "POS", "Cobros rápidos"),
                new AccesoGrupoRow(true, "DOC", "Emitir comprobantes"),
                new AccesoGrupoRow(false, "ADM", "Parámetros")
            ],
            "CON" =>
            [
                new AccesoGrupoRow(true, "CONS", "Consultas"),
                new AccesoGrupoRow(true, "REP", "Reportes"),
                new AccesoGrupoRow(false, "USR", "Usuarios")
            ],
            _ =>
            [
                new AccesoGrupoRow(true, "USR", "Usuarios"),
                new AccesoGrupoRow(true, "CFG", "Parámetros"),
                new AccesoGrupoRow(true, "SEC", "Seguridad")
            ]
        };
    }

    private sealed record GrupoUsuarioRow(string Codigo, string Nombre);
    private sealed record AccesoGrupoRow(bool Habilitado, string CodigoAcceso, string Descripcion);
}
