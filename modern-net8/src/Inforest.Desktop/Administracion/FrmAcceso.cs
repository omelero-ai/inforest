using Inforest.Application.Interfaces;

namespace Inforest.Desktop.Administracion;

/// <summary>
/// Legacy: <c>frmAcceso.frm</c>.
/// </summary>
public class FrmAcceso : Form
{
    private readonly IRbacService _rbacService;
    private readonly TreeView _tree;

    public FrmAcceso(IRbacService rbacService)
    {
        _rbacService = rbacService;
        Text = "Permisos por Grupo";
        _tree = new TreeView { Dock = DockStyle.Fill, CheckBoxes = true };
        Controls.Add(_tree);
        Load += FrmAcceso_Load;
    }

    private async void FrmAcceso_Load(object? sender, EventArgs e)
    {
        await _rbacService.ObtenerPermisosAsync("ADMIN", "02");
        _tree.Nodes.Clear();
        _tree.Nodes.Add(CrearModulo("Administración", ["Usuarios", "Parámetros", "Cajas"]));
        _tree.Nodes.Add(CrearModulo("Consultas", ["Ventas", "Documentos", "Exportar"]));
        _tree.Nodes.Add(CrearModulo("POS", ["Mesas", "Pedidos", "Reportes"]));
        _tree.ExpandAll();
    }

    private static TreeNode CrearModulo(string nombre, IEnumerable<string> formularios)
    {
        var modulo = new TreeNode(nombre) { Checked = true };
        foreach (var formulario in formularios)
            modulo.Nodes.Add(new TreeNode(formulario) { Checked = true });
        return modulo;
    }
}
