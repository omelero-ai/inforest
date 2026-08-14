using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Maestros;

/// <summary>
/// Lista de insumos/platos de stock crítico — equivalente a frmInsumo.frm.
/// Legacy: frmInsumo.frm / USP_LISTARINSUMOS.
/// BR-INSUMO-001, BR-INSUMO-002, BR-INSUMO-003, BR-INSUMO-004.
/// </summary>
public class FrmInsumo : Form
{
    private readonly ListarInsumosHandler _listarHandler;
    private readonly EliminarInsumoHandler _eliminarHandler;
    private readonly IInsumoRepository _repository;
    private readonly bool _moduloInfoRest;

    private readonly DataGridView dgv = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AutoGenerateColumns = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false
    };
    private readonly Button btnNuevo    = new() { Text = "Nuevo",    Width = 80 };
    private readonly Button btnEditar   = new() { Text = "Editar",   Width = 80 };
    private readonly Button btnEliminar = new() { Text = "Eliminar", Width = 80 };
    private readonly Button btnActualizar = new() { Text = "Actualizar", Width = 90 };
    private readonly Label lblContador = new() { AutoSize = true };

    private List<Insumo> _items = [];

    public FrmInsumo(
        ListarInsumosHandler listarHandler,
        EliminarInsumoHandler eliminarHandler,
        IInsumoRepository repository,
        bool moduloInfoRest = true)
    {
        _listarHandler = listarHandler;
        _eliminarHandler = eliminarHandler;
        _repository = repository;
        _moduloInfoRest = moduloInfoRest;

        Text = "Mantenimiento de Insumos/Platos de Stock Crítico";
        Width = 1100;
        Height = 650;
        StartPosition = FormStartPosition.CenterScreen;

        ConfigurarColumnas();
        Controls.Add(BuildLayout());

        Load += async (_, _) => await CargarAsync();
        btnNuevo.Click += (_, _) => AbrirDetalle(nuevo: true);
        btnEditar.Click += (_, _) => AbrirDetalle(nuevo: false);
        btnEliminar.Click += async (_, _) => await EliminarAsync();
        btnActualizar.Click += async (_, _) => await CargarAsync();
        dgv.DoubleClick += (_, _) => AbrirDetalle(nuevo: false);
    }

    private void ConfigurarColumnas()
    {
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Codigo",        HeaderText = "Código",       Width = 90  });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Usuario",       HeaderText = "Usuario",      Width = 120 });
        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Descripcion",   HeaderText = "Descripción",  Width = 320 });

        if (_moduloInfoRest)
            dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Stock",     HeaderText = "Cant. Disp.",  Width = 90  });

        dgv.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "FechaRegistro", HeaderText = "Fecha",        Width = 150 });
        dgv.Columns.Add(new DataGridViewCheckBoxColumn { DataPropertyName = "Activo",       HeaderText = "Activo",       Width = 60  });
    }

    private Control BuildLayout()
    {
        var acciones = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 44, Padding = new Padding(6) };
        acciones.Controls.AddRange([btnNuevo, btnEditar, btnEliminar, btnActualizar, lblContador]);

        var panel = new Panel { Dock = DockStyle.Fill };
        panel.Controls.Add(dgv);
        panel.Controls.Add(acciones);
        return panel;
    }

    internal async Task CargarAsync()
    {
        var result = await _listarHandler.HandleAsync(new ListarInsumosQuery(_moduloInfoRest));
        if (!result.EsExitoso) { MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

        _items = result.Valor!.ToList();
        dgv.DataSource = null;
        dgv.DataSource = _items;
        lblContador.Text = $"Registros: {_items.Count}";
    }

    private void AbrirDetalle(bool nuevo)
    {
        Insumo? seleccionado = null;
        if (!nuevo)
        {
            if (dgv.CurrentRow?.DataBoundItem is not Insumo item)
            {
                MessageBox.Show("No hay registro seleccionado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            seleccionado = item;
        }

        // FrmInsumoDetalle requires the parent form to refresh after save
        using var detalle = new FrmInsumoDetalle(_repository, seleccionado, _moduloInfoRest);
        detalle.ShowDialog(this);
        _ = CargarAsync();
    }

    private async Task EliminarAsync()
    {
        if (dgv.CurrentRow?.DataBoundItem is not Insumo item)
        {
            MessageBox.Show("Seleccione un insumo para eliminar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (MessageBox.Show("¿Seguro de eliminar el insumo?", "Confirmar", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
            return;

        var result = await _eliminarHandler.HandleAsync(new EliminarInsumoCommand(item.Codigo));
        if (!result.EsExitoso) { MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

        await CargarAsync();
    }
}
