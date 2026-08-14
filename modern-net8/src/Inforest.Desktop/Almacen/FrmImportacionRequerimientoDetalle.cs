using Inforest.Application.Almacen;
using Inforest.Domain.Entities.Almacen;

namespace Inforest.Desktop.Almacen;

/// <summary>
/// Visualización del detalle de un requerimiento de almacén con enlace a INFOREST.
/// Legacy: frmImportacionRequerimientoDetalle.frm.
/// BR-IMPORT-003: Permite verificar que todos los artículos tienen enlace con INFOREST.
/// </summary>
public class FrmImportacionRequerimientoDetalle : Form
{
    private readonly ObtenerDetalleRequerimientoHandler _handler;
    private readonly string _rq;
    private readonly string _baseDatosInforest;

    private readonly DataGridView dgv = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AutoGenerateColumns = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false
    };
    private readonly Button btnSalir = new() { Text = "Salir", Width = 80 };
    private readonly Label lblTitulo = new() { AutoSize = true, Font = new Font("MS Sans Serif", 10, FontStyle.Bold) };

    public FrmImportacionRequerimientoDetalle(
        ObtenerDetalleRequerimientoHandler handler,
        string rq,
        string baseDatosInforest = "INFOREST")
    {
        _handler = handler;
        _rq = rq;
        _baseDatosInforest = baseDatosInforest;

        Text = "Visualización de Requerimiento";
        Width = 750;
        Height = 500;
        StartPosition = FormStartPosition.CenterParent;

        ConfigurarColumnas();
        Controls.Add(BuildLayout());

        Load += async (_, _) => await CargarAsync();
        btnSalir.Click += (_, _) => Close();
    }

    private async Task CargarAsync()
    {
        lblTitulo.Text = $"Detalle del Requerimiento: {_rq}";

        var result = await _handler.HandleAsync(
            new ObtenerDetalleRequerimientoQuery(_rq, _baseDatosInforest));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var detalle = result.Valor!;
        dgv.DataSource = detalle.Select(d => new
        {
            d.CodArt,
            d.Articulo,
            d.Cantidad,
            Enlace = d.EnlaceInforRest ?? string.Empty,
            TieneEnlace = d.TieneEnlace ? "✓" : "✗"
        }).ToList();
    }

    private void ConfigurarColumnas()
    {
        dgv.Columns.AddRange(
            new DataGridViewTextBoxColumn { DataPropertyName = "CodArt",      HeaderText = "Código",    Width = 100 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Articulo",     HeaderText = "Producto",  Width = 260 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Cantidad",     HeaderText = "Cantidad",  Width = 90  },
            new DataGridViewTextBoxColumn { DataPropertyName = "Enlace",       HeaderText = "Enlace POS", Width = 100 },
            new DataGridViewTextBoxColumn { DataPropertyName = "TieneEnlace",  HeaderText = "Mapeado",   Width = 80  }
        );
    }

    private Control BuildLayout()
    {
        var pnlTop = new Panel { Dock = DockStyle.Top, Height = 35, Padding = new Padding(5) };
        lblTitulo.Location = new Point(5, 8);
        pnlTop.Controls.Add(lblTitulo);

        var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(5) };
        btnSalir.Location = new Point(5, 5);
        pnlBottom.Controls.Add(btnSalir);

        var pnlGrid = new Panel { Dock = DockStyle.Fill };
        pnlGrid.Controls.Add(dgv);

        var container = new Panel { Dock = DockStyle.Fill };
        container.Controls.Add(pnlGrid);
        container.Controls.Add(pnlBottom);
        container.Controls.Add(pnlTop);
        return container;
    }
}
