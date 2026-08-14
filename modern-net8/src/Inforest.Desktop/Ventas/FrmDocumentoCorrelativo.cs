using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Pantalla de consulta y administración de correlativos de comprobante por caja.
/// Legacy: frmDocumentoCorrelativo.frm + frmDocumentoCorrelativoDetalle.frm. BR-POS-008.
/// </summary>
public class FrmDocumentoCorrelativo : Form
{
    private readonly ObtenerTodosCorrelativosHandler? _handler;
    private readonly DataGridView _grid;
    private readonly Label _lblStatus;
    private readonly Button _btnRefresh;

    public FrmDocumentoCorrelativo(ObtenerTodosCorrelativosHandler? handler = null)
    {
        _handler = handler;

        Text = "Correlativos de Comprobante";
        Width = 920;
        Height = 540;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;

        // ── Grid ──────────────────────────────────────────────────────────────
        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            RowHeadersVisible = false,
            BackgroundColor = SystemColors.Window
        };

        _grid.Columns.AddRange(
            new DataGridViewTextBoxColumn { Name = "Caja", HeaderText = "Caja", DataPropertyName = "CodigoCaja", FillWeight = 40 },
            new DataGridViewTextBoxColumn { Name = "Impresora", HeaderText = "Impresora", DataPropertyName = "CodigoImpresora", FillWeight = 50 },
            new DataGridViewTextBoxColumn { Name = "Tipo", HeaderText = "Tipo Emisión", DataPropertyName = "TipoEmision", FillWeight = 60 },
            new DataGridViewTextBoxColumn { Name = "Descripcion", HeaderText = "Descripción", DataPropertyName = "Descripcion", FillWeight = 140 },
            new DataGridViewTextBoxColumn { Name = "Serie", HeaderText = "Serie", DataPropertyName = "Serie", FillWeight = 50 },
            new DataGridViewTextBoxColumn { Name = "UltimoNumero", HeaderText = "Último Número", DataPropertyName = "UltimoNumero", FillWeight = 90 },
            new DataGridViewTextBoxColumn { Name = "Autorizacion", HeaderText = "N° Autorización", DataPropertyName = "NumeroAutorizacion", FillWeight = 100 },
            new DataGridViewTextBoxColumn { Name = "FechaInicio", HeaderText = "Inicio", DataPropertyName = "FechaInicioDisplay", FillWeight = 70 },
            new DataGridViewTextBoxColumn { Name = "FechaCaducidad", HeaderText = "Caducidad", DataPropertyName = "FechaCaducidadDisplay", FillWeight = 70 },
            new DataGridViewCheckBoxColumn { Name = "FE", HeaderText = "FE", DataPropertyName = "FacturacionElectronica", FillWeight = 30 }
        );

        // ── Bottom panel ──────────────────────────────────────────────────────
        _lblStatus = new Label
        {
            AutoSize = true,
            Text = "Cargando...",
            Padding = new Padding(6, 4, 0, 0)
        };

        _btnRefresh = new Button
        {
            Text = "Actualizar",
            Width = 100,
            Height = 30,
            Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        };
        _btnRefresh.Click += async (_, _) => await CargarAsync();

        var btnCerrar = new Button
        {
            Text = "Cerrar",
            Width = 100,
            Height = 30,
            Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
            DialogResult = DialogResult.Cancel
        };
        btnCerrar.Click += (_, _) => Close();

        var bottom = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 44,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(6, 6, 6, 6)
        };
        bottom.Controls.AddRange([_lblStatus, new Label { AutoSize = true, Width = 200 }, _btnRefresh, btnCerrar]);

        Controls.AddRange([_grid, bottom]);

        Load += async (_, _) => await CargarAsync();
    }

    private async Task CargarAsync()
    {
        _btnRefresh.Enabled = false;
        _lblStatus.Text = "Cargando correlativos...";

        try
        {
            if (_handler is null)
            {
                _lblStatus.Text = "Sin datos (handler no configurado).";
                return;
            }

            var result = await _handler.HandleAsync(new ObtenerTodosCorrelativosQuery(), CancellationToken.None);
            if (!result.EsExitoso)
            {
                _lblStatus.Text = $"Error: {result.MensajeError}";
                return;
            }

            var filas = result.Valor!.Select(c => new CorrelativoRow(
                c.CodigoCaja,
                c.CodigoImpresora,
                c.TipoEmision,
                c.Descripcion ?? "",
                c.Serie ?? "",
                c.ObtenerNumeroActual(),
                c.NumeroAutorizacion ?? "",
                c.FechaInicio.HasValue ? c.FechaInicio.Value.ToString("dd/MM/yyyy") : "",
                c.FechaCaducidad.HasValue ? c.FechaCaducidad.Value.ToString("dd/MM/yyyy") : "",
                c.FacturacionElectronica)).ToList();

            _grid.DataSource = filas;
            _lblStatus.Text = $"{filas.Count} correlativo(s).";
        }
        catch (Exception ex)
        {
            _lblStatus.Text = $"Error inesperado: {ex.Message}";
        }
        finally
        {
            _btnRefresh.Enabled = true;
        }
    }

    /// <summary>Fila de presentación para el grid.</summary>
    private sealed record CorrelativoRow(
        string CodigoCaja,
        string CodigoImpresora,
        string TipoEmision,
        string Descripcion,
        string Serie,
        string UltimoNumero,
        string NumeroAutorizacion,
        string FechaInicioDisplay,
        string FechaCaducidadDisplay,
        bool FacturacionElectronica);
}
