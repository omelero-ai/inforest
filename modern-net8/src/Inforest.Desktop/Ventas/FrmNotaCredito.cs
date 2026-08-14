using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Lista de notas de crédito filtrada por rango de fecha.
/// Legacy: frmNotaCredito.frm — cmdProcesa_Click (vNotaCredito).
/// BR-NC-001..006. POS-FUNC-006.
/// </summary>
public class FrmNotaCredito : Form
{
    private readonly ObtenerNotasCreditoPorFechaHandler _obtenerHandler;
    private readonly AnularNotaCreditoHandler _anularHandler;
    private readonly EmitirNotaCreditoHandler _emitirHandler;
    private readonly string _usuarioActual;

    private readonly DateTimePicker _dtpDesde;
    private readonly DateTimePicker _dtpHasta;
    private readonly DataGridView _grid;
    private readonly Label _lblRegistros;
    private List<NotaCredito> _lista = new();

    public FrmNotaCredito(
        ObtenerNotasCreditoPorFechaHandler obtenerHandler,
        AnularNotaCreditoHandler anularHandler,
        EmitirNotaCreditoHandler emitirHandler,
        string usuarioActual)
    {
        _obtenerHandler = obtenerHandler;
        _anularHandler = anularHandler;
        _emitirHandler = emitirHandler;
        _usuarioActual = usuarioActual;

        Text = "Notas de Crédito";
        Width = 900;
        Height = 540;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;

        var panel = new Panel { Dock = DockStyle.Top, Height = 56, Padding = new Padding(8, 8, 8, 0) };
        var lblDesde = new Label { Text = "Desde:", AutoSize = true, Top = 16, Left = 8 };
        _dtpDesde = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 110, Top = 12, Left = 60, Value = DateTime.Today };
        var lblHasta = new Label { Text = "Hasta:", AutoSize = true, Top = 16, Left = 180 };
        _dtpHasta = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 110, Top = 12, Left = 228, Value = DateTime.Today };
        var btnProcesar = new Button { Text = "Buscar", Width = 80, Top = 10, Left = 352 };
        var btnDetalle = new Button { Text = "Ver / Anular", Width = 100, Top = 10, Left = 440 };
        var btnNueva = new Button { Text = "Nueva NC", Width = 90, Top = 10, Left = 548 };
        var btnCerrar = new Button { Text = "Cerrar", Width = 80, Top = 10, Left = 646 };

        btnProcesar.Click += async (_, _) => await CargarAsync();
        btnDetalle.Click += (_, _) => VerDetalle();
        btnNueva.Click += (_, _) => NuevaNotaCredito();
        btnCerrar.Click += (_, _) => Close();

        panel.Controls.AddRange([lblDesde, _dtpDesde, lblHasta, _dtpHasta, btnProcesar, btnDetalle, btnNueva, btnCerrar]);

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AllowUserToAddRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            MultiSelect = false,
            BorderStyle = BorderStyle.None
        };
        _grid.Columns.AddRange(
            new DataGridViewTextBoxColumn { HeaderText = "NC", DataPropertyName = "CodigoNotaCredito", FillWeight = 20 },
            new DataGridViewTextBoxColumn { HeaderText = "Fecha", DataPropertyName = "Fecha", FillWeight = 14 },
            new DataGridViewTextBoxColumn { HeaderText = "Documento", DataPropertyName = "CodigoDocumento", FillWeight = 20 },
            new DataGridViewTextBoxColumn { HeaderText = "Neto", DataPropertyName = "Neto", FillWeight = 13 },
            new DataGridViewTextBoxColumn { HeaderText = "Total", DataPropertyName = "Total", FillWeight = 13 },
            new DataGridViewTextBoxColumn { HeaderText = "Estado", DataPropertyName = "Estado", FillWeight = 10 },
            new DataGridViewTextBoxColumn { HeaderText = "Observación", DataPropertyName = "Observacion", FillWeight = 30 });

        _lblRegistros = new Label { Dock = DockStyle.Bottom, Height = 22, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0) };

        Controls.Add(_grid);
        Controls.Add(_lblRegistros);
        Controls.Add(panel);
    }

    private async Task CargarAsync()
    {
        var result = await _obtenerHandler.HandleAsync(
            new ObtenerNotasCreditoPorFechaQuery(_dtpDesde.Value.Date, _dtpHasta.Value.Date));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _lista = result.Valor!.ToList();
        _grid.DataSource = _lista.Select(n => new
        {
            n.CodigoNotaCredito,
            Fecha = n.Fecha.ToString("dd/MM/yyyy"),
            n.CodigoDocumento,
            Neto = n.Neto.ToString("N2"),
            Total = n.Total.ToString("N2"),
            Estado = MapEstado(n.Estado),
            n.Observacion
        }).ToList();

        _lblRegistros.Text = $"Registros: {_lista.Count}";
    }

    private void VerDetalle()
    {
        if (_grid.CurrentRow is null || _grid.CurrentRow.Index < 0 || _grid.CurrentRow.Index >= _lista.Count)
            return;
        var nc = _lista[_grid.CurrentRow.Index];
        using var frm = new FrmNotaCreditoDetalle(nc, _anularHandler, _usuarioActual, onSaved: null);
        frm.ShowDialog(this);
    }

    private void NuevaNotaCredito()
    {
        using var frm = new FrmNotaCreditoDetalle(null, _anularHandler, _usuarioActual, onSaved: async _ => await CargarAsync(), emitirHandler: _emitirHandler);
        frm.ShowDialog(this);
    }

    private static string MapEstado(string estado) => estado switch
    {
        "PE" => "Pendiente",
        "PR" => "Procesado",
        "AN" => "Anulado",
        _ => estado
    };
}
