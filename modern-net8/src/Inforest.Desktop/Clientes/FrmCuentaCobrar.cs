using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Clientes;

/// <summary>
/// Formulario de documentos pendientes de cobro por cuenta corriente.
/// Legacy: frmCuentaCobrar.frm.
/// Muestra documentos de vDocumentoGrilla donde tEstadoDocumento='03' (CtaCte pendiente de cobro).
/// BR-CTACTE-002: Estado '03' = deuda ctacte pendiente de cobro.
/// </summary>
public class FrmCuentaCobrar : Form
{
    private readonly ICuentaCorrienteRepository _repository;

    private readonly DateTimePicker dtpFecIni = new() { Value = DateTime.Today.AddDays(-30), Width = 130 };
    private readonly DateTimePicker dtpFecFin = new() { Value = DateTime.Today, Width = 130 };
    private readonly Button btnProcesar = new() { Text = "Procesar", Width = 100 };
    private readonly Button btnSalir    = new() { Text = "Salir",    Width = 80 };
    private readonly Label lblTotal     = new() { Dock = DockStyle.Bottom, Height = 26, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(8, 0, 8, 0) };

    private readonly DataGridView dgv = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AutoGenerateColumns = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false
    };

    private List<DocumentoPendienteCobro> _datos = [];

    public FrmCuentaCobrar(ICuentaCorrienteRepository repository)
    {
        _repository = repository;
        Text = "Documentos Pendientes de Cobro — Cuenta Corriente";
        Width = 1100;
        Height = 640;
        StartPosition = FormStartPosition.CenterScreen;
        MdiParent = System.Windows.Forms.Application.OpenForms.OfType<Form>().FirstOrDefault(f => f.IsMdiContainer);

        ConfigurarGrilla();
        Controls.Add(dgv);
        Controls.Add(lblTotal);
        Controls.Add(BuildBarraFiltros());

        btnProcesar.Click += async (_, _) => await CargarAsync();
        btnSalir.Click    += (_, _) => Close();
    }

    private void ConfigurarGrilla()
    {
        dgv.Columns.AddRange(
        [
            new DataGridViewTextBoxColumn { DataPropertyName = "Documento",         HeaderText = "Documento",   Width = 130 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Caja",              HeaderText = "Caja",        Width = 70 },
            new DataGridViewTextBoxColumn { DataPropertyName = "NombreClientePago", HeaderText = "Cliente pago",Width = 220 },
            new DataGridViewTextBoxColumn { DataPropertyName = "NombreCliente",     HeaderText = "Cliente doc.", Width = 200 },
            new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Fecha",
                HeaderText = "Fecha",
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy" }
            },
            new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Venta",
                HeaderText = "Importe",
                Width = 110,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight }
            },
            new DataGridViewTextBoxColumn { DataPropertyName = "Turno",      HeaderText = "Turno",      Width = 80 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Usuario",    HeaderText = "Usuario",    Width = 90 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Vencimiento",HeaderText = "Vencimiento",Width = 100 }
        ]);
    }

    private async Task CargarAsync()
    {
        Cursor = Cursors.WaitCursor;
        try
        {
            var handler = new ObtenerDocumentosPendientesCobroHandler(_repository);
            var result  = await handler.HandleAsync(
                new ObtenerDocumentosPendientesCobroQuery(dtpFecIni.Value, dtpFecFin.Value));

            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, "Error");
                return;
            }

            _datos = [.. result.Valor!];
            dgv.DataSource = _datos;
            lblTotal.Text  = $"Total documentos: {_datos.Count}  |  Total importe: {_datos.Sum(x => x.Venta):N2}";
        }
        finally { Cursor = Cursors.Default; }
    }

    private Control BuildBarraFiltros()
    {
        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 44,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(8, 6, 8, 6)
        };
        panel.Controls.Add(new Label { Text = "Desde:", AutoSize = true, Margin = new Padding(0, 4, 4, 0) });
        panel.Controls.Add(dtpFecIni);
        panel.Controls.Add(new Label { Text = "Hasta:", AutoSize = true, Margin = new Padding(12, 4, 4, 0) });
        panel.Controls.Add(dtpFecFin);
        panel.Controls.Add(btnProcesar);

        var right = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 44,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(4)
        };
        right.Controls.Add(btnSalir);

        Controls.Add(right);
        return panel;
    }
}
