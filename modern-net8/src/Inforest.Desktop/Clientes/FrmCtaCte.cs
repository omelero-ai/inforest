using Inforest.Application.Maestros;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Clientes;

/// <summary>
/// Formulario de correlativo de cuentas corrientes.
/// Legacy: frmCtaCte.frm — "Correlativo de Cuentas Corrientes"
/// Muestra la lista de clientes con cuenta corriente habilitada (vCompania / TDELIVERY lClienteCtaCte=1).
/// BR-CTACTE-003: Solo se muestran clientes activos con ctacte habilitada.
/// </summary>
public class FrmCtaCte : Form
{
    private readonly ICuentaCorrienteRepository _repository;
    private readonly IClienteRepository _clienteRepository;

    private readonly DataGridView dgv = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AutoGenerateColumns = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AllowUserToAddRows = false
    };

    private readonly Label lblTotal = new() { Dock = DockStyle.Bottom, Height = 26, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(8, 0, 8, 0) };
    private readonly Button btnActualizar = new() { Text = "Actualizar", Width = 110 };
    private readonly Button btnNuevoCliente = new() { Text = "Nuevo cliente", Width = 130 };
    private readonly Button btnSalir = new() { Text = "Salir", Width = 80 };

    private List<CuentaCorriente> _datos = [];

    public FrmCtaCte(ICuentaCorrienteRepository repository, IClienteRepository clienteRepository)
    {
        _repository = repository;
        _clienteRepository = clienteRepository;

        Text = "Correlativo de Cuentas Corrientes";
        Width = 1050;
        Height = 620;
        StartPosition = FormStartPosition.CenterScreen;
        MdiParent = System.Windows.Forms.Application.OpenForms.OfType<Form>().FirstOrDefault(f => f.IsMdiContainer);

        ConfigurarGrilla();
        Controls.Add(dgv);
        Controls.Add(BuildBarraInferior());
        Controls.Add(lblTotal);

        btnActualizar.Click    += async (_, _) => await CargarAsync();
        btnNuevoCliente.Click  += OnNuevoCliente;
        btnSalir.Click         += (_, _) => Close();
        Load                   += async (_, _) => await CargarAsync();
    }

    private void ConfigurarGrilla()
    {
        dgv.Columns.AddRange(
        [
            new DataGridViewTextBoxColumn { DataPropertyName = "CodigoCtaCte", HeaderText = "Código",     Width = 80 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Descripcion",  HeaderText = "Cliente",    Width = 260 },
            new DataGridViewTextBoxColumn { DataPropertyName = "Identidad",    HeaderText = "Identidad",  Width = 130 },
            new DataGridViewTextBoxColumn { DataPropertyName = "TipoCtaCte",   HeaderText = "Tipo",       Width = 60 },
            new DataGridViewTextBoxColumn
            {
                DataPropertyName = "LineaCredito",
                HeaderText = "Línea",
                Width = 110,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight }
            },
            new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Consumo",
                HeaderText = "Consumo",
                Width = 110,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight }
            },
            new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Saldo",
                HeaderText = "Saldo",
                Width = 110,
                DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", Alignment = DataGridViewContentAlignment.MiddleRight }
            },
            new DataGridViewTextBoxColumn
            {
                DataPropertyName = "DiasCredito",
                HeaderText = "Días crédito",
                Width = 90,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
            }
        ]);
    }

    private async Task CargarAsync()
    {
        Cursor = Cursors.WaitCursor;
        try
        {
            var handler = new ObtenerCuentasCorrientesHandler(_repository);
            var result  = await handler.HandleAsync(new ObtenerCuentasCorrientesQuery());
            _datos = result.EsExitoso ? [.. result.Valor!] : [];
            dgv.DataSource = _datos;
            lblTotal.Text  = $"Total registros: {_datos.Count}  |  Consumo total: {_datos.Sum(x => x.Consumo):N2}";
        }
        finally { Cursor = Cursors.Default; }
    }

    private void OnNuevoCliente(object? sender, EventArgs e)
    {
        using var frm = new FrmNuevoCliente(_clienteRepository);
        if (frm.ShowDialog(this) == DialogResult.OK)
            _ = CargarAsync();
    }

    private Control BuildBarraInferior()
    {
        var bar = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 44,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(4)
        };
        bar.Controls.AddRange([btnSalir, btnNuevoCliente, btnActualizar]);
        return bar;
    }
}
