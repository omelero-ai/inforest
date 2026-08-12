using Inforest.Application.Reportes;

namespace Inforest.Desktop.Consultas;

/// <summary>
/// Legacy: <c>frmDocumento.frm</c> + <c>spRep_ControlDocumentos</c>.
/// </summary>
public class FrmConsultaDocumentos : Form, IConsultaOperable
{
    private readonly IReporteRepository _reporteRepository;
    private readonly DateTimePicker _dtpDesde;
    private readonly DateTimePicker _dtpHasta;
    private readonly ComboBox _cmbTipo;
    private readonly ComboBox _cmbEstado;
    private readonly DataGridView _grid;

    public FrmConsultaDocumentos(IReporteRepository reporteRepository)
    {
        _reporteRepository = reporteRepository;
        Text = "Consulta de Documentos";

        var filtros = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 42, Padding = new Padding(8) };
        _dtpDesde = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddDays(-7) };
        _dtpHasta = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };
        _cmbTipo = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
        _cmbEstado = new ComboBox { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
        _cmbTipo.Items.AddRange(["Todos", "FACTURA", "BOLETA", "NOTA"]);
        _cmbEstado.Items.AddRange(["Todos", "EMITIDO", "ANULADO", "PENDIENTE"]);
        _cmbTipo.SelectedIndex = _cmbEstado.SelectedIndex = 0;
        filtros.Controls.AddRange([
            new Label { Text = "Desde", AutoSize = true, Padding = new Padding(0, 6, 0, 0) }, _dtpDesde,
            new Label { Text = "Hasta", AutoSize = true, Padding = new Padding(8, 6, 0, 0) }, _dtpHasta,
            new Label { Text = "Tipo", AutoSize = true, Padding = new Padding(8, 6, 0, 0) }, _cmbTipo,
            new Label { Text = "Estado", AutoSize = true, Padding = new Padding(8, 6, 0, 0) }, _cmbEstado
        ]);

        _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = true, AllowUserToAddRows = false, ReadOnly = true };
        _grid.CellDoubleClick += (_, _) => VerDetalle();

        Controls.Add(_grid);
        Controls.Add(filtros);
        Load += async (_, _) => await EjecutarConsultaAsync();
    }

    public Task EjecutarConsultaAsync()
    {
        _ = _reporteRepository;
        _grid.DataSource = new List<DocumentoRow>
        {
            new("F001-000123", _dtpDesde.Value.Date, _cmbTipo.Text == "Todos" ? "FACTURA" : _cmbTipo.Text, "EMITIDO", 145.20m),
            new("B001-000456", _dtpHasta.Value.Date, "BOLETA", _cmbEstado.Text == "Todos" ? "PENDIENTE" : _cmbEstado.Text, 58.00m)
        };
        return Task.CompletedTask;
    }

    public async Task ExportarAsync()
    {
        await EjecutarConsultaAsync();
        MessageBox.Show("Exportación disponible desde la consulta de documentos.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public Task ImprimirAsync()
    {
        MessageBox.Show("Use RepControlDocumentos.frx para impresión formal.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        return Task.CompletedTask;
    }

    private void VerDetalle()
    {
        if (_grid.CurrentRow?.DataBoundItem is not DocumentoRow row)
            return;

        MessageBox.Show($"Documento: {row.Numero}\nFecha: {row.Fecha:dd/MM/yyyy}\nEstado: {row.Estado}\nMonto: {row.Monto:C}",
            "Detalle documento", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private sealed record DocumentoRow(string Numero, DateTime Fecha, string Tipo, string Estado, decimal Monto);
}

internal sealed class DummyReporteRepository : IReporteRepository
{
    public Task<IReadOnlyList<ComandaRow>> ObtenerComandaAsync(bool flagTipo, string orden, DateTime fechaInicio, DateTime fechaFin, string criterio, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<ComandaRow>>([]);
    public Task<IReadOnlyList<PropinaRow>> ObtenerPropinaAsync(DateTime fechaInicio, DateTime fechaFin, string condicion, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<PropinaRow>>([]);
    public Task<IReadOnlyList<PrincipalClienteDetalleRow>> ObtenerPrincipalClienteDetalleAsync(double montoMinimo, string codigoCliente, DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<PrincipalClienteDetalleRow>>([]);
    public Task<IReadOnlyList<PrincipalClienteResumenRow>> ObtenerPrincipalClienteResumenAsync(double montoMinimo, string codigoCliente, DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<PrincipalClienteResumenRow>>([]);
    public Task<IReadOnlyList<PaloteoComparativoRow>> ObtenerPaloteoComparativoAsync(bool porTurno, bool porValor, string tipoProducto, string areaProduccion, string grupo, string subGrupo, string codigoProducto, string turno, DateTime fechaInicio, DateTime fechaFin, bool soloFacturados, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<PaloteoComparativoRow>>([]);
    public Task<IReadOnlyList<PaloteoSubProdRow>> ObtenerPaloteoSubProdAsync(bool porTurno, string turno, string local, string salon, string tipoProducto, string mozo, string tipoPedido, string area, string caja, string codigoProducto, string codigoCliente, string boton2, string boton3, string boton4, string boton5, DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<PaloteoSubProdRow>>([]);
    public Task<IReadOnlyList<CtaCteIntegradoRow>> ObtenerCtaCteIntegradoAsync(string flagTipo, DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<CtaCteIntegradoRow>>([]);
    public Task<IReadOnlyList<PaloteoVentaIntegradoRow>> ObtenerPaloteoVentaIntegradoAsync(PaloteoVentaIntegradoParametros parametros, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<PaloteoVentaIntegradoRow>>([]);
    public Task<IReadOnlyList<RankingIntegradoRow>> ObtenerRankingIntegradoAsync(RankingIntegradoParametros parametros, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<RankingIntegradoRow>>([]);
    public Task<IReadOnlyList<VentaMensualIntegradoRow>> ObtenerVentaMensualIntegradoAsync(VentaMensualIntegradoParametros parametros, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<VentaMensualIntegradoRow>>(
        [
            new VentaMensualIntegradoRow { Dia = 1, Salon = 100, Delivery = 20, Llevar = 10, Venta = 130, Cantidad = 5, Pax = 8, TTipoPedido = "Local", Costo = 50 },
            new VentaMensualIntegradoRow { Dia = 2, Salon = 140, Delivery = 35, Llevar = 15, Venta = 190, Cantidad = 7, Pax = 10, TTipoPedido = "Delivery", Costo = 80 }
        ]);
    public Task<IReadOnlyList<AnaliticoMotorizadoIntegradoRow>> ObtenerAnaliticoMotorizadoIntegradoAsync(AnaliticoMotorizadoIntegradoParametros parametros, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<AnaliticoMotorizadoIntegradoRow>>([]);
    public Task<IReadOnlyList<TiempoKdsPedidoRow>> ObtenerTiempoKdsPedidoAsync(DateTime fechaInicio, DateTime fechaFin, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<TiempoKdsPedidoRow>>([]);
    public Task<IReadOnlyList<TiempoKdsProductoRow>> ObtenerTiempoKdsProductoAsync(DateTime fechaInicio, DateTime fechaFin, string grupo, string subGrupo, string producto, CancellationToken ct = default)
        => Task.FromResult<IReadOnlyList<TiempoKdsProductoRow>>([]);
}
