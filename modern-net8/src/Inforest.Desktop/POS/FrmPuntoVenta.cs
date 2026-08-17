using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Application.Turno;
using Inforest.Desktop.Forms.Reportes;
using Inforest.Desktop.Pedidos;
using Inforest.Desktop.Turno;
using Inforest.Domain.Entities.Maestros;
using Microsoft.Extensions.DependencyInjection;

namespace Inforest.Desktop.POS;

/// <summary>
/// Legacy: <c>mdiPuntoVenta.frm</c>, <c>InfoRest.exe</c>, <c>modPuntoVenta.bas</c>. Principal formulario del POS.
/// </summary>
public class FrmPuntoVenta : Form
{
    private readonly ISessionService _sessionService;
    private readonly IMesaRepository _mesaRepository;
    private readonly ISalonRepository _salonRepository;
    private readonly ObtenerConfiguracionSistemaHandler _configuracionHandler;
    private readonly AbrirTurnoHandler _abrirTurnoHandler;
    private readonly CerrarTurnoHandler _cerrarTurnoHandler;
    private readonly IDiaContableService _diaContableService;
    private readonly IServiceProvider _serviceProvider;
    private readonly FlowLayoutPanel _panelMesas;
    private readonly ComboBox _cmbSalon;
    private readonly Label _lblUsuario;
    private readonly Label _lblCaja;
    private readonly Label _lblTurno;
    private readonly Label _lblDiaContable;
    private readonly Label _lblHora;
    private readonly System.Windows.Forms.Timer _timer;
    private List<Mesa> _mesas = [];

    public FrmPuntoVenta(
        ISessionService sessionService,
        IMesaRepository mesaRepository,
        ISalonRepository salonRepository,
        ObtenerConfiguracionSistemaHandler configuracionHandler,
        AbrirTurnoHandler abrirTurnoHandler,
        CerrarTurnoHandler cerrarTurnoHandler,
        IDiaContableService diaContableService,
        IServiceProvider serviceProvider)
    {
        _sessionService = sessionService;
        _mesaRepository = mesaRepository;
        _salonRepository = salonRepository;
        _configuracionHandler = configuracionHandler;
        _abrirTurnoHandler = abrirTurnoHandler;
        _cerrarTurnoHandler = cerrarTurnoHandler;
        _diaContableService = diaContableService;
        _serviceProvider = serviceProvider;

        Text = "Punto de Venta";
        WindowState = FormWindowState.Maximized;

        var menu = new MenuStrip();
        var turno = new ToolStripMenuItem("Turno");
        turno.DropDownItems.Add("Abrir turno", null, (_, _) => AbrirTurno());
        turno.DropDownItems.Add("Cerrar turno", null, async (_, _) => await CerrarTurnoAsync());
        menu.Items.Add(turno);
        var reportes = new ToolStripMenuItem("Reportes");
        reportes.DropDownItems.Add("Estados Cta Cte", null, (_, _) => AbrirReporte<FrmRepCtaCteReporte>());
        reportes.DropDownItems.Add("Control de Transacciones", null, (_, _) => AbrirReporte<FrmRepAnuladoReporte>());
        reportes.DropDownItems.Add("Liquidación Ticketera", null, (_, _) => AbrirReporte<FrmRepLiquidacionTicketReporte>());
        reportes.DropDownItems.Add("Paloteo Ticketera", null, (_, _) => AbrirReporte<FrmRepPaloteoTicketReporte>());
        reportes.DropDownItems.Add("Cierre Cajeros Delivery", null, (_, _) => AbrirReporte<FrmRepDeliveryTicketReporte>());
        reportes.DropDownItems.Add("Reporte de Reservas", null, (_, _) => AbrirReporte<FrmRepReservasReporte>());
        reportes.DropDownItems.Add("Reporte de Entregas", null, (_, _) => AbrirReporte<FrmRepEntregaReporte>());
        reportes.DropDownItems.Add("Venta Mensual por Fechas", null, (_, _) => AbrirReporte<FrmRepVentaFechaReporte>());
        reportes.DropDownItems.Add("Cta Cte Integrado", null, (_, _) => AbrirReporte<FrmCtaCteIntegradoReporte>());
        menu.Items.Add(reportes);
        var delivery = new ToolStripMenuItem("Delivery");
        delivery.DropDownItems.Add("Clientes Frecuentes", null, (_, _) => AbrirReporte<FrmClienteDelivery>());
        menu.Items.Add(delivery);
        menu.Items.Add(new ToolStripMenuItem("Administración", null, (_, _) => MessageBox.Show("Abrir FrmAdministracion desde el shell principal.", Text)));
        menu.Items.Add(new ToolStripMenuItem("Salir", null, (_, _) => Close()));

        var barraSuperior = new TableLayoutPanel { Dock = DockStyle.Top, Height = 56, ColumnCount = 6, Padding = new Padding(8) };
        barraSuperior.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
        barraSuperior.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
        barraSuperior.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
        barraSuperior.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
        barraSuperior.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10));
        barraSuperior.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10));
        _lblUsuario = CrearBadge();
        _lblCaja = CrearBadge();
        _lblTurno = CrearBadge();
        _lblDiaContable = CrearBadge();
        _lblHora = CrearBadge();
        _cmbSalon = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        _cmbSalon.SelectedIndexChanged += (_, _) => RenderizarMesas();
        barraSuperior.Controls.AddRange([_lblUsuario, _lblCaja, _lblTurno, _lblDiaContable, _lblHora, _cmbSalon]);

        _panelMesas = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(16), BackColor = Color.WhiteSmoke };

        Controls.Add(_panelMesas);
        Controls.Add(barraSuperior);
        Controls.Add(menu);
        MainMenuStrip = menu;

        _timer = new System.Windows.Forms.Timer { Interval = 30_000 };
        _timer.Tick += async (_, _) => await CargarDatosAsync();
        Load += async (_, _) =>
        {
            _timer.Start();
            await CargarDatosAsync();
        };
        FormClosed += (_, _) => _timer.Stop();
    }

    private async Task CargarDatosAsync()
    {
        var sesion = _sessionService.SesionActual;
        _lblUsuario.Text = $"Usuario: {sesion?.CodigoUsuario ?? "SIN SESIÓN"}";
        _lblCaja.Text = $"Caja: {sesion?.CodigoCaja ?? "001"}";
        _lblHora.Text = $"Hora: {DateTime.Now:HH:mm:ss}";

        // Obtener turno actual
        if (sesion is not null)
        {
            try
            {
                var turnoSvc = _serviceProvider.GetRequiredService<ObtenerTurnoActualHandler>();
                var turnoResult = await turnoSvc.HandleAsync(new ObtenerTurnoActualQuery(sesion.CodigoCaja));
                _lblTurno.Text = turnoResult.EsExitoso && turnoResult.Valor is not null
                    ? $"Turno: {turnoResult.Valor.CodigoTurno}"
                    : "Turno: sin turno";
            }
            catch { _lblTurno.Text = "Turno: --"; }
        }
        else
        {
            _lblTurno.Text = "Turno: --";
        }

        var config = await _configuracionHandler.HandleAsync(new ObtenerConfiguracionSistemaQuery());
        _lblDiaContable.Text = $"Día contable: {(config.Valor?.tHoraCierreDiaContable ?? DateTime.Today.ToShortDateString())}";

        var salones = await CargarSalonesAsync();
        _cmbSalon.DataSource = salones;
        _cmbSalon.DisplayMember = nameof(Salon.Detallado);
        _cmbSalon.ValueMember = nameof(Salon.CodigoSalon);

        _mesas = await CargarMesasAsync();
        RenderizarMesas();
    }

    private async Task<List<Salon>> CargarSalonesAsync()
    {
        try
        {
            var salones = await _salonRepository.ObtenerTodosAsync();
            if (salones.Count > 0) return salones.ToList();
        }
        catch { }

        return
        [
            Salon.Crear("01", "Salón Principal", null, "SISTEMA"),
            Salon.Crear("02", "Terraza", null, "SISTEMA")
        ];
    }

    private async Task<List<Mesa>> CargarMesasAsync()
    {
        try
        {
            var mesas = await _mesaRepository.ObtenerTodosAsync();
            if (mesas.Count > 0) return mesas.ToList();
        }
        catch { }

        return
        [
            Mesa.Crear("M01", "Mesa 01", null, "01", false, "0", "0", 4, "SISTEMA", EstadoMesa.Libre),
            Mesa.Crear("M02", "Mesa 02", null, "01", false, "0", "0", 4, "SISTEMA", EstadoMesa.Ocupada),
            Mesa.Crear("M03", "Mesa 03", null, "01", false, "0", "0", 4, "SISTEMA", EstadoMesa.EnCuenta),
            Mesa.Crear("M04", "Mesa 04", null, "02", false, "0", "0", 6, "SISTEMA", EstadoMesa.Libre)
        ];
    }

    private void RenderizarMesas()
    {
        _panelMesas.Controls.Clear();
        var salon = _cmbSalon.SelectedValue?.ToString();
        foreach (var mesa in _mesas.Where(m => string.IsNullOrWhiteSpace(salon) || m.CodigoSalon == salon))
        {
            var boton = new Button
            {
                Width = 140,
                Height = 90,
                Margin = new Padding(8),
                Text = $"{mesa.Detallado}\nCap. {mesa.NumeroPersonas}",
                BackColor = mesa.Estado switch
                {
                    EstadoMesa.Ocupada => Color.IndianRed,
                    EstadoMesa.EnCuenta => Color.Khaki,
                    _ => Color.LightGreen
                }
            };
            boton.DoubleClick += (_, _) => AbrirPedidoMesa(mesa);
            _panelMesas.Controls.Add(boton);
        }
    }

    private void AbrirReporte<T>() where T : Form
    {
        using var scope = _serviceProvider.CreateScope();
        var form = scope.ServiceProvider.GetRequiredService<T>();
        form.ShowDialog(this);
    }

    private void AbrirTurno()
    {
        using var frm = new FrmAperturaTurno(_abrirTurnoHandler, _diaContableService);
        if (frm.ShowDialog(this) == DialogResult.OK)
            _ = CargarDatosAsync();
    }

    private async Task CerrarTurnoAsync()
    {
        var sesion = _sessionService.SesionActual;
        if (sesion is null) { MessageBox.Show("Sin sesión activa.", Text); return; }

        try
        {
            var turnoSvc = _serviceProvider.GetRequiredService<ObtenerTurnoActualHandler>();
            var turnoResult = await turnoSvc.HandleAsync(new ObtenerTurnoActualQuery(sesion.CodigoCaja));
            if (!turnoResult.EsExitoso || turnoResult.Valor is null)
            {
                MessageBox.Show("No hay turno abierto en esta caja.", Text);
                return;
            }

            using var frm = new FrmCierreTurno(turnoResult.Valor.CodigoTurno, sesion.CodigoCaja, _cerrarTurnoHandler);
            if (frm.ShowDialog(this) == DialogResult.OK)
                await CargarDatosAsync();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al obtener turno: {ex.Message}", Text);
        }
    }

    private void AbrirPedidoMesa(Mesa mesa)
    {
        using var pedido = ActivatorUtilities.CreateInstance<FrmPedido>(_serviceProvider, mesa);
        pedido.ShowDialog(this);
        _ = CargarDatosAsync();
    }

    private static Label CrearBadge() => new()
    {
        Dock = DockStyle.Fill,
        TextAlign = ContentAlignment.MiddleLeft,
        Font = new Font("Segoe UI", 9, FontStyle.Bold),
        BorderStyle = BorderStyle.FixedSingle,
        Padding = new Padding(8, 0, 0, 0)
    };
}
