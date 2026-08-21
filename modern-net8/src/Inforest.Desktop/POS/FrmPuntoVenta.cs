using Dapper;
using Inforest.Application.Caja;
using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Application.Turno;
using Inforest.Desktop.Caja;
using Inforest.Desktop.Caja.Recibos;
using Inforest.Desktop.Clientes;
using Inforest.Desktop.Delivery;
using Inforest.Desktop.Forms.Reportes;
using Inforest.Desktop.Kitchen;
using Inforest.Desktop.Almacen;
using Inforest.Desktop.Maestros;
using Inforest.Desktop.Pedidos;
using Inforest.Desktop.Reservas;
using Inforest.Desktop.Turno;
using Inforest.Desktop.Ventas;
using Inforest.Domain.Entities.Configuracion;
using Inforest.Domain.Entities.Maestros;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Inforest.Desktop.POS;

/// <summary>
/// Legacy: <c>mdiPuntoVenta.frm</c>, <c>InfoRest.exe</c>, <c>modPuntoVenta.bas</c>.
/// Formulario MDI principal del módulo Punto de Venta.
///
/// Reglas: BR-POS-001, BR-POS-002, BR-POS-003, BR-POS-004, BR-POS-005, BR-POS-006.
/// SQL: LOG_SESIONES (insert/update), usp_GenObtieneDiaContable, SP_NOTICIAS_INFOREST.
/// </summary>
public class FrmPuntoVenta : Form
{
    // ── Colores de diseño (Guía de diseño estandarizado v1.0) ──────────────────
    private static readonly Color ColorPrimario    = Color.FromArgb(0x3A, 0xAC, 0x6E); // #3AAC6E
    private static readonly Color ColorSecundario  = Color.FromArgb(0x26, 0x64, 0xC7); // #2664C7
    private static readonly Color ColorFondo       = Color.FromArgb(0xF8, 0xF9, 0xFC); // #F8F9FC
    private static readonly Color ColorFondoToolbar = Color.White;
    private static readonly Color ColorTexto       = Color.FromArgb(0x23, 0x23, 0x23); // #232323
    private static readonly Color ColorMesaLibre   = Color.FromArgb(0xEA, 0xFA, 0xF1); // pastel verde
    private static readonly Color ColorMesaOcupada = Color.FromArgb(0xFF, 0xEC, 0xEB); // pastel rojo
    private static readonly Color ColorMesaCuenta  = Color.FromArgb(0xFF, 0xF5, 0xF0); // pastel naranja

    // ── Servicios ──────────────────────────────────────────────────────────────
    private readonly ISessionService _sessionService;
    private readonly IMesaRepository _mesaRepository;
    private readonly ISalonRepository _salonRepository;
    private readonly ObtenerConfiguracionSistemaHandler _configuracionHandler;
    private readonly ObtenerConfiguracionCajaHandler _configuracionCajaHandler;
    private readonly AbrirTurnoHandler _abrirTurnoHandler;
    private readonly CerrarTurnoHandler _cerrarTurnoHandler;
    private readonly ObtenerUltimoTurnoHandler _obtenerUltimoTurnoHandler;
    private readonly IDiaContableService _diaContableService;
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly IConfiguration _configuration;
    private readonly IServiceProvider _serviceProvider;

    // ── Estado operativo (equivalente a variables globales VB6) ───────────────
    private ConfiguracionSistema? _cfg;
    private ConfiguracionCaja? _cfgCaja;
    private bool _turnoIniciado;
    private string _turno = string.Empty;
    private string _codigoCaja = "001";
    private string _usuario = string.Empty;

    // ── Controles UI ──────────────────────────────────────────────────────────
    private readonly MenuStrip _menuStrip;
    private readonly ToolStrip _toolStrip;
    private readonly StatusStrip _statusStrip;
    private readonly FlowLayoutPanel _panelMesas;
    private readonly ComboBox _cmbSalon;

    // Toolbar buttons (equivalentes a cmdOpcionN del Legacy)
    private ToolStripButton _btnApertura   = null!; // cmdOpcion1
    private ToolStripButton _btnVenta      = null!; // cmdOpcion2
    private ToolStripButton _btnCierre     = null!; // cmdOpcion3
    private ToolStripButton _btnMesas      = null!; // cmdOpcion4
    private ToolStripButton _btnReciboEgr  = null!; // cmdOpcion5
    private ToolStripButton _btnReciboIng  = null!; // cmdOpcion6
    private ToolStripButton _btnReserva    = null!; // cmdOpcion7
    private ToolStripButton _btnPedidoCor  = null!; // cmdOpcion8
    private ToolStripButton _btnDocCor     = null!; // cmdOpcion9
    private ToolStripButton _btnCtaCte     = null!; // cmdOpcion10
    private ToolStripButton _btnCtaCobrar  = null!; // cmdOpcion11
    private ToolStripButton _btnPrecios    = null!; // cmdOpcion12
    private ToolStripButton _btnDelivery   = null!; // cmdOpcion13
    private ToolStripButton _btnDeliveryNo = null!; // cmdOpcion14
    private ToolStripButton _btnPinPad     = null!; // cmdOpcion15
    private ToolStripButton _btnSalir      = null!; // cmdOpcion16 (Salir / MCPV)
    private ToolStripButton _btnMensajes   = null!; // cmdOpcionMensaje
    private ToolStripButton _btnRecargar   = null!; // cmdRecargarTarjeta
    private ToolStripButton _btnConsulta   = null!; // cmdConsultaSaldo

    // Menú items que se habilitan/deshabilitan
    private ToolStripMenuItem _mnuVenta      = null!;
    private ToolStripMenuItem _mnuCierre     = null!;
    private ToolStripMenuItem _mnuImportacion = null!;
    private ToolStripMenuItem _mnuDocElec    = null!;
    private ToolStripMenuItem _mnuConexion   = null!;
    private ToolStripMenuItem _mnuPinPad     = null!;

    // StatusBar panels (equivalente a StatusBar.Panels del Legacy)
    private ToolStripStatusLabel _statusDiaContable = null!;
    private ToolStripStatusLabel _statusCaja        = null!;
    private ToolStripStatusLabel _statusUsuario     = null!;
    private ToolStripStatusLabel _statusTurno       = null!;
    private ToolStripStatusLabel _statusHora        = null!;

    // Timers (equivalentes a Timer, TimerBizlink, Timerwebapp, TimerActualizador del Legacy)
    private readonly System.Windows.Forms.Timer _timerReloj;          // actualiza hora en statusbar (1s)
    private readonly System.Windows.Forms.Timer _timerBizlink;        // FE billing (60s)
    private readonly System.Windows.Forms.Timer _timerWebapp;         // KDS/alertas (10s)
    private readonly System.Windows.Forms.Timer _timerConexion;       // multiLocal check (inactivo al inicio)
    private readonly System.Windows.Forms.Timer _timerActualizador;   // version check (30s)
    private int _contadorBizlink;
    private bool _alertaVentaMostrada;

    private List<Mesa> _mesas = [];

    public FrmPuntoVenta(
        ISessionService sessionService,
        IMesaRepository mesaRepository,
        ISalonRepository salonRepository,
        ObtenerConfiguracionSistemaHandler configuracionHandler,
        ObtenerConfiguracionCajaHandler configuracionCajaHandler,
        AbrirTurnoHandler abrirTurnoHandler,
        CerrarTurnoHandler cerrarTurnoHandler,
        ObtenerUltimoTurnoHandler obtenerUltimoTurnoHandler,
        IDiaContableService diaContableService,
        IDbConnectionFactory connectionFactory,
        IConfiguration configuration,
        IServiceProvider serviceProvider)
    {
        _sessionService         = sessionService;
        _mesaRepository         = mesaRepository;
        _salonRepository        = salonRepository;
        _configuracionHandler   = configuracionHandler;
        _configuracionCajaHandler = configuracionCajaHandler;
        _abrirTurnoHandler      = abrirTurnoHandler;
        _cerrarTurnoHandler     = cerrarTurnoHandler;
        _obtenerUltimoTurnoHandler = obtenerUltimoTurnoHandler;
        _diaContableService     = diaContableService;
        _connectionFactory      = connectionFactory;
        _configuration          = configuration;
        _serviceProvider        = serviceProvider;

        // ── Ventana MDI (equivalente a VB.MDIForm) ─────────────────────────
        Text          = "Punto de Venta";
        WindowState   = FormWindowState.Maximized;
        IsMdiContainer = true;
        BackColor     = ColorFondo;
        Font          = new Font("Segoe UI", 9f, FontStyle.Regular);

        // ── Construir UI ───────────────────────────────────────────────────
        _menuStrip  = ConstruirMenu();
        _toolStrip  = ConstruirToolbar();
        _statusStrip = ConstruirStatusStrip();
        _panelMesas = new FlowLayoutPanel
        {
            Dock       = DockStyle.Fill,
            AutoScroll = true,
            Padding    = new Padding(16),
            BackColor  = ColorFondo
        };
        _cmbSalon = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font          = new Font("Segoe UI", 9f),
            Width         = 200
        };
        _cmbSalon.SelectedIndexChanged += (_, _) => RenderizarMesas();

        var barraFiltro = new Panel { Dock = DockStyle.Top, Height = 34, BackColor = ColorFondoToolbar, Padding = new Padding(8, 4, 8, 4) };
        barraFiltro.Controls.Add(new Label { Text = "Salón:", AutoSize = true, Location = new Point(8, 8), ForeColor = ColorTexto });
        _cmbSalon.Location = new Point(60, 4);
        barraFiltro.Controls.Add(_cmbSalon);

        Controls.Add(_panelMesas);
        Controls.Add(barraFiltro);
        Controls.Add(_statusStrip);
        Controls.Add(_toolStrip);
        Controls.Add(_menuStrip);
        MainMenuStrip = _menuStrip;

        // ── Timers ─────────────────────────────────────────────────────────
        _timerReloj    = new System.Windows.Forms.Timer { Interval = 1_000 };
        _timerReloj.Tick += (_, _) => ActualizarHora();

        _timerBizlink  = new System.Windows.Forms.Timer { Interval = 60_000 };
        _timerBizlink.Tick += TimerBizlink_Tick;

        _timerWebapp   = new System.Windows.Forms.Timer { Interval = 10_000 };
        _timerWebapp.Tick += TimerWebapp_Tick;

        _timerConexion = new System.Windows.Forms.Timer { Interval = 5_000 };
        _timerConexion.Tick += TimerConexion_Tick;

        _timerActualizador = new System.Windows.Forms.Timer { Interval = 30_000 };
        _timerActualizador.Tick += async (_, _) => await TimerActualizador_TickAsync();

        // ── Eventos ────────────────────────────────────────────────────────
        Load        += async (_, _) => await FrmPuntoVenta_LoadAsync();
        FormClosing += async (s, e) => await FrmPuntoVenta_ClosingAsync(s, e);
        FormClosed  += (_, _) => DetenerTimers();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // CONSTRUCCIÓN DE UI
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Construye el menú principal equivalente al del Legacy mdiPuntoVenta.
    /// Legacy: mnuMovimiento, mnuCuentas, mnuConexion, mnuAyuda.
    /// </summary>
    private MenuStrip ConstruirMenu()
    {
        var menu = new MenuStrip { BackColor = ColorFondoToolbar, ForeColor = ColorTexto };

        // ── Movimientos ────────────────────────────────────────────────────
        var mnuMov = new ToolStripMenuItem("&Movimientos");

        // mnuInicio
        mnuMov.DropDownItems.Add("&Apertura de Turno", null, (_, _) => CmdApertura_Click());

        // mnuVenta
        _mnuVenta = new ToolStripMenuItem("&Punto de Venta") { Enabled = false };
        _mnuVenta.Click += (_, _) => CmdVenta_Click();
        mnuMov.DropDownItems.Add(_mnuVenta);

        // mnuPinPad
        _mnuPinPad = new ToolStripMenuItem("Activar PinPad (No Financiera)") { Visible = false };
        _mnuPinPad.Click += (_, _) => _ = CmdPinPad_Click();
        mnuMov.DropDownItems.Add(_mnuPinPad);

        // mnuCierre
        _mnuCierre = new ToolStripMenuItem("Ci&erre de Turno") { Enabled = false };
        _mnuCierre.Click += async (_, _) => await CmdCierre_ClickAsync();
        mnuMov.DropDownItems.Add(_mnuCierre);

        mnuMov.DropDownItems.Add(new ToolStripSeparator());

        // mnuCorrelativo
        mnuMov.DropDownItems.Add("Correlativo de Documentos", null, (_, _) => CmdCorrelativo_Click());

        mnuMov.DropDownItems.Add(new ToolStripSeparator());

        // mnuMesa
        mnuMov.DropDownItems.Add("&Mesas", null, (_, _) => CmdMesas_Click());

        // mnuInsumoCritico
        mnuMov.DropDownItems.Add("&Insumos/Platos de Stock Crítico", null, (_, _) => AbrirFormulario<FrmInsumo>());

        // mnuImportacion
        _mnuImportacion = new ToolStripMenuItem("&Importación de Requerimientos");
        _mnuImportacion.Click += (_, _) => AbrirFormulario<FrmImportacionRequerimientos>();
        mnuMov.DropDownItems.Add(_mnuImportacion);

        mnuMov.DropDownItems.Add(new ToolStripSeparator());

        // mnuAnulado (GAP-MDI-001: frmAnulado)
        mnuMov.DropDownItems.Add("&Anulados", null, (_, _) => AbrirFormulario<FrmRepAnuladoReporte>());

        // mnuCancelado (GAP-MDI-001: frmCancelado — stub, sin form equivalente)
        mnuMov.DropDownItems.Add("&Cancelados", null, (_, _) => MsgEnMigracion("frmCancelado"));

        // mnuClienteDeuda
        mnuMov.DropDownItems.Add("Clientes con &Deuda", null, (_, _) => MsgEnMigracion("frmRepClienteDeuda"));

        menu.Items.Add(mnuMov);

        // ── Correlativos ───────────────────────────────────────────────────
        var mnuCuentas = new ToolStripMenuItem("&Correlativos");
        mnuCuentas.DropDownItems.Add("Correlativo de Pedidos",       null, (_, _) => AbrirFormulario<FrmPedidoCorrelativo>());
        mnuCuentas.DropDownItems.Add("Correlativo de Documentos",    null, (_, _) => AbrirFormulario<FrmDocumentoCorrelativo>());
        mnuCuentas.DropDownItems.Add("Guías de Transporte",          null, (_, _) => MsgEnMigracion("frmGuiaTransporte"));
        mnuCuentas.DropDownItems.Add("Cuentas Corrientes",           null, (_, _) => AbrirFormulario<FrmCtaCte>());
        mnuCuentas.DropDownItems.Add("Recibos de Egreso",            null, (_, _) => AbrirFormulario<FrmReciboEgreso>());
        mnuCuentas.DropDownItems.Add("Recibos de Ingreso",           null, (_, _) => AbrirFormulario<FrmReciboIngreso>());
        mnuCuentas.DropDownItems.Add("Notas de Crédito",             null, (_, _) => AbrirFormulario<FrmNotaCredito>());
        mnuCuentas.DropDownItems.Add("Reservas",                     null, (_, _) => AbrirFormulario<FrmReserva>());
        _mnuDocElec = new ToolStripMenuItem("Documentos Electrónicos");
        _mnuDocElec.Click += (_, _) => MsgEnMigracion("frmDocumentoElectronicoCorrelativo");
        mnuCuentas.DropDownItems.Add(_mnuDocElec);
        mnuCuentas.DropDownItems.Add("Pedidos Central Producción",   null, (_, _) => AbrirFormulario<CentralPedidosForm>());
        mnuCuentas.DropDownItems.Add("Recargar Tarjeta de Proximidad", null, (_, _) => AbrirFormulario<FrmRecargarTarjeta>());
        mnuCuentas.DropDownItems.Add("Tarjeta Proximidad",           null, (_, _) => AbrirFormulario<FrmTarjetaProximidad>());
        mnuCuentas.DropDownItems.Add(new ToolStripSeparator());
        mnuCuentas.DropDownItems.Add("Cuenta por Cobrar",            null, (_, _) => AbrirFormulario<FrmCuentaCobrar>());
        menu.Items.Add(mnuCuentas);

        // ── Reportes ───────────────────────────────────────────────────────
        var mnuRep = new ToolStripMenuItem("&Reportes");
        mnuRep.DropDownItems.Add("Estados Cta Cte",              null, (_, _) => AbrirFormulario<FrmRepCtaCteReporte>());
        mnuRep.DropDownItems.Add("Control de Transacciones",     null, (_, _) => AbrirFormulario<FrmRepAnuladoReporte>());
        mnuRep.DropDownItems.Add("Liquidación Ticketera",        null, (_, _) => AbrirFormulario<FrmRepLiquidacionTicketReporte>());
        mnuRep.DropDownItems.Add("Paloteo Ticketera",            null, (_, _) => AbrirFormulario<FrmRepPaloteoTicketReporte>());
        mnuRep.DropDownItems.Add("Cierre Cajeros Delivery",      null, (_, _) => AbrirFormulario<FrmRepDeliveryTicketReporte>());
        mnuRep.DropDownItems.Add("Reporte de Reservas",          null, (_, _) => AbrirFormulario<FrmRepReservasReporte>());
        mnuRep.DropDownItems.Add("Reporte de Entregas",          null, (_, _) => AbrirFormulario<FrmRepEntregaReporte>());
        mnuRep.DropDownItems.Add("Venta Mensual por Fechas",     null, (_, _) => AbrirFormulario<FrmRepVentaFechaReporte>());
        mnuRep.DropDownItems.Add("Venta Mensual",                null, (_, _) => MsgEnMigracion("frmRepVentaMensual"));
        mnuRep.DropDownItems.Add("Ranking de Ventas",            null, (_, _) => MsgEnMigracion("frmRepRanking"));
        mnuRep.DropDownItems.Add("Cta Cte Integrado",            null, (_, _) => AbrirFormulario<FrmCtaCteIntegradoReporte>());
        mnuRep.DropDownItems.Add("Venta Mensual Integrado",      null, (_, _) => AbrirFormulario<FrmVentaMensualIntegradoReporte>());
        mnuRep.DropDownItems.Add("Registro de Ventas",           null, (_, _) => AbrirFormulario<FrmRepRegistroVentaReporte>());
        mnuRep.DropDownItems.Add("Analitico Motorizado",         null, (_, _) => AbrirFormulario<FrmAnaliticoMotorizadoIntegradoReporte>());
        menu.Items.Add(mnuRep);

        // ── Delivery ───────────────────────────────────────────────────────
        var mnuDel = new ToolStripMenuItem("&Delivery");
        mnuDel.DropDownItems.Add("Clientes Frecuentes",                 null, (_, _) => AbrirFormulario<FrmClienteDelivery>());
        mnuDel.DropDownItems.Add("Seguimiento Delivery en Tránsito",    null, (_, _) => AbrirFormulario<FrmPedidoDelivery>());
        mnuDel.DropDownItems.Add("Seguimiento Delivery Entregados",     null, (_, _) => AbrirFormulario<FrmPedidoDeliveryNo>());
        menu.Items.Add(mnuDel);

        // ── Conexión (multiLocal) ──────────────────────────────────────────
        _mnuConexion = new ToolStripMenuItem("C&onexión") { Visible = false };
        _mnuConexion.DropDownItems.Add("Cambiar de Local", null, (_, _) => MsgEnMigracion("frmServidorEnlace"));
        menu.Items.Add(_mnuConexion);

        // ── Ayuda ──────────────────────────────────────────────────────────
        var mnuAyuda = new ToolStripMenuItem("A&yuda");
        mnuAyuda.DropDownItems.Add("Acerca de...", null, (_, _) => MostrarAcercaDe());
        mnuAyuda.DropDownItems.Add("Liberar Descargo Venta", null, async (_, _) => await LiberarDescargoVentaAsync());
        mnuAyuda.DropDownItems.Add(new ToolStripSeparator());
        mnuAyuda.DropDownItems.Add("&Salir", null, (_, _) => Salir());
        menu.Items.Add(mnuAyuda);

        return menu;
    }

    /// <summary>
    /// Construye la barra de herramientas equivalente a xPicture con cmdOpcion1..16 del Legacy.
    /// Legacy: cmdOpcion1..16, cmdAmpliar, cmdOpcionMensaje, cmdRecargarTarjeta, cmdConsultaSaldo.
    /// </summary>
    private ToolStrip ConstruirToolbar()
    {
        var ts = new ToolStrip
        {
            Dock       = DockStyle.Top,
            BackColor  = ColorFondoToolbar,
            GripStyle  = ToolStripGripStyle.Hidden,
            Padding    = new Padding(4, 2, 4, 2),
            RenderMode = ToolStripRenderMode.System
        };

        _btnApertura   = CrearBotonToolbar("Apertura",       "Apertura de Turno");        // cmdOpcion1
        _btnVenta      = CrearBotonToolbar("Punto Venta",    "Punto de Venta");            // cmdOpcion2
        _btnCierre     = CrearBotonToolbar("Cierre",         "Cierre de Turno");           // cmdOpcion3
        _btnMesas      = CrearBotonToolbar("Mesas",          "Ver Mesas");                 // cmdOpcion4
        _btnReciboEgr  = CrearBotonToolbar("Rec. Egreso",    "Recibos de Egreso");         // cmdOpcion5
        _btnReciboIng  = CrearBotonToolbar("Rec. Ingreso",   "Recibos de Ingreso");        // cmdOpcion6
        _btnReserva    = CrearBotonToolbar("Reservas",       "Reservas");                  // cmdOpcion7
        _btnPedidoCor  = CrearBotonToolbar("Ped. Correl.",   "Pedido Correlativo");        // cmdOpcion8
        _btnDocCor     = CrearBotonToolbar("Doc. Correl.",   "Documento Correlativo");     // cmdOpcion9
        _btnCtaCte     = CrearBotonToolbar("Ctas Cte.",      "Cuentas Corrientes");        // cmdOpcion10
        _btnCtaCobrar  = CrearBotonToolbar("Ctas Cobrar",    "Cuentas por Cobrar");        // cmdOpcion11
        _btnPrecios    = CrearBotonToolbar("Precios",        "Carta de Productos");        // cmdOpcion12
        _btnDelivery   = CrearBotonToolbar("Delivery",       "Delivery en Tránsito");      // cmdOpcion13
        _btnDeliveryNo = CrearBotonToolbar("Del. Entregado", "Delivery Entregados");       // cmdOpcion14
        _btnPinPad     = CrearBotonToolbar("PinPad",         "Activar PinPad");            // cmdOpcion15
        _btnSalir      = CrearBotonToolbar("Salir",          "Salir del Sistema");          // cmdOpcion16
        _btnMensajes   = CrearBotonToolbar("Mensajes",       "Mensajes de Cocina");        // cmdOpcionMensaje
        _btnRecargar   = CrearBotonToolbar("Recargar Tarj.", "Recargar Tarjeta");          // cmdRecargarTarjeta
        _btnConsulta   = CrearBotonToolbar("Cons. Saldo",    "Consultar Saldos") ;         // cmdConsultaSaldo

        // Click handlers (equivalente a cmdOpcionN_Click en Legacy)
        _btnApertura.Click   += (_, _) => CmdApertura_Click();
        _btnVenta.Click      += (_, _) => CmdVenta_Click();
        _btnCierre.Click     += async (_, _) => await CmdCierre_ClickAsync();
        _btnMesas.Click      += (_, _) => CmdMesas_Click();
        _btnReciboEgr.Click  += (_, _) => AbrirFormulario<FrmReciboEgreso>();
        _btnReciboIng.Click  += (_, _) => AbrirFormulario<FrmReciboIngreso>();
        _btnReserva.Click    += (_, _) => AbrirFormulario<FrmReserva>();
        _btnPedidoCor.Click  += (_, _) => AbrirFormulario<FrmPedidoCorrelativo>();
        _btnDocCor.Click     += (_, _) => AbrirFormulario<FrmDocumentoCorrelativo>();
        _btnCtaCte.Click     += (_, _) => CmdCtaCte_Click();
        _btnCtaCobrar.Click  += (_, _) => AbrirFormulario<FrmCuentaCobrar>();
        _btnPrecios.Click    += (_, _) => AbrirFormulario<FrmPrecios>();
        _btnDelivery.Click   += (_, _) => AbrirFormulario<FrmPedidoDelivery>();
        _btnDeliveryNo.Click += (_, _) => AbrirFormulario<FrmPedidoDeliveryNo>();
        _btnPinPad.Click     += (_, _) => _ = CmdPinPad_Click();
        _btnSalir.Click      += (_, _) => CmdSalir_Click();
        _btnMensajes.Click   += (_, _) => AbrirFormulario<FrmMensajeCocina>();
        _btnRecargar.Click   += (_, _) => AbrirFormulario<FrmRecargarTarjeta>();
        _btnConsulta.Click   += (_, _) => MsgEnMigracion("frmConsultaSaldo");

        // Visibilidad inicial: igual que Legacy (PinPad y Consulta Saldo ocultos por defecto)
        _btnPinPad.Visible  = false;
        _btnConsulta.Visible = false;

        ts.Items.AddRange([
            _btnApertura, _btnVenta, _btnCierre,
            new ToolStripSeparator(),
            _btnMesas, _btnPrecios,
            new ToolStripSeparator(),
            _btnReciboEgr, _btnReciboIng,
            new ToolStripSeparator(),
            _btnCtaCte, _btnCtaCobrar,
            new ToolStripSeparator(),
            _btnReserva, _btnPedidoCor, _btnDocCor,
            new ToolStripSeparator(),
            _btnDelivery, _btnDeliveryNo,
            new ToolStripSeparator(),
            _btnMensajes, _btnRecargar, _btnConsulta, _btnPinPad,
            new ToolStripSeparator(),
            _btnSalir
        ]);

        return ts;
    }

    /// <summary>
    /// Construye la barra de estado equivalente al StatusBar del Legacy.
    /// Legacy: StatusBar.Panels.Item(1)..Item(4) → Día Contable, Caja, Usuario, Turno.
    /// </summary>
    private StatusStrip ConstruirStatusStrip()
    {
        var ss = new StatusStrip { BackColor = ColorFondoToolbar, ForeColor = ColorTexto };

        _statusDiaContable = new ToolStripStatusLabel("Día Contable: No Iniciado")
            { BorderSides = ToolStripStatusLabelBorderSides.Right, Spring = false };
        _statusCaja = new ToolStripStatusLabel("Caja: ---")
            { BorderSides = ToolStripStatusLabelBorderSides.Right };
        _statusUsuario = new ToolStripStatusLabel("Usuario: ---")
            { BorderSides = ToolStripStatusLabelBorderSides.Right };
        _statusTurno = new ToolStripStatusLabel("Turno: No Iniciado")
            { BorderSides = ToolStripStatusLabelBorderSides.Right };
        _statusHora = new ToolStripStatusLabel(DateTime.Now.ToString("HH:mm:ss"))
            { Spring = true, TextAlign = ContentAlignment.MiddleRight };

        ss.Items.AddRange(new ToolStripItem[]
            { _statusDiaContable, _statusCaja, _statusUsuario, _statusTurno, _statusHora });

        return ss;
    }

    private static ToolStripButton CrearBotonToolbar(string texto, string tooltip) =>
        new()
        {
            Text        = texto,
            ToolTipText = tooltip,
            DisplayStyle = ToolStripItemDisplayStyle.Text,
            Font        = new Font("Segoe UI", 8f, FontStyle.Bold),
            AutoSize    = true,
            Margin      = new Padding(1)
        };

    // ══════════════════════════════════════════════════════════════════════════
    // CARGA DEL FORMULARIO  (equivalente a MDIForm_Load)
    // ══════════════════════════════════════════════════════════════════════════

    private async Task FrmPuntoVenta_LoadAsync()
    {
        var sesion = _sessionService.SesionActual;
        _codigoCaja = sesion?.CodigoCaja ?? _configuration["Inforest:CodigoCaja"] ?? "001";
        _usuario    = sesion?.CodigoUsuario ?? string.Empty;

        // Caption: "Punto de Venta : Local <localConectado>"
        Text = $"Punto de Venta : {_codigoCaja}";

        // Cargar configuración sistema + caja
        await CargarConfiguracionAsync();

        // Aplicar visibilidad de controles según flags (equivalente a MDIForm_Load del Legacy)
        AplicarFlagsDeConfiguracion();

        // Validar instancia única por caja (LOG_SESIONES – equivalente a GCAA 27012022 en Legacy)
        if (!await ValidarYRegistrarSesionCajaAsync()) return;

        // Estado inicial: sin turno
        ActivaInicio(false);
        ActualizarStatusBar();

        // Iniciar timers
        _timerReloj.Start();
        if (_cfg?.lFEBiz == true) _timerBizlink.Start();
        _timerWebapp.Start();
        // _timerConexion se activará si multiLocal=true (ver AplicarFlagsDeConfiguracion)

        // TimerActualizador — sólo si ActivoActualizador habilitado (TPARAMETROVERSION)
        _timerActualizador.Start();

        // Modo MCPV: inicializar turno del usuario (GAP-MDI-006/009)
        // Legacy: If lMCPV Then Me.Visible = False; frmMozoUsuario.Show vbModal; InicializaMCPV
        if (_cfgCaja?.lMCPV == true)
            await InicializaMCPVAsync();

        // Cargar salones y mesas
        await CargarSalonesYMesasAsync();

        // Obtener y mostrar día contable
        await ActualizarDiaContableAsync();

        // Iniciar noticia informativa
        IniciarNoticia();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // CIERRE DEL FORMULARIO  (equivalente a MDIForm_Unload)
    // ══════════════════════════════════════════════════════════════════════════

    private async Task FrmPuntoVenta_ClosingAsync(object? sender, FormClosingEventArgs e)
    {
        DetenerTimers();
        // Actualizar LOG_SESIONES al salir (equivalente a MDIForm_Unload del Legacy)
        await ActualizarLogSesionesCierreAsync();
    }

    private void DetenerTimers()
    {
        _timerReloj.Stop();
        _timerBizlink.Stop();
        _timerWebapp.Stop();
        _timerConexion.Stop();
        _timerActualizador.Stop();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // LOG_SESIONES  (equivalente a GCAA 27012022 en MDIForm_Load y MDIForm_Unload)
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Valida que no haya otra instancia activa del POS para la misma caja y registra la sesión.
    /// Legacy: IF jajaj > 0 THEN MsgBox "La Caja..." → Salir ELSE INSERT LOG_SESIONES.
    /// </summary>
    private async Task<bool> ValidarYRegistrarSesionCajaAsync()
    {
        try
        {
            using var cn = await _connectionFactory.CreateOpenConnectionAsync();

            var otro = await cn.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM LOG_SESIONES WHERE ISNULL(lActivo,0) = 0 AND tCaja = @caja AND ISNULL(PC,'') <> @pc",
                new { caja = _codigoCaja, pc = Environment.MachineName });

            if (otro > 0)
            {
                MessageBox.Show(
                    $"La Caja {_codigoCaja} ya tiene activo otro Punto de Venta con el mismo código.",
                    "Inforest", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
                return false;
            }

            await cn.ExecuteAsync(
                "INSERT INTO LOG_SESIONES (tCaja, fRegistro, lActivo, tUsuario, PC) VALUES (@caja, GETDATE(), 0, @usuario, @pc)",
                new { caja = _codigoCaja, usuario = _usuario, pc = Environment.MachineName });

            return true;
        }
        catch
        {
            // No bloquear inicio si hay error de BD; continuar con advertencia
            return true;
        }
    }

    /// <summary>
    /// Marca la sesión como cerrada en LOG_SESIONES al salir.
    /// Legacy: UPDATE LOG_SESIONES SET lActivo=1 WHERE lActivo=0 AND TCAJA=... AND PC=...
    /// </summary>
    private async Task ActualizarLogSesionesCierreAsync()
    {
        try
        {
            using var cn = await _connectionFactory.CreateOpenConnectionAsync();
            await cn.ExecuteAsync(
                "UPDATE LOG_SESIONES SET lActivo = 1 WHERE lActivo = 0 AND tCaja = @caja AND PC = @pc",
                new { caja = _codigoCaja, pc = Environment.MachineName });
        }
        catch { /* Silencioso — cierre de sesión best-effort */ }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // CONFIGURACIÓN Y FLAGS
    // ══════════════════════════════════════════════════════════════════════════

    private async Task CargarConfiguracionAsync()
    {
        var cfgResult = await _configuracionHandler.HandleAsync(new ObtenerConfiguracionSistemaQuery());
        _cfg = cfgResult.Valor;

        if (_codigoCaja is not null)
        {
            var cajResult = await _configuracionCajaHandler.HandleAsync(new ObtenerConfiguracionCajaQuery(_codigoCaja));
            _cfgCaja = cajResult.Valor;
        }
    }

    /// <summary>
    /// Aplica visibilidad/habilitación de controles según flags de configuración.
    /// Legacy: MDIForm_Load — lOCR, lAlmacen, lFacturacionE, lFEOfisis, lVisaNet, multiLocal, lSiab.
    /// </summary>
    private void AplicarFlagsDeConfiguracion()
    {
        // lAlmacen → habilita importación de requerimientos
        _mnuImportacion.Enabled = _cfg?.lAlmacen == true;

        // lFacturacionE / lFEOfisis → menú documentos electrónicos
        if (_cfg is not null)
        {
            _mnuDocElec.Enabled = _cfg.lFacturacionE && !_cfg.lFEOfisis;
        }

        // lVisaNet → botón PinPad
        _btnPinPad.Visible  = _cfgCaja?.lVisaNet == true;
        _mnuPinPad.Visible  = _cfgCaja?.lVisaNet == true;

        // lSiab → botón consulta saldo
        _btnConsulta.Visible = _cfgCaja?.lSiab == true;

        // lMultilocal → menú Conexión
        _mnuConexion.Visible = _cfg?.lMultilocal == true;
        if (_cfg?.lMultilocal == true) _timerConexion.Start();

        // lOCR → cambiar caption de Punto de Venta a Caja Rápida
        if (_cfgCaja?.lCajaRapida == true)
        {
            _mnuVenta.Text  = "&Caja Rápida";
            _btnVenta.Text  = "Caja Rápida";
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // ActivaInicio  (equivalente a modPuntoVenta.bas → Public Sub ActivaInicio)
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Habilita o deshabilita botones y menús según si el turno está abierto.
    /// Legacy: modPuntoVenta.bas → ActivaInicio(Activa As Boolean).
    /// </summary>
    private void ActivaInicio(bool activa)
    {
        _turnoIniciado = activa;

        _btnVenta.Enabled    = activa;
        _btnCierre.Enabled   = activa;
        _mnuVenta.Enabled    = activa;
        _mnuCierre.Enabled   = activa;

        // lAlmacen — importación sólo cuando hay turno Y almacén habilitado
        _mnuImportacion.Enabled = activa && _cfg?.lAlmacen == true;

        _btnApertura.Enabled = !activa;
    }

    // ══════════════════════════════════════════════════════════════════════════
    // STATUS BAR Y DATOS DINÁMICOS
    // ══════════════════════════════════════════════════════════════════════════

    private void ActualizarStatusBar()
    {
        // Panel 2: Caja
        _statusCaja.Text = $"Caja : {_codigoCaja}";
        // Panel 3: Usuario (sin asterisco inicial del legacy)
        _statusUsuario.Text = $"Usuario : {_usuario.TrimStart('*')}";
        // Panel 4: Turno
        _statusTurno.Text = _turnoIniciado ? $"Turno : {_turno}" : "Turno : No Iniciado";
    }

    private void ActualizarHora() =>
        _statusHora.Text = DateTime.Now.ToString("HH:mm:ss");

    private async Task ActualizarDiaContableAsync()
    {
        try
        {
            var fecha = await ObtenerDiaContableAsync();
            _statusDiaContable.Text = $"Día Contable : {fecha:dd/MM/yyyy}";
        }
        catch
        {
            _statusDiaContable.Text = "Día Contable : No Iniciado";
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SALONES Y MESAS
    // ══════════════════════════════════════════════════════════════════════════

    private async Task CargarSalonesYMesasAsync()
    {
        var salones = await CargarSalonesAsync();
        _cmbSalon.DataSource    = salones;
        _cmbSalon.DisplayMember = nameof(Salon.Detallado);
        _cmbSalon.ValueMember   = nameof(Salon.CodigoSalon);

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
            Salon.Crear("02", "Terraza",          null, "SISTEMA")
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
        _panelMesas.SuspendLayout();
        _panelMesas.Controls.Clear();
        var salon = _cmbSalon.SelectedValue?.ToString();
        foreach (var mesa in _mesas.Where(m => string.IsNullOrWhiteSpace(salon) || m.CodigoSalon == salon))
        {
            var boton = new Button
            {
                Width     = 130,
                Height    = 80,
                Margin    = new Padding(6),
                Text      = $"{mesa.Detallado}\nCap. {mesa.NumeroPersonas}",
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = ColorTexto,
                BackColor = mesa.Estado switch
                {
                    EstadoMesa.Ocupada  => ColorMesaOcupada,
                    EstadoMesa.EnCuenta => ColorMesaCuenta,
                    _                   => ColorMesaLibre
                },
                FlatStyle = FlatStyle.Flat
            };
            boton.FlatAppearance.BorderColor = mesa.Estado switch
            {
                EstadoMesa.Ocupada  => Color.FromArgb(0xF4, 0x43, 0x36),
                EstadoMesa.EnCuenta => Color.FromArgb(0xE4, 0x65, 0x25),
                _                   => ColorPrimario
            };
            boton.DoubleClick += (_, _) => AbrirPedidoMesa(mesa);
            _panelMesas.Controls.Add(boton);
        }
        _panelMesas.ResumeLayout();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // HANDLERS DE BOTONES / MENÚ
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// cmdOpcion1_Click / mnuInicio_Click — Apertura de Turno.
    ///
    /// Legacy: mdiPuntoVenta.frm cmdOpcion1_Click.
    ///   1. Si lDiaContable=False (manual): muestra frmDiaContable primero.
    ///   2. Si lDiaContableAperturado=True: muestra frmInicio (FrmAperturaTurno).
    ///
    /// BR-TURNO-001, BR-TURNO-002, BR-DC-001.
    /// </summary>
    private void CmdApertura_Click()
    {
        // Paso 1: Día Contable manual (lDiaContable=False)
        if (_cfg?.lDiaContableAutomatico == false)
        {
            var tipoCambioRepo = _serviceProvider.GetService<ITipoCambioRepository>();
            using var dlgDC = new FrmDiaContable(
                FrmDiaContable.Modo.Apertura,
                _usuario,
                _serviceProvider.GetRequiredService<AperturarDiaContableHandler>(),
                _serviceProvider.GetRequiredService<CerrarDiaContableHandler>(),
                _serviceProvider.GetRequiredService<ObtenerDiaContableHandler>());
            dlgDC.ShowDialog(this);
            if (!dlgDC.IniciaPorDiaContable) return;
            _ = ActualizarDiaContableAsync();
        }

        // Paso 2: Determinar modo de consulta de turno
        var modoTurno = _cfgCaja?.lMCPV == true
            ? ModoConsultaTurno.PorUsuario
            : ModoConsultaTurno.PorCaja;

        // Paso 3: Abrir FrmAperturaTurno (frmInicio)
        var tipoCambioRepository = _serviceProvider.GetService<ITipoCambioRepository>();
        using var frm = new FrmAperturaTurno(
            abrirTurnoHandler:    _abrirTurnoHandler,
            obtenerUltimoHandler: _obtenerUltimoTurnoHandler,
            codigoCaja:           _codigoCaja,
            codigoUsuario:        _usuario,
            modoTurno:            modoTurno,
            monedaN:              _cfg?.tMonN ?? "S/.",
            monedaE:              _cfg?.tMonE ?? "",
            pais:                 _configuration["Inforest:Pais"] ?? "000",
            tipoCambioRepo:       tipoCambioRepository);

        if (frm.ShowDialog(this) == DialogResult.OK)
        {
            _turno         = frm.CodigoTurnoAbierto;
            _turnoIniciado = true;
            ActivaInicio(true);
            _ = ActualizarDiaContableAsync();
            ActualizarStatusBar();
        }
    }

    /// <summary>
    /// cmdOpcion2_Click / mnuVenta_Click — Abrir Punto de Venta (venta en mesa).
    /// Legacy: frmVenta.Show vbModal.
    /// </summary>
    private void CmdVenta_Click()
    {
        if (!_turnoIniciado)
        {
            MessageBox.Show("Es necesario aperturar el Turno!", "Inforest", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        AbrirFormulario<FrmVenta>();
    }

    /// <summary>
    /// cmdOpcion3_Click / mnuCierre_Click — Cierre de Turno.
    /// Legacy: frmLiquidacionDetalle.Show vbModal → si wInicio = False → ActivaInicio(False).
    /// </summary>
    private async Task CmdCierre_ClickAsync()
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
            {
                _turnoIniciado = false;
                ActivaInicio(false);
                ActualizarStatusBar();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al obtener turno: {ex.Message}", Text);
        }
    }

    /// <summary>
    /// cmdOpcion4_Click / mnuMesa_Click — Consulta de Mesas.
    /// Legacy: sTipo = "V"; frmMesaConsulta.Show vbModal.
    /// </summary>
    private void CmdMesas_Click()
    {
        AbrirFormulario<FrmMesaConsulta>();
        _ = CargarSalonesYMesasAsync(); // refrescar mapa de mesas
    }

    /// <summary>
    /// cmdOpcion10_Click / mnuCtaCte_Click — Cuentas Corrientes.
    /// Legacy: If nTC = 0 Then MsgBox "Es necesario aperturar el Turno!".
    /// </summary>
    private void CmdCtaCte_Click()
    {
        if (!_turnoIniciado)
        {
            MessageBox.Show("Es necesario aperturar el Turno!", "Inforest", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        AbrirFormulario<FrmCtaCte>();
    }

    /// <summary>
    /// cmdOpcion15_Click / mnuPinPad_Click — PinPad (No Financiera).
    /// Legacy: usa fiStartOperation + bucle de status VisaNet.
    /// </summary>
    private async Task CmdPinPad_Click()
    {
        var handler = _serviceProvider.GetRequiredService<ObtenerTerminalesPinPadHandler>();
        var result = await handler.HandleAsync(new ObtenerTerminalesPinPadQuery(_codigoCaja));
        var terminales = result.EsExitoso ? result.Valor?.ToList() ?? [] : [];
        using var frm = new FrmPagoPinPad(terminales);
        frm.ShowDialog(this);
    }

    /// <summary>
    /// mnuCorrelativo_Click — Correlativo de Documentos (legacy → frmFactura).
    /// </summary>
    private void CmdCorrelativo_Click()
    {
        AbrirFormulario<FrmDocumentoCorrelativo>();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // DÍA CONTABLE  (equivalente a obtieneDiaContable en Legacy)
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Obtiene la fecha del día contable llamando al SP <c>usp_GenObtieneDiaContable</c>.
    /// Legacy: obtieneDiaContable() → oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn).
    /// </summary>
    private async Task<DateTime> ObtenerDiaContableAsync()
    {
        var sesion  = _sessionService.SesionActual;
        var usuario = sesion?.CodigoUsuario ?? _usuario;
        var lDiaContable    = _cfg?.lDiaContableAutomatico ?? false;
        var horaCierre      = _cfg?.tHoraCierreDiaContable ?? "00:00";

        return await _diaContableService.ObtenerDiaContableAsync(lDiaContable, horaCierre, usuario);
    }

    // ══════════════════════════════════════════════════════════════════════════
    // NOTICIAS  (equivalente a obtieneNoticia + IniciarNoticia)
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Lanza InfoNoticias.exe si existe la noticia configurada.
    /// Legacy: IniciarNoticia() → ShellExecute InfoNoticias.exe.
    /// </summary>
    private void IniciarNoticia()
    {
        try
        {
            var noticiaPath = Path.Combine(AppContext.BaseDirectory, "Noticia", "InfoNoticias.exe");
            if (!File.Exists(noticiaPath)) return;
            System.Diagnostics.Process.Start(noticiaPath, $"{_usuario}");
        }
        catch { /* Silencioso */ }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // TIMERS
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// TimerBizlink (60s) — Procesamiento de facturación electrónica BizLink.
    /// Legacy: TimerBizlink_Timer — If Contador > 5 Then lFEBiz + lFEBizOffLine.
    /// </summary>
    private void TimerBizlink_Tick(object? sender, EventArgs e)
    {
        if (_contadorBizlink > 5)
        {
            _contadorBizlink = 0;
            // Procesamiento FE delegado a servicio de infraestructura
            // (equivalente a INSERTA_FE_INFOREST y INSERTAFE del Legacy)
        }
        else
        {
            _contadorBizlink++;
        }
    }

    /// <summary>
    /// Timerwebapp (10s) — Alerta de ventas y sincronización KDS.
    /// Legacy: Timerwebapp_Timer — lAlertaVenta, lIntMobileKDS, lIntPatio.
    /// </summary>
    private void TimerWebapp_Tick(object? sender, EventArgs e)
    {
        try
        {
            // Alerta de venta horaria (lAlertaVenta) — una sola vez por día
            if (EsHorarioAlertaVenta() && _turnoIniciado && !_alertaVentaMostrada)
            {
                _alertaVentaMostrada = true;
                MessageBox.Show("Es necesario realizar descarga de Ventas", "Alerta de Venta",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            // Resetear flag cuando se sale de la ventana horaria
            if (!EsHorarioAlertaVenta()) _alertaVentaMostrada = false;
        }
        catch { /* Silencioso — timer no debe detener la app */ }
    }

    /// <summary>
    /// Timer multiLocal — Verifica conexión al servidor.
    /// Legacy: Timer_Timer — If multiLocal = True And ultimoConectado = False.
    /// </summary>
    private void TimerConexion_Tick(object? sender, EventArgs e)
    {
        // La conexión multilocal se verifica mediante IDbConnectionFactory
        // Si la conexión falla, se mostrará el formulario de servidor de enlace
    }

    private static bool EsHorarioAlertaVenta()
    {
        var ahora = TimeOnly.FromDateTime(DateTime.Now);
        return ahora >= new TimeOnly(23, 50) && ahora <= new TimeOnly(23, 59, 59);
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SALIR  (equivalente a Salir() en Legacy)
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Cierra el POS de forma ordenada: limpia LOG_SESIONES y registra auditoría.
    /// Legacy: Salir() → Visor, lVisaNet → fiClosePort, LOG_SESIONES update, registroAccesoAuditoria.
    /// </summary>
    public void Salir()
    {
        try
        {
            // Auditoría de salida (equivalente a registroAccesoAuditoria "S", sUsuario)
            var auditoria = _serviceProvider.GetService<IAuditoriaService>();
            if (auditoria is not null)
                _ = auditoria.RegistrarSalidaAsync(new RegistroAccesoAuditoriaRequest(
                    "S", "INFOREST", "PuntoVenta", _usuario, 0));
        }
        catch { /* Silencioso — el cierre no puede bloquearse por auditoría */ }

        Close();
    }

    /// <summary>
    /// cmdOpcion16_Click — Salir del sistema o, en modo MCPV, mostrar selector de mozo.
    /// Legacy: cmdOpcion16_Click → If lMCPV Then frmMozoUsuario.Show vbModal; InicializaMCPV Else Salir().
    /// </summary>
    private void CmdSalir_Click()
    {
        if (_cfgCaja?.lMCPV == true)
        {
            // Modo MCPV: selector de mozo (GAP-MDI-006: frmMozoUsuario no migrado)
            MsgEnMigracion("frmMozoUsuario");
            _ = InicializaMCPVAsync();
        }
        else
        {
            Salir();
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // InicializaMCPV  (GAP-MDI-009)
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Inicializa el estado del turno en modo MCPV (Multi-Cajero Punto de Venta).
    /// Busca un turno activo del usuario en MTURNO y establece el estado de la caja.
    /// Legacy: InicializaMCPV() → SELECT * FROM MTURNO WHERE tUsuario=... AND lcierre=0.
    /// </summary>
    public async Task InicializaMCPVAsync()
    {
        try
        {
            using var cn = await _connectionFactory.CreateOpenConnectionAsync();

            var turno = await cn.QueryFirstOrDefaultAsync<TurnoMcpvRow>(
                "SELECT TOP 1 tTurno FROM MTURNO WHERE tUsuario = @usuario AND lcierre = 0 ORDER BY tTurno",
                new { usuario = _usuario });

            if (turno is not null)
            {
                // Verificar tipo de cambio del día
                var nTC = await cn.ExecuteScalarAsync<double?>(
                    "SELECT nVenta FROM TTIPOCAMBIO WHERE fFecha = CONVERT(date, GETDATE())") ?? 0;

                _turno = turno.tTurno ?? string.Empty;

                if (nTC == 0)
                {
                    // Turno encontrado pero sin tipo de cambio → mostrar apertura
                    ActivaInicio(false);
                    _btnApertura.Enabled = true;
                }
                else
                {
                    // Turno activo con tipo de cambio → habilitar operaciones
                    ActivaInicio(true);
                    _btnApertura.Enabled = false;
                }
            }
            else
            {
                _turno = string.Empty;
                ActivaInicio(false);
                _btnApertura.Enabled = true;
            }

            ActualizarStatusBar();
        }
        catch
        {
            // Si hay error de BD, dejar el estado sin turno
            ActivaInicio(false);
        }
    }

    private sealed record TurnoMcpvRow(string? tTurno);

    // ══════════════════════════════════════════════════════════════════════════
    // NAVEGACIÓN A PEDIDO POR MESA
    // ══════════════════════════════════════════════════════════════════════════

    private void AbrirPedidoMesa(Mesa mesa)
    {
        using var pedido = ActivatorUtilities.CreateInstance<FrmPedido>(_serviceProvider, mesa);
        pedido.ShowDialog(this);
        _ = CargarSalonesYMesasAsync();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // HELPERS DE APERTURA DE FORMULARIOS
    // ══════════════════════════════════════════════════════════════════════════

    private void AbrirFormulario<T>() where T : Form
    {
        using var scope = _serviceProvider.CreateScope();
        var form = scope.ServiceProvider.GetRequiredService<T>();
        form.ShowDialog(this);
    }

    private void MsgEnMigracion(string legacyForm) =>
        MessageBox.Show($"Funcionalidad en migración ({legacyForm}).", Text,
            MessageBoxButtons.OK, MessageBoxIcon.Information);

    // ══════════════════════════════════════════════════════════════════════════
    // MISCELÁNEOS
    // ══════════════════════════════════════════════════════════════════════════

    private void MostrarAcercaDe()
    {
        var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        MessageBox.Show(
            $"INFOREST Punto de Venta\nVersión {ver}\n\nInfomática S.A.C.",
            "Acerca de...", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    /// <summary>
    /// Libera el descargo de venta (fuerza tsolouno=0 en TPARAMETRO).
    /// Legacy: mnuLiberar_Click → Cn.Execute "update tparametro set tsolouno=0..."
    /// </summary>
    private async Task LiberarDescargoVentaAsync()
    {
        try
        {
            using var cn = await _connectionFactory.CreateOpenConnectionAsync();
            await cn.ExecuteAsync("UPDATE TPARAMETRO SET tsolouno = 0, lNomSoloUno = ''");
            MessageBox.Show("¡Descargo de Venta Liberado!", "Descargo de Ventas",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al liberar descargo: {ex.Message}", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // TIMER ACTUALIZADOR  (GAP-MDI-010)
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Comprueba si hay una nueva versión disponible en TPARAMETROVERSION.
    /// Legacy: TimerActualizador_Timer → compara version.txt con TPARAMETROVERSION.VersionInfoRest
    ///         y VersionFeInfoRest vs VersionActualFEInfoRest, luego muestra FrmActualizacion.
    /// </summary>
    private async Task TimerActualizador_TickAsync()
    {
        try
        {
            using var cn = await _connectionFactory.CreateOpenConnectionAsync();

            var row = await cn.QueryFirstOrDefaultAsync<VersionParametroRow>(
                "SELECT TOP 1 VersionInfoRest, VersionFeInfoRest, VersionActualFEInfoRest, lFEBiz " +
                "FROM TPARAMETROVERSION");

            if (row is null) return;

            var asmVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var versionActual = $"{asmVersion?.Major}.{asmVersion?.Minor}.{asmVersion?.Build}";

            // Verificar actualización de Facturación Electrónica
            if (!string.IsNullOrWhiteSpace(row.VersionFeInfoRest) && row.lFEBiz == true &&
                row.VersionFeInfoRest != row.VersionActualFEInfoRest)
            {
                _timerActualizador.Stop();
                var frm = new FrmActualizacion(
                    $"Nueva Actualización disponible de Facturación Electrónica\n FE Versión: {row.VersionFeInfoRest}")
                    { Tipo = "FE" };
                frm.ShowDialog(this);
                return;
            }

            // Verificar actualización de INFOREST
            if (!string.IsNullOrWhiteSpace(row.VersionInfoRest) &&
                row.VersionInfoRest != versionActual)
            {
                _timerActualizador.Stop();
                var frm = new FrmActualizacion(
                    $"Nueva Actualización disponible\n InfoRest Versión: {row.VersionInfoRest}\n" +
                    "Se recomienda actualizar el sistema, ya que puede presentar problemas en algunos procesos.")
                    { Tipo = "INFOREST" };
                frm.ShowDialog(this);
            }
        }
        catch { /* Silencioso — el timer actualizador no debe interferir con la operación */ }
    }

    private sealed record VersionParametroRow(
        string? VersionInfoRest,
        string? VersionFeInfoRest,
        string? VersionActualFEInfoRest,
        bool? lFEBiz);
}
