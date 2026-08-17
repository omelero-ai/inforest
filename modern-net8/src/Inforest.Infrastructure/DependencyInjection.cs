using Inforest.Application.Caja;
using Inforest.Application.Impresion;
using Inforest.Infrastructure.Almacen;
using Inforest.Infrastructure.Impresion;
using Inforest.Application.Kitchen;
using Inforest.Infrastructure.Notifications;
using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Application.Interfaces.Country;
using Inforest.Application.Interfaces.Hardware;
using Inforest.Application.Maestros;
using Inforest.Application.Pedidos;
using Inforest.Application.Reportes;
using Inforest.Application.Turno;
using Inforest.Application.Ventas;
using Inforest.Domain.Repositories;
using Inforest.Infrastructure.Caja;
using Inforest.Infrastructure.Configuracion;
using Inforest.Infrastructure.Country;
using Inforest.Infrastructure.Data;
using Inforest.Infrastructure.Delivery;
using Inforest.Infrastructure.Hardware;
using Inforest.Infrastructure.Kitchen;
using Inforest.Infrastructure.Maestros;
using Inforest.Infrastructure.Motorizado;
using Inforest.Infrastructure.Pedidos;
using Inforest.Infrastructure.Reportes;
using Inforest.Infrastructure.Security;
using Inforest.Infrastructure.Turno;
using Inforest.Infrastructure.Ventas;
using Inforest.Infrastructure.Reservas;
using Inforest.Infrastructure.Runtime;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace Inforest.Infrastructure;

/// <summary>
/// Configuración de DI para la capa Infrastructure.
/// </summary>
public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        // Configurar Serilog
        Log.Logger = new LoggerConfiguration()
            .ReadFrom.Configuration(configuration)
            .Enrich.FromLogContext()
            .WriteTo.Console()
            .WriteTo.File("logs/inforest-.log", rollingInterval: RollingInterval.Day)
            .CreateLogger();

        services.AddLogging(loggingBuilder =>
            loggingBuilder.AddSerilog(dispose: true));

        // P3-02: Contrato SQL y ejecución de datos Legacy
        services.AddSingleton<IDbConnectionFactory, SqlConnectionFactory>();
        services.AddSingleton<ISpExecutor, SpExecutor>();
        services.AddSingleton<IModernPasswordHashStore, ModernPasswordHashStore>();
        services.AddSingleton<ISessionService, SessionService>();
        services.AddSingleton<IAppInstanceGuard, AppInstanceGuard>();
        services.AddScoped<IAuthService, AuthService>();
        services.AddSingleton<IRbacService, RbacService>();
        services.AddScoped<IAuditoriaService, AuditoriaService>();
        services.AddScoped<ILicenseService, LicenseService>();
        services.AddSingleton<IImpresoraService, NullImpresoraService>();
        services.AddScoped<IKitchenConfigurationProvider, KitchenConfigurationProvider>();
        services.AddScoped<IProduccionCocinaService, ProduccionCocinaService>();
        services.AddScoped<IMensajeCocinaRepository, MensajeCocinaRepository>();
        services.AddScoped<IKdsLegacyGateway, KdsLegacyGateway>();
        services.AddScoped<IComandaAreaDispatcher, ComandaAreaDispatcher>();
        services.AddScoped<IKdsDispatcher, KdsXmlDispatcher>();

        // P3-09: Delivery, Despacho, Motorizados y Centralización
        services.AddScoped<IClienteDeliveryRepository, ClienteDeliveryRepository>();
        services.AddScoped<IPedidoDeliveryRepository, PedidoDeliveryRepository>();
        services.AddScoped<IMotorizadoRepository, MotorizadoRepository>();
        services.AddScoped<ILocalRepository, LocalRepository>();
        services.AddScoped<ICentralPedidosRepository, CentralPedidosRepository>();
        services.AddScoped<IRappiOrderService, RappiOrderAdapter>();

        // W1: Configuración — TPARAMETRO / TCAJA (BR-SQL-003, BR-SQL-004)
        services.AddScoped<IParametroRepository, ParametroRepository>();
        services.AddSingleton<IParametroService, ParametroService>();

        // W2: Maestros — Productos, Grupos, Subgrupos, Clientes, Mesas, Salones, Mozos
        services.AddScoped<IGrupoProductoRepository, GrupoProductoRepository>();
        services.AddScoped<ISubGrupoProductoRepository, SubGrupoProductoRepository>();
        services.AddScoped<IClienteRepository, ClienteRepository>();
        services.AddScoped<ICuentaCorrienteRepository, CuentaCorrienteRepository>();
        services.AddScoped<IMesaRepository, MesaRepository>();
        services.AddScoped<ISalonRepository, SalonRepository>();
        services.AddScoped<IMozoRepository, MozoRepository>();
        services.AddScoped<IProductoMaestroRepository, ProductoMaestroRepository>();
        services.AddScoped<IInsumoRepository, InsumoRepository>();

        // W3: Turno / Día Contable (BR-004, BR-005)
        services.AddScoped<ITurnoRepository, TurnoRepository>();
        services.AddScoped<IDiaContableService, DiaContableService>();

        // W4: Pedidos (BR-SQL-001, BR-SQL-002, BR-003)
        services.AddScoped<IPedidoRepository, PedidoRepository>();
        services.AddScoped<IPedidoReadRepository, PedidoRepository>();

        // W5: Venta / Documentos (BR-002, BR-013)
        services.AddScoped<IDocumentoRepository, DocumentoRepository>();
        services.AddScoped<ICorrelativoRepository, CorrelativoRepository>();
        // POS-FUNC-006: Notas de Crédito (BR-NC-001..006)
        services.AddScoped<INotaCreditoRepository, NotaCreditoRepository>();
        // POS-FUNC-016: Impresión pre-cuenta — TIMPRESORA (BR-PRECUENTA-001..003)
        services.AddScoped<IImpresoraRepository, ImpresoraRepository>();
        // POS-FUNC-027: Cambio de Documento (BR-CAMBIO-001..005)
        services.AddScoped<ICambioDocumentoRepository, CambioDocumentoRepository>();
        // POS-FUNC-028: Actualizar Datos Pedido — usp_ActualizarCabPeDoc (BR-ACTPED-001..003)
        services.AddScoped<IActualizarDatosPedidoRepository, ActualizarDatosPedidoRepository>();
        // POS-FUNC-029: Motivos Anulación — vMotivoAnulacion / TTABLA (BR-MOTIVO-001..003)
        services.AddScoped<IMotivoAnulacionRepository, MotivoAnulacionRepository>();

        // W6: Caja / Pagos (BR-007, BR-013)
        services.AddScoped<IPagoRepository, PagoRepository>();
        services.AddScoped<IMedioPagoRepository, MedioPagoRepository>();

        // POS-FUNC-018: Recibo Ingresos / Egresos — MINGRESO / MEGRESO (BR-RECIBO-001..012)
        services.AddScoped<IReciboIngresoRepository, ReciboIngresoRepository>();
        services.AddScoped<IReciboEgresoRepository, ReciboEgresoRepository>();

        // POS-FUNC-025: División de Pedidos — MPEDIDO/DPEDIDO (BR-DIV-001..005)
        services.AddScoped<IDivisionPedidoRepository, DivisionPedidoRepository>();

        // POS-FUNC-011: Reservas — TRESERVA
        services.AddScoped<IReservaRepository, ReservaRepository>();

        // W14: Almacén — BR-008 (descargo de inventario en venta)
        services.AddScoped<IInventoryGateway, InventoryGateway>();

        // POS-FUNC-014: Importación de requerimientos de almacén al POS
        services.AddScoped<IRequerimientoAlmacenRepository, RequerimientoAlmacenRepository>();
        services.AddScoped<IImportacionPedidoGateway, ImportacionPedidoGateway>();

        // W14: Notificaciones email — claCorreoElectronico.cls
        services.Configure<SmtpOptions>(configuration.GetSection(SmtpOptions.SectionName));
        services.AddScoped<INotificacionEmailService, SmtpEmailService>();

        // P3-10: Reportes — IReporteRepository + ReporteRepository (ADR-007)
        services.AddScoped<IReporteRepository, ReporteRepository>();

        // P3-11: Periféricos especializados e integraciones por país
        services.AddHardwareServices(configuration);
        services.AddCountryServices(configuration);

        return services;
    }

    private static void AddHardwareServices(this IServiceCollection services, IConfiguration configuration)
    {
        // PinPad — BR-PERIPH-001, ADR-009
        services.Configure<PinPadOptions>(configuration.GetSection(PinPadOptions.SectionName));
        services.AddScoped<IPinPadTerminalRepository, PinPadTerminalRepository>();
        var pinPadHabilitado = configuration.GetValue<bool>($"{PinPadOptions.SectionName}:Habilitado");
        if (pinPadHabilitado)
            services.AddSingleton<IPinPadService, PinPadService>();
        else
            services.AddSingleton<IPinPadService, NullPinPadService>();

        // CashDro — BR-PERIPH-002, ADR-011
        services.Configure<CashDroOptions>(configuration.GetSection(CashDroOptions.SectionName));
        var cashDroHabilitado = configuration.GetValue<bool>($"{CashDroOptions.SectionName}:Habilitado");
        if (cashDroHabilitado)
            services.AddSingleton<ICashDroService, CashDroService>();
        else
            services.AddSingleton<ICashDroService, NullCashDroService>();

        // BlueVision — BR-PERIPH-003, ADR-012
        services.Configure<BlueVisionOptions>(configuration.GetSection(BlueVisionOptions.SectionName));
        var blueVisionHabilitado = configuration.GetValue<bool>($"{BlueVisionOptions.SectionName}:Habilitado");
        if (blueVisionHabilitado)
        {
            services.AddHttpClient<IBlueVisionService, BlueVisionHttpClient>(client =>
            {
                var url = configuration.GetValue<string>($"{BlueVisionOptions.SectionName}:Url") ?? "http://localhost/";
                client.BaseAddress = new Uri(url);
            });
        }
        else
        {
            services.AddSingleton<IBlueVisionService, NullBlueVisionService>();
        }

        // Biometría SecuGen — BR-PERIPH-004, ADR-009 (GAP: OCX 32-bit sin SDK .NET)
        services.Configure<BiometriaOptions>(configuration.GetSection(BiometriaOptions.SectionName));
        services.AddSingleton<IBiometriaService, NullBiometriaService>();

        // Impresora Fiscal Epson — BR-PERIPH-005, ADR-009 (GAP: OCX 32-bit sin SDK .NET)
        services.Configure<ImpresoraFiscalOptions>(configuration.GetSection(ImpresoraFiscalOptions.SectionName));
        services.AddSingleton<IImpresoraFiscalService, NullImpresoraFiscalService>();
    }

    private static void AddCountryServices(this IServiceCollection services, IConfiguration configuration)
    {
        // FE Options — BR-PERIPH-006, ADR-010
        services.Configure<PeruFEOptions>(configuration.GetSection(PeruFEOptions.SectionName));
        services.Configure<ArgentinaFEOptions>(configuration.GetSection(ArgentinaFEOptions.SectionName));

        // Gateways de FE por país (todos registrados para resolución por fábrica)
        services.AddSingleton<IFacturacionElectronicaGateway, NullFEGateway>();
        var feHabilitado = configuration.GetValue<bool>("FE:Habilitado");
        if (feHabilitado)
        {
            services.AddSingleton<IFacturacionElectronicaGateway, PeruFEGateway>();
            services.AddSingleton<IFacturacionElectronicaGateway, ArgentinaFEGateway>();
            services.AddSingleton<IFacturacionElectronicaGateway, ChileFEGateway>();
            services.AddSingleton<IFacturacionElectronicaGateway, EcuadorFEGateway>();
            services.AddSingleton<IFacturacionElectronicaGateway, BoliviaFEGateway>();
        }
        services.AddSingleton<FacturacionElectronicaFactory>();

        // Políticas de país — BR-PERIPH-007, ADR-008
        services.AddSingleton<IPaisPolicy, PeruPaisPolicy>();
        services.AddSingleton<IPaisPolicy, ArgentinaPaisPolicy>();
        services.AddSingleton<IPaisPolicy, ChilePaisPolicy>();
        services.AddSingleton<IPaisPolicy, EcuadorPaisPolicy>();
        services.AddSingleton<IPaisPolicy, BoliviaPaisPolicy>();
        services.AddSingleton<IPaisPolicy, EspanaPaisPolicy>();
        services.AddSingleton<PaisPolicyFactory>();
    }
}
