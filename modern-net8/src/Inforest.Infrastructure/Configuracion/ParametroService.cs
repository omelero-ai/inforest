using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Infrastructure.Configuracion;

/// <summary>
/// Legacy: TPARAMETRO expone flags globales consumidos en todo el sistema. BR-SQL-003
/// </summary>
internal sealed class ParametroService : IParametroService
{
    private static readonly IReadOnlyDictionary<string, PropertyInfo> Propiedades =
        typeof(ConfiguracionSistema)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .ToDictionary(property => property.Name, StringComparer.OrdinalIgnoreCase);

    private readonly IParametroRepository _repository;
    private readonly SemaphoreSlim _cacheLock = new(1, 1);
    private ConfiguracionSistema? _cache;

    public ParametroService(IParametroRepository repository)
    {
        _repository = repository;
    }

    public async Task<string?> ObtenerAsync(string clave, CancellationToken cancellationToken = default)
    {
        var valor = await ObtenerValorAsync(clave, cancellationToken);
        return ConvertirAString(valor);
    }

    public async Task<decimal?> ObtenerNumericoAsync(string clave, CancellationToken cancellationToken = default)
    {
        var valor = await ObtenerValorAsync(clave, cancellationToken);
        return TryConvertToDecimal(valor, out var numero) ? numero : null;
    }

    public async Task<bool> ObtenerBoolAsync(string clave, bool valorDefecto = false, CancellationToken cancellationToken = default)
    {
        var valor = await ObtenerValorAsync(clave, cancellationToken);
        return TryConvertToBoolean(valor, out var booleano) ? booleano : valorDefecto;
    }

    public async Task<bool> TieneConfiguracionAsync(CancellationToken cancellationToken = default)
        => await ObtenerConfiguracionAsync(cancellationToken) is not null;

    public void InvalidarCache()
    {
        _cache = null;
    }

    private async Task<object?> ObtenerValorAsync(string clave, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(clave))
            return null;

        var configuracion = await ObtenerConfiguracionAsync(cancellationToken);
        if (configuracion is null || !Propiedades.TryGetValue(clave, out var property))
            return null;

        return property.GetValue(configuracion);
    }

    private async Task<ConfiguracionSistema?> ObtenerConfiguracionAsync(CancellationToken cancellationToken)
    {
        if (_cache is not null)
            return _cache;

        await _cacheLock.WaitAsync(cancellationToken);
        try
        {
            if (_cache is null)
                _cache = await _repository.ObtenerConfiguracionAsync(cancellationToken);

            return _cache;
        }
        finally
        {
            _cacheLock.Release();
        }
    }

    private static string? ConvertirAString(object? valor)
        => valor switch
        {
            null => null,
            bool booleano => booleano ? "1" : "0",
            DateTime fecha => fecha.ToString("O", CultureInfo.InvariantCulture),
            _ => Convert.ToString(valor, CultureInfo.InvariantCulture)
        };

    private static bool TryConvertToDecimal(object? valor, out decimal numero)
    {
        if (valor is null)
        {
            numero = default;
            return false;
        }

        if (valor is decimal decimalValue)
        {
            numero = decimalValue;
            return true;
        }

        if (valor is bool booleano)
        {
            numero = booleano ? 1m : 0m;
            return true;
        }

        if (valor is IConvertible convertible)
        {
            try
            {
                numero = convertible.ToDecimal(CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
            }
        }

        if (decimal.TryParse(Convert.ToString(valor, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out numero))
            return true;

        numero = default;
        return false;
    }

    private static bool TryConvertToBoolean(object? valor, out bool booleano)
    {
        switch (valor)
        {
            case bool boolValue:
                booleano = boolValue;
                return true;
            case null:
                booleano = default;
                return false;
            case string text:
                var normalized = text.Trim();
                if (string.IsNullOrEmpty(normalized))
                    break;

                if (normalized.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                    normalized.Equals("SI", StringComparison.OrdinalIgnoreCase) ||
                    normalized.Equals("S", StringComparison.OrdinalIgnoreCase) ||
                    normalized.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                    normalized.Equals("ON", StringComparison.OrdinalIgnoreCase) ||
                    normalized.Equals("YES", StringComparison.OrdinalIgnoreCase))
                {
                    booleano = true;
                    return true;
                }

                if (normalized.Equals("0", StringComparison.OrdinalIgnoreCase) ||
                    normalized.Equals("NO", StringComparison.OrdinalIgnoreCase) ||
                    normalized.Equals("N", StringComparison.OrdinalIgnoreCase) ||
                    normalized.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ||
                    normalized.Equals("OFF", StringComparison.OrdinalIgnoreCase))
                {
                    booleano = false;
                    return true;
                }
                break;
            case sbyte or byte or short or ushort or int or uint or long or ulong:
                booleano = Convert.ToInt64(valor, CultureInfo.InvariantCulture) != 0;
                return true;
            case float or double or decimal:
                booleano = Convert.ToDecimal(valor, CultureInfo.InvariantCulture) != 0m;
                return true;
        }

        booleano = default;
        return false;
    }
}
