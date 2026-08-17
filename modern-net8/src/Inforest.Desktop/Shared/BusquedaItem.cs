namespace Inforest.Desktop.Shared;

/// <summary>
/// Ítem genérico para el diálogo de búsqueda rápida.
/// Equivale a las columnas <c>Codigo</c> y <c>Descripcion</c> del Recordset de VB6.
/// </summary>
/// <param name="Codigo">Código identificador del ítem (campo <c>Codigo</c>).</param>
/// <param name="Descripcion">Texto descriptivo del ítem (campo <c>Descripcion</c>).</param>
/// <param name="Extra">Campo adicional opcional expuesto como columna extra en el grid.</param>
public sealed record BusquedaItem(string Codigo, string Descripcion, string Extra = "");

/// <summary>
/// Resultado de la búsqueda rápida.
/// </summary>
/// <param name="Codigo">Código del ítem seleccionado.</param>
/// <param name="Descripcion">Descripción del ítem seleccionado.</param>
public sealed record BusquedaResultado(string Codigo, string Descripcion);
