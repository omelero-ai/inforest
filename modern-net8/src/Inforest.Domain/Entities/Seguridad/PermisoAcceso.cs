namespace Inforest.Domain.Entities.Seguridad;

/// <summary>
/// Permiso explícito cargado desde TACCESO/TGRUPOACCESO.
/// Legacy: TACCESO.tCodigoAcceso, tModulo, tFormulario, tNombreObjeto.
/// Regla BR-006.
/// </summary>
public sealed record PermisoAcceso(
    string CodigoAcceso,
    string Modulo,
    string Formulario,
    string NombreObjeto,
    string Descripcion);
