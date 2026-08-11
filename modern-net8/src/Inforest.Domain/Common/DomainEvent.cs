namespace Inforest.Domain.Common;

/// <summary>
/// Evento de dominio base para INFOREST.
/// </summary>
public abstract record DomainEvent(DateTime OccurredAt)
{
    protected DomainEvent() : this(DateTime.UtcNow) { }
}
