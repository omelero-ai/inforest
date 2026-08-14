# Módulos — INFOREST

> Documentación de los módulos funcionales del sistema INFOREST.

---

## Módulos del Sistema

| Módulo | Ejecutable Legacy | Documentación | Estado |
|---|---|---|---|
| Restaurante (POS + Admin + Delivery) | Todos los ejecutables | [restaurante/](restaurante/README.md) | PARTIAL — análisis en curso |

---

## Estructura de Módulos Legacy

El sistema Legacy está organizado en módulos funcionales que corresponden a 7 ejecutables. Todos comparten la misma base de datos y componentes comunes.

```mermaid
graph TB
    subgraph POS["Punto de Venta"]
        IR[InfoRest.exe\nSalón, Delivery, Llevar]
        CR[CajaRapida.exe\nCaja Rápida / Fast Food]
        AD[Adicion.exe\nAdición de ítems]
    end

    subgraph Admin["Administración"]
        AM[Administracion.exe\nMaestros, Config, Reportes Admin]
    end

    subgraph Reports["Consultas"]
        CO[Consulta.exe\nReportes Gerenciales]
    end

    subgraph Delivery["Delivery"]
        DE[Despachador.exe\nCentral Pedidos / Despacho]
        MO[Motorizado.exe\nSeguimiento Entregas]
    end

    subgraph Shared["Componentes Compartidos"]
        DB[(SQL Server\nINFOREST)]
        SEC[(INFSEGURIDAD)]
        DEC[modDeclaracion\nEstado global]
        PROC[modProcedimiento\nUtilidades]
    end

    POS --> Shared
    Admin --> Shared
    Reports --> Shared
    Delivery --> Shared
```

---

## Notas

- Todos los módulos están documentados inicialmente en [restaurante/README.md](restaurante/README.md)
- A medida que avance la migración, cada módulo tendrá su propia subcarpeta
- La documentación detallada del Legacy está en [legacy-restaurant/README.md](../../legacy-restaurant/README.md)
