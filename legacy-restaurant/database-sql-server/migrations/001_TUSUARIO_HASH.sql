-- ============================================================
-- Migración 001: Tabla sidecar TUSUARIO_HASH
-- Base de datos: INFSEGURIDAD
-- Propósito: Almacenar hashes BCrypt para usuarios sin modificar
--            TUSUARIO.tPassword (varchar(8)) que no admite BCrypt.
-- Origen Legacy: TUSUARIO.tPassword varchar(8) — SEC-006
-- ADR: ADR-013 (ver docs/architecture/architecture-decisions.md)
-- ============================================================

USE [INFSEGURIDAD]
GO

IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'TUSUARIO_HASH'
)
BEGIN
    CREATE TABLE [dbo].[TUSUARIO_HASH] (
        [tCodigoUsuario]  VARCHAR(15)   NOT NULL,
        [tHashBCrypt]     VARCHAR(72)   NOT NULL,
        [fActualizacion]  DATETIME      NOT NULL CONSTRAINT [DF_TUSUARIO_HASH_fActualizacion] DEFAULT (GETDATE()),
        CONSTRAINT [PK_TUSUARIO_HASH] PRIMARY KEY CLUSTERED ([tCodigoUsuario] ASC)
    )

    -- FK opcional — comentar si TUSUARIO no está en INFSEGURIDAD
    -- ALTER TABLE [dbo].[TUSUARIO_HASH]
    --     ADD CONSTRAINT [FK_TUSUARIO_HASH_TUSUARIO]
    --     FOREIGN KEY ([tCodigoUsuario]) REFERENCES [dbo].[TUSUARIO] ([tCodigoUsuario])
    --     ON DELETE CASCADE;

    PRINT 'Tabla TUSUARIO_HASH creada exitosamente en INFSEGURIDAD.'
END
ELSE
BEGIN
    PRINT 'Tabla TUSUARIO_HASH ya existe. No se realizaron cambios.'
END
GO
