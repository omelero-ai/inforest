IF not EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'INFSEGURIDAD')
BEGIN
CREATE DATABASE INFSEGURIDAD
end
Go

USE INFSEGURIDAD
GO

IF NOT EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'THARDKEY')
CREATE TABLE [dbo].[THARDKEY] (
	[id] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tLicencia] [nvarchar] (1000) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fRegistro] [datetime] NULL ,
	[tCliente] [nvarchar] (500) NULL  
) ON [PRIMARY]
GO

IF NOT EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'TPARAMETRO')
CREATE TABLE [dbo].[TPARAMETRO] (
	[tIdentificacionTributaria] [nvarchar] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[tRazonSocial] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tRazonComercial] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tDireccion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tTelefono] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tEmail] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tWebPage] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMonedaN] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMonN] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMonedaE] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[tMonE] [nvarchar] (3) COLLATE Modern_Spanish_CI_AS NULL ,
	[nTiempo] [smallint] NULL ,
	[nChkTiempo] [smallint] NULL ,
	[Impuesto1] [float] NULL ,
	[Impuesto2] [float] NULL ,
	[Impuesto3] [float] NULL ,
	[tImpuesto1] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tImpuesto2] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[tImpuesto3] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[lLongitud] [bit] NULL ,
	[nLongitud] [int] NULL ,
	[lPrinter] [bit] NULL ,
	[lRapido] [bit] NULL ,
	[tBoton1] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBoton2] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBoton3] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL , 
	[tBoton4] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tBoton5] [nvarchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[tPiePreCuenta] [nvarchar] (255) COLLATE Modern_Spanish_CI_AS NULL ,
	[lInforest] [bit] NULL ,
	[lAlmacen] [bit] NULL ,
	[lInfhotel] [bit] NULL , 
	[tClub] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[nPunto] [float] NULL ,
	[tVersion] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[lDiaContableAutomatico] [bit] NULL ,
	[tHoraCierreDiaContable] [nvarchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[lDiaContableManual] [bit] NULL ,
	[lClub] [bit] NULL 
) ON [PRIMARY]
GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'MMODULO')
CREATE TABLE [dbo].[MMODULO](
	[tModulo] [nvarchar](2) NOT NULL,
	[tDetallado] [nvarchar](200) NULL,
	[tResumido] [nvarchar](50) NULL,
 CONSTRAINT [PK_MMODULO] PRIMARY KEY CLUSTERED 
(
	[tModulo] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

go
IF NOT EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'MMOVIMIENTO')
CREATE   TABLE [dbo].[MMOVIMIENTO](
	[nCorrelativo] [bigint] NOT NULL,
	[nCorrelativoCampo] [bigint] NOT NULL,
	[tModulo] [nvarchar](2) NULL,
	[tBaseDatos] [nvarchar](150) NULL,
	[tTabla] [nvarchar](20) NULL,
	[tAccion] [nvarchar](50) NULL,
	[ffecha] [datetime] NULL,
	[tUsuario] [nvarchar](50) NULL,
	[tId_RegistroUnico] [nvarchar](20) NULL,
	[tId_RegistroDetalle] [nvarchar](20) NULL,
	[tCampo] [nvarchar](70) NULL,
	[tValorAnterior] [nvarchar](100) NULL,
	[tValorActual] [nvarchar](100) NULL,
 CONSTRAINT [PK_MMOVIMIENTO] PRIMARY KEY CLUSTERED 
(
	[nCorrelativo] ASC,
	[nCorrelativoCampo] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

IF NOT EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'MMOVIMIENTOACCESO')
CREATE TABLE [dbo].[MMOVIMIENTOACCESO] (
	[tBaseDatos] [nvarchar] (150) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tModulo] [nvarchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[tUsuario] [nvarchar] (100) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[nCorrelativo] [bigint] NOT NULL ,
	[fIngreso] [datetime] NULL ,
	[fSalida] [datetime] NULL 
) ON [PRIMARY]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'MACCION')
CREATE TABLE [dbo].[MACCION](
	[tAccion] [nvarchar](2) NULL,
	[tDescripcion] [nvarchar](50) NULL
) ON [PRIMARY]

GO


IF NOT EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'NOTICIAS')
CREATE TABLE NOTICIAS(
ID_NOT INT IDENTITY(1,1) PRIMARY KEY,
APLICACION VARCHAR(30) DEFAULT NULL,
URI_IMG1 VARCHAR(200) DEFAULT NULL,
URI_IMG2 VARCHAR(200) DEFAULT NULL,
URI_IMG3 VARCHAR(200) DEFAULT NULL,
FECHA_REGISTRO DATETIME DEFAULT NULL
)
GO


IF NOT EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'TPARAMETROVERSION')
CREATE TABLE [dbo].[TPARAMETROVERSION](
	[VersionInfoRest] [nvarchar](50) NULL,
	[VersionInfoHotel] [nvarchar](50) NULL,
	[VersionInfoBack] [nvarchar](50) NULL,
	[VersionFeInfoRest] [nvarchar](50) NULL,
	[VersionFeInfoHotel] [nvarchar](50) NULL,
	[VersionActualFEInfoRest] [nvarchar](50) NULL,
	[VersionActualFEInfoHotel] [nvarchar](50) NULL
) ON [PRIMARY]
GO


ALTER TABLE THARDKEY ALTER COLUMN TLICENCIA nvarchar(1000)
GO

ALTER TABLE MMOVIMIENTO ALTER COLUMN tValorAnterior nvarchar(250)
GO

ALTER TABLE MMOVIMIENTO ALTER COLUMN tValorActual nvarchar(250)
GO

ALTER TABLE [DBO].[MMOVIMIENTOACCESO] DROP CONSTRAINT [PK_MMOVIMIENTOACCESO]
GO
ALTER TABLE [dbo].[MMOVIMIENTOACCESO] WITH NOCHECK ADD 
	CONSTRAINT [PK_MMOVIMIENTOACCESO] PRIMARY KEY  CLUSTERED 
	(
		[tBaseDatos],
		[tModulo],
		[tUsuario],
		[nCorrelativo]
	)  ON [PRIMARY] 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ups_Aud_RegistraIngSalUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ups_Aud_RegistraIngSalUsuario]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ups_Aud_RegistraMovimientos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ups_Aud_RegistraMovimientos]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarMModulo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarMModulo]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarMMOVIMIENTO]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarMMOVIMIENTO]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarMMOVIMIENTOREP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarMMOVIMIENTOREP]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarMMOVIMIENTOACCESO]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarMMOVIMIENTOACCESO]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarMMOVIMIENTOACCESOREP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarMMOVIMIENTOACCESOREP]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarTodosAccion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarTodosAccion]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarTodosBaseDatos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarTodosBaseDatos]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarTodosMModulo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarTodosMModulo]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarTodosTabla]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarTodosTabla]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_Aud_SeleccionarTodosUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_Aud_SeleccionarTodosUsuario]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_AutSol_AuditoriaTipoDocumentoImpresora]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_AutSol_AuditoriaTipoDocumentoImpresora]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_INSERTAR_LICENCIAS]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_INSERTAR_LICENCIAS]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_CONSULTAR_LICENCIAS]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_CONSULTAR_LICENCIAS]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_CONSULTAR_LICENCIA_MAS_RECIENTE]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_CONSULTAR_LICENCIA_MAS_RECIENTE]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_ACTUALIZAR_LICENCIA]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_ACTUALIZAR_LICENCIA]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_INSERTAR_NOTICIAS]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_INSERTAR_NOTICIAS]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_CONSULTAR_NOTICIAS]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_CONSULTAR_NOTICIAS]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_CONSULTAR_NOTICIA_MAS_RECIENTE]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_CONSULTAR_NOTICIA_MAS_RECIENTE]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_ACTUALIZAR_NOTICIAS]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_ACTUALIZAR_NOTICIAS]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_CONSULTAR_ESTADO]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_CONSULTAR_ESTADO]
GO

CREATE PROCEDURE SP_CONSULTAR_ESTADO
AS
BEGIN

	DROP TABLE who2

	CREATE TABLE who2(
	  SPID INT,Status VARCHAR(255),
      Login  VARCHAR(255),
	  HostName  VARCHAR(255), 
      BlkBy  VARCHAR(255),
	  DBName  VARCHAR(255), 
      Command VARCHAR(255),
	  CPUTime INT, 
      DiskIO INT,
	  LastBatch VARCHAR(255), 
      ProgramName VARCHAR(255),
	  SPID2 INT, 
      REQUESTID INT
	  ) 
	  
	INSERT INTO who2 EXEC sp_who2

	DECLARE @INFHOTEL INT = (SELECT COUNT(*) FROM who2 WHERE ProgramName = 'infhotel')
	DECLARE @INFOREST INT = (SELECT COUNT(*) FROM who2 WHERE ProgramName = 'inforest')
	DECLARE @ALMACEN INT = (SELECT COUNT(*) FROM who2 WHERE ProgramName = 'almacen')

	SELECT @INFHOTEL AS INFHOTEL, @INFOREST AS INFOREST, @ALMACEN AS ALMACEN
END
GO


CREATE PROCEDURE SP_ACTUALIZAR_NOTICIAS(
@APLICACION VARCHAR(30), 
@URI_IMG1 VARCHAR(200), 
@URI_IMG2 VARCHAR(200), 
@URI_IMG3 VARCHAR(200), 
@FECHA DATETIME
)
AS
BEGIN
	UPDATE NOTICIAS SET URI_IMG1 = @URI_IMG1, URI_IMG2 = @URI_IMG2, URI_IMG3 = @URI_IMG3, FECHA_REGISTRO = @FECHA where APLICACION = @APLICACION
END

GO
CREATE PROCEDURE SP_CONSULTAR_NOTICIA_MAS_RECIENTE(
@FECHA DATETIME, 
@APLICACION VARCHAR(30)
)
AS
BEGIN
	SELECT COUNT(*) FROM NOTICIAS WHERE FECHA_REGISTRO >= @FECHA AND APLICACION = @APLICACION --MUESTRAME LAS NOTICIAS MAS ANTIGUAS QUE LA DE LA NUBE
END
GO

CREATE PROCEDURE SP_CONSULTAR_NOTICIAS(
@APLICACION VARCHAR(30)
)
AS
BEGIN
	SELECT COUNT(*) FROM NOTICIAS where APLICACION = @APLICACION
END
GO

CREATE PROCEDURE SP_INSERTAR_NOTICIAS(
@APLICACION VARCHAR(30), 
@URI_IMG1 VARCHAR(200), 
@URI_IMG2 VARCHAR(200), 
@URI_IMG3 VARCHAR(200), 
@FECHA DATETIME
)
AS
BEGIN
	INSERT NOTICIAS(APLICACION, URI_IMG1, URI_IMG2, URI_IMG3, FECHA_REGISTRO) VALUES(@APLICACION,@URI_IMG1,@URI_IMG2,@URI_IMG3,@FECHA)
END
GO

CREATE PROCEDURE SP_ACTUALIZAR_LICENCIA(
@RUC NVARCHAR(500), 
@LICENCIA NVARCHAR(1000), 
@FECHA DATETIME
)
AS
BEGIN
	UPDATE THARDKEY SET tLicencia = @LICENCIA, fRegistro = @FECHA where id = @RUC
END
GO

CREATE PROCEDURE SP_CONSULTAR_LICENCIA_MAS_RECIENTE(
@FECHA DATETIME
)
AS
BEGIN
	SELECT * FROM THARDKEY WHERE fRegistro < @FECHA --MUESTRAME LA LICENCIA DONDE LA FECHA SEA MENOR A LA DE LA NUBE
END
GO

CREATE PROCEDURE SP_CONSULTAR_LICENCIAS(
@RUC NVARCHAR(500)
)
as
BEGIN
	SELECT * FROM THARDKEY WHERE id = @RUC
END
GO

CREATE PROCEDURE SP_INSERTAR_LICENCIAS(
@RUC NVARCHAR(500), 
@LICENCIA NVARCHAR(1000), 
@FEC_REG DATETIME,@CLIENTE nvarchar(100)
)
AS
begin
	INSERT INTO THARDKEY(id,tLicencia,fRegistro,tCliente) VALUES (@RUC,@LICENCIA,@FEC_REG ,@CLIENTE)
end
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE proc ups_Aud_RegistraIngSalUsuario
@tTipo as nvarchar(1),
@tBaseDatos as nvarchar(150),
@tModulo as nvarchar(2),
@tUsuario as nvarchar(100),
@nCorrelativoUsuario as bigint,
@nCorrelativo as bigint output
as
begin
DECLARE @nSecuencia bigint

if UPPER(@tTipo)='I'
BEGIN
set @nSecuencia=(select isnull(MAX(ncorrelativo),0)  from mmovimientoacceso where 
	tBaseDatos=@tBaseDatos and
	tModulo=@tModulo and
	tUsuario=@tUsuario 
) + 1
set @nCorrelativo=@nSecuencia
insert into mmovimientoacceso
values(upper(@tBaseDatos),@tModulo, upper(@tUsuario),@nCorrelativo,GETDATE(),null)

END



IF UPPER(@tTipo)='S'
BEGIN

	
update mmovimientoacceso
set fsalida=GETDATE()
where 	tBaseDatos=@tBaseDatos and
	tModulo=@tModulo and
	tUsuario=@tUsuario and
	nCorrelativo= @nCorrelativoUsuario
	
set @nCorrelativo=8
end
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE [dbo].[ups_Aud_RegistraMovimientos]
@vch_nCorrelativo int,
@vch_nCorrelativoCampo int,
@vch_tModulo nvarchar(2),
@vch_tBaseDatos nvarchar(20),
@vch_tTabla nvarchar(20),
@vch_tAccion nvarchar(50),
@vch_tUsuario nvarchar(50),
@vch_tId_RegistroUnico nvarchar(20),
@vch_tId_RegistroDetalle nvarchar(20),
@vch_tCampo nvarchar(70),
@vch_tValorAnterior nvarchar(100),
@vch_tValorActual nvarchar(100)

AS
SET NOCOUNT ON
BEGIN
      DECLARE @vchMsgError varchar(100)
      BEGIN TRANSACTION

		INSERT INTO MMOVIMIENTO (	  nCorrelativo,		nCorrelativoCampo,	   tModulo,     tBaseDatos,	    tTabla,	    tAccion,   ffecha,	   tUsuario,	 tId_RegistroUnico,	    tId_RegistroDetalle,	 tCampo,	 tValorAnterior,	 tValorActual)
						VALUES  (@vch_nCorrelativo,@vch_nCorrelativoCampo,@vch_tModulo,@vch_tBaseDatos,@vch_tTabla,@vch_tAccion,GETDATE(),@vch_tUsuario,@vch_tId_RegistroUnico,@vch_tId_RegistroDetalle,@vch_tCampo,@vch_tValorAnterior,@vch_tValorActual)
	      IF @@ERROR <> 0
	      BEGIN
                    ROLLBACK TRANSACTION
	            SET @vchMsgError = 'Error al Actualizar datos'
	            RAISERROR(@vchMsgError,16,1)
	            RETURN(-1)
	      END
      COMMIT TRANSACTION 
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


create PROCEDURE [dbo].[usp_Aud_SeleccionarMModulo] 
@Var_tModulo as nvarchar(2)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT tModulo, tDetallado, tResumido from dbo.MMODULO where tModulo=@Var_tModulo
	
END

GO
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[usp_Aud_SeleccionarMMOVIMIENTO]
	@fec_Inicio as datetime, 
	@fec_Final as datetime
AS

BEGIN
	SET NOCOUNT ON;
	

	SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
	FROM MMOVIMIENTO
	where fFecha between @fec_Inicio and @fec_Final

END

GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[usp_Aud_SeleccionarMMOVIMIENTOREP]
	@vch_tBaseDatos as varchar(150), 
	@vch_tModulo as varchar(2),
	@vch_tTabla as varchar(20),
	@vch_tAccion as varchar(50),
	@fec_Inicio as datetime, 
	@fec_Final as datetime,
	@vch_tUsuario as varchar(100)
AS

BEGIN
	SET NOCOUNT ON;
	
	if @vch_tBaseDatos <>''
		begin
			if @vch_tModulo <>''
				begin
					if @vch_tUsuario <>''
						begin
							if @vch_tTabla <> ''
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tTabla=@vch_tTabla and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin

													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tTabla=@vch_tTabla and 
													fFecha between @fec_Inicio and @fec_Final

										end
								end
							else
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
						end
					else
						begin
							if @vch_tTabla <> ''
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tTabla=@vch_tTabla and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tTabla=@vch_tTabla and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
							else
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
						end
				end
			else
				begin
					if @vch_tUsuario <>''
						begin
							if @vch_tTabla <> ''
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tTabla=@vch_tTabla and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final

										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tTabla=@vch_tTabla and 
													fFecha between @fec_Inicio and @fec_Final

										end
								end
							else
								begin
									if @vch_tAccion <> ''
										begin

													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
						end
					else
						begin
							if @vch_tTabla <> ''
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tTabla=@vch_tTabla and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tTabla=@vch_tTabla and 
													fFecha between @fec_Inicio and @fec_Final

										end
								end
							else
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final

										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tBaseDatos=@vch_tBaseDatos and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
						end
				end
		end
	else
		begin
			if @vch_tModulo <>''
				begin
					if @vch_tUsuario <>''
						begin
							if @vch_tTabla <> ''
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tTabla=@vch_tTabla and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tTabla=@vch_tTabla and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
							else
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
						end
					else
						begin
							if @vch_tTabla <> ''
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tTabla=@vch_tTabla and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tTabla=@vch_tTabla and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
							else
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
						end
				end
			else
				begin
					if @vch_tUsuario <>''
						begin
							if @vch_tTabla <> ''
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tTabla=@vch_tTabla and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tTabla=@vch_tTabla and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
							else
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tModulo=@vch_tModulo and 
													tUsuario=@vch_tUsuario and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
						end
					else
						begin
							if @vch_tTabla <> ''
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tTabla=@vch_tTabla and 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tTabla=@vch_tTabla and 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
							else
								begin
									if @vch_tAccion <> ''
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													tAccion=@vch_tAccion and 
													fFecha between @fec_Inicio and @fec_Final
										end
									else
										begin
													SELECT nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual
													FROM MMOVIMIENTO
													where 
													fFecha between @fec_Inicio and @fec_Final
										end
								end
						end
				end

		end
	

END

GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[usp_Aud_SeleccionarMMOVIMIENTOACCESO]
	@vch_tTipoFecha as varchar(1),
	@fec_Inicio as datetime, 
	@fec_Final as datetime 
AS

BEGIN

	SET NOCOUNT ON;
	if @vch_tTipoFecha='I'
		begin
			SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
			FROM MMOVIMIENTOACCESO  
			where fIngreso between @fec_Inicio and @fec_Final
		end
	else
		begin
			if @vch_tTipoFecha='S'
				begin
					SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
					FROM MMOVIMIENTOACCESO  
					where fSalida between @fec_Inicio and @fec_Final
				end
			else
				begin
					SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
					FROM MMOVIMIENTOACCESO  
					where ((fIngreso between @fec_Inicio and @fec_Final) or (fSalida between @fec_Inicio and @fec_Final))
				end
		end

END

GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


create PROCEDURE [dbo].[usp_Aud_SeleccionarMMOVIMIENTOACCESOREP]
	@vch_tBaseDatos as varchar(150), 
	@vch_tModulo as varchar(2),
	@vch_tUsuario as varchar(100), 
	@vch_tTipoFecha as varchar(1),
	@fec_Inicio as datetime, 
	@fec_Final as datetime 
AS

BEGIN

	SET NOCOUNT ON;
	if @vch_tBaseDatos <>''
		begin
			if @vch_tModulo <>''
				begin
					if @vch_tUsuario <>''
						begin
							if @vch_tTipoFecha='I'
								begin
									SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
									FROM MMOVIMIENTOACCESO  
									where tBaseDatos=@vch_tBaseDatos and MMOVIMIENTOACCESO.tModulo=@vch_tModulo and tUsuario=@vch_tUsuario and fIngreso between @fec_Inicio and @fec_Final
								end
							else
								begin
									if @vch_tTipoFecha='S'
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where tBaseDatos=@vch_tBaseDatos and MMOVIMIENTOACCESO.tModulo=@vch_tModulo and tUsuario=@vch_tUsuario and fSalida between @fec_Inicio and @fec_Final
										end
									else
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where tBaseDatos=@vch_tBaseDatos and MMOVIMIENTOACCESO.tModulo=@vch_tModulo and tUsuario=@vch_tUsuario and ((fIngreso between @fec_Inicio and @fec_Final) or (fSalida between @fec_Inicio and @fec_Final))
										end
								end
						end
					else
						begin
							if @vch_tTipoFecha='I'
								begin
									SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
									FROM MMOVIMIENTOACCESO  
									where tBaseDatos=@vch_tBaseDatos and MMOVIMIENTOACCESO.tModulo=@vch_tModulo and fIngreso between @fec_Inicio and @fec_Final
								end
							else
								begin
									if @vch_tTipoFecha='S'
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO   
											where tBaseDatos=@vch_tBaseDatos and MMOVIMIENTOACCESO.tModulo=@vch_tModulo and fSalida between @fec_Inicio and @fec_Final
										end
									else
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO   
											where tBaseDatos=@vch_tBaseDatos and MMOVIMIENTOACCESO.tModulo=@vch_tModulo  and ((fIngreso between @fec_Inicio and @fec_Final) or (fSalida between @fec_Inicio and @fec_Final))
										end
								end
						end
				end
			else
				begin
					if @vch_tUsuario <>''
						begin
							
							if @vch_tTipoFecha='I'
								begin
									SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
									FROM MMOVIMIENTOACCESO  
									where tBaseDatos=@vch_tBaseDatos and tUsuario=@vch_tUsuario and fIngreso between @fec_Inicio and @fec_Final
								end
							else
								begin
									if @vch_tTipoFecha='S'
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where tBaseDatos=@vch_tBaseDatos and tUsuario=@vch_tUsuario and fSalida between @fec_Inicio and @fec_Final
										end
									else
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where tBaseDatos=@vch_tBaseDatos and tUsuario=@vch_tUsuario  and ((fIngreso between @fec_Inicio and @fec_Final) or (fSalida between @fec_Inicio and @fec_Final))
										end
								end
						end
					else
						begin
							
							if @vch_tTipoFecha='I'
								begin
									SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
									FROM MMOVIMIENTOACCESO  
									where tBaseDatos=@vch_tBaseDatos and fIngreso between @fec_Inicio and @fec_Final
								end
							else
								begin
									if @vch_tTipoFecha='S'
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where tBaseDatos=@vch_tBaseDatos and fSalida between @fec_Inicio and @fec_Final
										end
									else
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where tBaseDatos=@vch_tBaseDatos and ((fIngreso between @fec_Inicio and @fec_Final) or (fSalida between @fec_Inicio and @fec_Final))
										end
								end
						end
				end
		end
	else
		begin
			if @vch_tModulo <>''
				begin
					if @vch_tUsuario <>''
						begin
							
							if @vch_tTipoFecha='I'
								begin
									SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
									FROM MMOVIMIENTOACCESO  
									where MMOVIMIENTOACCESO.tModulo=@vch_tModulo and tUsuario=@vch_tUsuario and fIngreso between @fec_Inicio and @fec_Final
								end
							else
								begin
									if @vch_tTipoFecha='S'
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where MMOVIMIENTOACCESO.tModulo=@vch_tModulo and tUsuario=@vch_tUsuario and fSalida between @fec_Inicio and @fec_Final
										end
									else
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where MMOVIMIENTOACCESO.tModulo=@vch_tModulo and tUsuario=@vch_tUsuario and ((fIngreso between @fec_Inicio and @fec_Final) or (fSalida between @fec_Inicio and @fec_Final))
										end
								end
						end
					else
						begin
							
							if @vch_tTipoFecha='I'
								begin
									SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
									FROM MMOVIMIENTOACCESO  
									where MMOVIMIENTOACCESO.tModulo=@vch_tModulo and fIngreso between @fec_Inicio and @fec_Final
								end
							else
								begin
									if @vch_tTipoFecha='S'
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where MMOVIMIENTOACCESO.tModulo=@vch_tModulo and fSalida between @fec_Inicio and @fec_Final
										end
									else
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where MMOVIMIENTOACCESO.tModulo=@vch_tModulo and ((fIngreso between @fec_Inicio and @fec_Final) or (fSalida between @fec_Inicio and @fec_Final))
										end
								end
						end
				end
			else
				begin
					if @vch_tUsuario <>''
						begin
							
							if @vch_tTipoFecha='I'
								begin
									SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
									FROM MMOVIMIENTOACCESO  
									where tUsuario=@vch_tUsuario and fIngreso between @fec_Inicio and @fec_Final
								end
							else
								begin
									if @vch_tTipoFecha='S'
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where tUsuario=@vch_tUsuario and fSalida between @fec_Inicio and @fec_Final
										end
									else
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where tUsuario=@vch_tUsuario and ((fIngreso between @fec_Inicio and @fec_Final) or (fSalida between @fec_Inicio and @fec_Final))
										end
								end
						end
					else
						begin
							if @vch_tTipoFecha='I'
								begin
									SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
									FROM MMOVIMIENTOACCESO  
									where fIngreso between @fec_Inicio and @fec_Final
								end
							else
								begin
									if @vch_tTipoFecha='S'
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where fSalida between @fec_Inicio and @fec_Final
										end
									else
										begin
											SELECT tBaseDatos, tmodulo , tUsuario, nCorrelativo, fIngreso, fSalida 
											FROM MMOVIMIENTOACCESO  
											where ((fIngreso between @fec_Inicio and @fec_Final) or (fSalida between @fec_Inicio and @fec_Final))
										end
								end
						end
				end
		end
END

GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[usp_Aud_SeleccionarTodosAccion] 

AS
BEGIN

	SET NOCOUNT ON;

	SELECT tAccion,tDescripcion from MAccion
	
END

GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[usp_Aud_SeleccionarTodosBaseDatos] 

AS
BEGIN

	SET NOCOUNT ON;

	SELECT tBaseDatos from vBaseDatos
	
END

GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[usp_Aud_SeleccionarTodosMModulo] 
AS
BEGIN

	SET NOCOUNT ON;

	SELECT tModulo, tDetallado, tResumido from dbo.MMODULO
	
END

GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


create PROCEDURE [dbo].[usp_Aud_SeleccionarTodosTabla] 

AS
BEGIN

	SET NOCOUNT ON;

	SELECT tTabla from vTabla
	
END

GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[usp_Aud_SeleccionarTodosUsuario] 
AS
BEGIN

	SET NOCOUNT ON;
	SELECT tUsuario from vUsuario
	
END
GO

CREATE PROCEDURE [usp_AutSol_AuditoriaTipoDocumentoImpresora]
@nCorrelativoCampo bigint,
@tModulo nvarchar(2),
@tBaseDatos nvarchar(250),
@tTabla nvarchar(250),
@tAccion nvarchar(2),
@tUsuario nvarchar(250),
@tId_RegistroUnico nvarchar(250),
@tId_RegistroDetalle nvarchar(250),
@tCampo nvarchar(250),
@tValorAnterior nvarchar(250),
@tValorActual nvarchar(250),
@nCorrelativoCampo1 bigint,
@tCampo1 nvarchar(250),
@tValorAnterior1 nvarchar(250),
@tValorActual1 nvarchar(250),
@nCorrelativoCampo2 bigint,
@tCampo2 nvarchar(250),
@tValorAnterior2 nvarchar(250),
@tValorActual2 nvarchar(250),
@nCorrelativoCampo3 bigint,
@tCampo3 nvarchar(250),
@tValorAnterior3 nvarchar(250),
@tValorActual3 nvarchar(250)
as
begin

declare @correlativo bigint
declare @campo nvarchar(500)

set @correlativo =(select isnull(MAX(nCorrelativo),0)+1 from mmovimiento )


--
 insert into MMOVIMIENTO (nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual)
 values (@correlativo,@nCorrelativoCampo,@tModulo,@tBaseDatos,@tTabla,@tAccion,GETDATE(),@tUsuario,@tId_RegistroUnico,@tId_RegistroDetalle,@tCampo,@tValorAnterior,@tValorActual)
 
 insert into MMOVIMIENTO (nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual)
 values (@correlativo,@nCorrelativoCampo1,@tModulo,@tBaseDatos,@tTabla,@tAccion,GETDATE(),@tUsuario,@tId_RegistroUnico,@tId_RegistroDetalle,@tCampo1,@tValorAnterior1,@tValorActual1)
 
 insert into MMOVIMIENTO (nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual)
 values (@correlativo,@nCorrelativoCampo2,@tModulo,@tBaseDatos,@tTabla,@tAccion,GETDATE(),@tUsuario,@tId_RegistroUnico,@tId_RegistroDetalle,@tCampo2,@tValorAnterior2,@tValorActual2)
 
 insert into MMOVIMIENTO (nCorrelativo,nCorrelativoCampo,tModulo,tBaseDatos,tTabla,tAccion,ffecha,tUsuario,tId_RegistroUnico,tId_RegistroDetalle,tCampo,tValorAnterior,tValorActual)
 values (@correlativo,@nCorrelativoCampo3,@tModulo,@tBaseDatos,@tTabla,@tAccion,GETDATE(),@tUsuario,@tId_RegistroUnico,@tId_RegistroDetalle,@tCampo3,@tValorAnterior3,@tValorActual3)

end 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vBaseDatos]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vBaseDatos]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTabla]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTabla]
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vUsuario]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vUsuario]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vBaseDatos]
AS
SELECT DISTINCT tBaseDatos
FROM         dbo.MMOVIMIENTOACCESO

GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vTabla]
AS
SELECT DISTINCT tTabla
FROM         dbo.MMOVIMIENTO

GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vUsuario]
AS
SELECT DISTINCT tUsuario
FROM         dbo.MMOVIMIENTOACCESO

GO

DELETE FROM MACCION

DELETE FROM MACCION
insert into MACCION values ('01','AGREGAR')
insert into MACCION values ('02','ACTUALIZAR')
insert into MACCION values ('03','ELIMINAR')

DELETE FROM MMODULO

insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('01','Inforest','Inforest')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('02','Adición','Adición')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('03','Chef Control','Chef Control')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('04','Despachador','Despachador')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('05','Anfitriona','Anfitriona')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('06','Reportes','Reportes')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('07','Almacen','Almacen')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('08','Costo','Costo')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('09','Infhotel','Infhotel')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('10','Eventos','Eventos')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('11','Promociones','Promociones')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('12','CentralDelivery','CentralDelivery')
insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('13','Administracion','Administracion')
Insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('14','Consultas','Consultas')
Insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('15','Auditoria','Auditoria')
Insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('16','Autorizaciones','Autorizaciones')
Insert into MMODULO ([tModulo],[tDetallado],[tResumido]) values ('17','Adm. Centralizada','Adm. Centralizada')



GO

IF (select COUNT(*) from TPARAMETROVERSION)=0 
BEGIN

INSERT INTO TPARAMETROVERSION (VersionInfoRest,	VersionInfoHotel,	VersionInfoBack,	VersionFeInfoRest,	VersionFeInfoHotel,	VersionActualFEInfoRest,	VersionActualFEInfoHotel)
values ('4.95.5','2.4.1','4.94.9011','6.0.0','6.0.0','6.0.0','6.0.0')

END

GO



PRINT ' VERIFICAR QUE EN LA BD INFSEGURIDAD, TABLA TPARAMETRO, EL CAMPO TIDENTIFICACIONTRIBUTARIA CONTENGA EL RUC DEL CLIENTE, Y EN EL CAMPO NLONGITUD
ESTE LA CANTIDAD DE DIGITOS QUE FORMAN ESTE CAMPO'







