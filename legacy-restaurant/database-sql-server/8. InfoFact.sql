--****************************************************************INFOFACT INICIO **********************************************
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vNotaCreditoCabeceraInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vNotaCreditoCabeceraInfoFact]
GO
CREATE VIEW vNotaCreditoCabeceraInfoFact
as
	select 
		MN.tNotaCredito AS tDocumento, 
		(case when ISNULL(TC.lCajaContingencia,0)=1 then '0' else  LEFT(MD.tdocumento, 1) end)  + LEFT(RIGHT(MN.tNotaCredito, 12), 3) + '-' + RIGHT(MN.tNotaCredito, 8) AS IdDocumento,
		isnull(vtp.Sunat,'00') AS TipoDocumento,
		'0101'  TipoOperacion,-- revisar retencion y operacion onerosa
		--CONVERT(NVARCHAR(10), MN.fRegistro, 120) AS FechaEmision,
		CASE
			WHEN ISNULL(TC.lCajaContingencia,0)=1 THEN ISNULL(CONVERT(NVARCHAR(10), MN.fFecha, 120), CONVERT(NVARCHAR(10), MN.fregistro, 120))
			ELSE CONVERT(NVARCHAR(10), MN.fregistro, 120)
		END AS FechaEmision,
		--CONVERT(NVARCHAR(10), MN.fRegistro, 108) AS HoraEmision,
		CASE
			WHEN ISNULL(TC.lCajaContingencia,0)=1 THEN ISNULL(CONVERT(NVARCHAR(10), MN.fFecha, 108), CONVERT(NVARCHAR(10), MN.fRegistro, 108))
			ELSE CONVERT(NVARCHAR(10), MN.fRegistro, 108)
		END AS HoraEmision,
		--CONVERT(NVARCHAR(10), MN.fRegistro, 120) AS FechaVencimiento, -- VERIFICAR AL CREDITO
		CONVERT(NVARCHAR(10),
			CASE
				WHEN ISNULL(TC.lCajaContingencia,0)=1 THEN ISNULL(MN.fFecha, MN.fregistro)
				ELSE  MN.fregistro
			END
			+case when MD.tEstadoDocumento ='03' then (case when isnull(vd.nDiasCredito,0) = 0 then 1 else isnull(vd.nDiasCredito,0)end) 
				else 0 end , 120) AS FechaVencimiento, -- VERIFICAR AL CREDITO
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then 'USD' else 'PEN' end AS Moneda,
		ltrim(rtrim(TP.tIdentificacionTributaria)) AS Emi_NroDocumento,
		'6' AS Emi_TipoDocumento,
		ltrim(rtrim(TP.tRazonSocial)) AS Emi_NombreLegal,
		ltrim(rtrim(TP.tRazonComercial)) AS Emi_NombreComercial,
		TP.tCodigoUbigeo AS Emi_Ubigeo,
		ltrim(rtrim(TP.tDireccion)) AS Emi_Direccion,
		ltrim(rtrim(TP.tDireccion2)) AS Emi_Urbanizacion,
		TP.tDepartamento AS Emi_Departamento,
		TP.tProvincia AS Emi_Provincia,
		TP.tDistrito AS Emi_Distrito,
		ltrim(rtrim(TP.tCodigoTienda)) AS Emi_CodigoAnexo,
		'PE' as Emi_Pais,
		case when isnull(TP.Correo,'')='' then 'facturacion.electronica@infhotel.com.pe' else isnull(TP.Correo,'') end  AS Emi_Correo,
		ltrim(rtrim(vc.tIdentidad)) AS Rec_NroDocumento,
		vti.tvalor AS Rec_TipoDocumento,
		case when len(ltrim(rtrim(vc.Descripcion)))=1 then 'Cliente' else ltrim(rtrim(vc.Descripcion)) end  AS Rec_NombreLegal,
		case when len(ltrim(rtrim(vc.Descripcion)))=1 then 'Cliente' else ltrim(rtrim(vc.Descripcion)) end  AS Rec_NombreComercial,
		vc.CodigoUbigeo AS Rec_Ubigeo,
		ltrim(rtrim(vc.tDireccion)) AS Rec_Direccion,
		ltrim(rtrim(vc.Urbanizacion)) AS Rec_Urbanizacion,
		TUB.tDepartamento AS Rec_Departamento,
		TUB.tProvincia AS Rec_Provincia,
		TUB.tDistrito AS Rec_Distrito,
		'' AS Rec_CodigoAnexo,
		'PE' as Rec_Pais,
		case when isnull(vc.tCorreo,'')='' then 'facturacion.electronica@infhotel.com.pe' else isnull(vc.tCorreo,'') end  AS Rec_Correo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Gravado/TT.nVenta as decimal(30,5)) else Cast(DC.Gravado as decimal(30,5)) end  AS Gravadas,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Gratuita/TT.nVenta as decimal(30,5)) else Cast(DC.Gratuita as decimal(30,5)) end  AS Gratuitas,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Inafecto/TT.nVenta as decimal(30,5)) else Cast(DC.Inafecto as decimal(30,5)) end  AS Inafectas,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Exonerado/TT.nVenta as decimal(30,5)) else Cast(DC.Exonerado as decimal(30,5)) end AS Exoneradas,
		0.00 AS Exportacion,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Neto/TT.nVenta as decimal(30,5)) else Cast(DC.Neto as decimal(30,5)) end AS LineExtensionAmount,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((DC.Neto +DC.nImpuesto1)/TT.nVenta as decimal(30,5)) else Cast(DC.Neto+DC.nImpuesto1 as decimal(30,5)) end AS TaxInclusiveAmount,
		0.00000  AS FactorMultiplicadorDscto,
		0.00 AS DescuentoGlobal,
		'' AS CodigoRazonDcto,
		0.00000  AS MontoBaseParaDcto,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nVenta/TT.nVenta as decimal(30,5)) else Cast(DC.nVenta as decimal(30,5)) end TotalVenta,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nImpuesto1/TT.nVenta as decimal(30,5)) else Cast(DC.nImpuesto1 as decimal(30,5)) end TotalIgv,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.IgvGratuita/TT.nVenta as decimal(30,5)) else Cast(DC.IgvGratuita as decimal(30,5)) end AS TotalIgvGratuita,
		0.00 AS TotalIsc,
		--case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nLey1/TT.nVenta as decimal(30,5)) else Cast(DC.nLey1 as decimal(30,5)) end AS  TotalOtrosTributos,
		Cast(DC.nLey1 as decimal(30,5))  AS  TotalOtrosTributos,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nImpuesto2/TT.nVenta as decimal(30,5)) else Cast(DC.nImpuesto2 as decimal(30,5)) end AS TotalRecargo,
		Cast(DC.nImpuesto2/(case when DC.Neto<>0 then DC.Neto else 1 end ) as decimal(30,5)) AS PorcentajeRecargo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Neto/TT.nVenta as decimal(30,5)) else Cast(DC.Neto as decimal(30,5)) end AS MontoBaseRecargo,
		0.00 AS MontoPercepcion,
		0.00 AS MontoDetraccion,
		'0.00' AS TasaDetraccion,
		'' AS CodigoMedioPago,
		'' AS CuentaBancoNacion,
		'' AS CodigoBienOServicio,
		'' AS TipoDocAnticipo,
		'' AS MonedaAnticipo,
		0.00 AS MontoAnticipo,
		case when ltrim(rtrim(isnull(MN.tObservacion,''))) ='' then 'ANULACION DE LA OPERACION - Documento Referencia: ' + LEFT(MD.tdocumento, 1) + LEFT(RIGHT(MD.tDocumento, 12), 3) + '-' + RIGHT(MD.tDocumento, 8)  else ltrim(rtrim(isnull(MN.tObservacion,''))) + ' - Documento Referencia: ' + LEFT(MD.tdocumento, 1) + LEFT(RIGHT(MD.tDocumento, 12), 3) + '-' + RIGHT(MD.tDocumento, 8) end AS Notas,
		--case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then [dbo].[FnCantidadConLetra](Cast(MD.nVenta/TT.nVenta as decimal(30,5)),1) else [dbo].[FnCantidadConLetra](Cast(MD.nVenta as decimal(30,5)),0) end AS MontoEnLetras,
		'' AS MontoEnLetras,
		isnull(MD.tordenCompra,'0') AS NroOrdenCompra,
		case when MD.tEstadoDocumento ='03' and ISNULL(MD.tClientePago,'')<>'' then CONVERT(bit,1) else CONVERT(bit,0) end AS Credito,
		case 
			when MD.tEstadoDocumento ='03' and ISNULL(MD.tClientePago,'')<>'' then (
			case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((MD.nVenta- (
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			) -0.001)/TT.nVenta as decimal(30,5))  else Cast(MD.nVenta- (
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			) -0.001 as decimal(30,5))  end
			)else 0.00 end  AS MontoCredito,
		ISNULL(MN.lEstadoFacturacion,0) AS lEstadoFacturacion,
		ISNULL(MN.tRespFacturacion,'') AS RespuestaFacturacion,
		ISNULL(MN.tImprTermica,'') AS ImpresionTerminca,
		TP.Token as Token, 
		isnull(TT.nVenta,0.000) as TipoCambio,
		cast(0.00 as decimal(30,5))  as MontoRetencion,
		cast(0.00 as decimal(30,5))  as TasaRetencion,
		
		MD.tDocumento as DocumentoReferencia,
		(case when ISNULL(TC.lCajaContingencia,0)=1 then '0' else  LEFT(MD.tdocumento, 1) end)  + LEFT(RIGHT(MD.tdocumento, 12), 3) + '-' + RIGHT(MD.tDocumento, 8) AS NroReferencia,
		isnull(vmtnc.Codigo,'01') AS Tipo,
		ltrim(rtrim(isnull(vmtnc.Descripcion,'ANULACION DE LA OPERACION'))) AS Descripcion,
		MN.fRegistroAnulado AS FechaAnulado,
		MN.fRegistro AS FechaRegistro,
		'ANULACION NC: ' + LEFT(MD.tdocumento, 1) + LEFT(RIGHT(mn.tNotaCredito, 12), 3) + '-' + RIGHT(MN.tNotaCredito, 8)  as motivoBaja,
		MN.tEstadoDocumento as EstadoDocumento,
		isnull(vtp2.Sunat,'00') AS TipoDocumentoReferencia
	--MD.* 
	from MNOTACREDITO MN  
		INNER JOIN MDOCUMENTO MD on MN.tDocumento = MD.tDocumento
		INNER JOIN vCliente vc on MD.tCodigoCliente = vc.Codigo
		LEFT JOIN vTipoDocumento vtp on MN.tTipoDocumento= vtp.Codigo
		LEFT JOIN vTipoDocumento vtp2 on MD.tTipoDocumento= vtp2.Codigo
		left join TUBIGEO TUB on vc.CodigoUbigeo= TUB.tCodigo
		LEFT JOIN vTipoIdentidad vti on vc.tTipoIdentidad =vti.Codigo
		LEFT JOIN vMotivoNotaCredito vmtnc on MN.tMotivo = vmtnc.Codigo
		LEFT JOIN vDelivery vd on MD.tClientePago = vd.Codigo
		left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha)
		LEFT JOIN TCAJA TC on MD.tCaja = TC.tCaja
		LEFT JOIN (
			select D.tDocumento,
			sum(case when (ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaInafectacion,0)=1) OR (ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaInafectacion,0)=0 AND isnull(D.tAplicaExoneracion,0)=0) then nPrecioNeto*nCantidad ELSE 0.00 end)  AS Inafecto,
			sum(case when ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaExoneracion,0)=1 then D.nPrecioNeto*D.nCantidad ELSE 0.00 end) as Exonerado,
			sum(case when ROUND(D.nImpuesto1,2)<>0 and ISNULL(M.tCortesia,'')='' then D.nPrecioNeto*D.nCantidad ELSE 0.00 end) as Gravado,
			sum(case 
				when ROUND(D.nprecioneto,2)=0  then  CAST(D.nCantidad + ( D.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,5)) 
				when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(D.nPrecioNeto*D.nCantidad AS decimal(30,5)) 
			else 0.00 end) as Gratuita,
			sum(case 
				when ROUND(D.nprecioneto,2)=0  then  CAST((D.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,5)) 
				when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(D.nImpuesto1 AS decimal(30,5)) 
			else 0.00 end) as IgvGratuita,
			sum(case when ISNULL(M.tCortesia,'')='' then D.nPrecioNeto*D.nCantidad else 0.00 end) as Neto,
			sum(Case when ISNULL(M.tCortesia,'')='' then  D.nImpuesto1 else 0.00 end) as nImpuesto1,
			sum(Case when ISNULL(M.tCortesia,'')='' then  isnull(D.nImpuesto2,0) + isnull(D.nImpuesto3,0) else 0.00 end) as nImpuesto2,-- sum(nImpuesto3)as nImpuesto3,
			sum(case when ISNULL(M.tCortesia,'')='' then ((D.nPrecioNeto*D.nCantidad)+D.nImpuesto1) else 0.00 end ) as  nVenta,
			sum(isnull(D.nLey1,0)) as  nLey1, MN.tNotaCredito
		from
			MNOTACREDITO MN
			INNER JOIN MDOCUMENTO M on MN.tDocumento= M.tDocumento 
			INNER JOIN DDOCUMENTO D  ON m.tDocumento = D.tDocumento,
			(select Impuesto1 from TPARAMETRO) TP
		where (MN.nVenta between M.nVenta-0.05 and M.nVenta+0.05 )
		Group BY D.tDocumento, MN.tNotaCredito
		UNION
				select M.tDocumento,
			(case when ROUND(MN.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' then MN.nNeto ELSE 0.00 end)  AS Inafecto,
			0.00 as Exonerado,
			(case when ROUND(MN.nImpuesto1,2)<>0 and ISNULL(M.tCortesia,'')='' then MN.nNeto ELSE 0.00 end) as Gravado,
			(case 
				when ROUND(MN.nNeto,2)=0  then  CAST(1 + (TP.Impuesto1/100) AS decimal(30,5)) 
				when ROUND(MN.nNeto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(MN.nNeto AS decimal(30,5)) 
			else 0.00 end) as Gratuita,
			(case 
				when ROUND(MN.nNeto,2)=0  then  CAST((TP.Impuesto1/100) AS decimal(30,5)) 
				when ROUND(MN.nNeto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(MN.nImpuesto1 AS decimal(30,5)) 
			else 0.00 end) as IgvGratuita,
			(case when ISNULL(M.tCortesia,'')='' then MN.nNeto else 0.00 end) as Neto,
			(Case when ISNULL(M.tCortesia,'')='' then  MN.nImpuesto1 else 0.00 end) as nImpuesto1,
			(Case when ISNULL(M.tCortesia,'')='' then  isnull(MN.nImpuesto2,0) + isnull(MN.nImpuesto3,0) else 0.00 end) as nImpuesto2,-- sum(nImpuesto3)as nImpuesto3,
			(case when ISNULL(M.tCortesia,'')='' then ((MN.nNeto)+MN.nImpuesto1) else 0.00 end ) as  nVenta,
			(0) as  nLey1, MN.tNotaCredito
		from
			MNOTACREDITO MN
			INNER JOIN MDOCUMENTO M on MN.tDocumento= M.tDocumento ,
			(select Impuesto1 from TPARAMETRO) TP
		where  MN.nVenta < M.nVenta-0.05   and not  exists ( select * from DNOTACREDITO where tNotaCredito=MN.tNotaCredito)
		UNION
		select MN.tDocumento,
			sum(case when (ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaInafectacion,0)=1) OR (ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaInafectacion,0)=0 AND isnull(D.tAplicaExoneracion,0)=0) then nPrecioNeto*nCantidad ELSE 0.00 end)  AS Inafecto,
			sum(case when ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaExoneracion,0)=1 then D.nPrecioNeto*D.nCantidad ELSE 0.00 end) as Exonerado,
			sum(case when ROUND(D.nImpuesto1,2)<>0 and ISNULL(M.tCortesia,'')='' then D.nPrecioNeto*D.nCantidad ELSE 0.00 end) as Gravado,
			sum(case 
				when ROUND(D.nprecioneto,2)=0  then  CAST(D.nCantidad + ( D.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,5)) 
				when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(D.nPrecioNeto*D.nCantidad AS decimal(30,5)) 
			else 0.00 end) as Gratuita,
			sum(case 
				when ROUND(D.nprecioneto,2)=0  then  CAST((D.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,5)) 
				when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(D.nImpuesto1 AS decimal(30,5)) 
			else 0.00 end) as IgvGratuita,
			sum(case when ISNULL(M.tCortesia,'')='' then D.nPrecioNeto*D.nCantidad else 0.00 end) as Neto,
			sum(Case when ISNULL(M.tCortesia,'')='' then  D.nImpuesto1 else 0.00 end) as nImpuesto1,
			sum(Case when ISNULL(M.tCortesia,'')='' then  isnull(D.nImpuesto2,0) + isnull(D.nImpuesto3,0) else 0.00 end) as nImpuesto2,-- sum(nImpuesto3)as nImpuesto3,
			sum(case when ISNULL(M.tCortesia,'')='' then ((D.nPrecioNeto*D.nCantidad)+D.nImpuesto1) else 0.00 end ) as  nVenta,
			sum(isnull(D.nLey1,0)) as  nLey1, MN.tNotaCredito
		from
			MNOTACREDITO MN
			INNER JOIN MDOCUMENTO M on MN.tDocumento= M.tDocumento 
			INNER JOIN DNOTACREDITO D  ON MN.tNotaCredito = D.tNotaCredito,
			(select Impuesto1 from TPARAMETRO) TP
		where  MN.nVenta < M.nVenta-0.05   and  exists ( select * from DNOTACREDITO where tNotaCredito=MN.tNotaCredito)
		Group BY MN.tDocumento, MN.tNotaCredito
		) DC ON MN.tNotaCredito = DC.tNotaCredito ,
		(
			select top 1 tIdentificacionTributaria, tRazonSocial, tRazonComercial, 
						tCodigoUbigeo,tDireccion, tDireccion2, tCodigoTienda, TokenInfoFact as Token, tEmail as Correo,
						isnull(TPARAMETRO.lDetraccion,0) lDetraccion, isnull(TPARAMETRO.lMontoDetraccion,0.00) nMontoDetraccion, 
						isnull(TPARAMETRO.lPorcentajeDetraccion,0) nPorcentajeDetracccion, isnull(TPARAMETRO.tMsjDetraccion,'') Detraccion,
						isnull(convert(float,isnull(TPARAMETRO.tMontoRetencion,0)),0.00) nMontoRetencion, 
						isnull(convert(float,isnull(TPARAMETRO.lRetencion,0)),0.00) nPorcentajeRetencion,
						TUBIGEO.*
			from TPARAMETRO left join  TUBIGEO on TPARAMETRO.tCodigoUbigeo= TUBIGEO.tCodigo
		) TP
	where vtp.lFacturacionE=1

GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoCabeceraInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoCabeceraInfoFact]
GO
CREATE  VIEW vDocumentoCabeceraInfoFact
as
	select 
		MD.tDocumento AS tDocumento, 
		(case when ISNULL(TC.lCajaContingencia,0)=1 then '0' else  LEFT(MD.tdocumento, 1) end) + LEFT(RIGHT(MD.tdocumento, 12), 3) + '-' + RIGHT(MD.tdocumento, 8) AS IdDocumento,
		isnull(vtp.Sunat,'00') AS TipoDocumento,
		Case when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then '1001' else '0101' end  TipoOperacion, -- revisar retencion y operacion onerosa
		--CONVERT(NVARCHAR(10), MD.fregistro, 120) AS FechaEmision,
		CASE
			WHEN ISNULL(TC.lCajaContingencia,0)=1 THEN ISNULL(CONVERT(NVARCHAR(10), MD.femision, 120), CONVERT(NVARCHAR(10), MD.fregistro, 120))
			ELSE CONVERT(NVARCHAR(10), MD.fregistro, 120)
		END AS FechaEmision,
		--CONVERT(NVARCHAR(10), MD.fregistro, 108) AS HoraEmision,
		CASE
			WHEN ISNULL(TC.lCajaContingencia,0)=1 THEN ISNULL(CONVERT(NVARCHAR(10), MD.femision, 108), CONVERT(NVARCHAR(10), MD.fregistro, 108))
			ELSE CONVERT(NVARCHAR(10), MD.fregistro, 108)
		END AS HoraEmision,
		--CONVERT(NVARCHAR(10), MD.fregistro, 120) AS FechaVencimiento, -- VERIFICAR AL CREDITO
		CONVERT(NVARCHAR(10), 
		CASE
			WHEN ISNULL(TC.lCajaContingencia,0)=1 THEN ISNULL(MD.femision, MD.fregistro)
			ELSE MD.fregistro
		END
		+case when MD.tEstadoDocumento ='03' then (case when isnull(vd.nDiasCredito,0) = 0 then 1 else isnull(vd.nDiasCredito,0)end) else 0 end , 120) AS FechaVencimiento, -- VERIFICAR AL CREDITO
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then 'USD' else 'PEN' end AS Moneda,
		TP.tIdentificacionTributaria AS Emi_NroDocumento,
		'6' AS Emi_TipoDocumento,
		ltrim(rtrim(TP.tRazonSocial)) AS Emi_NombreLegal,
		ltrim(rtrim(TP.tRazonComercial)) AS Emi_NombreComercial,
		TP.tCodigoUbigeo AS Emi_Ubigeo,
		ltrim(rtrim(TP.tDireccion)) AS Emi_Direccion,
		ltrim(rtrim(TP.tDireccion2)) AS Emi_Urbanizacion,
		TP.tDepartamento AS Emi_Departamento,
		TP.tProvincia AS Emi_Provincia,
		TP.tDistrito AS Emi_Distrito,
		ltrim(rtrim(TP.tCodigoTienda)) AS Emi_CodigoAnexo,
		'PE' as Emi_Pais,
		case when isnull(TP.Correo,'')='' then 'facturacion.electronica@infhotel.com.pe' else isnull(TP.Correo,'') end  AS Emi_Correo,
		ltrim(rtrim(vc.tIdentidad)) AS Rec_NroDocumento,
		vti.tvalor AS Rec_TipoDocumento,
		case when len(ltrim(rtrim(vc.Descripcion)))=1 then 'Cliente' else ltrim(rtrim(vc.Descripcion)) end  AS Rec_NombreLegal,
		case when len(ltrim(rtrim(vc.Descripcion)))=1 then 'Cliente' else ltrim(rtrim(vc.Descripcion)) end  AS Rec_NombreComercial,
		vc.CodigoUbigeo AS Rec_Ubigeo,
		ltrim(rtrim(vc.tDireccion)) AS Rec_Direccion,
		ltrim(rtrim(vc.Urbanizacion)) AS Rec_Urbanizacion,
		TUB.tDepartamento AS Rec_Departamento,
		TUB.tProvincia AS Rec_Provincia,
		TUB.tDistrito AS Rec_Distrito,
		'' AS Rec_CodigoAnexo,
		'PE' as Rec_Pais,
		case when ltrim(rtrim(isnull(vc.tCorreo,'')))='' then 'facturacion.electronica@infhotel.com.pe' else ltrim(rtrim(isnull(vc.tCorreo,''))) end  AS Rec_Correo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Gravado/TT.nVenta as decimal(30,5)) else Cast(DC.Gravado as decimal(30,5)) end  AS Gravadas,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Gratuita/TT.nVenta as decimal(30,5)) else Cast(DC.Gratuita as decimal(30,5)) end  AS Gratuitas,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Inafecto/TT.nVenta as decimal(30,5)) else Cast(DC.Inafecto as decimal(30,5)) end  AS Inafectas,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Exonerado/TT.nVenta as decimal(30,5)) else Cast(DC.Exonerado as decimal(30,5)) end AS Exoneradas,
		0.00 AS Exportacion,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((DC.Neto + DC.NetoAnticipo)/TT.nVenta as decimal(30,5)) else Cast((DC.Neto+DC.NetoAnticipo) as decimal(30,5)) end AS LineExtensionAmount,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((DC.Neto+DC.TotalAnticipo + DC.nImpuesto1)/TT.nVenta as decimal(30,5)) else Cast(DC.Neto+DC.TotalAnticipo+DC.nImpuesto1 as decimal(30,5)) end AS TaxInclusiveAmount,
		
		cast( DC.NetoAnticipo as decimal(30,5)) AS DescuentoGlobal,
		Case when DC.NetoAnticipo>0 then
				cast(DC.NetoAnticipo/(DC.Neto+DC.NetoAnticipo) AS decimal(30,8))
			else 0.00000 end AS FactorMultiplicadorDscto,
		Case when DC.NetoAnticipo>0 then
				case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((DC.Neto+DC.NetoAnticipo)/TT.nVenta as decimal(30,5)) else Cast(DC.Neto+DC.NetoAnticipo as decimal(30,5)) end 
			else 0.00000 end AS MontoBaseParaDcto,
		Case when DC.NetoAnticipo>0 then '04' else '' end AS CodigoRazonDcto,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nVenta/TT.nVenta as decimal(30,5)) else Cast(DC.nVenta as decimal(30,5)) end TotalVenta,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nImpuesto1/TT.nVenta as decimal(30,5)) else Cast(DC.nImpuesto1 as decimal(30,5)) end TotalIgv,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.IgvGratuita/TT.nVenta as decimal(30,5)) else Cast(DC.IgvGratuita as decimal(30,5)) end AS TotalIgvGratuita,
		0.00 AS TotalIsc,
		Cast(DC.nLey1 as decimal(30,5))  AS  TotalOtrosTributos,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nImpuesto2/TT.nVenta as decimal(30,5)) else Cast(DC.nImpuesto2 as decimal(30,5)) end AS TotalRecargo,
		Cast(DC.nImpuesto2/(case when DC.Neto>0 then DC.Neto else 1 end ) as decimal(30,5)) AS PorcentajeRecargo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Neto/TT.nVenta as decimal(30,5)) else Cast(DC.Neto as decimal(30,5)) end AS MontoBaseRecargo,
		0.00 AS MontoPercepcion,
		Case when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100) else 0.00 end AS MontoDetraccion,
		Case when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then TP.nPorcentajeDetracccion else 0 end AS TasaDetraccion,
		SUBSTRING(TP.Detraccion, 5, 3) AS CodigoMedioPago,
		ltrim(rtrim(SUBSTRING(TP.Detraccion, 9, 30))) AS CuentaBancoNacion,
		SUBSTRING(TP.Detraccion, 1, 3) AS CodigoBienOServicio,
		'' AS TipoDocAnticipo,
		'' AS MonedaAnticipo,
		0.00 AS MontoAnticipo,
		cast(DC.TotalAnticipo as decimal(30,5)) AS MontoTotalAnticipo,
		ltrim(rtrim(ISNULL((SELECT top 1 tObservacion FROM MPEDIDO WHERE tCodigoPedido IN (SELECT TOP 1 tCodigoPedido FROM DDOCUMENTO WHERE tDocumento=MD.tDocumento)),''))) AS Notas,
		'' AS MontoEnLetras,
		--ltrim(rtrim(MD.tordenCompra)) AS NroOrdenCompra,
		case when  ISNULL(ltrim(rtrim(MD.tordenCompra)),'') = '' then isnull((select top 1 tGuiaTransporte from MGUIATRANSPORTE where tDocumento = Md.tDocumento),'') else  ISNULL(ltrim(rtrim(MD.tordenCompra)),'') end  AS NroOrdenCompra,
		case when MD.tEstadoDocumento ='03' and ISNULL(MD.tClientePago,'')<>'' then CONVERT(bit,1) else CONVERT(bit,0) end AS Credito,
		case 
			when MD.tEstadoDocumento ='03' and ISNULL(MD.tClientePago,'')<>'' then (
				case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((MD.nVenta- (
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			) -0.001)/TT.nVenta as decimal(30,5))  else Cast(MD.nVenta- (
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			) -0.001 as decimal(30,5))  end
			)else 0.00 end  AS MontoCredito,
		ISNULL(MD.lEstadoFacturacion,0) AS lEstadoFacturacion,
		ISNULL(MD.tRespFacturacion,'') AS RespuestaFacturacion,
		ISNULL(MD.tImprTermica,'') AS ImpresionTerminca,
		TP.Token As Token, isnull(TT.nVenta,0.000) as TipoCambio,
		cast(0.00 as decimal(30,5))  as MontoRetencion,
		cast(0.00 as decimal(30,5))  as TasaRetencion,
		
		MD.fRegistroAnulado AS FechaAnulado,
		MD.fRegistro AS FechaRegistro,
		case when ISNULL(MD.tObservacion,'') ='' and  MD.tEstadoDocumento ='04' then 'ANULACION DE LA OPERACION' else ISNULL(MD.tObservacion,'') end as motivoBaja,
		MD.tEstadoDocumento as EstadoDocumento
	--MD.* 
	from MDOCUMENTO MD
		INNER JOIN vCliente vc on MD.tCodigoCliente = vc.Codigo
		LEFT JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo
		left join TUBIGEO TUB on vc.CodigoUbigeo= TUB.tCodigo
		LEFT JOIN vTipoIdentidad vti on vc.tTipoIdentidad =vti.Codigo
		left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha) 
		LEFT JOIN vDelivery vd on MD.tClientePago = vd.Codigo
		LEFT JOIN TCAJA TC on MD.tCaja = TC.tCaja
		LEFT JOIN (
			select D.tDocumento,
			sum(case when ( ROUND(D.nprecioneto,2)>0 and ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaInafectacion,0)=1) OR (ROUND(D.nprecioneto,2)>0 and ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaInafectacion,0)=0 AND isnull(D.tAplicaExoneracion,0)=0) then nPrecioNeto*nCantidad ELSE 0.00 end)  AS Inafecto,
			sum(case when ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaExoneracion,0)=1 then D.nPrecioNeto*D.nCantidad ELSE 0.00 end) as Exonerado,
			sum(case when ROUND(D.nImpuesto1,2)<>0 and ISNULL(M.tCortesia,'')='' then D.nPrecioNeto*D.nCantidad ELSE 0.00 end) as Gravado,
			sum(case 
				when ROUND(D.nprecioneto,2)=0  then  CAST(D.nCantidad + ( D.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,5)) 
				when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(D.nPrecioNeto*D.nCantidad AS decimal(30,5)) 
			else 0.00 end) as Gratuita,
			sum(case 
				when ROUND(D.nprecioneto,2)=0  then  CAST((D.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,5)) 
				when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(D.nImpuesto1 AS decimal(30,5)) 
			else 0.00 end) as IgvGratuita,
			sum(case when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')='' then D.nPrecioNeto*D.nCantidad else 0.00 end) as Neto,
			sum(Case when ISNULL(M.tCortesia,'')='' then  D.nImpuesto1 else 0.00 end) as nImpuesto1,
			sum(Case when ISNULL(M.tCortesia,'')='' then  isnull(D.nImpuesto2,0) + isnull(D.nImpuesto3,0) else 0.00 end) as nImpuesto2,-- sum(nImpuesto3)as nImpuesto3,
			sum(case when ISNULL(M.tCortesia,'')='' then ((D.nPrecioNeto*D.nCantidad)+D.nImpuesto1) else 0.00 end ) as  nVenta,
			sum(isnull(D.nLey1,0)) as  nLey1,
			isnull(SUM(case when ISNULL(D.tDocumentoAnticipo,'')<>'' then ABS( (D.nPrecioNeto + D.nPrecioImpuesto1)*D.nCantidad ) else 0.00000 end ),0.00000000) TotalAnticipo,
			isnull(SUM(case when ISNULL(D.tDocumentoAnticipo,'')<>'' then ABS( (D.nPrecioNeto)*D.nCantidad ) else 0.00000 end ),0.00000000) NetoAnticipo
			from DDOCUMENTO D
			left join MDOCUMENTO M on D.tDocumento= M.tDocumento , (select Impuesto1 from TPARAMETRO) TP
			--where d.tDocumento='F00100000000553'
			Group BY D.tDocumento
		) DC ON MD.tDocumento = DC.tDocumento ,

		(
			select top 1 tIdentificacionTributaria, tRazonSocial, tRazonComercial, 
						tCodigoUbigeo,tDireccion, tDireccion2, tCodigoTienda,TokenInfoFact as Token, tEmail as Correo, 
						isnull(TPARAMETRO.lDetraccion,0) lDetraccion, isnull(TPARAMETRO.lMontoDetraccion,0.00) nMontoDetraccion, 
						isnull(TPARAMETRO.lPorcentajeDetraccion,0) nPorcentajeDetracccion, isnull(TPARAMETRO.tMsjDetraccion,'') Detraccion,
						isnull(convert(float,isnull(TPARAMETRO.tMontoRetencion,0)),0.00) nMontoRetencion, 
						isnull(convert(float,isnull(TPARAMETRO.lRetencion,0)),0.00) nPorcentajeRetencion,
						TUBIGEO.*
			from TPARAMETRO left join  TUBIGEO on TPARAMETRO.tCodigoUbigeo= TUBIGEO.tCodigo
		) TP
	where vtp.lFacturacionE=1
GO

--- Modificacion, validacion emision en dolares - cliente tacamas

ALTER  VIEW [dbo].[vDocumentoCabeceraInfoFact]
as
	select 
		MD.tDocumento AS tDocumento, 
		(case when ISNULL(TC.lCajaContingencia,0)=1 then '0' else  LEFT(MD.tdocumento, 1) end) + LEFT(RIGHT(MD.tdocumento, 12), 3) + '-' + RIGHT(MD.tdocumento, 8) AS IdDocumento,
		isnull(vtp.Sunat,'00') AS TipoDocumento,
		Case when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then '1001' else '0101' end  TipoOperacion, -- revisar retencion y operacion onerosa
		--CONVERT(NVARCHAR(10), MD.fregistro, 120) AS FechaEmision,
		CASE
			WHEN ISNULL(TC.lCajaContingencia,0)=1 THEN ISNULL(CONVERT(NVARCHAR(10), MD.femision, 120), CONVERT(NVARCHAR(10), MD.fregistro, 120))
			ELSE CONVERT(NVARCHAR(10), MD.fregistro, 120)
		END AS FechaEmision,
		--CONVERT(NVARCHAR(10), MD.fregistro, 108) AS HoraEmision,
		CASE
			WHEN ISNULL(TC.lCajaContingencia,0)=1 THEN ISNULL(CONVERT(NVARCHAR(10), MD.femision, 108), CONVERT(NVARCHAR(10), MD.fregistro, 108))
			ELSE CONVERT(NVARCHAR(10), MD.fregistro, 108)
		END AS HoraEmision,
		--CONVERT(NVARCHAR(10), MD.fregistro, 120) AS FechaVencimiento, -- VERIFICAR AL CREDITO
		CONVERT(NVARCHAR(10), 
		CASE
			WHEN ISNULL(TC.lCajaContingencia,0)=1 THEN ISNULL(MD.femision, MD.fregistro)
			ELSE MD.fregistro
		END
		+case when MD.tEstadoDocumento ='03' then (case when isnull(vd.nDiasCredito,0) = 0 then 1 else isnull(vd.nDiasCredito,0)end) else 0 end , 120) AS FechaVencimiento, -- VERIFICAR AL CREDITO
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then 'USD' else 'PEN' end AS Moneda,
		TP.tIdentificacionTributaria AS Emi_NroDocumento,
		'6' AS Emi_TipoDocumento,
		ltrim(rtrim(TP.tRazonSocial)) AS Emi_NombreLegal,
		ltrim(rtrim(TP.tRazonComercial)) AS Emi_NombreComercial,
		TP.tCodigoUbigeo AS Emi_Ubigeo,
		ltrim(rtrim(TP.tDireccion)) AS Emi_Direccion,
		ltrim(rtrim(TP.tDireccion2)) AS Emi_Urbanizacion,
		TP.tDepartamento AS Emi_Departamento,
		TP.tProvincia AS Emi_Provincia,
		TP.tDistrito AS Emi_Distrito,
		ltrim(rtrim(TP.tCodigoTienda)) AS Emi_CodigoAnexo,
		'PE' as Emi_Pais,
		case when isnull(TP.Correo,'')='' then 'facturacion.electronica@infhotel.com.pe' else isnull(TP.Correo,'') end  AS Emi_Correo,
		ltrim(rtrim(vc.tIdentidad)) AS Rec_NroDocumento,
		vti.tvalor AS Rec_TipoDocumento,
		case when len(ltrim(rtrim(vc.Descripcion)))=1 then 'Cliente' else ltrim(rtrim(vc.Descripcion)) end  AS Rec_NombreLegal,
		case when len(ltrim(rtrim(vc.Descripcion)))=1 then 'Cliente' else ltrim(rtrim(vc.Descripcion)) end  AS Rec_NombreComercial,
		vc.CodigoUbigeo AS Rec_Ubigeo,
		ltrim(rtrim(vc.tDireccion)) AS Rec_Direccion,
		ltrim(rtrim(vc.Urbanizacion)) AS Rec_Urbanizacion,
		TUB.tDepartamento AS Rec_Departamento,
		TUB.tProvincia AS Rec_Provincia,
		TUB.tDistrito AS Rec_Distrito,
		'' AS Rec_CodigoAnexo,
		'PE' as Rec_Pais,
		case when ltrim(rtrim(isnull(vc.tCorreo,'')))='' then 'facturacion.electronica@infhotel.com.pe' else ltrim(rtrim(isnull(vc.tCorreo,''))) end  AS Rec_Correo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Gravado/TT.nVenta as decimal(30,5)) else Cast(DC.Gravado as decimal(30,5)) end  AS Gravadas,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Gratuita/TT.nVenta as decimal(30,5)) else Cast(DC.Gratuita as decimal(30,5)) end  AS Gratuitas,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Inafecto/TT.nVenta as decimal(30,5)) else Cast(DC.Inafecto as decimal(30,5)) end  AS Inafectas,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Exonerado/TT.nVenta as decimal(30,5)) else Cast(DC.Exonerado as decimal(30,5)) end AS Exoneradas,
		0.00 AS Exportacion,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((DC.Neto + DC.NetoAnticipo)/TT.nVenta as decimal(30,5)) else Cast((DC.Neto+DC.NetoAnticipo) as decimal(30,5)) end AS LineExtensionAmount,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((DC.Neto+DC.TotalAnticipo + DC.nImpuesto1)/TT.nVenta as decimal(30,5)) else Cast(DC.Neto+DC.TotalAnticipo+DC.nImpuesto1 as decimal(30,5)) end AS TaxInclusiveAmount,
		
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then cast( DC.NetoAnticipo/TT.nVenta as decimal(30,5)) else cast( DC.NetoAnticipo as decimal(30,5)) end AS DescuentoGlobal,
		Case when DC.NetoAnticipo>0 then
				cast(DC.NetoAnticipo/(DC.Neto+DC.NetoAnticipo) AS decimal(30,8))
			else 0.00000 end AS FactorMultiplicadorDscto,
		Case when DC.NetoAnticipo>0 then
				case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((DC.Neto+DC.NetoAnticipo)/TT.nVenta as decimal(30,5)) else Cast(DC.Neto+DC.NetoAnticipo as decimal(30,5)) end 
			else 0.00000 end AS MontoBaseParaDcto,
		Case when DC.NetoAnticipo>0 then '04' else '' end AS CodigoRazonDcto,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nVenta/TT.nVenta as decimal(30,5)) else Cast(DC.nVenta as decimal(30,5)) end TotalVenta,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nImpuesto1/TT.nVenta as decimal(30,5)) else Cast(DC.nImpuesto1 as decimal(30,5)) end TotalIgv,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.IgvGratuita/TT.nVenta as decimal(30,5)) else Cast(DC.IgvGratuita as decimal(30,5)) end AS TotalIgvGratuita,
		0.00 AS TotalIsc,
		Cast(DC.nLey1 as decimal(30,5))  AS  TotalOtrosTributos,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.nImpuesto2/TT.nVenta as decimal(30,5)) else Cast(DC.nImpuesto2 as decimal(30,5)) end AS TotalRecargo,
		Cast(DC.nImpuesto2/(case when DC.Neto>0 then DC.Neto else 1 end ) as decimal(30,5)) AS PorcentajeRecargo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(DC.Neto/TT.nVenta as decimal(30,5)) else Cast(DC.Neto as decimal(30,5)) end AS MontoBaseRecargo,
		0.00 AS MontoPercepcion,
		Case when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100) else 0.00 end AS MontoDetraccion,
		Case when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then TP.nPorcentajeDetracccion else 0 end AS TasaDetraccion,
		SUBSTRING(TP.Detraccion, 5, 3) AS CodigoMedioPago,
		ltrim(rtrim(SUBSTRING(TP.Detraccion, 9, 30))) AS CuentaBancoNacion,
		SUBSTRING(TP.Detraccion, 1, 3) AS CodigoBienOServicio,
		'' AS TipoDocAnticipo,
		'' AS MonedaAnticipo,
		0.00 AS MontoAnticipo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then cast(DC.TotalAnticipo/TT.nVenta as decimal(30,5)) else cast(DC.TotalAnticipo as decimal(30,5)) end AS MontoTotalAnticipo,
		ltrim(rtrim(ISNULL((SELECT top 1 tObservacion FROM MPEDIDO WHERE tCodigoPedido IN (SELECT TOP 1 tCodigoPedido FROM DDOCUMENTO WHERE tDocumento=MD.tDocumento)),''))) AS Notas,
		'' AS MontoEnLetras,
		ltrim(rtrim(MD.tordenCompra)) AS NroOrdenCompra,
		case when MD.tEstadoDocumento ='03' and ISNULL(MD.tClientePago,'')<>'' then CONVERT(bit,1) else CONVERT(bit,0) end AS Credito,
		case 
			when MD.tEstadoDocumento ='03' and ISNULL(MD.tClientePago,'')<>'' then (
				case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((MD.nVenta- (
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			) -0.001)/TT.nVenta as decimal(30,5))  else Cast(MD.nVenta- (
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			) -0.001 as decimal(30,5))  end
			)else 0.00 end  AS MontoCredito,
		ISNULL(MD.lEstadoFacturacion,0) AS lEstadoFacturacion,
		ISNULL(MD.tRespFacturacion,'') AS RespuestaFacturacion,
		ISNULL(MD.tImprTermica,'') AS ImpresionTerminca,
		TP.Token As Token, isnull(TT.nVenta,0.000) as TipoCambio,
		
		cast(0.00 as decimal(30,5))  as MontoRetencion,
		cast(0.00 as decimal(30,5))  as TasaRetencion,
		
		MD.fRegistroAnulado AS FechaAnulado,
		MD.fRegistro AS FechaRegistro,
		case when ISNULL(MD.tObservacion,'') ='' and  MD.tEstadoDocumento ='04' then 'ANULACION DE LA OPERACION' else ISNULL(MD.tObservacion,'') end as motivoBaja,
		MD.tEstadoDocumento as EstadoDocumento
	--MD.* 
	from MDOCUMENTO MD
		INNER JOIN vCliente vc on MD.tCodigoCliente = vc.Codigo
		LEFT JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo
		left join TUBIGEO TUB on vc.CodigoUbigeo= TUB.tCodigo
		LEFT JOIN vTipoIdentidad vti on vc.tTipoIdentidad =vti.Codigo
		left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha) 
		LEFT JOIN vDelivery vd on MD.tClientePago = vd.Codigo
		LEFT JOIN TCAJA TC on MD.tCaja = TC.tCaja
		LEFT JOIN (
			select D.tDocumento,
			sum(case when ( ROUND(D.nprecioneto,2)>0 and ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaInafectacion,0)=1) OR (ROUND(D.nprecioneto,2)>0 and ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaInafectacion,0)=0 AND isnull(D.tAplicaExoneracion,0)=0) then nPrecioNeto*nCantidad ELSE 0.00 end)  AS Inafecto,
			sum(case when ROUND(D.nImpuesto1,2)=0 and ISNULL(M.tCortesia,'')='' and isnull(D.tAplicaExoneracion,0)=1 then D.nPrecioNeto*D.nCantidad ELSE 0.00 end) as Exonerado,
			sum(case when ROUND(D.nImpuesto1,2)<>0 and ISNULL(M.tCortesia,'')='' then D.nPrecioNeto*D.nCantidad ELSE 0.00 end) as Gravado,
			sum(case 
				when ROUND(D.nprecioneto,2)=0  then  CAST(D.nCantidad + ( D.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,5)) 
				when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(D.nPrecioNeto*D.nCantidad AS decimal(30,5)) 
			else 0.00 end) as Gratuita,
			sum(case 
				when ROUND(D.nprecioneto,2)=0  then  CAST((D.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,5)) 
				when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')<>'' then  CAST(D.nImpuesto1 AS decimal(30,5)) 
			else 0.00 end) as IgvGratuita,
			sum(case when ROUND(D.nprecioneto,2)<>0 and ISNULL(M.tCortesia,'')='' then D.nPrecioNeto*D.nCantidad else 0.00 end) as Neto,
			sum(Case when ISNULL(M.tCortesia,'')='' then  D.nImpuesto1 else 0.00 end) as nImpuesto1,
			sum(Case when ISNULL(M.tCortesia,'')='' then  isnull(D.nImpuesto2,0) + isnull(D.nImpuesto3,0) else 0.00 end) as nImpuesto2,-- sum(nImpuesto3)as nImpuesto3,
			sum(case when ISNULL(M.tCortesia,'')='' then ((D.nPrecioNeto*D.nCantidad)+D.nImpuesto1) else 0.00 end ) as  nVenta,
			sum(isnull(D.nLey1,0)) as  nLey1,
			isnull(SUM(case when ISNULL(D.tDocumentoAnticipo,'')<>'' then ABS( (D.nPrecioNeto + D.nPrecioImpuesto1)*D.nCantidad ) else 0.00000 end ),0.00000000) TotalAnticipo,
			isnull(SUM(case when ISNULL(D.tDocumentoAnticipo,'')<>'' then ABS( (D.nPrecioNeto)*D.nCantidad ) else 0.00000 end ),0.00000000) NetoAnticipo
			from DDOCUMENTO D
			left join MDOCUMENTO M on D.tDocumento= M.tDocumento , (select Impuesto1 from TPARAMETRO) TP
			--where d.tDocumento='F00100000000553'
			Group BY D.tDocumento
		) DC ON MD.tDocumento = DC.tDocumento ,

		(
			select top 1 tIdentificacionTributaria, tRazonSocial, tRazonComercial, 
						tCodigoUbigeo,tDireccion, tDireccion2, tCodigoTienda,TokenInfoFact as Token, tEmail as Correo, 
						isnull(TPARAMETRO.lDetraccion,0) lDetraccion, isnull(TPARAMETRO.lMontoDetraccion,0.00) nMontoDetraccion, 
						isnull(TPARAMETRO.lPorcentajeDetraccion,0) nPorcentajeDetracccion, isnull(TPARAMETRO.tMsjDetraccion,'') Detraccion,
						isnull(convert(float,isnull(TPARAMETRO.tMontoRetencion,0)),0.00) nMontoRetencion, 
						isnull(convert(float,isnull(TPARAMETRO.lRetencion,0)),0.00) nPorcentajeRetencion,
						TUBIGEO.*
			from TPARAMETRO left join  TUBIGEO on TPARAMETRO.tCodigoUbigeo= TUBIGEO.tCodigo
		) TP
	where vtp.lFacturacionE=1
GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vNotaCreditoDetalleInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vNotaCreditoDetalleInfoFact]
GO
CREATE VIEW vNotaCreditoDetalleInfoFact
AS
Select tDocumento,	Id,	Cantidad,	UnidadMedida,	CodigoItem,	Descripcion,	PrecioUnitario/TC as PrecioUnitario,	PrecioReferencial/TC as PrecioReferencial,	TipoPrecio,	
TipoImpuesto,	PorcentajeImpuesto,	Impuesto/TC as Impuesto,	ImpuestoSelectivo,	OtroImpuesto,	PrecioUnitarioBolsa,	Descuento,	
TotalVenta/TC as TotalVenta,	PlacaVehiculo,	DetalleViaje,	UbigeoOrigen,	UbigeoDestino,	DireccionOrigen,	DireccionDestino,	CodigoProductoSunat,	ConfiguracionVehicular
 from (
	Select
		MN.tNotaCredito  AS tDocumento, 
		CONVERT(int, DD.tItem) AS Id,
		CAST(DD.nCantidad AS decimal(30,5)) AS Cantidad,
		'NIU' AS UnidadMedida,
		DD.tCodigoProducto AS CodigoItem,
		ltrim(rtrim(vp.Descripcion)) AS Descripcion,
		case when ROUND(DD.nprecioneto,2)=0 OR ISNULL(MD.tCortesia,'')<>'' then  CAST(0 AS decimal(30,10))  else CAST(DD.nPrecioNeto AS decimal(30,10)) end AS PrecioUnitario,
		case 
			when ROUND(DD.nprecioneto,2)=0  then CAST(1 + (TP.Impuesto1/100) AS decimal(30,10)) 
			when ISNULL(MD.tCortesia,'')<>''  then CAST(DD.nPrecioNeto AS decimal(30,10)) 
			else CAST(DD.nPrecioNeto + (DD.nImpuesto1/DD.nCantidad) AS decimal(30,10)) end AS PrecioReferencial,
		CASE WHEN round(DD.nPrecioNeto,2)=0 OR ISNULL(MD.tCortesia,'')<>''then '02' else '01' end AS TipoPrecio,
		case
			when ROUND(DD.nprecioneto,2)<>0 and ISNULL(MD.tCortesia,'')='' and ((ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=1) OR (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=0 AND isnull(DD.tAplicaExoneracion,0)=0)) then '30'  --Inafecto
			when ROUND(DD.nprecioneto,2)<>0 and ISNULL(MD.tCortesia,'')<>'' and ((ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=1) OR (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=0 AND isnull(DD.tAplicaExoneracion,0)=0)) then '31'  --Inafecto
			when ROUND(DD.nprecioneto,2)<>0 and (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaExoneracion,0)=1)  and ISNULL(MD.tCortesia,'')='' then '20' --Exonerado
			when ROUND(DD.nprecioneto,2)<>0 and (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaExoneracion,0)=1)  and ISNULL(MD.tCortesia,'')<>'' then '21' --Exonerado
			when ROUND(DD.nprecioneto,2)=0 then '15'
			when ROUND(DD.nImpuesto1,2)<>0 and ISNULL(MD.tCortesia,'')='' then '10'
			when ROUND(DD.nImpuesto1,2)<>0 and ISNULL(MD.tCortesia,'')<>'' then '15' 
			Else '10' 
		end AS TipoImpuesto,
		case 
			when ROUND(DD.nprecioneto,2)=0 then Cast(TP.Impuesto1  as decimal(30,5))
			else  Cast(round((DD.nPrecioImpuesto1/DD.nprecioneto)*100,2)   as decimal(30,5))
		end AS PorcentajeImpuesto,
		case 
			when ROUND(DD.nprecioneto,2)=0 then CAST(DD.nCantidad*(TP.Impuesto1/100) AS decimal(30,5))
			else CAST(DD.nImpuesto1 AS decimal(30,5))
		end  AS Impuesto,
		0.00 as ImpuestoSelectivo,
		CAST(ISNULL(DD.nLey1,0) AS decimal(30,5)) AS OtroImpuesto,
		CAST((ISNULL(DD.nLey1,0)/DD.nCantidad) AS decimal(30,5)) AS PrecioUnitarioBolsa,
		0.00 as Descuento,
		case 
			when ROUND(DD.nprecioneto,2)=0 then  CAST( DD.nCantidad+ ( DD.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,10))
			else CAST(DD.nPrecioNeto*DD.nCantidad AS decimal(30,10))
		end  AS TotalVenta,
		'' AS PlacaVehiculo,
		'' AS DetalleViaje,
		'' AS UbigeoOrigen,
		'' AS UbigeoDestino,
		'' AS DireccionOrigen,
		'' AS DireccionDestino,
		'90101500' AS CodigoProductoSunat,
		'' AS ConfiguracionVehicular,
		Case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(TT.nVenta AS decimal(30,5)) else 1.000 end TC
		from  MNOTACREDITO MN  
		INNER JOIN DDOCUMENTO DD on MN.tDocumento = DD.tDocumento
		left join MDOCUMENTO MD on DD.tDocumento= MD.tDocumento
		inner JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo 
		left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha) 
		LEFT JOIN vProducto vp ON  DD.tCodigoProducto=vp.Codigo
		, (select Impuesto1 from TPARAMETRO) TP 
	Where  vtp.lFacturacionE=1  and (MN.nVenta between MD.nVenta-0.05 and MD.nVenta+0.05 ) and ISNULL(MD.lAnticipo,0)=0
	
	UNION

	Select
		MN.tNotaCredito  AS tDocumento, 
		1 AS Id,
		CAST(1 AS decimal(30,5)) AS Cantidad,
		'NIU' AS UnidadMedida,
		'0000000' AS CodigoItem,
		ltrim(rtrim(isnull(MN.tObservacion,'ANULACION'))) AS Descripcion,
		case when ROUND(MN.nNeto,2)=0 OR ISNULL(MD.tCortesia,'')<>'' then  CAST(0 AS decimal(30,10))  else CAST(MN.nNeto AS decimal(30,10)) end AS PrecioUnitario,
		case 
			when ROUND(MN.nNeto,2)=0  then CAST(1 + (TP.Impuesto1/100) AS decimal(30,10)) 
			when ISNULL(MD.tCortesia,'')<>''  then CAST(MN.nNeto AS decimal(30,10)) 
			else CAST(MN.nNeto + (MN.nImpuesto1) AS decimal(30,10)) end AS PrecioReferencial,
		CASE WHEN round(MN.nNeto,2)=0 OR ISNULL(MD.tCortesia,'')<>''then '02' else '01' end AS TipoPrecio,
		case
			when ROUND(MN.nNeto,2)<>0 and ISNULL(MD.tCortesia,'')='' and ROUND(MN.nImpuesto1,2)=0 then '30'  --Inafecto
			when ROUND(MN.nNeto,2)<>0 and ISNULL(MD.tCortesia,'')<>'' and ROUND(MN.nImpuesto1,2)=0 then '31'  --Inafecto
			when ROUND(MN.nNeto,2)=0 then '15'
			when ROUND(MN.nNeto,2)<>0 and ISNULL(MD.tCortesia,'')='' then '10'
			when ROUND(MN.nNeto,2)<>0 and ISNULL(MD.tCortesia,'')<>'' then '15' 
			Else '10' 
		end AS TipoImpuesto,
		case 
			when ROUND(MN.nNeto,2)=0 then Cast(TP.Impuesto1  as decimal(30,5))
			else  Cast(round(((MN.nImpuesto1/MN.nNeto))*100,0)   as decimal(30,5))
		end AS PorcentajeImpuesto,
		case 
			when ROUND(MN.nNeto,2)=0 then CAST((TP.Impuesto1/100) AS decimal(30,5))
			else CAST(MN.nImpuesto1 AS decimal(30,5))
		end  AS Impuesto,
		0.00 as ImpuestoSelectivo,
		CAST(0 AS decimal(30,5)) AS OtroImpuesto,
		CAST(0 AS decimal(30,5)) AS PrecioUnitarioBolsa,
		0.00 as Descuento,
		case 
			when ROUND(MN.nNeto,2)=0 then  CAST( (TP.Impuesto1/100)  AS decimal(30,10))
			else CAST(MN.nNeto AS decimal(30,10))
		end  AS TotalVenta,
		'' AS PlacaVehiculo,
		'' AS DetalleViaje,
		'' AS UbigeoOrigen,
		'' AS UbigeoDestino,
		'' AS DireccionOrigen,
		'' AS DireccionDestino,
		'90101500' AS CodigoProductoSunat,
		'' AS ConfiguracionVehicular,
		Case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(TT.nVenta AS decimal(30,5)) else 1.000 end TC
		from  MNOTACREDITO MN  
		left join MDOCUMENTO MD on MN.tDocumento= MD.tDocumento
		inner JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo 
		left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha) 
		, (select Impuesto1 from TPARAMETRO) TP 
	where  (vtp.lFacturacionE=1  and  MN.nVenta < MD.nVenta-0.05
		and not  exists ( select * from DNOTACREDITO where tNotaCredito=MN.tNotaCredito)) OR  ISNULL(MD.lAnticipo,0)=1
	
	UNION

	Select
		MN.tNotaCredito  AS tDocumento, 
		CONVERT(int, DD.tItem) AS Id,
		CAST(DD.nCantidad AS decimal(30,5)) AS Cantidad,
		'NIU' AS UnidadMedida,
		DD.tCodigoProducto AS CodigoItem,
		ltrim(rtrim(vp.Descripcion)) AS Descripcion,
		case when ROUND(DD.nprecioneto,2)=0 OR ISNULL(MD.tCortesia,'')<>'' then  CAST(0 AS decimal(30,10))  else CAST(DD.nPrecioNeto AS decimal(30,10)) end AS PrecioUnitario,
		case 
			when ROUND(DD.nprecioneto,2)=0  then CAST(1 + (TP.Impuesto1/100) AS decimal(30,10)) 
			when ISNULL(MD.tCortesia,'')<>''  then CAST(DD.nPrecioNeto AS decimal(30,10)) 
			else CAST(DD.nPrecioNeto + (DD.nImpuesto1/DD.nCantidad) AS decimal(30,10)) end AS PrecioReferencial,
		CASE WHEN round(DD.nPrecioNeto,2)=0 OR ISNULL(MD.tCortesia,'')<>''then '02' else '01' end AS TipoPrecio,
		case
			when ROUND(DD.nprecioneto,2)<>0 and ISNULL(MD.tCortesia,'')='' and ((ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=1) OR (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=0 AND isnull(DD.tAplicaExoneracion,0)=0)) then '30'  --Inafecto
			when ROUND(DD.nprecioneto,2)<>0 and ISNULL(MD.tCortesia,'')<>'' and ((ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=1) OR (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=0 AND isnull(DD.tAplicaExoneracion,0)=0)) then '31'  --Inafecto
			when ROUND(DD.nprecioneto,2)<>0 and (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaExoneracion,0)=1)  and ISNULL(MD.tCortesia,'')='' then '20' --Exonerado
			when ROUND(DD.nprecioneto,2)<>0 and (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaExoneracion,0)=1)  and ISNULL(MD.tCortesia,'')<>'' then '21' --Exonerado
			when ROUND(DD.nprecioneto,2)=0 then '15'
			when ROUND(DD.nImpuesto1,2)<>0 and ISNULL(MD.tCortesia,'')='' then '10'
			when ROUND(DD.nImpuesto1,2)<>0 and ISNULL(MD.tCortesia,'')<>'' then '15' 
			Else '10' 
		end AS TipoImpuesto,
		case 
			when ROUND(DD.nprecioneto,2)=0 then Cast(TP.Impuesto1  as decimal(30,5))
			else  Cast(round((DD.nPrecioImpuesto1/DD.nprecioneto)*100,2)   as decimal(30,5))
		end AS PorcentajeImpuesto,
		case 
			when ROUND(DD.nprecioneto,2)=0 then CAST(DD.nCantidad*(TP.Impuesto1/100) AS decimal(30,5))
			else CAST(DD.nImpuesto1 AS decimal(30,5))
		end  AS Impuesto,
		0.00 as ImpuestoSelectivo,
		CAST(ISNULL(DD.nLey1,0) AS decimal(30,5)) AS OtroImpuesto,
		CAST((ISNULL(DD.nLey1,0)/DD.nCantidad) AS decimal(30,5)) AS PrecioUnitarioBolsa,
		0.00 as Descuento,
		case 
			when ROUND(DD.nprecioneto,2)=0 then  CAST( DD.nCantidad+ ( DD.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,10))
			else CAST(DD.nPrecioNeto*DD.nCantidad AS decimal(30,10))
		end  AS TotalVenta,
		'' AS PlacaVehiculo,
		'' AS DetalleViaje,
		'' AS UbigeoOrigen,
		'' AS UbigeoDestino,
		'' AS DireccionOrigen,
		'' AS DireccionDestino,
		'90101500' AS CodigoProductoSunat,
		'' AS ConfiguracionVehicular,
		Case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(TT.nVenta AS decimal(30,5)) else 1.000 end TC
		from  MNOTACREDITO MN  
		INNER JOIN DNOTACREDITO DD on MN.tNotaCredito = DD.tNotaCredito
		left join MDOCUMENTO MD on MN.tDocumento= MD.tDocumento
		inner JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo 
		left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha) 
		LEFT JOIN vProducto vp ON  DD.tCodigoProducto=vp.Codigo
		, (select Impuesto1 from TPARAMETRO) TP 
	where vtp.lFacturacionE=1  and MN.nVenta < MD.nVenta-0.05 
		and  exists ( select * from DNOTACREDITO where tNotaCredito=MN.tNotaCredito) and ISNULL(MD.lAnticipo,0)=0

	) AS DOCUMENTO

GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoDetalleInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoDetalleInfoFact]
GO
CREATE  VIEW vDocumentoDetalleInfoFact
AS
Select tDocumento,	Id,	Cantidad,	UnidadMedida,	CodigoItem,	Descripcion,	PrecioUnitario/TC as PrecioUnitario,	PrecioReferencial/TC as PrecioReferencial,	TipoPrecio,	
TipoImpuesto,	PorcentajeImpuesto,	Impuesto/TC as Impuesto,	ImpuestoSelectivo,	OtroImpuesto,	PrecioUnitarioBolsa,	Descuento,	
TotalVenta/TC as TotalVenta,	PlacaVehiculo,	DetalleViaje,	UbigeoOrigen,	UbigeoDestino,	DireccionOrigen,	DireccionDestino,	CodigoProductoSunat,	ConfiguracionVehicular,
DescuentoInfoRest
 from (
	Select
	DD.tDocumento, 
	CONVERT(int, DD.tItem) AS Id,
	CAST(DD.nCantidad AS decimal(30,5)) AS Cantidad,
	'NIU' AS UnidadMedida,
	DD.tCodigoProducto AS CodigoItem,
	ltrim(rtrim(vp.Descripcion)) AS Descripcion,
	case when ROUND(DD.nprecioneto,2)=0 OR ISNULL(MD.tCortesia,'')<>'' then  
	CAST(0 AS decimal(30,10))  
	else CAST(DD.nPrecioNeto AS decimal(30,10)) end 
	AS PrecioUnitario,
	case 
		when ROUND(DD.nprecioneto,2)=0  then CAST(1 + (TP.Impuesto1/100) AS decimal(30,10)) 
		when ISNULL(MD.tCortesia,'')<>''  then CAST(DD.nPrecioNeto AS decimal(30,10)) 
		else CAST(DD.nPrecioNeto + (DD.nImpuesto1/DD.nCantidad) AS decimal(30,10)) end AS PrecioReferencial,
	CASE WHEN round(DD.nPrecioNeto,2)=0 OR ISNULL(MD.tCortesia,'')<>''then '02' else '01' end AS TipoPrecio,
	case
		--when ISNULL(MD.tCortesia,'')<>'' then '15' 
		when ROUND(DD.nprecioneto,2)>0 and ISNULL(MD.tCortesia,'')='' and ((ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=1) OR (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=0 AND isnull(DD.tAplicaExoneracion,0)=0)) then '30'  --Inafecto
		when ROUND(DD.nprecioneto,2)>0 and ISNULL(MD.tCortesia,'')<>'' and ((ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=1) OR (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaInafectacion,0)=0 AND isnull(DD.tAplicaExoneracion,0)=0)) then '31'  --Inafecto
		when ROUND(DD.nprecioneto,2)>0 and (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaExoneracion,0)=1)  and ISNULL(MD.tCortesia,'')='' then '20' --Exonerado
		when ROUND(DD.nprecioneto,2)>0 and (ROUND(DD.nImpuesto1,2)=0 and isnull(DD.tAplicaExoneracion,0)=1)  and ISNULL(MD.tCortesia,'')<>'' then '21' --Exonerado
		when ROUND(DD.nprecioneto,2)=0 then '15'
		when ROUND(DD.nImpuesto1,2)>0 and ISNULL(MD.tCortesia,'')='' then '10'
		when ROUND(DD.nImpuesto1,2)>0 and ISNULL(MD.tCortesia,'')<>'' then '15' 
		Else '10' 
	end AS TipoImpuesto,
	case 
		when ROUND(DD.nprecioneto,2)=0 then Cast(TP.Impuesto1  as decimal(30,5))
		else  Cast(round((DD.nPrecioImpuesto1/DD.nprecioneto)*100,2)   as decimal(30,5))
	end AS PorcentajeImpuesto,
	case 
		when ROUND(DD.nprecioneto,2)=0 then CAST(DD.nCantidad*(TP.Impuesto1/100) AS decimal(30,5))
		else CAST(DD.nImpuesto1 AS decimal(30,5))
	end  AS Impuesto,
	0.00 as ImpuestoSelectivo,
	CAST(ISNULL(DD.nLey1,0) AS decimal(30,5)) AS OtroImpuesto,
	CAST((ISNULL(DD.nLey1,0)/DD.nCantidad) AS decimal(30,5)) AS PrecioUnitarioBolsa,
	0.00 as Descuento,
	case 
		when ROUND(DD.nprecioneto,2)=0 then  
		CAST( DD.nCantidad+ ( DD.nCantidad*(TP.Impuesto1/100) ) AS decimal(30,10))
		else CAST(DD.nPrecioNeto as decimal (30,10))*CAST(DD.nCantidad AS decimal(30,5))
	end  AS TotalVenta,
	'' AS PlacaVehiculo,
	'' AS DetalleViaje,
	'' AS UbigeoOrigen,
	'' AS UbigeoDestino,
	'' AS DireccionOrigen,
	'' AS DireccionDestino,
	'90101500' AS CodigoProductoSunat,
	'' AS ConfiguracionVehicular,
	Case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast(TT.nVenta AS decimal(30,5)) else 1.000 end TC,
	CAST(DD.nDescuento*DD.nCantidad AS decimal(30,10)) DescuentoInfoRest
	from DDOCUMENTO DD 
	left join MDOCUMENTO MD on DD.tDocumento= MD.tDocumento
	inner JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo 
	left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha) 
	LEFT JOIN vProducto vp ON  DD.tCodigoProducto=vp.Codigo
	, (select Impuesto1 from TPARAMETRO) TP 
	Where  vtp.lFacturacionE=1  and ISNULL(DD.tDocumentoAnticipo,'')=''

) AS DOCUMENTO

GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoLeyendasInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoLeyendasInfoFact]
GO
Create View vDocumentoLeyendasInfoFact
As
	select 
		tDocumento,'1002' AS Codigo , 'TRANSFERENCIA GRATUITA DE UN BIEN Y/O SERVICIO PRESTADO GRATUITAMENTE' as Descripcion
	from MDOCUMENTO MD 
		inner JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo 
	where ISNULL(tCortesia,'')<>'' and vtp.lFacturacionE=1
	UNION ALL
	select 
		MN.tNotaCredito as  tDocumento,'1002' AS Codigo , 'TRANSFERENCIA GRATUITA DE UN BIEN Y/O SERVICIO PRESTADO GRATUITAMENTE' as Descripcion
	from MNOTACREDITO MN  
		INNER JOIN MDOCUMENTO MD on MN.tDocumento = MD.tDocumento
		inner JOIN vTipoDocumento vtp on MN.tTipoDocumento= vtp.Codigo 
	where ISNULL(MD.tCortesia,'')<>'' and vtp.lFacturacionE=1
	UNION ALL
	select 
		tDocumento,'2006' AS Codigo , 'Operacion sujeta a detraccion' as Descripcion
	from MDOCUMENTO MD 
		inner JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo ,
		(
			select top 1 tIdentificacionTributaria, tRazonSocial, tRazonComercial, 
						tCodigoUbigeo,tDireccion, tDireccion2, tCodigoTienda,TokenInfoFact as Token, tEmail as Correo, 
						isnull(TPARAMETRO.lDetraccion,0) lDetraccion, isnull(TPARAMETRO.lMontoDetraccion,0.00) nMontoDetraccion, 
						isnull(TPARAMETRO.lPorcentajeDetraccion,0) nPorcentajeDetracccion, isnull(TPARAMETRO.tMsjDetraccion,'') Detraccion
						
			from TPARAMETRO 
		) TP
	where vtp.lFacturacionE=1 and TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01'
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoCreditosInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoCreditosInfoFact]
GO
Create  View vDocumentoCreditosInfoFact
As
	select 
		tDocumento, 1 AS NroCuota, 
		case 
			when MD.tEstadoDocumento ='03' and ISNULL(MD.tClientePago,'')<>'' then (
			case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((MD.nVenta - (
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			) -0.001)/TT.nVenta as decimal(30,5))  else Cast(MD.nVenta - (
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			)  -0.001 as decimal(30,5))  end) else 0.00 end AS MontoCuota,
			case when ISNULL(vd.nDiasCredito,0)=0 then CONVERT(NVARCHAR(10), MD.fregistro + 1, 120) else CONVERT(NVARCHAR(10), MD.fregistro + isnull(vd.nDiasCredito,0) , 120) end  as FechaCredito, 
			CONVERT(NVARCHAR(10), MD.fregistro , 120) as FechaRegistro
	from MDOCUMENTO MD 
		INNER JOIN vCliente vc on MD.tCodigoCliente = vc.Codigo
		inner JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo
		left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha) 
		LEFT JOIN vDelivery vd on MD.tClientePago = vd.Codigo,
		(
			select top 1 tIdentificacionTributaria, tRazonSocial, tRazonComercial, 
						tCodigoUbigeo,tDireccion, tDireccion2, tCodigoTienda,TokenInfoFact as Token, tEmail as Correo, 
						isnull(lDetraccion,0) lDetraccion, isnull(lMontoDetraccion,0.00) nMontoDetraccion, 
						isnull(lPorcentajeDetraccion,0) nPorcentajeDetracccion, isnull(tMsjDetraccion,'') Detraccion,
						 isnull(convert(float,isnull(tMontoRetencion,0)),0.00) nMontoRetencion, 
						isnull(convert(float,isnull(lRetencion,0)),0.00) nPorcentajeRetencion
			from TPARAMETRO
		) TP
	where MD.tEstadoDocumento='03' and vtp.lFacturacionE=1 and ISNULL(MD.tClientePago,'')<>''
	UNION ALL
	select 
		MN.tNotaCredito as tDocumento, 1 AS NroCuota, 
		case 
			when MD.tEstadoDocumento ='03' and ISNULL(MD.tClientePago,'')<>'' then (
			case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then Cast((MD.nVenta -(
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			) -0.001)/TT.nVenta as decimal(30,5))  else Cast(MD.nVenta- (
				Case 
					when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeDetracccion/100)
					when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then MD.nVenta* (TP.nPorcentajeRetencion/100)
					else 0.00 end
			)  -0.001 as decimal(30,5))  end)else 0.00 end AS MontoCuota,
			case when ISNULL(vd.nDiasCredito,0)=0 then CONVERT(NVARCHAR(10), MD.fregistro + 1, 120) else CONVERT(NVARCHAR(10), MD.fregistro +  isnull(vd.nDiasCredito,0) , 120) end  as FechaCredito, 
			CONVERT(NVARCHAR(10), MD.fregistro , 120) as FechaRegistro
	from MNOTACREDITO MN  
		INNER JOIN MDOCUMENTO MD on MN.tDocumento = MD.tDocumento
		INNER JOIN vCliente vc on MD.tCodigoCliente = vc.Codigo
		inner JOIN vTipoDocumento vtp on MN.tTipoDocumento= vtp.Codigo
		left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha) 
		LEFT JOIN vDelivery vd on MD.tClientePago = vd.Codigo,
		(
			select top 1 tIdentificacionTributaria, tRazonSocial, tRazonComercial, 
						tCodigoUbigeo,tDireccion, tDireccion2, tCodigoTienda,TokenInfoFact as Token, tEmail as Correo, 
						isnull(lDetraccion,0) lDetraccion, isnull(lMontoDetraccion,0.00) nMontoDetraccion, 
						isnull(lPorcentajeDetraccion,0) nPorcentajeDetracccion, isnull(tMsjDetraccion,'') Detraccion,
						 isnull(convert(float,isnull(tMontoRetencion,0)),0.00) nMontoRetencion, 
						isnull(convert(float,isnull(lRetencion,0)),0.00) nPorcentajeRetencion
			from TPARAMETRO
		) TP
	where MD.tEstadoDocumento='03' and vtp.lFacturacionE=1 and ISNULL(MD.tClientePago,'')<>''

GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vBajasDetalleInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vBajasDetalleInfoFact]
GO
Create View vBajasDetalleInfoFact
AS
	select tDocumento,	fRegistro,	DocSunat,	EstadoInfoFact,	NroTicketCdr,	ZipCdr,	mensajes,	IdInfoFact 
	from TESTADOINFOFACT WHERE tDocumento LIKE 'RA%' OR tDocumento LIKE 'RC%'
	--where LEFT(tDocumento,2) in ('RA', 'RC')
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vBajaFacturaInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vBajaFacturaInfoFact]
GO
CREATE VIEW vBajaFacturaInfoFact
as
SELECT 
	'RA-'+CONVERT(NVARCHAR(8),FechaAnulado, 112) + '-'+ right('0000'+RIGHT(Emi_CodigoAnexo,2)+ convert(nvarchar(4),DATEPART(DAY, FechaRegistro) ) ,4) AS idDocumento,
	fechaEmision,	fechaReferencia,	Emi_NroDocumento,	Emi_TipoDocumento,	Emi_NombreLegal,	
	Emi_NombreComercial,	Emi_CodigoAnexo,	
	ROW_NUMBER() OVER(PARTITION BY 'RA-'+CONVERT(NVARCHAR(8),FechaAnulado, 112) + '-'+ RIGHT(Emi_CodigoAnexo,2)+ convert(nvarchar(4),DATEPART(DAY, FechaRegistro) )  ORDER BY BAJAS.tdocumento+BAJAS.tipoDocumento ASC) AS id,
	tipoDocumento,	serie,	correlativo,	motivoBaja,	BAJAS.tdocumento,	Token , convert(date,FechaAnulado) AS FechaAnulado,
	convert(bit,(CASE WHEN isnull(FACT.tDocumento,'') = '' then 0 else 1 end )) lEstadoFacturacion, ISNULL(FACT.EstadoInfoFact,'') as RespuestaFacturacion
	FROM 
	(
		select  MD.tDocumento,MD.FechaAnulado,  MD.FechaRegistro, MD.Emi_CodigoAnexo,			
			CONVERT(NVARCHAR(10), MD.FechaAnulado, 120) AS fechaEmision,
			CONVERT(NVARCHAR(10), MD.FechaRegistro, 120) AS fechaReferencia, 
			MD.Emi_NroDocumento,MD.Emi_TipoDocumento, MD.Emi_NombreLegal,
			MD.Emi_NombreComercial, MD.tipoDocumento, 
			LEFT(MD.tdocumento, 1) + LEFT(RIGHT(MD.tdocumento, 12), 3)  AS serie,
			RIGHT(MD.tdocumento, 8) AS correlativo,
			MD.motivoBaja, MD.Token
		from vDocumentoCabeceraInfoFact MD
		where  MD.EstadoDocumento ='04' and LEFT(MD.tDocumento,1)='F' and  ISNULL(MD.lEstadoFacturacion,0) =1
		UNION ALL
		select MN.tDocumento,MN.FechaAnulado,  MN.FechaRegistro, MN.Emi_CodigoAnexo,			
			CONVERT(NVARCHAR(10), MN.FechaAnulado, 120) AS fechaEmision,
			CONVERT(NVARCHAR(10), MN.FechaRegistro, 120) AS fechaReferencia, 
			MN.Emi_NroDocumento,MN.Emi_TipoDocumento, MN.Emi_NombreLegal,
			MN.Emi_NombreComercial, MN.tipoDocumento, 
			LEFT(MN.IdDocumento, 1) + LEFT(RIGHT(MN.tDocumento, 12), 3)  AS serie,
			RIGHT(MN.tdocumento, 8) AS correlativo,
			MN.motivoBaja, MN.Token
		from vNotaCreditoCabeceraInfoFact MN 
		where   MN.EstadoDocumento ='04' and LEFT(MN.IdDocumento, 1)='F' and  ISNULL(MN.lEstadoFacturacion,0) =1
	) AS BAJAS LEFT JOIN vBajasDetalleInfoFact FACT on 'RA-'+CONVERT(NVARCHAR(8),FechaAnulado, 112) + '-'+ RIGHT(Emi_CodigoAnexo,2)+ convert(nvarchar(4),DATEPART(DAY, FechaRegistro) ) =FACT.TDOCUMENTO  
	where convert(date,FechaAnulado) <= convert(date,GETDATE()-1)

GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vBajaBoletaInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vBajaBoletaInfoFact]
GO
CREATE VIEW vBajaBoletaInfoFact
as
	SELECT 
		'RC-'+CONVERT(NVARCHAR(8),FechaAnulado, 112) + '-'+ right('0000'+RIGHT(Emi_CodigoAnexo,2)+ convert(nvarchar(4),DATEPART(DAY, FechaRegistro) ),4) AS idDocumento,
		fechaEmision,	fechaReferencia,	Emi_NroDocumento,	Emi_TipoDocumento,	Emi_NombreLegal,	
		Emi_NombreComercial,	Emi_CodigoAnexo,	
		ROW_NUMBER() OVER(PARTITION BY 'RC-'+CONVERT(NVARCHAR(8),FechaAnulado, 112) + '-'+ RIGHT(Emi_CodigoAnexo,2)+ convert(nvarchar(4),DATEPART(DAY, FechaRegistro) )  ORDER BY bajas.tdocumento+bajas.tipoDocumento ASC) AS id,
		tipoDocumento,	serie,	correlativo as correlativoInicio, correlativo as correlativoFin,
		moneda,  totalVenta,  totalDescuentos, totalIgv,	totalIsc,totalOtrosImpuestos,totalImpuestoBolsas,
		gravadas,  exoneradas,  inafectas, exportacion,	gratuitas,   idDocumento as Res_idDocumento, tipoDocumentoReceptor,
		nroDocumentoReceptor ,  codigoEstadoItem,  documentoRelacionado,  tipoDocumentoRelacionado,
		motivoBaja,	bajas.tdocumento,	Token , convert(date,FechaAnulado) AS FechaAnulado,
		convert(bit,(CASE WHEN isnull(FACT.tDocumento,'') = '' then 0 else 1 end )) lEstadoFacturacion, ISNULL(FACT.EstadoInfoFact,'') as RespuestaFacturacion
	FROM 
	(
		select  MD.tDocumento,MD.FechaAnulado,  MD.FechaRegistro, MD.Emi_CodigoAnexo,			
			CONVERT(NVARCHAR(10), MD.FechaAnulado, 120) AS fechaEmision,
			CONVERT(NVARCHAR(10), MD.FechaRegistro, 120) AS fechaReferencia, 
			MD.Emi_NroDocumento,MD.Emi_TipoDocumento, MD.Emi_NombreLegal,
			MD.Emi_NombreComercial, MD.tipoDocumento, 
			LEFT(MD.tdocumento, 1) + LEFT(RIGHT(MD.tdocumento, 12), 3)  AS serie,
			CONVERT(nvarchar(8),convert(bigint,RIGHT(MD.tdocumento, 8))) AS correlativo,
			MD.Moneda as moneda, MD.TotalVenta as totalVenta, 0 totalDescuentos, MD.TotalIgv as totalIgv,
			0 totalIsc, MD.TotalRecargo as totalOtrosImpuestos, MD.TotalOtrosTributos  as totalImpuestoBolsas,
			MD.Gravadas as gravadas, MD.Exoneradas as exoneradas, MD.Inafectas as inafectas, MD.Exportacion as exportacion,
			MD.Gratuitas as gratuitas, MD.IdDocumento as  idDocumento, MD.Rec_TipoDocumento as tipoDocumentoReceptor,
			MD.Rec_NroDocumento as nroDocumentoReceptor , 3  as codigoEstadoItem, '' as documentoRelacionado, '' as tipoDocumentoRelacionado,
			MD.motivoBaja, MD.Token
		from vDocumentoCabeceraInfoFact MD
		where  MD.EstadoDocumento ='04' and LEFT(MD.tDocumento,1)='B' and  ISNULL(MD.lEstadoFacturacion,0) =1
		UNION ALL
		select MN.tDocumento,MN.FechaAnulado,  MN.FechaRegistro, MN.Emi_CodigoAnexo,			
			CONVERT(NVARCHAR(10), MN.FechaAnulado, 120) AS fechaEmision,
			CONVERT(NVARCHAR(10), MN.FechaRegistro, 120) AS fechaReferencia, 
			MN.Emi_NroDocumento,MN.Emi_TipoDocumento, MN.Emi_NombreLegal,
			MN.Emi_NombreComercial, MN.tipoDocumento, 
			LEFT(MN.IdDocumento, 1) + LEFT(RIGHT(MN.tDocumento, 12), 3)  AS serie,
			CONVERT(nvarchar(8),convert(bigint,RIGHT(MN.tdocumento, 8))) AS correlativo,
			MN.Moneda as moneda, MN.TotalVenta as totalVenta, 0 totalDescuentos, MN.TotalIgv as totalIgv,
			0 totalIsc, MN.TotalRecargo as totalOtrosImpuestos, MN.TotalOtrosTributos  as totalImpuestoBolsas,
			MN.Gravadas as gravadas, MN.Exoneradas as exoneradas, MN.Inafectas as inafectas, MN.Exportacion as exportacion,
			MN.Gratuitas as gratuitas, MN.IdDocumento as  idDocumento, MN.Rec_TipoDocumento as tipoDocumentoReceptor,
			MN.Rec_NroDocumento as nroDocumentoReceptor , 3  as codigoEstadoItem, MN.NroReferencia as documentoRelacionado, MN.TipoDocumentoReferencia as tipoDocumentoRelacionado,
			MN.motivoBaja, MN.Token
		from vNotaCreditoCabeceraInfoFact MN 
		where   MN.EstadoDocumento ='04' and LEFT(MN.IdDocumento, 1)='B' and ISNULL(MN.lEstadoFacturacion,0) =1
	) AS BAJAS 
	LEFT JOIN vBajasDetalleInfoFact FACT on'RC-'+CONVERT(NVARCHAR(8),FechaAnulado, 112) + '-'+ RIGHT(Emi_CodigoAnexo,2)+ convert(nvarchar(4),DATEPART(DAY, FechaRegistro) ) =FACT.TDOCUMENTO   
	where  convert(date,FechaAnulado) <= convert(date,GETDATE()-1) 
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDatosAdicionalesInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDatosAdicionalesInfoFact]
GO
CREATE VIEW vDatosAdicionalesInfoFact
AS
	select 
		MD.tDocumento, TP.tDireccion2 as ExtraDireccionEmisor,
		TP.tTelefono as ExtraTelefonoEmisor, TP.Correo ExtraEmailEmisor, TP.tWebPage ExtraWebEmisor,
		convert(bit,Case 
			when TP.lDetraccion =1 and MD.nVenta>= TP.nMontoDetraccion and  isnull(vtp.Sunat,'00') ='01' then 0
			when ISNULL(vc.AgenteRetencion,0) =1 and MD.nVenta>= TP.nMontoRetencion and  isnull(vtp.Sunat,'00') ='01' then 1
			else 0 end) IsRetencion,
			Case when round(isnull(MD.nDescuento,0),2)=0 then ''
			else isnull((select Descripcion from vMotivoDescuento where Codigo=MD.tDescuento),'') +' ' + CAst( CAST( isnull(nDescuento,0.00)  as decimal(30,2)) AS nvarchar(20))
			end as Descuento
	from MDOCUMENTO MD 
		INNER JOIN vCliente vc on MD.tCodigoCliente = vc.Codigo
		inner JOIN vTipoDocumento vtp on MD.tTipoDocumento= vtp.Codigo,
		(
			select top 1 tIdentificacionTributaria, tRazonSocial, tRazonComercial, isnull(tTelefono,'') tTelefono, isnull(tWebPage,'') as tWebPage,
						tCodigoUbigeo,tDireccion, isnull(tDireccion2,'') tDireccion2, tCodigoTienda,TokenInfoFact as Token, isnull(tEmail,'') as Correo, 
						isnull(lDetraccion,0) lDetraccion, isnull(lMontoDetraccion,0.00) nMontoDetraccion, 
						isnull(lPorcentajeDetraccion,0) nPorcentajeDetracccion, isnull(tMsjDetraccion,'') Detraccion,
						 isnull(convert(float,isnull(tMontoRetencion,0)),0.00) nMontoRetencion, 
						isnull(convert(float,isnull(lRetencion,0)),0.00) nPorcentajeRetencion
			from TPARAMETRO
		) TP
		where vtp.lFacturacionE=1 
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vDocumentoAnticiposInfoFact]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vDocumentoAnticiposInfoFact]
GO
CREATE VIEW [dbo].[vDocumentoAnticiposInfoFact] 
AS
	Select 
		DD.tDocumento,
		DD.tDocumentoAnticipo,
		case
			when left(DD.tDocumentoAnticipo,1)='F' then '02'
			when left(DD.tDocumentoAnticipo,1)='B' then '03'
			else '99'
		end as TipoDocAnticipo, -- Catalogo 12 sunat
		LEFT(DD.tDocumentoAnticipo, 1) + LEFT(RIGHT(DD.tDocumentoAnticipo, 12), 3) + '-' + RIGHT(DD.tDocumentoAnticipo, 8)  as DocAnticipo ,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then 'USD' else 'PEN' end AS MonedaAnticipo,
		cast(ABS((DD.nPrecioNeto + DD.nPrecioImpuesto1)*DD.nCantidad ) as decimal(30,5))  as MontoAnticipo,
		cast(ABS((DD.nPrecioNeto )*DD.nCantidad ) as decimal(30,5)) as AnticipoNeto,
		cast(ABS((DD.nPrecioImpuesto1 )*DD.nCantidad ) as decimal(30,5)) as IgvAnticipo,
		cast(ABS((DD.nPrecioImpuesto2 + DD.nPrecioImpuesto3 )*DD.nCantidad ) as decimal(30,5)) as OtroImpuestoAnticipo
	from DDOCUMENTO DD LEFT join MDOCUMENTO MD ON DD.tDocumentoAnticipo= MD.tDocumento where ISNULL(DD.tDocumentoAnticipo,'')<>'' 

GO

--- Modificacion, validacion emision en dolares - cliente tacamas

ALTER VIEW [dbo].[vDocumentoAnticiposInfoFact] 
AS
	Select 
		DD.tDocumento,
		DD.tDocumentoAnticipo,
		case
			when left(DD.tDocumentoAnticipo,1)='F' then '02'
			when left(DD.tDocumentoAnticipo,1)='B' then '03'
			else '99'
		end as TipoDocAnticipo, -- Catalogo 12 sunat
		LEFT(DD.tDocumentoAnticipo, 1) + LEFT(RIGHT(DD.tDocumentoAnticipo, 12), 3) + '-' + RIGHT(DD.tDocumentoAnticipo, 8)  as DocAnticipo ,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then 'USD' else 'PEN' end AS MonedaAnticipo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then cast(ABS( (DD.nPrecioNeto + DD.nPrecioImpuesto1)*DD.nCantidad )/TT.nVenta as decimal(30,5))  else cast(ABS((DD.nPrecioNeto + DD.nPrecioImpuesto1)*DD.nCantidad ) as decimal(30,5)) end  as MontoAnticipo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then cast(ABS((DD.nPrecioNeto )*DD.nCantidad )/TT.nVenta as decimal(30,5))  else cast(ABS((DD.nPrecioNeto )*DD.nCantidad ) as decimal(30,5))  end as AnticipoNeto,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then cast(ABS((DD.nPrecioImpuesto1 )*DD.nCantidad )/TT.nVenta as decimal(30,5)) else cast(ABS((DD.nPrecioImpuesto1 )*DD.nCantidad ) as decimal(30,5)) end as IgvAnticipo,
		case when ISNULL(MD.lImpresionMonedaExtranjera,0)=1 then cast(ABS((DD.nPrecioImpuesto2 + DD.nPrecioImpuesto3 )*DD.nCantidad )/TT.nVenta as decimal(30,5)) else cast(ABS((DD.nPrecioImpuesto2 + DD.nPrecioImpuesto3 )*DD.nCantidad ) as decimal(30,5)) end as OtroImpuestoAnticipo
	from DDOCUMENTO DD 
	LEFT join MDOCUMENTO MD ON DD.tDocumentoAnticipo= MD.tDocumento 
	left join (select fFecha, nVenta from TTIPOCAMBIO) TT on CONVERT(date,MD.fRegistro) = CONVERT(date,TT.fFecha) 
	where ISNULL(DD.tDocumentoAnticipo,'')<>'' 


GO


update  TCLIENTE set lActivo=0 where len(ltrim(rtrim(tEmpresa)))<2 or len(ltrim(rtrim(tIdentidad)))<2
--****************************************************************INFOFACT FIN **********************************************
GO
