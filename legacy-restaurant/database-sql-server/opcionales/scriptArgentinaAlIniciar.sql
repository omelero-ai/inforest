
UPDATE TTABLA SET nValor=0  WHERE ttabla='PAISORIGEN'
UPDATE TTABLA SET nValor=1  WHERE ttabla='PAISORIGEN' AND TCODIGO='003'

DELETE FROM TTABLA WHERE TTABLA='MONEDA'
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) 
VALUES('MONEDA', '01',2,'MONEDA NACIONAL','$',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) 
VALUES('MONEDA', '02',2,'MONEDA EXTRANJERA','US$',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) 
VALUES('MONEDA', '03',2,'MONEDA EXTRANJERA2','€',1)
insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) 
VALUES('MONEDA', '04',2,'MONEDA EXTRANJERA3','R$',1)


UPDATE TPARAMETRO SET tMonedaN='Pesos' ,tMonN='$', tMonedaE='Dolares', tMonE='US$', tMonedaE2='Euros', tMonE2='€', tMonedaE3='Reales', tMonE3='R$'

GO
DELETE TTABLA WHERE TTABLA='TIPOIDENTIDAD'
GO
IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOIDENTIDAD'AND TCODIGO='01')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPOIDENTIDAD',	'01',	NULL, 'CUIT',	'CUIT',	NULL,	NULL,	'',	1,	1)

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOIDENTIDAD'AND TCODIGO='02')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPOIDENTIDAD',	'02',	NULL, 'CUIL',	'CUIL',	NULL,	NULL,	'',	1,	1)

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOIDENTIDAD'AND TCODIGO='03')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPOIDENTIDAD',	'03',	NULL, 'CDI',	'CDI',	NULL,	NULL,	'',	0,	1)

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOIDENTIDAD'AND TCODIGO='04')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPOIDENTIDAD',	'04',	NULL, 'DNI',	'DNI',	NULL,	NULL,	'',	0,	1)

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOIDENTIDAD'AND TCODIGO='05')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPOIDENTIDAD',	'05',	NULL, 'CI Extranjera',	'CI Extranjera',	NULL,	NULL,	'',	0,	1)

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOIDENTIDAD'AND TCODIGO='06')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPOIDENTIDAD',	'06',	NULL, 'Pasaporte',	'Pasaporte',	NULL,	NULL,	'',	0,	1)

GO
DELETE TTABLA WHERE TTABLA='TIPORESPONSABLE'
GO

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='01')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPORESPONSABLE',	'01',	NULL, 'Responsable Inscripto',	'Responsable Inscripto',	NULL,	NULL,	'I',	0,	1)

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='02')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPORESPONSABLE',	'02',	NULL, 'Res. no inscripto',	'Res. no inscripto',	NULL,	NULL,	'R',	0,	1)

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='03')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPORESPONSABLE',	'03',	NULL, 'Exento',	'Exento',	NULL,	NULL,	'E',	0,	1)

--IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='04')=0
--INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
--VALUES              ('TIPORESPONSABLE',	'04',	NULL, 'No responsable',	'No responsable',	NULL,	NULL,	'N',	0,	0)

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='05')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPORESPONSABLE',	'05',	NULL, 'Consumidor Final',	'Consumidor Final',	NULL,	NULL,	'F',	0,	1)

--IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='06')=0
--INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
--VALUES              ('TIPORESPONSABLE',	'06',	NULL, 'Resp. no inscripto, vta. bienes de uso',	'Resp. no inscripto',	NULL,	NULL,	'B',	0, 0)

--IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='07')=0
--INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
--VALUES              ('TIPORESPONSABLE',	'07',	NULL, 'Resp. monotributo',	'Resp. monotributo',	NULL,	NULL,	'M',	0,	1) 

--IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='08')=0
--INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
--VALUES              ('TIPORESPONSABLE',	'08',	NULL, 'Monotributista social',	'Monotributista social',	NULL,	NULL,	'S',	0,	0) 

--IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='09')=0
--INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
--VALUES              ('TIPORESPONSABLE',	'09',	NULL, 'Pequeño contribuyente eventual',	'Pequeño contribuyente e',	NULL,	NULL,	'W',	0,	0) 

--IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='10')=0
--INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
--VALUES              ('TIPORESPONSABLE',	'10',	NULL, 'Pequeño contribuyente eventual social',	'Pequeño contribuyente e',	NULL,	NULL,	'V',	0,	0) 

IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPORESPONSABLE'AND TCODIGO='11')=0
INSERT INTO TTABLA	(TTABLA	,TCODIGO,	nTamano,	tDetallado,	tResumido,	tIcono,	nBoton,	tValor,	nValor,	lActivo)
VALUES              ('TIPORESPONSABLE',	'11',	NULL, 'No categorizado',	'No categorizado',	NULL,	NULL,	'S',	0,	1) 
