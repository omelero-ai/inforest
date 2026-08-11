	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MGUIA]') and OBJECTPROPERTY(id, N'IsTable') = 1)
	   DROP TABLE MGUIA
GO
	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DGUIA]') and OBJECTPROPERTY(id, N'IsTable') = 1)
	   DROP TABLE DGUIA
GO
		truncate table tubigeo
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010101','CHACHAPOYAS','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010102','ASUNCION','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010103','BALSAS','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010104','CHETO','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010105','CHILIQUIN','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010106','CHUQUIBAMBA','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010107','GRANADA','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010108','HUANCAS','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010109','LA JALCA','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010110','LEIMEBAMBA','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010111','LEVANTO','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010112','MAGDALENA','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010113','MARISCAL CASTILLA','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010114','MOLINOPAMPA','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010115','MONTEVIDEO','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010116','OLLEROS','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010117','QUINJALCA','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010118','SAN FRANCISCO DE DAGUAS','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010119','SAN ISIDRO DE MAINO','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010120','SOLOCO','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010121','SONCHE','CHACHAPOYAS','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010201','BAGUA','BAGUA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010202','ARAMANGO','BAGUA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010203','COPALLIN','BAGUA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010204','EL PARCO','BAGUA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010205','IMAZA','BAGUA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010206','LA PECA','BAGUA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010301','JUMBILLA','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010302','CHISQUILLA','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010303','CHURUJA','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010304','COROSHA','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010305','CUISPES','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010306','FLORIDA','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010307','JAZAN','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010308','RECTA','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010309','SAN CARLOS','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010310','SHIPASBAMBA','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010311','VALERA','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010312','YAMBRASBAMBA','BONGARA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010401','NIEVA','CONDORCANQUI','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010402','EL CENEPA','CONDORCANQUI','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010403','RIO SANTIAGO','CONDORCANQUI','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010501','LAMUD','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010502','CAMPORREDONDO','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010503','COCABAMBA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010504','COLCAMAR','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010505','CONILA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010506','INGUILPATA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010507','LONGUITA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010508','LONYA CHICO','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010509','LUYA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010510','LUYA VIEJO','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010511','MARIA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010512','OCALLI','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010513','OCUMAL','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010514','PISUQUIA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010515','PROVIDENCIA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010516','SAN CRISTOBAL','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010517','SAN FRANCISCO DEL YESO','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010518','SAN JERONIMO','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010519','SAN JUAN DE LOPECANCHA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010520','SANTA CATALINA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010521','SANTO TOMAS','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010522','TINGO','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010523','TRITA','LUYA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010601','SAN NICOLAS','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010602','CHIRIMOTO','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010603','COCHAMAL','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010604','HUAMBO','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010605','LIMABAMBA','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010606','LONGAR','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010607','MARISCAL BENAVIDES','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010608','MILPUC','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010609','OMIA','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010610','SANTA ROSA','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010611','TOTORA','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010612','VISTA ALEGRE','RODRIGUEZ DE MENDOZA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010701','BAGUA GRANDE','UTCUBAMBA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010702','CAJARURO','UTCUBAMBA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010703','CUMBA','UTCUBAMBA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010704','EL MILAGRO','UTCUBAMBA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010705','JAMALCA','UTCUBAMBA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010706','LONYA GRANDE','UTCUBAMBA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('010707','YAMON','UTCUBAMBA','AMAZONAS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020101','HUARAZ','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020102','COCHABAMBA','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020103','COLCABAMBA','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020104','HUANCHAY','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020105','INDEPENDENCIA','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020106','JANGAS','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020107','LA LIBERTAD','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020108','OLLEROS','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020109','PAMPAS','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020110','PARIACOTO','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020111','PIRA','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020112','TARICA','HUARAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020201','AIJA','AIJA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020202','CORIS','AIJA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020203','HUACLLAN','AIJA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020204','LA MERCED','AIJA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020205','SUCCHA','AIJA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020301','LLAMELLIN','ANTONIO RAYMONDI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020302','ACZO','ANTONIO RAYMONDI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020303','CHACCHO','ANTONIO RAYMONDI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020304','CHINGAS','ANTONIO RAYMONDI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020305','MIRGAS','ANTONIO RAYMONDI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020306','SAN JUAN DE RONTOY','ANTONIO RAYMONDI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020401','CHACAS','ASUNCION','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020402','ACOCHACA','ASUNCION','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020501','CHIQUIAN','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020502','ABELARDO PARDO LEZAMETA','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020503','ANTONIO RAYMONDI','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020504','AQUIA','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020505','CAJACAY','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020506','CANIS','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020507','COLQUIOC','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020508','HUALLANCA','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020509','HUASTA','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020510','HUAYLLACAYAN','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020511','LA PRIMAVERA','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020512','MANGAS','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020513','PACLLON','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020514','SAN MIGUEL DE CORPANQUI','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020515','TICLLOS','BOLOGNESI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020601','CARHUAZ','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020602','ACOPAMPA','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020603','AMASHCA','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020604','ANTA','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020605','ATAQUERO','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020606','MARCARA','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020607','PARIAHUANCA','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020608','SAN MIGUEL DE ACO','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020609','SHILLA','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020610','TINCO','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020611','YUNGAR','CARHUAZ','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020701','SAN LUIS','CARLOS FERMIN FITZCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020702','SAN NICOLAS','CARLOS FERMIN FITZCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020703','YAUYA','CARLOS FERMIN FITZCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020801','CASMA','CASMA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020802','BUENA VISTA ALTA','CASMA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020803','COMANDANTE NOEL','CASMA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020804','YAUTAN','CASMA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020901','CORONGO','CORONGO','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020902','ACO','CORONGO','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020903','BAMBAS','CORONGO','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020904','CUSCA','CORONGO','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020905','LA PAMPA','CORONGO','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020906','YANAC','CORONGO','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('020907','YUPAN','CORONGO','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021001','HUARI','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021002','ANRA','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021003','CAJAY','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021004','CHAVIN DE HUANTAR','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021005','HUACACHI','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021006','HUACCHIS','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021007','HUACHIS','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021008','HUANTAR','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021009','MASIN','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021010','PAUCAS','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021011','PONTO','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021012','RAHUAPAMPA','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021013','RAPAYAN','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021014','SAN MARCOS','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021015','SAN PEDRO DE CHANA','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021016','UCO','HUARI','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021101','HUARMEY','HUARMEY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021102','COCHAPETI','HUARMEY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021103','CULEBRAS','HUARMEY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021104','HUAYAN','HUARMEY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021105','MALVAS','HUARMEY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021201','CARAZ','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021202','HUALLANCA','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021203','HUATA','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021204','HUAYLAS','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021205','MATO','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021206','PAMPAROMAS','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021207','PUEBLO LIBRE','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021208','SANTA CRUZ','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021209','SANTO TORIBIO','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021210','YURACMARCA','HUAYLAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021301','PISCOBAMBA','MARISCAL LUZURIAGA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021302','CASCA','MARISCAL LUZURIAGA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021303','ELEAZAR GUZMAN BARRON','MARISCAL LUZURIAGA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021304','FIDEL OLIVAS ESCUDERO','MARISCAL LUZURIAGA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021305','LLAMA','MARISCAL LUZURIAGA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021306','LLUMPA','MARISCAL LUZURIAGA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021307','LUCMA','MARISCAL LUZURIAGA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021308','MUSGA','MARISCAL LUZURIAGA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021401','OCROS','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021402','ACAS','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021403','CAJAMARQUILLA','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021404','CARHUAPAMPA','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021405','COCHAS','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021406','CONGAS','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021407','LLIPA','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021408','SAN CRISTOBAL DE RAJAN','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021409','SAN PEDRO','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021410','SANTIAGO DE CHILCAS','OCROS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021501','CABANA','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021502','BOLOGNESI','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021503','CONCHUCOS','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021504','HUACASCHUQUE','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021505','HUANDOVAL','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021506','LACABAMBA','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021507','LLAPO','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021508','PALLASCA','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021509','PAMPAS','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021510','SANTA ROSA','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021511','TAUCA','PALLASCA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021601','POMABAMBA','POMABAMBA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021602','HUAYLLAN','POMABAMBA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021603','PAROBAMBA','POMABAMBA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021604','QUINUABAMBA','POMABAMBA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021701','RECUAY','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021702','CATAC','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021703','COTAPARACO','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021704','HUAYLLAPAMPA','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021705','LLACLLIN','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021706','MARCA','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021707','PAMPAS CHICO','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021708','PARARIN','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021709','TAPACOCHA','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021710','TICAPAMPA','RECUAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021801','CHIMBOTE','SANTA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021802','CACERES DEL PERU','SANTA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021803','COISHCO','SANTA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021804','MACATE','SANTA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021805','MORO','SANTA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021806','NEPEÑA','SANTA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021807','SAMANCO','SANTA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021808','SANTA','SANTA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021809','NUEVO CHIMBOTE','SANTA','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021901','SIHUAS','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021902','ACOBAMBA','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021903','ALFONSO UGARTE','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021904','CASHAPAMPA','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021905','CHINGALPO','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021906','HUAYLLABAMBA','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021907','QUICHES','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021908','RAGASH','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021909','SAN JUAN','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('021910','SICSIBAMBA','SIHUAS','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('022001','YUNGAY','YUNGAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('022002','CASCAPARA','YUNGAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('022003','MANCOS','YUNGAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('022004','MATACOTO','YUNGAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('022005','QUILLO','YUNGAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('022006','RANRAHIRCA','YUNGAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('022007','SHUPLUY','YUNGAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('022008','YANAMA','YUNGAY','ANCASH')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030101','ABANCAY','ABANCAY','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030102','CHACOCHE','ABANCAY','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030103','CIRCA','ABANCAY','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030104','CURAHUASI','ABANCAY','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030105','HUANIPACA','ABANCAY','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030106','LAMBRAMA','ABANCAY','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030107','PICHIRHUA','ABANCAY','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030108','SAN PEDRO DE CACHORA','ABANCAY','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030109','TAMBURCO','ABANCAY','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030201','ANDAHUAYLAS','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030202','ANDARAPA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030203','CHIARA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030204','HUANCARAMA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030205','HUANCARAY','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030206','HUAYANA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030207','KISHUARA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030208','PACOBAMBA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030209','PACUCHA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030210','PAMPACHIRI','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030211','POMACOCHA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030212','SAN ANTONIO DE CACHI','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030213','SAN JERONIMO','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030214','SAN MIGUEL DE CHACCRAMPA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030215','SANTA MARIA DE CHICMO','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030216','TALAVERA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030217','TUMAY HUARACA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030218','TURPO','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030219','KAQUIABAMBA','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030220','JOSÉ MARÍA ARGUEDAS','ANDAHUAYLAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030301','ANTABAMBA','ANTABAMBA','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030302','EL ORO','ANTABAMBA','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030303','HUAQUIRCA','ANTABAMBA','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030304','JUAN ESPINOZA MEDRANO','ANTABAMBA','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030305','OROPESA','ANTABAMBA','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030306','PACHACONAS','ANTABAMBA','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030307','SABAINO','ANTABAMBA','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030401','CHALHUANCA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030402','CAPAYA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030403','CARAYBAMBA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030404','CHAPIMARCA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030405','COLCABAMBA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030406','COTARUSE','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030407','HUAYLLO','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030408','JUSTO APU SAHUARAURA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030409','LUCRE','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030410','POCOHUANCA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030411','SAN JUAN DE CHACÑA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030412','SAÑAYCA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030413','SORAYA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030414','TAPAIRIHUA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030415','TINTAY','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030416','TORAYA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030417','YANACA','AYMARAES','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030501','TAMBOBAMBA','COTABAMBAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030502','COTABAMBAS','COTABAMBAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030503','COYLLURQUI','COTABAMBAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030504','HAQUIRA','COTABAMBAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030505','MARA','COTABAMBAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030506','CHALLHUAHUACHO','COTABAMBAS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030601','CHINCHEROS','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030602','ANCO_HUALLO','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030603','COCHARCAS','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030604','HUACCANA','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030605','OCOBAMBA','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030606','ONGOY','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030607','URANMARCA','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030608','RANRACANCHA','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030609','ROCCHACC','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030610','EL PORVENIR','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030611','LOS CHANKAS','CHINCHEROS','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030701','CHUQUIBAMBILLA','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030702','CURPAHUASI','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030703','GAMARRA','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030704','HUAYLLATI','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030705','MAMARA','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030706','MICAELA BASTIDAS','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030707','PATAYPAMPA','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030708','PROGRESO','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030709','SAN ANTONIO','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030710','SANTA ROSA','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030711','TURPAY','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030712','VILCABAMBA','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030713','VIRUNDO','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('030714','CURASCO','GRAU','APURIMAC')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040101','AREQUIPA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040102','ALTO SELVA ALEGRE','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040103','CAYMA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040104','CERRO COLORADO','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040105','CHARACATO','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040106','CHIGUATA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040107','JACOBO HUNTER','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040108','LA JOYA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040109','MARIANO MELGAR','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040110','MIRAFLORES','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040111','MOLLEBAYA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040112','PAUCARPATA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040113','POCSI','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040114','POLOBAYA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040115','QUEQUEÑA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040116','SABANDIA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040117','SACHACA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040118','SAN JUAN DE SIGUAS','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040119','SAN JUAN DE TARUCANI','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040120','SANTA ISABEL DE SIGUAS','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040121','SANTA RITA DE SIGUAS','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040122','SOCABAYA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040123','TIABAYA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040124','UCHUMAYO','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040125','VITOR','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040126','YANAHUARA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040127','YARABAMBA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040128','YURA','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040129','JOSE LUIS BUSTAMANTE Y RIVERO','AREQUIPA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040201','CAMANA','CAMANA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040202','JOSE MARIA QUIMPER','CAMANA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040203','MARIANO NICOLAS VALCARCEL','CAMANA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040204','MARISCAL CACERES','CAMANA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040205','NICOLAS DE PIEROLA','CAMANA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040206','OCOÑA','CAMANA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040207','QUILCA','CAMANA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040208','SAMUEL PASTOR','CAMANA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040301','CARAVELI','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040302','ACARI','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040303','ATICO','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040304','ATIQUIPA','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040305','BELLA UNION','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040306','CAHUACHO','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040307','CHALA','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040308','CHAPARRA','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040309','HUANUHUANU','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040310','JAQUI','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040311','LOMAS','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040312','QUICACHA','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040313','YAUCA','CARAVELI','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040401','APLAO','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040402','ANDAGUA','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040403','AYO','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040404','CHACHAS','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040405','CHILCAYMARCA','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040406','CHOCO','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040407','HUANCARQUI','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040408','MACHAGUAY','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040409','ORCOPAMPA','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040410','PAMPACOLCA','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040411','TIPAN','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040412','UÑON','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040413','URACA','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040414','VIRACO','CASTILLA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040501','CHIVAY','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040502','ACHOMA','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040503','CABANACONDE','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040504','CALLALLI','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040505','CAYLLOMA','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040506','COPORAQUE','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040507','HUAMBO','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040508','HUANCA','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040509','ICHUPAMPA','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040510','LARI','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040511','LLUTA','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040512','MACA','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040513','MADRIGAL','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040514','SAN ANTONIO DE CHUCA','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040515','SIBAYO','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040516','TAPAY','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040517','TISCO','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040518','TUTI','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040519','YANQUE','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040520','MAJES','CAYLLOMA','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040601','CHUQUIBAMBA','CONDESUYOS','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040602','ANDARAY','CONDESUYOS','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040603','CAYARANI','CONDESUYOS','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040604','CHICHAS','CONDESUYOS','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040605','IRAY','CONDESUYOS','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040606','RIO GRANDE','CONDESUYOS','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040607','SALAMANCA','CONDESUYOS','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040608','YANAQUIHUA','CONDESUYOS','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040701','MOLLENDO','ISLAY','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040702','COCACHACRA','ISLAY','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040703','DEAN VALDIVIA','ISLAY','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040704','ISLAY','ISLAY','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040705','MEJIA','ISLAY','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040706','PUNTA DE BOMBON','ISLAY','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040801','COTAHUASI','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040802','ALCA','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040803','CHARCANA','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040804','HUAYNACOTAS','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040805','PAMPAMARCA','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040806','PUYCA','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040807','QUECHUALLA','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040808','SAYLA','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040809','TAURIA','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040810','TOMEPAMPA','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('040811','TORO','LA UNION','AREQUIPA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050101','AYACUCHO','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050102','ACOCRO','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050103','ACOS VINCHOS','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050104','CARMEN ALTO','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050105','CHIARA','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050106','OCROS','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050107','PACAYCASA','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050108','QUINUA','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050109','SAN JOSE DE TICLLAS','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050110','SAN JUAN BAUTISTA','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050111','SANTIAGO DE PISCHA','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050112','SOCOS','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050113','TAMBILLO','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050114','VINCHOS','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050115','JESUS NAZARENO','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050116','ANDRÉS AVELINO CÁCERES DORREGARAY','HUAMANGA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050201','CANGALLO','CANGALLO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050202','CHUSCHI','CANGALLO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050203','LOS MOROCHUCOS','CANGALLO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050204','MARIA PARADO DE BELLIDO','CANGALLO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050205','PARAS','CANGALLO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050206','TOTOS','CANGALLO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050301','SANCOS','HUANCA SANCOS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050302','CARAPO','HUANCA SANCOS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050303','SACSAMARCA','HUANCA SANCOS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050304','SANTIAGO DE LUCANAMARCA','HUANCA SANCOS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050401','HUANTA','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050402','AYAHUANCO','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050403','HUAMANGUILLA','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050404','IGUAIN','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050405','LURICOCHA','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050406','SANTILLANA','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050407','SIVIA','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050408','LLOCHEGUA','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050409','CANAYRE','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050410','UCHURACCAY','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050411','PUCACOLPA','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050412','CHACA','HUANTA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050501','SAN MIGUEL','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050502','ANCO','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050503','AYNA','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050504','CHILCAS','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050505','CHUNGUI','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050506','LUIS CARRANZA','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050507','SANTA ROSA','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050508','TAMBO','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050509','SAMUGARI','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050510','ANCHIHUAY','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050511','ORONCCOY','LA MAR','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050601','PUQUIO','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050602','AUCARA','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050603','CABANA','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050604','CARMEN SALCEDO','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050605','CHAVIÑA','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050606','CHIPAO','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050607','HUAC-HUAS','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050608','LARAMATE','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050609','LEONCIO PRADO','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050610','LLAUTA','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050611','LUCANAS','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050612','OCAÑA','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050613','OTOCA','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050614','SAISA','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050615','SAN CRISTOBAL','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050616','SAN JUAN','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050617','SAN PEDRO','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050618','SAN PEDRO DE PALCO','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050619','SANCOS','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050620','SANTA ANA DE HUAYCAHUACHO','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050621','SANTA LUCIA','LUCANAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050701','CORACORA','PARINACOCHAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050702','CHUMPI','PARINACOCHAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050703','CORONEL CASTAÑEDA','PARINACOCHAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050704','PACAPAUSA','PARINACOCHAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050705','PULLO','PARINACOCHAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050706','PUYUSCA','PARINACOCHAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050707','SAN FRANCISCO DE RAVACAYCO','PARINACOCHAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050708','UPAHUACHO','PARINACOCHAS','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050801','PAUSA','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050802','COLTA','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050803','CORCULLA','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050804','LAMPA','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050805','MARCABAMBA','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050806','OYOLO','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050807','PARARCA','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050808','SAN JAVIER DE ALPABAMBA','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050809','SAN JOSE DE USHUA','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050810','SARA SARA','PAUCAR DEL SARA SARA','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050901','QUEROBAMBA','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050902','BELEN','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050903','CHALCOS','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050904','CHILCAYOC','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050905','HUACAÑA','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050906','MORCOLLA','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050907','PAICO','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050908','SAN PEDRO DE LARCAY','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050909','SAN SALVADOR DE QUIJE','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050910','SANTIAGO DE PAUCARAY','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('050911','SORAS','SUCRE','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051001','HUANCAPI','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051002','ALCAMENCA','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051003','APONGO','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051004','ASQUIPATA','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051005','CANARIA','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051006','CAYARA','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051007','COLCA','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051008','HUAMANQUIQUIA','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051009','HUANCARAYLLA','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051010','HUAYA','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051011','SARHUA','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051012','VILCANCHOS','VICTOR FAJARDO','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051101','VILCAS HUAMAN','VILCAS HUAMAN','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051102','ACCOMARCA','VILCAS HUAMAN','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051103','CARHUANCA','VILCAS HUAMAN','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051104','CONCEPCION','VILCAS HUAMAN','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051105','HUAMBALPA','VILCAS HUAMAN','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051106','INDEPENDENCIA','VILCAS HUAMAN','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051107','SAURAMA','VILCAS HUAMAN','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('051108','VISCHONGO','VILCAS HUAMAN','AYACUCHO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060101','CAJAMARCA','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060102','ASUNCION','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060103','CHETILLA','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060104','COSPAN','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060105','ENCAÑADA','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060106','JESUS','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060107','LLACANORA','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060108','LOS BAÑOS DEL INCA','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060109','MAGDALENA','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060110','MATARA','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060111','NAMORA','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060112','SAN JUAN','CAJAMARCA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060201','CAJABAMBA','CAJABAMBA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060202','CACHACHI','CAJABAMBA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060203','CONDEBAMBA','CAJABAMBA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060204','SITACOCHA','CAJABAMBA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060301','CELENDIN','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060302','CHUMUCH','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060303','CORTEGANA','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060304','HUASMIN','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060305','JORGE CHAVEZ','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060306','JOSE GALVEZ','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060307','MIGUEL IGLESIAS','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060308','OXAMARCA','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060309','SOROCHUCO','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060310','SUCRE','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060311','UTCO','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060312','LA LIBERTAD DE PALLAN','CELENDIN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060401','CHOTA','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060402','ANGUIA','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060403','CHADIN','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060404','CHIGUIRIP','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060405','CHIMBAN','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060406','CHOROPAMPA','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060407','COCHABAMBA','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060408','CONCHAN','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060409','HUAMBOS','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060410','LAJAS','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060411','LLAMA','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060412','MIRACOSTA','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060413','PACCHA','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060414','PION','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060415','QUEROCOTO','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060416','SAN JUAN DE LICUPIS','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060417','TACABAMBA','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060418','TOCMOCHE','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060419','CHALAMARCA','CHOTA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060501','CONTUMAZA','CONTUMAZA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060502','CHILETE','CONTUMAZA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060503','CUPISNIQUE','CONTUMAZA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060504','GUZMANGO','CONTUMAZA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060505','SAN BENITO','CONTUMAZA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060506','SANTA CRUZ DE TOLED','CONTUMAZA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060507','TANTARICA','CONTUMAZA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060508','YONAN','CONTUMAZA','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060601','CUTERVO','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060602','CALLAYUC','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060603','CHOROS','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060604','CUJILLO','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060605','LA RAMADA','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060606','PIMPINGOS','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060607','QUEROCOTILLO','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060608','SAN ANDRES DE CUTERVO','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060609','SAN JUAN DE CUTERVO','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060610','SAN LUIS DE LUCMA','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060611','SANTA CRUZ','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060612','SANTO DOMINGO DE LA CAPILLA','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060613','SANTO TOMAS','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060614','SOCOTA','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060615','TORIBIO CASANOVA','CUTERVO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060701','BAMBAMARCA','HUALGAYOC','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060702','CHUGUR','HUALGAYOC','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060703','HUALGAYOC','HUALGAYOC','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060801','JAEN','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060802','BELLAVISTA','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060803','CHONTALI','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060804','COLASAY','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060805','HUABAL','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060806','LAS PIRIAS','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060807','POMAHUACA','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060808','PUCARA','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060809','SALLIQUE','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060810','SAN FELIPE','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060811','SAN JOSE DEL ALTO','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060812','SANTA ROSA','JAEN','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060901','SAN IGNACIO','SAN IGNACIO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060902','CHIRINOS','SAN IGNACIO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060903','HUARANGO','SAN IGNACIO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060904','LA COIPA','SAN IGNACIO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060905','NAMBALLE','SAN IGNACIO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060906','SAN JOSE DE LOURDES','SAN IGNACIO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('060907','TABACONAS','SAN IGNACIO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061001','PEDRO GALVEZ','SAN MARCOS','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061002','CHANCAY','SAN MARCOS','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061003','EDUARDO VILLANUEVA','SAN MARCOS','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061004','GREGORIO PITA','SAN MARCOS','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061005','ICHOCAN','SAN MARCOS','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061006','JOSE MANUEL QUIROZ','SAN MARCOS','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061007','JOSE SABOGAL','SAN MARCOS','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061101','SAN MIGUEL','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061102','BOLIVAR','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061103','CALQUIS','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061104','CATILLUC','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061105','EL PRADO','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061106','LA FLORIDA','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061107','LLAPA','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061108','NANCHOC','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061109','NIEPOS','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061110','SAN GREGORIO','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061111','SAN SILVESTRE DE COCHAN','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061112','TONGOD','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061113','UNION AGUA BLANCA','SAN MIGUEL','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061201','SAN PABLO','SAN PABLO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061202','SAN BERNARDINO','SAN PABLO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061203','SAN LUIS','SAN PABLO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061204','TUMBADEN','SAN PABLO','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061301','SANTA CRUZ','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061302','ANDABAMBA','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061303','CATACHE','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061304','CHANCAYBAÑOS','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061305','LA ESPERANZA','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061306','NINABAMBA','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061307','PULAN','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061308','SAUCEPAMPA','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061309','SEXI','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061310','UTICYACU','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('061311','YAUYUCAN','SANTA CRUZ','CAJAMARCA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('070101','CALLAO','CALLAO','CALLAO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('070102','BELLAVISTA','CALLAO','CALLAO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('070103','CARMEN DE LA LEGUA','CALLAO','CALLAO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('070104','LA PERLA','CALLAO','CALLAO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('070105','LA PUNTA','CALLAO','CALLAO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('070106','VENTANILLA','CALLAO','CALLAO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('070107','MI PERU','CALLAO','CALLAO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080101','CUSCO','CUSCO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080102','CCORCA','CUSCO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080103','POROY','CUSCO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080104','SAN JERONIMO','CUSCO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080105','SAN SEBASTIAN','CUSCO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080106','SANTIAGO','CUSCO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080107','SAYLLA','CUSCO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080108','WANCHAQ','CUSCO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080201','ACOMAYO','ACOMAYO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080202','ACOPIA','ACOMAYO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080203','ACOS','ACOMAYO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080204','MOSOC LLACTA','ACOMAYO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080205','POMACANCHI','ACOMAYO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080206','RONDOCAN','ACOMAYO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080207','SANGARARA','ACOMAYO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080301','ANTA','ANTA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080302','ANCAHUASI','ANTA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080303','CACHIMAYO','ANTA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080304','CHINCHAYPUJIO','ANTA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080305','HUAROCONDO','ANTA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080306','LIMATAMBO','ANTA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080307','MOLLEPATA','ANTA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080308','PUCYURA','ANTA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080309','ZURITE','ANTA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080401','CALCA','CALCA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080402','COYA','CALCA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080403','LAMAY','CALCA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080404','LARES','CALCA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080405','PISAC','CALCA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080406','SAN SALVADOR','CALCA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080407','TARAY','CALCA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080408','YANATILE','CALCA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080501','YANAOCA','CANAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080502','CHECCA','CANAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080503','KUNTURKANKI','CANAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080504','LANGUI','CANAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080505','LAYO','CANAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080506','PAMPAMARCA','CANAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080507','QUEHUE','CANAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080508','TUPAC AMARU','CANAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080601','SICUANI','CANCHIS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080602','CHECACUPE','CANCHIS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080603','COMBAPATA','CANCHIS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080604','MARANGANI','CANCHIS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080605','PITUMARCA','CANCHIS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080606','SAN PABLO','CANCHIS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080607','SAN PEDRO','CANCHIS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080608','TINTA','CANCHIS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080701','SANTO TOMAS','CHUMBIVILCAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080702','CAPACMARCA','CHUMBIVILCAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080703','CHAMACA','CHUMBIVILCAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080704','COLQUEMARCA','CHUMBIVILCAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080705','LIVITACA','CHUMBIVILCAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080706','LLUSCO','CHUMBIVILCAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080707','QUIÑOTA','CHUMBIVILCAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080708','VELILLE','CHUMBIVILCAS','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080801','ESPINAR','ESPINAR','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080802','CONDOROMA','ESPINAR','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080803','COPORAQUE','ESPINAR','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080804','OCORURO','ESPINAR','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080805','PALLPATA','ESPINAR','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080806','PICHIGUA','ESPINAR','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080807','SUYCKUTAMBO','ESPINAR','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080808','ALTO PICHIGUA','ESPINAR','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080901','SANTA ANA','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080902','ECHARATE','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080903','HUAYOPATA','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080904','MARANURA','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080905','OCOBAMBA','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080906','QUELLOUNO','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080907','KIMBIRI','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080908','SANTA TERESA','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080909','VILCABAMBA','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080910','PICHARI','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080911','INKAWASI','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080912','VILLA VIRGEN','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080913','VILLA KINTIARINA','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('080914','MEGANTONI','LA CONVENCION','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081001','PARURO','PARURO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081002','ACCHA','PARURO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081003','CCAPI','PARURO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081004','COLCHA','PARURO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081005','HUANOQUITE','PARURO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081006','OMACHA','PARURO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081007','PACCARITAMBO','PARURO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081008','PILLPINTO','PARURO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081009','YAURISQUE','PARURO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081101','PAUCARTAMBO','PAUCARTAMBO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081102','CAICAY','PAUCARTAMBO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081103','CHALLABAMBA','PAUCARTAMBO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081104','COLQUEPATA','PAUCARTAMBO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081105','HUANCARANI','PAUCARTAMBO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081106','KOSÑIPATA','PAUCARTAMBO','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081201','URCOS','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081202','ANDAHUAYLILLAS','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081203','CAMANTI','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081204','CCARHUAYO','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081205','CCATCA','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081206','CUSIPATA','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081207','HUARO','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081208','LUCRE','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081209','MARCAPATA','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081210','OCONGATE','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081211','OROPESA','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081212','QUIQUIJANA','QUISPICANCHI','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081301','URUBAMBA','URUBAMBA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081302','CHINCHERO','URUBAMBA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081303','HUAYLLABAMBA','URUBAMBA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081304','MACHUPICCHU','URUBAMBA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081305','MARAS','URUBAMBA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081306','OLLANTAYTAMBO','URUBAMBA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('081307','YUCAY','URUBAMBA','CUSCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090101','HUANCAVELICA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090102','ACOBAMBILLA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090103','ACORIA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090104','CONAYCA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090105','CUENCA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090106','HUACHOCOLPA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090107','HUAYLLAHUARA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090108','IZCUCHACA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090109','LARIA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090110','MANTA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090111','MARISCAL CACERES','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090112','MOYA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090113','NUEVO OCCORO','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090114','PALCA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090115','PILCHACA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090116','VILCA','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090117','YAULI','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090118','ASCENSION','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090119','HUANDO','HUANCAVELICA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090201','ACOBAMBA','ACOBAMBA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090202','ANDABAMBA','ACOBAMBA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090203','ANTA','ACOBAMBA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090204','CAJA','ACOBAMBA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090205','MARCAS','ACOBAMBA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090206','PAUCARA','ACOBAMBA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090207','POMACOCHA','ACOBAMBA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090208','ROSARIO','ACOBAMBA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090301','LIRCAY','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090302','ANCHONGA','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090303','CALLANMARCA','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090304','CCOCHACCASA','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090305','CHINCHO','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090306','CONGALLA','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090307','HUANCA-HUANCA','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090308','HUAYLLAY GRANDE','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090309','JULCAMARCA','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090310','SAN ANTONIO DE ANTAPARCO','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090311','SANTO TOMAS DE PATA','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090312','SECCLLA','ANGARAES','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090401','CASTROVIRREYNA','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090402','ARMA','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090403','AURAHUA','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090404','CAPILLAS','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090405','CHUPAMARCA','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090406','COCAS','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090407','HUACHOS','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090408','HUAMATAMBO','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090409','MOLLEPAMPA','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090410','SAN JUAN','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090411','SANTA ANA','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090412','TANTARA','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090413','TICRAPO','CASTROVIRREYNA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090501','CHURCAMPA','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090502','ANCO','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090503','CHINCHIHUASI','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090504','EL CARMEN','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090505','LA MERCED','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090506','LOCROJA','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090507','PAUCARBAMBA','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090508','SAN MIGUEL DE MAYOCC','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090509','SAN PEDRO DE CORIS','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090510','PACHAMARCA','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090511','COSME','CHURCAMPA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090601','HUAYTARA','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090602','AYAVI','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090603','CORDOVA','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090604','HUAYACUNDO ARMA','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090605','LARAMARCA','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090606','OCOYO','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090607','PILPICHACA','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090608','QUERCO','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090609','QUITO-ARMA','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090610','SAN ANTONIO DE CUSICANCHA','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090611','SAN FRANCISCO DE SANGAYAICO','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090612','SAN ISIDRO','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090613','SANTIAGO DE CHOCORVOS','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090614','SANTIAGO DE QUIRAHUARA','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090615','SANTO DOMINGO DE CAPILLAS','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090616','TAMBO','HUAYTARA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090701','PAMPAS','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090702','ACOSTAMBO','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090703','ACRAQUIA','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090704','AHUAYCHA','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090705','COLCABAMBA','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090706','DANIEL HERNANDEZ','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090707','HUACHOCOLPA','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090709','HUARIBAMBA','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090710','ÑAHUIMPUQUIO','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090711','PAZOS','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090713','QUISHUAR','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090714','SALCABAMBA','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090715','SALCAHUASI','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090716','SAN MARCOS DE ROCCHAC','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090717','SURCUBAMBA','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090718','TINTAY PUNCU','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090719','QUICHUAS','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090720','ANDAYMARCA','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090721','ROBLE','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090722','PICHOS','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('090723','SANTIAGO DE TÚCUMA','TAYACAJA','HUANCAVELICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100101','HUANUCO','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100102','AMARILIS','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100103','CHINCHAO','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100104','CHURUBAMBA','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100105','MARGOS','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100106','QUISQUI','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100107','SAN FRANCISCO DE CAYRAN','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100108','SAN PEDRO DE CHAULAN','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100109','SANTA MARIA DEL VALLE','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100110','YARUMAYO','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100111','PILLCO MARCA','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100112','YACUS','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100113','SAN PABLO DE PILLAO','HUANUCO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100201','AMBO','AMBO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100202','CAYNA','AMBO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100203','COLPAS','AMBO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100204','CONCHAMARCA','AMBO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100205','HUACAR','AMBO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100206','SAN FRANCISCO','AMBO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100207','SAN RAFAEL','AMBO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100208','TOMAY KICHWA','AMBO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100301','LA UNION','DOS DE MAYO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100307','CHUQUIS','DOS DE MAYO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100311','MARIAS','DOS DE MAYO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100313','PACHAS','DOS DE MAYO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100316','QUIVILLA','DOS DE MAYO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100317','RIPAN','DOS DE MAYO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100321','SHUNQUI','DOS DE MAYO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100322','SILLAPATA','DOS DE MAYO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100323','YANAS','DOS DE MAYO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100401','HUACAYBAMBA','HUACAYBAMBA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100402','CANCHABAMBA','HUACAYBAMBA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100403','COCHABAMBA','HUACAYBAMBA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100404','PINRA','HUACAYBAMBA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100501','LLATA','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100502','ARANCAY','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100503','CHAVIN DE PARIARCA','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100504','JACAS GRANDE','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100505','JIRCAN','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100506','MIRAFLORES','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100507','MONZON','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100508','PUNCHAO','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100509','PUÑOS','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100510','SINGA','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100511','TANTAMAYO','HUAMALIES','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100601','RUPA-RUPA','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100602','DANIEL ALOMIAS ROBLES','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100603','HERMILIO VALDIZAN','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100604','JOSE CRESPO Y CASTILLO','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100605','LUYANDO','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100606','MARIANO DAMASO BERAUN','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100607','PUCAYACU','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100608','CASTILLO GRANDE','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100609','PUEBLO NUEVO','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100610','SANTO DOMINGO DE ANDA','LEONCIO PRADO','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100701','HUACRACHUCO','MARAÑON','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100702','CHOLON','MARAÑON','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100703','SAN BUENAVENTURA','MARAÑON','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100704','LA MORADA','MARAÑON','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100705','SANTA ROSA DE ALTO YANAJANCA','MARAÑON','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100801','PANAO','PACHITEA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100802','CHAGLLA','PACHITEA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100803','MOLINO','PACHITEA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100804','UMARI','PACHITEA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100901','PUERTO INCA','PUERTO INCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100902','CODO DEL POZUZO','PUERTO INCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100903','HONORIA','PUERTO INCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100904','TOURNAVISTA','PUERTO INCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('100905','YUYAPICHIS','PUERTO INCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101001','JESUS','LAURICOCHA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101002','BAÑOS','LAURICOCHA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101003','JIVIA','LAURICOCHA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101004','QUEROPALCA','LAURICOCHA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101005','RONDOS','LAURICOCHA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101006','SAN FRANCISCO DE ASIS','LAURICOCHA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101007','SAN MIGUEL DE CAURI','LAURICOCHA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101101','CHAVINILLO','YAROWILCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101102','CAHUAC','YAROWILCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101103','CHACABAMBA','YAROWILCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101104','APARICIO POMARES','YAROWILCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101105','JACAS CHICO','YAROWILCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101106','OBAS','YAROWILCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101107','PAMPAMARCA','YAROWILCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('101108','CHORAS','YAROWILCA','HUANUCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110101','ICA','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110102','LA TINGUIÑA','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110103','LOS AQUIJES','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110104','OCUCAJE','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110105','PACHACUTEC','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110106','PARCONA','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110107','PUEBLO NUEVO','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110108','SALAS','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110109','SAN JOSE DE LOS MOLINOS','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110110','SAN JUAN BAUTISTA','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110111','SANTIAGO','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110112','SUBTANJALLA','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110113','TATE','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110114','YAUCA DEL ROSARIO','ICA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110201','CHINCHA ALTA','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110202','ALTO LARAN','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110203','CHAVIN','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110204','CHINCHA BAJA','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110205','EL CARMEN','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110206','GROCIO PRADO','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110207','PUEBLO NUEVO','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110208','SAN JUAN DE YANAC','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110209','SAN PEDRO DE HUACARPANA','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110210','SUNAMPE','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110211','TAMBO DE MORA','CHINCHA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110301','NAZCA','NAZCA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110302','CHANGUILLO','NAZCA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110303','EL INGENIO','NAZCA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110304','MARCONA','NAZCA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110305','VISTA ALEGRE','NAZCA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110401','PALPA','PALPA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110402','LLIPATA','PALPA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110403','RIO GRANDE','PALPA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110404','SANTA CRUZ','PALPA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110405','TIBILLO','PALPA','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110501','PISCO','PISCO','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110502','HUANCANO','PISCO','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110503','HUMAY','PISCO','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110504','INDEPENDENCIA','PISCO','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110505','PARACAS','PISCO','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110506','SAN ANDRES','PISCO','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110507','SAN CLEMENTE','PISCO','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('110508','TUPAC AMARU INCA','PISCO','ICA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120101','HUANCAYO','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120104','CARHUACALLANGA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120105','CHACAPAMPA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120106','CHICCHE','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120107','CHILCA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120108','CHONGOS ALTO','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120111','CHUPURO','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120112','COLCA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120113','CULLHUAS','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120114','EL TAMBO','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120116','HUACRAPUQUIO','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120117','HUALHUAS','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120119','HUANCAN','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120120','HUASICANCHA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120121','HUAYUCACHI','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120122','INGENIO','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120124','PARIAHUANCA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120125','PILCOMAYO','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120126','PUCARA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120127','QUICHUAY','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120128','QUILCAS','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120129','SAN AGUSTIN','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120130','SAN JERONIMO DE TUNAN','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120132','SAÑO','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120133','SAPALLANGA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120134','SICAYA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120135','SANTO DOMINGO DE ACOBAMBA','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120136','VIQUES','HUANCAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120201','CONCEPCION','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120202','ACO','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120203','ANDAMARCA','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120204','CHAMBARA','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120205','COCHAS','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120206','COMAS','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120207','HEROINAS TOLEDO','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120208','MANZANARES','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120209','MARISCAL CASTILLA','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120210','MATAHUASI','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120211','MITO','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120212','NUEVE DE JULIO','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120213','ORCOTUNA','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120214','SAN JOSE DE QUERO','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120215','SANTA ROSA DE OCOPA','CONCEPCION','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120301','CHANCHAMAYO','CHANCHAMAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120302','PERENE','CHANCHAMAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120303','PICHANAQUI','CHANCHAMAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120304','SAN LUIS DE SHUARO','CHANCHAMAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120305','SAN RAMON','CHANCHAMAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120306','VITOC','CHANCHAMAYO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120401','JAUJA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120402','ACOLLA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120403','APATA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120404','ATAURA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120405','CANCHAYLLO','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120406','CURICACA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120407','EL MANTARO','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120408','HUAMALI','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120409','HUARIPAMPA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120410','HUERTAS','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120411','JANJAILLO','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120412','JULCAN','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120413','LEONOR ORDOÑEZ','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120414','LLOCLLAPAMPA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120415','MARCO','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120416','MASMA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120417','MASMA CHICCHE','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120418','MOLINOS','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120419','MONOBAMBA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120420','MUQUI','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120421','MUQUIYAUYO','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120422','PACA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120423','PACCHA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120424','PANCAN','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120425','PARCO','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120426','POMACANCHA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120427','RICRAN','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120428','SAN LORENZO','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120429','SAN PEDRO DE CHUNAN','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120430','SAUSA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120431','SINCOS','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120432','TUNAN MARCA','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120433','YAULI','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120434','YAUYOS','JAUJA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120501','JUNIN','JUNIN','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120502','CARHUAMAYO','JUNIN','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120503','ONDORES','JUNIN','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120504','ULCUMAYO','JUNIN','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120601','SATIPO','SATIPO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120602','COVIRIALI','SATIPO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120603','LLAYLLA','SATIPO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120604','MAZAMARI','SATIPO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120605','PAMPA HERMOSA','SATIPO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120606','PANGOA','SATIPO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120607','RIO NEGRO','SATIPO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120608','RIO TAMBO','SATIPO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120609','VIZCATÁN DEL ENE','SATIPO','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120701','TARMA','TARMA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120702','ACOBAMBA','TARMA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120703','HUARICOLCA','TARMA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120704','HUASAHUASI','TARMA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120705','LA UNION','TARMA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120706','PALCA','TARMA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120707','PALCAMAYO','TARMA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120708','SAN PEDRO DE CAJAS','TARMA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120709','TAPO','TARMA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120801','LA OROYA','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120802','CHACAPALPA','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120803','HUAY-HUAY','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120804','MARCAPOMACOCHA','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120805','MOROCOCHA','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120806','PACCHA','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120807','SANTA BARBARA DE CARHUACAYAN','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120808','SANTA ROSA DE SACCO','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120809','SUITUCANCHA','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120810','YAULI','YAULI','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120901','CHUPACA','CHUPACA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120902','AHUAC','CHUPACA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120903','CHONGOS BAJO','CHUPACA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120904','HUACHAC','CHUPACA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120905','HUAMANCACA CHICO','CHUPACA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120906','SAN JUAN DE YSCOS','CHUPACA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120907','SAN JUAN DE JARPA','CHUPACA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120908','TRES DE DICIEMBRE','CHUPACA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('120909','YANACANCHA','CHUPACA','JUNIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130101','TRUJILLO','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130102','EL PORVENIR','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130103','FLORENCIA DE MORA','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130104','HUANCHACO','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130105','LA ESPERANZA','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130106','LAREDO','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130107','MOCHE','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130108','POROTO','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130109','SALAVERRY','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130110','SIMBAL','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130111','VICTOR LARCO HERRERA','TRUJILLO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130201','ASCOPE','ASCOPE','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130202','CHICAMA','ASCOPE','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130203','CHOCOPE','ASCOPE','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130204','MAGDALENA DE CAO','ASCOPE','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130205','PAIJAN','ASCOPE','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130206','RAZURI','ASCOPE','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130207','SANTIAGO DE CAO','ASCOPE','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130208','CASA GRANDE','ASCOPE','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130301','BOLIVAR','BOLIVAR','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130302','BAMBAMARCA','BOLIVAR','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130303','CONDORMARCA','BOLIVAR','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130304','LONGOTEA','BOLIVAR','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130305','UCHUMARCA','BOLIVAR','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130306','UCUNCHA','BOLIVAR','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130401','CHEPEN','CHEPEN','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130402','PACANGA','CHEPEN','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130403','PUEBLO NUEVO','CHEPEN','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130501','JULCAN','JULCAN','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130502','CALAMARCA','JULCAN','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130503','CARABAMBA','JULCAN','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130504','HUASO','JULCAN','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130601','OTUZCO','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130602','AGALLPAMPA','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130604','CHARAT','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130605','HUARANCHAL','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130606','LA CUESTA','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130608','MACHE','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130610','PARANDAY','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130611','SALPO','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130613','SINSICAP','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130614','USQUIL','OTUZCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130701','SAN PEDRO DE LLOC','PACASMAYO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130702','GUADALUPE','PACASMAYO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130703','JEQUETEPEQUE','PACASMAYO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130704','PACASMAYO','PACASMAYO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130705','SAN JOSE','PACASMAYO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130801','TAYABAMBA','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130802','BULDIBUYO','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130803','CHILLIA','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130804','HUANCASPATA','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130805','HUAYLILLAS','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130806','HUAYO','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130807','ONGON','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130808','PARCOY','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130809','PATAZ','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130810','PIAS','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130811','SANTIAGO DE CHALLAS','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130812','TAURIJA','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130813','URPAY','PATAZ','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130901','HUAMACHUCO','SANCHEZ CARRION','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130902','CHUGAY','SANCHEZ CARRION','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130903','COCHORCO','SANCHEZ CARRION','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130904','CURGOS','SANCHEZ CARRION','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130905','MARCABAL','SANCHEZ CARRION','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130906','SANAGORAN','SANCHEZ CARRION','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130907','SARIN','SANCHEZ CARRION','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('130908','SARTIMBAMBA','SANCHEZ CARRION','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131001','SANTIAGO DE CHUCO','SANTIAGO DE CHUCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131002','ANGASMARCA','SANTIAGO DE CHUCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131003','CACHICADAN','SANTIAGO DE CHUCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131004','MOLLEBAMBA','SANTIAGO DE CHUCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131005','MOLLEPATA','SANTIAGO DE CHUCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131006','QUIRUVILCA','SANTIAGO DE CHUCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131007','SANTA CRUZ DE CHUCA','SANTIAGO DE CHUCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131008','SITABAMBA','SANTIAGO DE CHUCO','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131101','CASCAS','GRAN CHIMU','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131102','LUCMA','GRAN CHIMU','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131103','COMPIN','GRAN CHIMU','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131104','SAYAPULLO','GRAN CHIMU','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131201','VIRU','VIRU','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131202','CHAO','VIRU','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('131203','GUADALUPITO','VIRU','LA LIBERTAD')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140101','CHICLAYO','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140102','CHONGOYAPE','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140103','ETEN','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140104','ETEN PUERTO','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140105','JOSE LEONARDO ORTIZ','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140106','LA VICTORIA','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140107','LAGUNAS','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140108','MONSEFU','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140109','NUEVA ARICA','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140110','OYOTUN','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140111','PICSI','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140112','PIMENTEL','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140113','REQUE','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140114','SANTA ROSA','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140115','SAÑA','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140116','CAYALTI','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140117','PATAPO','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140118','POMALCA','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140119','PUCALA','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140120','TUMAN','CHICLAYO','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140201','FERREÑAFE','FERREÑAFE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140202','CAÑARIS','FERREÑAFE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140203','INCAHUASI','FERREÑAFE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140204','MANUEL ANTONIO MESONES MURO','FERREÑAFE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140205','PITIPO','FERREÑAFE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140206','PUEBLO NUEVO','FERREÑAFE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140301','LAMBAYEQUE','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140302','CHOCHOPE','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140303','ILLIMO','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140304','JAYANCA','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140305','MOCHUMI','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140306','MORROPE','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140307','MOTUPE','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140308','OLMOS','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140309','PACORA','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140310','SALAS','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140311','SAN JOSE','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('140312','TUCUME','LAMBAYEQUE','LAMBAYEQUE')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150101','LIMA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150102','ANCON','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150103','ATE','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150104','BARRANCO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150105','BREÑA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150106','CARABAYLLO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150107','CHACLACAYO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150108','CHORRILLOS','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150109','CIENEGUILLA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150110','COMAS','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150111','EL AGUSTINO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150112','INDEPENDENCIA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150113','JESUS MARIA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150114','LA MOLINA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150115','LA VICTORIA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150116','LINCE','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150117','LOS OLIVOS','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150118','LURIGANCHO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150119','LURIN','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150120','MAGDALENA DEL MAR','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150121','PUEBLO LIBRE','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150122','MIRAFLORES','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150123','PACHACAMAC','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150124','PUCUSANA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150125','PUENTE PIEDRA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150126','PUNTA HERMOSA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150127','PUNTA NEGRA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150128','RIMAC','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150129','SAN BARTOLO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150130','SAN BORJA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150131','SAN ISIDRO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150132','SAN JUAN DE LURIGANCHO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150133','SAN JUAN DE MIRAFLORES','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150134','SAN LUIS','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150135','SAN MARTIN DE PORRES','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150136','SAN MIGUEL','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150137','SANTA ANITA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150138','SANTA MARIA DEL MAR','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150139','SANTA ROSA','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150140','SANTIAGO DE SURCO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150141','SURQUILLO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150142','VILLA EL SALVADOR','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150143','VILLA MARIA DEL TRIUNFO','LIMA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150201','BARRANCA','BARRANCA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150202','PARAMONGA','BARRANCA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150203','PATIVILCA','BARRANCA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150204','SUPE','BARRANCA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150205','SUPE PUERTO','BARRANCA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150301','CAJATAMBO','CAJATAMBO','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150302','COPA','CAJATAMBO','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150303','GORGOR','CAJATAMBO','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150304','HUANCAPON','CAJATAMBO','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150305','MANAS','CAJATAMBO','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150401','CANTA','CANTA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150402','ARAHUAY','CANTA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150403','HUAMANTANGA','CANTA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150404','HUAROS','CANTA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150405','LACHAQUI','CANTA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150406','SAN BUENAVENTURA','CANTA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150407','SANTA ROSA DE QUIVES','CANTA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150501','SAN VICENTE DE CAÑETE','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150502','ASIA','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150503','CALANGO','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150504','CERRO AZUL','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150505','CHILCA','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150506','COAYLLO','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150507','IMPERIAL','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150508','LUNAHUANA','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150509','MALA','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150510','NUEVO IMPERIAL','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150511','PACARAN','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150512','QUILMANA','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150513','SAN ANTONIO','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150514','SAN LUIS','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150515','SANTA CRUZ DE FLORES','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150516','ZUÑIGA','CAÑETE','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150601','HUARAL','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150602','ATAVILLOS ALTO','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150603','ATAVILLOS BAJO','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150604','AUCALLAMA','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150605','CHANCAY','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150606','IHUARI','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150607','LAMPIAN','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150608','PACARAOS','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150609','SAN MIGUEL DE ACOS','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150610','SANTA CRUZ DE ANDAMARCA','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150611','SUMBILCA','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150612','VEINTISIETE DE NOVIEMBRE','HUARAL','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150701','MATUCANA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150702','ANTIOQUIA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150703','CALLAHUANCA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150704','CARAMPOMA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150705','CHICLA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150706','CUENCA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150707','HUACHUPAMPA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150708','HUANZA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150709','HUAROCHIRI','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150710','LAHUAYTAMBO','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150711','LANGA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150712','LARAOS','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150713','MARIATANA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150714','RICARDO PALMA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150715','SAN ANDRES DE TUPICOCHA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150716','SAN ANTONIO','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150717','SAN BARTOLOME','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150718','SAN DAMIAN','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150719','SAN JUAN DE IRIS','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150720','SAN JUAN DE TANTARANCHE','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150721','SAN LORENZO DE QUINTI','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150722','SAN MATEO','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150723','SAN MATEO DE OTAO','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150724','SAN PEDRO DE CASTA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150725','SAN PEDRO DE HUANCAYRE','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150726','SANGALLAYA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150727','SANTA CRUZ DE COCACHACRA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150728','SANTA EULALIA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150729','SANTIAGO DE ANCHUCAYA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150730','SANTIAGO DE TUNA','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150731','SANTO DOMINGO DE LOS OLLEROS','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150732','SURCO','HUAROCHIRI','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150801','HUACHO','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150802','AMBAR','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150803','CALETA DE CARQUIN','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150804','CHECRAS','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150805','HUALMAY','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150806','HUAURA','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150807','LEONCIO PRADO','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150808','PACCHO','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150809','SANTA LEONOR','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150810','SANTA MARIA','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150811','SAYAN','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150812','VEGUETA','HUAURA','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150901','OYON','OYON','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150902','ANDAJES','OYON','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150903','CAUJUL','OYON','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150904','COCHAMARCA','OYON','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150905','NAVAN','OYON','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('150906','PACHANGARA','OYON','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151001','YAUYOS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151002','ALIS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151003','AYAUCA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151004','AYAVIRI','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151005','AZANGARO','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151006','CACRA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151007','CARANIA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151008','CATAHUASI','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151009','CHOCOS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151010','COCHAS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151011','COLONIA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151012','HONGOS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151013','HUAMPARA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151014','HUANCAYA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151015','HUANGASCAR','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151016','HUANTAN','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151017','HUAÑEC','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151018','LARAOS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151019','LINCHA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151020','MADEAN','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151021','MIRAFLORES','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151022','OMAS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151023','PUTINZA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151024','QUINCHES','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151025','QUINOCAY','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151026','SAN JOAQUIN','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151027','SAN PEDRO DE PILAS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151028','TANTA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151029','TAURIPAMPA','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151030','TOMAS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151031','TUPE','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151032','VIÑAC','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('151033','VITIS','YAUYOS','LIMA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160101','IQUITOS','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160102','ALTO NANAY','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160103','FERNANDO LORES','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160104','INDIANA','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160105','LAS AMAZONAS','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160106','MAZAN','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160107','NAPO','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160108','PUNCHANA','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160110','TORRES CAUSANA','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160112','BELEN','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160113','SAN JUAN BAUTISTA','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160201','YURIMAGUAS','ALTO AMAZONAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160202','BALSAPUERTO','ALTO AMAZONAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160205','JEBEROS','ALTO AMAZONAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160206','LAGUNAS','ALTO AMAZONAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160210','SANTA CRUZ','ALTO AMAZONAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160211','TENIENTE CESAR LOPEZ ROJAS','ALTO AMAZONAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160301','NAUTA','LORETO','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160302','PARINARI','LORETO','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160303','TIGRE','LORETO','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160304','TROMPETEROS','LORETO','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160305','URARINAS','LORETO','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160401','RAMON CASTILLA','MARISCAL RAMON CASTILLA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160402','PEBAS','MARISCAL RAMON CASTILLA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160403','YAVARI','MARISCAL RAMON CASTILLA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160404','SAN PABLO','MARISCAL RAMON CASTILLA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160501','REQUENA','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160502','ALTO TAPICHE','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160503','CAPELO','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160504','EMILIO SAN MARTIN','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160505','MAQUIA','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160506','PUINAHUA','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160507','SAQUENA','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160508','SOPLIN','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160509','TAPICHE','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160510','JENARO HERRERA','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160511','YAQUERANA','REQUENA','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160601','CONTAMANA','UCAYALI','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160602','INAHUAYA','UCAYALI','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160603','PADRE MARQUEZ','UCAYALI','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160604','PAMPA HERMOSA','UCAYALI','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160605','SARAYACU','UCAYALI','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160606','VARGAS GUERRA','UCAYALI','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160701','BARRANCA','DATEM DEL MARAÑON','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160702','CAHUAPANAS','DATEM DEL MARAÑON','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160703','MANSERICHE','DATEM DEL MARAÑON','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160704','MORONA','DATEM DEL MARAÑON','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160705','PASTAZA','DATEM DEL MARAÑON','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160706','ANDOAS','DATEM DEL MARAÑON','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160801','PUTUMAYO','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160802','ROSA PANDURO','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160803','TENIENTE MANUEL CLAVERO','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('160804','YAGUAS','MAYNAS','LORETO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170101','TAMBOPATA','TAMBOPATA','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170102','INAMBARI','TAMBOPATA','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170103','LAS PIEDRAS','TAMBOPATA','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170104','LABERINTO','TAMBOPATA','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170201','MANU','MANU','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170202','FITZCARRALD','MANU','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170203','MADRE DE DIOS','MANU','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170204','HUEPETUHE','MANU','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170301','IÑAPARI','TAHUAMANU','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170302','IBERIA','TAHUAMANU','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('170303','TAHUAMANU','TAHUAMANU','MADRE DE DIOS')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180101','MOQUEGUA','MARISCAL NIETO','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180102','CARUMAS','MARISCAL NIETO','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180103','CUCHUMBAYA','MARISCAL NIETO','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180104','SAMEGUA','MARISCAL NIETO','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180105','SAN CRISTOBAL','MARISCAL NIETO','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180106','TORATA','MARISCAL NIETO','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180201','OMATE','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180202','CHOJATA','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180203','COALAQUE','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180204','ICHUÑA','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180205','LA CAPILLA','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180206','LLOQUE','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180207','MATALAQUE','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180208','PUQUINA','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180209','QUINISTAQUILLAS','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180210','UBINAS','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180211','YUNGA','GENERAL SANCHEZ CERR','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180301','ILO','ILO','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180302','EL ALGARROBAL','ILO','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('180303','PACOCHA','ILO','MOQUEGUA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190101','CHAUPIMARCA','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190102','HUACHON','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190103','HUARIACA','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190104','HUAYLLAY','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190105','NINACACA','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190106','PALLANCHACRA','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190107','PAUCARTAMBO','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190108','SAN FRANCISCO DE ASIS DE YARUSYACAN','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190109','SIMON BOLIVAR','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190110','TICLACAYAN','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190111','TINYAHUARCO','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190112','VICCO','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190113','YANACANCHA','PASCO','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190201','YANAHUANCA','DANIEL ALCIDES CARRI','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190202','CHACAYAN','DANIEL ALCIDES CARRI','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190203','GOYLLARISQUIZGA','DANIEL ALCIDES CARRI','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190204','PAUCAR','DANIEL ALCIDES CARRI','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190205','SAN PEDRO DE PILLAO','DANIEL ALCIDES CARRI','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190206','SANTA ANA DE TUSI','DANIEL ALCIDES CARRI','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190207','TAPUC','DANIEL ALCIDES CARRI','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190208','VILCABAMBA','DANIEL ALCIDES CARRI','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190301','OXAPAMPA','OXAPAMPA','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190302','CHONTABAMBA','OXAPAMPA','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190303','HUANCABAMBA','OXAPAMPA','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190304','PALCAZU','OXAPAMPA','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190305','POZUZO','OXAPAMPA','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190306','PUERTO BERMUDEZ','OXAPAMPA','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190307','VILLA RICA','OXAPAMPA','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('190308','CONSTITUCIÓN','OXAPAMPA','PASCO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200101','PIURA','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200104','CASTILLA','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200105','CATACAOS','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200107','CURA MORI','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200108','EL TALLAN','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200109','LA ARENA','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200110','LA UNION','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200111','LAS LOMAS','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200114','TAMBO GRANDE','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200115','26 DE OCTUBRE','PIURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200201','AYABACA','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200202','FRIAS','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200203','JILILI','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200204','LAGUNAS','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200205','MONTERO','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200206','PACAIPAMPA','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200207','PAIMAS','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200208','SAPILLICA','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200209','SICCHEZ','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200210','SUYO','AYABACA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200301','HUANCABAMBA','HUANCABAMBA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200302','CANCHAQUE','HUANCABAMBA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200303','EL CARMEN DE LA FRONTERA','HUANCABAMBA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200304','HUARMACA','HUANCABAMBA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200305','LALAQUIZ','HUANCABAMBA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200306','SAN MIGUEL DE EL FAIQUE','HUANCABAMBA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200307','SONDOR','HUANCABAMBA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200308','SONDORILLO','HUANCABAMBA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200401','CHULUCANAS','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200402','BUENOS AIRES','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200403','CHALACO','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200404','LA MATANZA','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200405','MORROPON','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200406','SALITRAL','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200407','SAN JUAN DE BIGOTE','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200408','SANTA CATALINA DE MOSSA','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200409','SANTO DOMINGO','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200410','YAMANGO','MORROPON','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200501','PAITA','PAITA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200502','AMOTAPE','PAITA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200503','ARENAL','PAITA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200504','COLAN','PAITA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200505','LA HUACA','PAITA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200506','TAMARINDO','PAITA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200507','VICHAYAL','PAITA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200601','SULLANA','SULLANA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200602','BELLAVISTA','SULLANA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200603','IGNACIO ESCUDERO','SULLANA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200604','LANCONES','SULLANA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200605','MARCAVELICA','SULLANA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200606','MIGUEL CHECA','SULLANA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200607','QUERECOTILLO','SULLANA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200608','SALITRAL','SULLANA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200701','PARIÑAS','TALARA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200702','EL ALTO','TALARA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200703','LA BREA','TALARA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200704','LOBITOS','TALARA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200705','LOS ORGANOS','TALARA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200706','MANCORA','TALARA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200801','SECHURA','SECHURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200802','BELLAVISTA DE LA UNION','SECHURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200803','BERNAL','SECHURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200804','CRISTO NOS VALGA','SECHURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200805','VICE','SECHURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('200806','RINCONADA LLICUAR','SECHURA','PIURA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210101','PUNO','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210102','ACORA','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210103','AMANTANI','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210104','ATUNCOLLA','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210105','CAPACHICA','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210106','CHUCUITO','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210107','COATA','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210108','HUATA','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210109','MAÑAZO','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210110','PAUCARCOLLA','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210111','PICHACANI','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210112','PLATERIA','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210113','SAN ANTONIO','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210114','TIQUILLACA','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210115','VILQUE','PUNO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210201','AZANGARO','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210202','ACHAYA','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210203','ARAPA','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210204','ASILLO','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210205','CAMINACA','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210206','CHUPA','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210207','JOSE DOMINGO CHOQUEHUANCA','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210208','MUÑANI','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210209','POTONI','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210210','SAMAN','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210211','SAN ANTON','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210212','SAN JOSE','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210213','SAN JUAN DE SALINAS','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210214','SANTIAGO DE PUPUJA','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210215','TIRAPATA','AZANGARO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210301','MACUSANI','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210302','AJOYANI','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210303','AYAPATA','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210304','COASA','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210305','CORANI','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210306','CRUCERO','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210307','ITUATA','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210308','OLLACHEA','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210309','SAN GABAN','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210310','USICAYOS','CARABAYA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210401','JULI','CHUCUITO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210402','DESAGUADERO','CHUCUITO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210403','HUACULLANI','CHUCUITO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210404','KELLUYO','CHUCUITO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210405','PISACOMA','CHUCUITO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210406','POMATA','CHUCUITO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210407','ZEPITA','CHUCUITO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210501','ILAVE','EL COLLAO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210502','CAPAZO','EL COLLAO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210503','PILCUYO','EL COLLAO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210504','SANTA ROSA','EL COLLAO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210505','CONDURIRI','EL COLLAO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210601','HUANCANE','HUANCANE','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210602','COJATA','HUANCANE','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210603','HUATASANI','HUANCANE','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210604','INCHUPALLA','HUANCANE','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210605','PUSI','HUANCANE','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210606','ROSASPATA','HUANCANE','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210607','TARACO','HUANCANE','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210608','VILQUE CHICO','HUANCANE','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210701','LAMPA','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210702','CABANILLA','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210703','CALAPUJA','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210704','NICASIO','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210705','OCUVIRI','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210706','PALCA','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210707','PARATIA','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210708','PUCARA','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210709','SANTA LUCIA','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210710','VILAVILA','LAMPA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210801','AYAVIRI','MELGAR','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210802','ANTAUTA','MELGAR','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210803','CUPI','MELGAR','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210804','LLALLI','MELGAR','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210805','MACARI','MELGAR','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210806','NUÑOA','MELGAR','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210807','ORURILLO','MELGAR','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210808','SANTA ROSA','MELGAR','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210809','UMACHIRI','MELGAR','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210901','MOHO','MOHO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210902','CONIMA','MOHO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210903','HUAYRAPATA','MOHO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('210904','TILALI','MOHO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211001','PUTINA','SAN ANTONIO DE PUTIN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211002','ANANEA','SAN ANTONIO DE PUTIN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211003','PEDRO VILCA APAZA','SAN ANTONIO DE PUTIN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211004','QUILCAPUNCU','SAN ANTONIO DE PUTIN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211005','SINA','SAN ANTONIO DE PUTIN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211101','JULIACA','SAN ROMAN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211102','CABANA','SAN ROMAN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211103','CABANILLAS','SAN ROMAN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211104','CARACOTO','SAN ROMAN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211105','SAN MIGUEL','SAN ROMAN','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211201','SANDIA','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211202','CUYOCUYO','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211203','LIMBANI','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211204','PATAMBUCO','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211205','PHARA','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211206','QUIACA','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211207','SAN JUAN DEL ORO','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211208','YANAHUAYA','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211209','ALTO INAMBARI','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211210','SAN PEDRO DE PUTINA PUNCO','SANDIA','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211301','YUNGUYO','YUNGUYO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211302','ANAPIA','YUNGUYO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211303','COPANI','YUNGUYO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211304','CUTURAPI','YUNGUYO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211305','OLLARAYA','YUNGUYO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211306','TINICACHI','YUNGUYO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('211307','UNICACHI','YUNGUYO','PUNO')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220101','MOYOBAMBA','MOYOBAMBA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220102','CALZADA','MOYOBAMBA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220103','HABANA','MOYOBAMBA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220104','JEPELACIO','MOYOBAMBA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220105','SORITOR','MOYOBAMBA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220106','YANTALO','MOYOBAMBA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220201','BELLAVISTA','BELLAVISTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220202','ALTO BIAVO','BELLAVISTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220203','BAJO BIAVO','BELLAVISTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220204','HUALLAGA','BELLAVISTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220205','SAN PABLO','BELLAVISTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220206','SAN RAFAEL','BELLAVISTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220301','SAN JOSE DE SISA','EL DORADO','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220302','AGUA BLANCA','EL DORADO','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220303','SAN MARTIN','EL DORADO','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220304','SANTA ROSA','EL DORADO','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220305','SHATOJA','EL DORADO','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220401','SAPOSOA','HUALLAGA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220402','ALTO SAPOSOA','HUALLAGA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220403','EL ESLABON','HUALLAGA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220404','PISCOYACU','HUALLAGA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220405','SACANCHE','HUALLAGA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220406','TINGO DE SAPOSOA','HUALLAGA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220501','LAMAS','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220502','ALONSO DE ALVARADO','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220503','BARRANQUITA','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220504','CAYNARACHI','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220505','CUÑUMBUQUI','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220506','PINTO RECODO','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220507','RUMISAPA','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220508','SAN ROQUE DE CUMBAZA','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220509','SHANAO','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220510','TABALOSOS','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220511','ZAPATERO','LAMAS','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220601','JUANJUI','MARISCAL CACERES','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220602','CAMPANILLA','MARISCAL CACERES','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220603','HUICUNGO','MARISCAL CACERES','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220604','PACHIZA','MARISCAL CACERES','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220605','PAJARILLO','MARISCAL CACERES','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220701','PICOTA','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220702','BUENOS AIRES','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220703','CASPISAPA','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220704','PILLUANA','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220705','PUCACACA','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220706','SAN CRISTOBAL','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220707','SAN HILARION','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220708','SHAMBOYACU','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220709','TINGO DE PONASA','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220710','TRES UNIDOS','PICOTA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220801','RIOJA','RIOJA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220802','AWAJUN','RIOJA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220803','ELIAS SOPLIN VARGAS','RIOJA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220804','NUEVA CAJAMARCA','RIOJA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220805','PARDO MIGUEL','RIOJA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220806','POSIC','RIOJA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220807','SAN FERNANDO','RIOJA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220808','YORONGOS','RIOJA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220809','YURACYACU','RIOJA','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220901','TARAPOTO','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220902','ALBERTO LEVEAU','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220903','CACATACHI','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220904','CHAZUTA','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220905','CHIPURANA','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220906','EL PORVENIR','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220907','HUIMBAYOC','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220908','JUAN GUERRA','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220909','LA BANDA DE SHILCAYO','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220910','MORALES','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220911','PAPAPLAYA','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220912','SAN ANTONIO','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220913','SAUCE','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('220914','SHAPAJA','SAN MARTIN','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('221001','TOCACHE','TOCACHE','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('221002','NUEVO PROGRESO','TOCACHE','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('221003','POLVORA','TOCACHE','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('221004','SHUNTE','TOCACHE','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('221005','UCHIZA','TOCACHE','SAN MARTIN')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230101','TACNA','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230102','ALTO DE LA ALIANZA','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230103','CALANA','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230104','CIUDAD NUEVA','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230105','INCLAN','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230106','PACHIA','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230107','PALCA','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230108','POCOLLAY','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230109','SAMA','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230110','CORONEL GREGORIO ALBARRACIN LANCHIPA','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230111','LA YARADA-LOS PALOS','TACNA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230201','CANDARAVE','CANDARAVE','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230202','CAIRANI','CANDARAVE','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230203','CAMILACA','CANDARAVE','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230204','CURIBAYA','CANDARAVE','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230205','HUANUARA','CANDARAVE','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230206','QUILAHUANI','CANDARAVE','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230301','LOCUMBA','JORGE BASADRE','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230302','ILABAYA','JORGE BASADRE','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230303','ITE','JORGE BASADRE','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230401','TARATA','TARATA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230402','HEROES ALBARRACIN','TARATA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230403','ESTIQUE','TARATA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230404','ESTIQUE-PAMPA','TARATA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230405','SITAJARA','TARATA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230406','SUSAPAYA','TARATA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230407','TARUCACHI','TARATA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('230408','TICACO','TARATA','TACNA')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240101','TUMBES','TUMBES','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240102','CORRALES','TUMBES','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240103','LA CRUZ','TUMBES','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240104','PAMPAS DE HOSPITAL','TUMBES','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240105','SAN JACINTO','TUMBES','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240106','SAN JUAN DE LA VIRGEN','TUMBES','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240201','ZORRITOS','CONTRALMIRANTE VILLA','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240202','CASITAS','CONTRALMIRANTE VILLA','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240203','CANOAS DE PUNTA SAL','CONTRALMIRANTE VILLA','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240301','ZARUMILLA','ZARUMILLA','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240302','AGUAS VERDES','ZARUMILLA','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240303','MATAPALO','ZARUMILLA','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('240304','PAPAYAL','ZARUMILLA','TUMBES')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250101','CALLERIA','CORONEL PORTILLO','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250102','CAMPOVERDE','CORONEL PORTILLO','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250103','IPARIA','CORONEL PORTILLO','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250104','MASISEA','CORONEL PORTILLO','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250105','YARINACOCHA','CORONEL PORTILLO','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250106','NUEVA REQUENA','CORONEL PORTILLO','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250107','MANANTAY','CORONEL PORTILLO','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250201','RAYMONDI','ATALAYA','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250202','SEPAHUA','ATALAYA','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250203','TAHUANIA','ATALAYA','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250204','YURUA','ATALAYA','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250301','PADRE ABAD','PADRE ABAD','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250302','IRAZOLA','PADRE ABAD','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250303','CURIMANA','PADRE ABAD','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250304','NESHUYA','PADRE ABAD','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250305','ALEXANDER VON HUMBOLDT','PADRE ABAD','UCAYALI')
		insert into tubigeo (tCodigo,tdistrito,tprovincia,tdepartamento ) values ('250401','PURUS','PURUS','UCAYALI')
GO
	delete from TTABLA where tTabla='FRECUENCIA'
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','00','Diario','Diario',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','01','Lunes','Lunes',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','02','Martes','Martes',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','03','Miercoles','Miercoles',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','04','Jueves','Jueves',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','05','Viernes','Viernes',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','06','Sabados','Sabados',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','07','Domingos','Domingo',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('FRECUENCIA','99','Fecha Especial','Fecha Espcial',0,1)
GO
	delete from TTABLA where tTabla='ESTADODOCUMENTO'
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nValor, nTamano,lActivo) values ('ESTADODOCUMENTO','01','EMITIDO','EMITIDO',0,2,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nValor, nTamano,lActivo) values ('ESTADODOCUMENTO','02','PAGADO','PAGADO',0,2,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nValor, nTamano,lActivo) values ('ESTADODOCUMENTO','03','POR COBRAR','POR COBRAR',0,2,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nValor, nTamano,lActivo) values ('ESTADODOCUMENTO','04','ANULADO','ANULADO',0,2,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nValor, nTamano,lActivo) values ('ESTADODOCUMENTO','05','PROCESADO','PROCESADO',0,2,1)
GO
	delete FROM TTABLA where TTABLA='MOTIVOELIMINACION' AND tCodigo='000'
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('MOTIVOELIMINACION','000','Otro','Otro',0,1)
GO
	delete FROM TTABLA where TTABLA='MOTIVODESCUENTO' AND tCodigo='000'
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, nValor, lActivo) values ('MOTIVODESCUENTO','000','Otro','Otro',0,0,1)
GO
	if (select count(tCodigo) from TTABLA where TTABLA='TIPOCANCELACION' AND tCodigo='000') = 0 
	Begin
		delete FROM TTABLA where TTABLA='TIPOCANCELACION' AND tCodigo='000'
		INSERT TTABLA (TTABLA, TCODIGO, NTAMANO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCANCELACION','000',2,'Otro','Otro',0,1)
	end
GO
	if (select count(tCodigo) from TTABLA where TTABLA='TIPOCANCELACION' AND tCodigo='001') = 0 
	Begin
		delete FROM TTABLA where TTABLA='TIPOCANCELACION' AND tCodigo='001'
		INSERT TTABLA (TTABLA, TCODIGO, NTAMANO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCANCELACION','001',2,'Recibo Ingreso','Recibo Ingreso',0,1)
	end
GO
	if (select count(tCodigo) from TTABLA where TTABLA='TIPOCANCELACION' AND tCodigo='002') = 0 
	Begin
		delete FROM TTABLA where TTABLA='TIPOCANCELACION' AND tCodigo='002'
		INSERT TTABLA (TTABLA, TCODIGO, NTAMANO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCANCELACION','002',2,'Nota de Crédito','Nota de Crédito',0,1)
	end
GO
	if (select count(tCodigo) from TTABLA where tTabla='PAISORIGEN' and TCODIGO='000') = 0 
		Begin
			insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO,NVALOR, LACTIVO) 
			VALUES('PAISORIGEN', '000',2,'PERU','PERU',1,1)
		end
GO
	if (select count(tCodigo) from TTABLA where tTabla='PAISORIGEN' and TCODIGO='001') = 0 
		Begin
			insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO,NVALOR, LACTIVO) 
			VALUES('PAISORIGEN', '001',2,'BOLIVIA','BOLIVIA',0,1)
		end
GO
	if (select count(tCodigo) from TTABLA where tTabla='PAISORIGEN' and TCODIGO='002') = 0 
		Begin
			insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO,NVALOR, LACTIVO) 
			VALUES('PAISORIGEN', '002',2,'ECUADOR','ECUADOR',0,1)
		end
GO
	if (select count(tCodigo) from TTABLA where tTabla='PAISORIGEN' and TCODIGO='003') = 0 
		Begin
			insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO,NVALOR, LACTIVO) 
			VALUES('PAISORIGEN', '003',2,'ARGENTINA','ARGENTINA',0,1)
		end
GO
	if (select count(tCodigo) from TTABLA where tTabla='PAISORIGEN' and TCODIGO='004') = 0 
		Begin
			insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO,NVALOR, LACTIVO) 
			VALUES('PAISORIGEN', '004',2,'CHILE','CHILE',0,1)
		end
GO
	if (select count(tCodigo) from TTABLA where ttabla='PAISORIGEN' and TCODIGO='005')=0
		Begin
			insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO,NVALOR, LACTIVO) 
			VALUES('PAISORIGEN','005',2,'ESPAÑA','ESPAÑA',0,1)
		end
GO
IF (SELECT TCODIGO FROM  TTABLA where tTabla='PAISORIGEN' AND lActivo = 1 AND nValor = 1) <> '000'
BEGIN
	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='003' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 1' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='003'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='003' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','003','','',0,1,'Otro 1')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='004' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 2' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='004'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='004' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','004','','',0,1,'Otro 2')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='005' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 3' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='005'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='005' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','005','','',0,1,'Otro 3')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='006' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 4' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='006'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='006' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','006','','',0,1,'Otro 4')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='007' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 5' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='007'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='007' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','007','','',0,1,'Otro 5')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='008' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 6' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='008'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='008' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','008','','',0,1,'Otro 6')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='009' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 7' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='009'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='009' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','009','','',0,1,'Otro 7')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='010' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 8' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='010'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='010' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','010','','',0,1,'Otro 8')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='011' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 9' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='011'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='011' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','011','','',0,1,'Otro 9')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='012' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 10' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='012'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='012' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','012','','',0,1,'Otro 10')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='013' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 11' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='013'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='013' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','013','','',0,1,'Otro 11')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='014' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 12' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='014'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='014' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','014','','',0,1,'Otro 12')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='015' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 13' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='015'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='015' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','015','','',0,1,'Otro 13')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='016' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 14' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='016'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='016' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','016','','',0,1,'Otro 14')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='017' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 15' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='017'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='017' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','017','','',0,1,'Otro 15')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='018' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 16' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='018'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='018' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','018','','',0,1,'Otro 16')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='019' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 17' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='019'
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='019' ) = 0
		begin
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','019','','',0,1,'Otro 17')
		end

	if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='020' ) > 0
		begin
			update TTABLA set tValor2 ='Otro 18' WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='020'
		end
END
GO
if (select count(TCODIGO) from TTABLA WHERE TTABLA = N'TIPOCANCELACION' and TCODIGO ='020' ) = 0
begin
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo,tValor2) values ('TIPOCANCELACION','020','','',0,1,'Otro 18')
end
GO
	delete FROM TTABLA where TTABLA='ESTADOGUIA'
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOGUIA','01','EMITIDO','EMITIDO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOGUIA','02','FACTURADO','FACTURADO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOGUIA','03','ANULADO','ANULADO',0,1)
GO
	delete FROM TTABLA where TTABLA='ESTADOPEDIDO' AND tCodigo='01'
	delete FROM TTABLA where TTABLA='ESTADOPEDIDO' AND tCodigo='02'
	delete FROM TTABLA where TTABLA='ESTADOPEDIDO' AND tCodigo='03'
	delete FROM TTABLA where TTABLA='ESTADOPEDIDO' AND tCodigo='04'
	delete FROM TTABLA where TTABLA='ESTADOPEDIDO' AND tCodigo='05'

	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOPEDIDO','01','EN PROCESO','EN PROCESO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOPEDIDO','02','FACTURADO','FACTURADO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOPEDIDO','03','ANULADO','ANULADO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOPEDIDO','04','CUENTA CORRIENTE','CUENTA CORRIENTE',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOPEDIDO','05','CARGADO','CARGADO',0,1)
GO

	--delete FROM TTABLA where TTABLA='TIPOPAGO' and TCODIGO='04'
	--delete FROM TTABLA where TTABLA='TIPOPAGO' and TCODIGO='05'
	--INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPAGO','04','Varios','Varios',0,1)
	--INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPAGO','05','Puntos','Puntos',0,1)
GO
	delete FROM TTABLA where TTABLA='TIPOCLIENTEFRECUENTE' and TCODIGO='00'
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCLIENTEFRECUENTE','00','SIN TIPO','SIN TIPO',0,1)
GO
	DELETE FROM TTABLA WHERE TTABLA='TIPOUBICACION'
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '01',2,'Avenida','AV',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '02',2,'Jiron','JR',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '03',2,'Calle','CL',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '04',2,'Pasaje','PJ',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '05',2,'Alameda','AL',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '06',2,'Malecon','MA',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '07',2,'Ovalo','OV',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '08',2,'Parque','PQ',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '09',2,'Plaza','PL',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '10',2,'Carretera','CA',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '11',2,'Plazuela','PL',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '12',2,'Bajada','BA',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '13',2,'Conjunto Residencial','RS',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '14',2,'Barrio Fiscal','BF',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '15',2,'Urbanizacion','UR',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '16',2,'Urbanizacion Popular','UP',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '17',2,'Cooperativa','CO',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '18',2,'Villa Militar','VM',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '19',2,'Unidad Vecinal','UV',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '20',2,'Asociacion','AS',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '21',2,'Paseo','PS',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '22',2,'Prolongacion','PR',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '23',2,'Asentamiento Humano','AH',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '24',2,'Agrupacion','AG',1)
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('TIPOUBICACION', '25',2,'Otros','OT',1)
GO
	DELETE FROM TTABLA WHERE TTABLA='UNIDADNEGOCIO' AND TCODIGO='00'
	insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('UNIDADNEGOCIO', '00',2,'SIN UNIDAD NEGOCIO','SIN UNIDAD NEGOCIO',1)
GO
	delete from TACCESO
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000001','03','Boton Platos','mdiAdministracion','BT','cmdopcion1',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000002','03','Boton Clientes','mdiAdministracion','BT','cmdopcion2',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000003','03','Boton Mesas','mdiAdministracion','BT','cmdopcion3',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000004','03','Boton Usuarios','mdiAdministracion','BT','cmdopcion4',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000005','03','Boton Backup','mdiAdministracion','BT','cmdopcion5',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000006','03','Boton Restore','mdiAdministracion','BT','cmdopcion6',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000007','04','Vistas de Mesas','mdiConsulta','BT','cmdopcion1',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000008','04','Coorelativo de Pedidos','mdiConsulta','BT','cmdopcion2',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000009','04','Coorelativo de Documentos','mdiConsulta','BT','cmdopcion3',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000010','04','Liquidacion de Cajero','mdiConsulta','BT','cmdopcion4',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000011','04','Registro de Ventas	','mdiConsulta','BT','cmdopcion5',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000012','04','Paloteo de Produccion','mdiConsulta','BT','cmdopcion6',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000013','04','Propinas','mdiConsulta','BT','cmdopcion7',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000014','02','Apertura','mdiPuntoVenta','BT','cmdopcion1',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000015','02','Punto de Venta','mdiPuntoVenta','BT','cmdopcion2',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden)  values ('00000016','02','Cierre','mdiPuntoVenta','BT','cmdopcion3',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000017','02','Mesas','mdiPuntoVenta','BT','cmdopcion4',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000018','02','Recibo de Egresos','mdiPuntoVenta','BT','cmdopcion5',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000019','02','Recibo de Ingresos','mdiPuntoVenta','BT','cmdopcion6',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000020','02','Reservas','mdiPuntoVenta','BT','cmdopcion7',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000021','02','Correlativo de Pedidos','mdiPuntoVenta','BT','cmdopcion8',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000022','02','Correlativo de Documentos','mdiPuntoVenta','BT','cmdopcion9',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000023','02','Cuentas Corrientes','mdiPuntoVenta','BT','cmdopcion10',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000024','02','Cuentas por Cobrar','mdiPuntoVenta','BT','cmdopcion11',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000025','02','Carta de Productos','mdiPuntoVenta','BT','cmdopcion12',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000026','02','Delivery en Tránsito','mdiPuntoVenta','BT','cmdopcion13',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('00000027','02','Delivery Entregados','mdiPuntoVenta','BT','cmdopcion14',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10100000','03','Configuración','mdiAdministracion','MN','mnuConfiguracion',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10110000','03','Parámetros Generales','mdiAdministracion','MN','mnuParametro',NULL,12)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10120000','03','Canal de Venta','mdiAdministracion','MN','mnuTipoPedido',NULL,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10130000','03','Mantenimientos de Impresora','mdiAdministracion','MN','mnuImpresora',null,6)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10140000','03','Configuración de Cajas','mdiAdministracion','MN','mnuConfiguraCaja',null,2)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10150000','03','Configuración de Impresoras','mdiAdministracion','MN','mnuConfiguracionImpresora',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10160000','03','Configura Mensaje','mdiAdministracion','MN','mnuMensaje','TMENSAJE',7)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10161000','03','Mantenimiento de Establecimientos (Locales)','mdiAdministracion','MN','mnuManteEstable',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10162000','03','Mantenimiento de Agrupacion de Puntos de Ventas','mdiAdministracion','MN','mnuMantSectorVentas',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10170000','03','Cierre de Periodo','mdiAdministracion','MN','mnuCierre',null,35)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10190000','03','Actualización de Tablas','mdiAdministracion','MN','mnuTablaReplica',null,null)
	--CESAR
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10175000','03','Usuarios','mdiAdministracion','MN','mnuUsuario','TUSUARIO',17)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10175010','03','Opción Agregar','frmUsuario','MN','cmdopcion0','TUSUARIO',null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10175020','03','Opción Modificar','frmUsuario','MN','cmdopcion1','TUSUARIO',null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10175030','03','Opción Eliminar','frmUsuario','MN','cmdOpcion2','TUSUARIO',null)
	
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10180000','03','Grupo de Usuarios','mdiAdministracion','MN','mnuGrupoUsuario','GRUPOUSUARIO',1)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10185000','03','Tipos de Cambio','mdiAdministracion','MN','mnuTipoCambio','TTIPOCAMBIO',16)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10200000','03','Tablas','mdiAdministracion','MN','mnuTabla',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10201000','03','Tipos de Identidad','mdiAdministracion','MN','mnuIdentidad','TTABLA-TIPOIDENTIDAD',28)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10202000','03','Tipos de Documento','mdiAdministracion','MN','mnuTipoDocumento','TTABLA-TIPODOCUMENTO',27)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10203000','03','Tipos de Clientes Frecuentes','mdiAdministracion','MN','mnuTipocliente',null,26)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10204000','03','Tipos de Cuenta Corriente','mdiAdministracion','MN','mnuTipoCtaCte',null,24)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10205000','03','Cuentas Contables de Cancelación','mdiAdministracion','MN','mnuCuentaContable','TTABLA-TIPOPAGO',31)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10206000','03','Otros Tipos de Cancelación','mdiAdministracion','MN','mnuCancelacion','TTABLA-TIPOCANCELACION',25)
	--INSERT INTO TACCESO values ('10207000','03','Clientes Cuentas Corrientes','mdiAdministracion','MN','mnuCliente',null,4)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10208000','03','Clientes Frecuentes','mdiAdministracion','MN','mnuDelivery',null,4)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10209000','03','Clientes Facturados','mdiAdministracion','MN','mnuClienteFactura',null,3)
	--INSERT INTO TACCESO values ('10209500','03','Maitres','mdiAdministracion','MN','mnuMaitre','TTABLA-MAITRE',39)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10210000','03','Mozos','mdiAdministracion','MN','mnuMozo','TTABLA-MOZO',22)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10211000','03','Motorizados','mdiAdministracion','MN','mnuMotorizados','TTABLA-MOTORIZADO',23)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10212000','03','Empacadores','mdiAdministracion','MN','mnuEmpacador','TTABLA-EMPACADOR',32)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10213000','03','Zonas','mdiAdministracion','MN','mnuZona','TTABLA-ZONA',33)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10214000','03','Distritos','mdiAdministracion','MN','mnuDistritos','TTABLA-DISTRITO',20)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10215000','03','Mesas','mdiAdministracion','MN','mnuMesas','TMESA',8)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10216000','03','Motivos de Cortesías','mdiAdministracion','MN','mnuCortesia','TTABLA-CORTESIA',19)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10217000','03','Motivos de Eliminación','mdiAdministracion','MN','mnuEliminacion','TTABLA-MOTIVOELIMINACION',21)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10218000','03','Motivos de Descuentos','mdiAdministracion','MN','mnuDescuento','TMOTIVODESCUENTO',9)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10218100','03','Estado de Clientes Frecuentes','mdiAdministracion','MN','mnuEstadoClienteFrecuente','TTABLA-ESTADOFRECUENTE',38)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10218200','03','Tipos de Egreso','mdiAdministracion','MN','mnuTipoEgreso',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10219000','03','Tarjetas Bancarias','mdiAdministracion','MN','mnuTarjetaCredito','TTARJETACREDITO',15)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10220000','03','Areas de Producción','mdiAdministracion','MN','mnuArea','TTABLA-AREA',34)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10300000','03','Productos de Venta','mdiAdministracion','MN','mnuProd',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10301000','03','Tipos de Producto','mdiAdministracion','MN','mnuTipoProducto','TTABLA-TIPOPRODUCTO',29)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10301500','03','Unidad de Negocio','mdiAdministracion','MN','mnuUnidadNegocio','TTABLA-UNIDADNEGOCIO',30)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10301550','03','Sucursales','mdiAdministracion','MN','mnuSucursales','TTABLA-SUCURSAL',36)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10302000','03','Operadores','mdiAdministracion','MN','mnuOperador','TOPERADOR',11)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10303000','03','Propiedades','mdiAdministracion','MN','mnuPropiedad','TPROPIEDAD',14)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10304000','03','Grupos y SubGrupos','mdiAdministracion','MN','mnuGrupo','GRUPO',5)

	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10305000','03','Productos y Precios','mdiAdministracion','MN','mnuProducto','PRODUCTO',13)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10305010','03','Opción Agregar','frmProducto','MN','cmdopcion0','PRODUCTO', null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10305020','03','Opción Modificar','frmProducto','MN','cmdopcion1','PRODUCTO', null)

	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10305500','03','Agrupacion de Caja Rápida','mdiAdministracion','MN','mnuAgrupacion','AGRUPACION',18)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10305600','03','Insumos/Platos de Stock Crítico','mdiAdministracion','MN','mnuInsumoCritico','TINSUMO',37)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10306000','03','Ofertas','mdiAdministracion','MN','mnuOferta','TOFERTA',10)
	--INSERT INTO TACCESO values ('10308000','03','Equivalencia de Productos','mdiAdministracion','MN','mnuEquivalencias','TPRODUCTOXPRODUCTO',15)
	--INSERT INTO TACCESO values ('10308000','03','Clientes Frecuentes','mdiAdministracion','MN','mnuDelivery',null,15)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10400000','03','Utilitario','mdiAdministracion','MN','mnuUtilitario',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10401000','03','Optimizador de la BD','mdiAdministracion','MN','mnuOptimizador',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10402000','03','Backup','mdiAdministracion','MN','mnuBackup',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10403000','03','Restore','mdiAdministracion','MN','mnuRestore',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10404000','03','Transferencia Almacén','mdiAdministracion','MN','mnuAlmacen',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10406000','03','Cambiar de Local','mdiAdministracion','MN','mnuCambiaLocal',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('10405000','03','Eliminar Cortesía','mdiAdministracion','MN','mnuElimina',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20100000','04','Correlativos','mdiConsulta','MN','mnuCuentas',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20101000','04','Pedidos','mdiConsulta','MN','mnuCorrelativoPedido',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20102000','04','Documentos','mdiConsulta','MN','mnuCorrelativoDocumento',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20103000','04','Cuentas Corrientes','mdiConsulta','MN','mnuCtaCte',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20104000','04','Cuentas por Cobrar','mdiConsulta','MN','mnuCuentaCobrar',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20105000','04','Notas de Crédito','mdiConsulta','MN','mnuNotaCredito',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20106000','04','Reservas','mdiConsulta','MN','mnuReserva',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20107000','04','Recibos de Egreso','mdiConsulta','MN','mnuRecibo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20108000','04','Recibos de Ingreso','mdiConsulta','MN','mnuIngreso',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20109000','04','Turnos','mdiConsulta','MN','mnuTurno',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20200000','04','Reportes','mdiConsulta','MN','mnuReporte',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201000','04','De Control','mdiConsulta','MN','mnuControl',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201010','04','Liquidacion de Cajero','mdiConsulta','MN','mnuLiquidacion',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201015','04','Paloteo de Producción','mdiConsulta','MN','mnuPaloteo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201020','04','Paloteo de Propiedades','mdiConsulta','MN','mnuPropiedades',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201025','04','Paloteo de Insumos','mdiConsulta','MN','mnuPaloteoInsumo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201026','04','Paloteo de Equivalencias','mdiConsulta','MN','mnuRepEquivalencias',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201027','04','Paloteo de Ofertas','mdiConsulta','MN','mnuPaloteoOfertas',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201028','04','Paloteo de Productos por Meses','mdiConsulta','MN','mnuPaloteoProductoMes',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201030','04','Comandas','mdiConsulta','MN','mnuComanda',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201035','04','Pedidos','mdiConsulta','MN','mnuCorrelativo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201040','04','Propinas','mdiConsulta','MN','mnuPropina',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201045','04','Cortesias','mdiConsulta','MN','mnuCortesia',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201050','04','Descuentos','mdiConsulta','MN','mnuDescuento',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201055','04','Cuentas Corrientes','mdiConsulta','MN','mnuRepCtaCte',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201060','04','Cuentas por Cobrar','mdiConsulta','MN','mnuClienteDeuda',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201065','04','Contactos','mdiConsulta','MN','mnuContacto',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201070','04','Control de Transacciones','mdiConsulta','MN','mnuAnulado',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20201075','04','Diferencias entre Paloteo vs Liquidacion','mdiConsulta','MN','mnuDiferencias',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20202000','04','Contables','mdiConsulta','MN','mnuContables',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20202010','04','Registro de Ventas','mdiConsulta','MN','mnuRegistroVenta',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20202020','04','Principales Clientes','mdiConsulta','MN','mnuPrincipal',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20202030','04','Cobranzas','mdiConsulta','MN','mnuCobranza',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203000','04','Estadisticos','mdiConsulta','MN','mnuEstadistico',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203010','04','Ranking de Produccion','mdiConsulta','MN','mnuRanking',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203012','04','Registro(Central de Pedidos)','mdiConsulta','MN','mnuRepEntregasReg',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203014','04','Entregas (Central de Pedidos)','mdiConsulta','MN','mnuRepEntregas',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203015','04','Analítico de Productos por Mozos','mdiConsulta','MN','mnuMozo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203020','04','Analítico de Productos por Motorizados','mdiConsulta','MN','mnuMotorizado',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203030','04','Analítico de Clientes Frecuentes','mdiConsulta','MN','mnuFrecuente',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203035','04','Producción por Mozos','mdiConsulta','MN','mnuEstPropina',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203036','04','Planilla de Motorizados','mdiConsulta','MN','mnuPlanillaMotorizados',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203040','04','Tiempos en Salon','mdiConsulta','MN','mnuTiempoSalon',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203045','04','Tiempos Delivery','mdiConsulta','MN','mnuTiempoDelivery',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203046','04','Tiempos KDS','mdiConsulta','MN','mnuTiempoKDS',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203047','04','Tiempos Chef Control','mdiConsulta','MN','mnuTiempoChefControl',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203050','04','Diferencias de Tiempos Delivery','mdiConsulta','MN','mnuDiferenciaDelivery',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203055','04','Rotación de Mesas','mdiConsulta','MN','mnuRotacion',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203060','04','Ocupabilidad de Mesas','mdiConsulta','MN','mnuOcupabilidad',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203065','04','Paloteo Comparativo','mdiConsulta','MN','mnuPaloteoComparativo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20203070','04','Resultados Operativos de Ventas','mdiConsulta','MN','mnuResultadoOperativo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20204000','04','Gerencial','mdiConsulta','MN','mnuAnalitico',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20204010','04','Venta Anual por Meses','mdiConsulta','MN','mnuVenta',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20204020','04','Venta Mensual por Fechas','mdiConsulta','MN','mnuVentasFechas',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20204030','04','Venta Comparativa Anual','mdiConsulta','MN','mnuVentaComparada',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20204040','04','Venta Comparativo Mensual','mdiConsulta','MN','mnuVentaComparadaMensual',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20204050','04','Venta por Turnos','mdiConsulta','MN','mnuVentaTurno',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20204060','04','Cobranza Mensual por Fechas','mdiConsulta','MN','mnuCobranzaFecha',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20300000','04','Ticketera	','mdiConsulta','MN','mnuTicketera',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20301000','04','Liquidación de Cajero','mdiConsulta','MN','mnuLiquidacionTicket',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20302000','04','Paloteo de Producción','mdiConsulta','MN','mnuPaloteoTicket',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20500000','04','Conexión	','mdiConsulta','MN','mnuConexion',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('20501000','04','Cambiar de Local','mdiConsulta','MN','mnuCambiaLocal',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30100000','02','Movimientos','mdiConsulta','MN','mnuMovimiento',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30101000','02','Apertura de Turno','mdiPuntoVenta','MN','mnuInicio',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30102000','02','Punto de Venta','mdiPuntoVenta','MN','mnuVenta',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30103000','02','Activar PinPad','mdiPuntoVenta','MN','mnuPinPad',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30104000','02','Cierre de Turno','mdiPuntoVenta','MN','mnuCierre',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30105000','02','Enumeración de Documentos','mdiPuntoVenta','MN','mnuCorrelativo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30106000','02','Mesas','mdiPuntoVenta','MN','mnuMesa',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30106100','02','Insumos/Platos de Stock Crítico','mdiPuntoVenta','MN','mnuInsumoCritico',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30200000','02','Correlativos','mdiPuntoVenta','MN','mnuCorrelativo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30201000','02','Correlativo de Pedidos','mdiPuntoVenta','MN','mnuCorrelativoPedido',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30202000','02','Correlativo de Documentos','mdiPuntoVenta','MN','mnuCorrelativoDocumento',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30203000','02','Cuentas Corrientes','mdiPuntoVenta','MN','mnuCtaCte',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30204000','02','Recibo de Egresos','mdiPuntoVenta','MN','mnuRecibo',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30204010','02','Opción Modificar','frmReciboEgresoDetalle','MN','cmdopcion1','MEGRESO',null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30205000','02','Recibo de Ingresos','mdiPuntoVenta','MN','mnuReciboIngreso',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30206000','02','Notas de Crédito','mdiPuntoVenta','MN','mnuNotaCredito',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30207000','02','Reservas','mdiPuntoVenta','MN','mnuReserva',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30207100','02','Pedidos Central Producción','mdiPuntoVenta','MN','mnuCorrelativoCentralPedidos',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30208000','02','Cuentas por Cobrar','mdiPuntoVenta','MN','mnuCuentaCobrar',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30300000','02','Conexión','mdiPuntoVenta','MN','mnuConexion',null,null)
	INSERT INTO TACCESO(tCodigoAcceso,tModulo,tDescripcion,tFormulario,tTipoObjeto,tNombreObjeto,tTabla,nOrden) values ('30301000','02','Cambiar de Local','mdiPuntoVenta','MN','mnuCambiaLocal',null,null)
GO
	if (select count(tCodigo) from TTABLA where tTabla='ETIQUETA') = 0 
		BEGIN
			insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('ETIQUETA', '01',2,'','',1)
			insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('ETIQUETA', '02',2,'','',1)
			insert into TTABLA (TTABLA, TCODIGO, NTAMANO, TDETALLADO, TRESUMIDO, LACTIVO) VALUES('ETIQUETA', '03',2,'','',1)
		END
GO
	if (select count(tCodigo) from TTABLA where tTabla='CAJARAPIDA') = 0 -- select *from TTABLA
		BEGIN
		   insert into TTABLA (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor,	lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','11',null,'','',null,1,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','12',null,'','',null,2,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','13',null,'','',null,3,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','14',null,'','',null,4,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','15',null,'','',null,5,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','16',null,'','',null,6,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','17',null,'','',null,7,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','18',null,'','',null,8,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','19',null,'','',null,9,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','21',null,'','',null,1,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','22',null,'','',null,2,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','23',null,'','',null,3,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','24',null,'','',null,4,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','25',null,'','',null,5,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','26',null,'','',null,6,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','27',null,'','',null,7,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','28',null,'','',null,8,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','29',null,'','',null,9,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','31',null,'','',null,1,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','32',null,'','',null,2,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','33',null,'','',null,3,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','34',null,'','',null,4,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','35',null,'','',null,5,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','36',null,'','',null,6,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','37',null,'','',null,7,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','38',null,'','',null,8,'',0,1,null,'',0
		   insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'CAJARAPIDA','39',null,'','',null,9,'',0,1,null,'',0
		end
GO
	if (select count(tCodigo) from TTABLA where tTabla='SALON') = 0  
		BEGIN
			insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'SALON','01',null,'SALON 01','SALON 01',null,0,'001',0,1,null,null,null
			insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'SALON','02',null,'SALON 02','SALON 02',null,1,'001',0,1,null,null,null
		END
GO
	if (select count(TCODIGO) from TTABLA where  TTABLA='TIPOPRODUCTO')=0 
		begin
			insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'TIPOPRODUCTO','01',2   ,'ALIMENTOS','ALIMENTOS',null,null,'',0   ,1,1,null,null
			insert into ttabla (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo) select 'TIPOPRODUCTO','02',null,'BEBIDAS'   ,'BEBIDAS' ,null,null,'',null,1,1,null,null
		end
GO
	if (select count(TCODIGO) from TTABLA where  TTABLA='TIPODESCARGO')=0 
		begin
			INSERT TTABLA(TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor,lActivo,lReplica,tValor2,lAutorizaEnviosMozo,lNuevoModificado)VALUES('TIPODESCARGO','D',1,'Directo','Directo',NULL,0,NULL,0,1,0,NULL,NULL,NULL)
			INSERT TTABLA(TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor,lActivo,lReplica,tValor2,lAutorizaEnviosMozo,lNuevoModificado)VALUES('TIPODESCARGO','M',1,'Menu','Menu',NULL,0,NULL,0,1,0,NULL,NULL,NULL)
			INSERT TTABLA(TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor,lActivo,lReplica,tValor2,lAutorizaEnviosMozo,lNuevoModificado)VALUES('TIPODESCARGO','N',1,'SinDescargo','SinDescargo',NULL,0,NULL,0,1,0,NULL,NULL,NULL)
			INSERT TTABLA(TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor,lActivo,lReplica,tValor2,lAutorizaEnviosMozo,lNuevoModificado)VALUES('TIPODESCARGO','R',1,'Receta','Receta',NULL,0,NULL,0,1,0,NULL,NULL,NULL)
		end

GO
	if (select count(*) from TMESA ) = 0  
		BEGIN
			INSERT INTO TMESA ( tCodigoMesa,tDetallado,tResumido,tSalon,lFumador,tX,tY,nPersona,lActivo,tUsuario,fRegistro,tEstadoMesa,lReplica) 
			VALUES 	('000','Sin Mesa','Sin Mesa','00',1,0,0,1,1,'MASTER','2004-03-30 16:07:00',	'01',NULL)
		END
GO
	if (select count(TGRUPOUSUARIO) from TGRUPOUSUARIO where TGRUPOUSUARIO='00')=0 
		begin
			insert into TGRUPOUSUARIO 
			(tGrupoUsuario,	tDetallado,	lModulo01,	lModulo02,	lModulo03,	lOpcion01,	lOpcion02,	lOpcion03,	
			lOpcion04,	lOpcion05,	lOpcion06,	lOpcion07,	lOpcion08,	lOpcion09,	lOpcion10,	lOpcion11,	lOpcion12,
			lOpcion13,	lOpcion14,	lOpcion15,	lOpcion16,	lOpcion17,	lOpcion18,	lActivo,	lOpcion19,	lOpcion20,	
			lOpcion21,	lModulo04,	lModulo05,	tNivel,	lControlNivel,	lOpcion22,	lOpcion24,	lOpcion25,	lOpcion26,	
			lOpcion27,	lOpcion28,	lOpcion29,	lOpcion30,	lOpcion31,	lOpcion32,	lOpcion33)
			select '00','INFHOTEL',1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,'00',1,1,1,1,1,1,1,1,1,1,1,1
			--insert into TGRUPOUSUARIO select '01','SUPERVISOR',1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,'00',1,1,1,1,1,1,1,1,1,1,1
		END
	else
		begin
			update TGRUPOUSUARIO
			set tDetallado='INFHOTEL', lModulo01=1,lModulo02=1,lModulo03=1, lModulo04=1,  lOpcion01=1,
			lOpcion02=1,lOpcion03=1,lOpcion04=1,lOpcion05=1,lOpcion06=1,lOpcion07=1,lOpcion08=1,lOpcion09=1,lOpcion10=1,lOpcion11=1,
			lOpcion12=1,lOpcion13=1,lOpcion14=1,lOpcion15=1,lOpcion16=1,lOpcion17=1,lOpcion18=1,lOpcion19=1,lOpcion20=1,lOpcion21=1,
			lactivo=1, lModulo05=1, lOpcion22=1, lOpcion24 = 1, lOpcion25 = 1, lOpcion26 = 1, lOpcion27 = 1
			where TGRUPOUSUARIO='00'
		end
GO
	IF(SELECT COUNT(*) FROM TTIPOIDENTIDAD)=0  
		BEGIN
			INSERT INTO TTIPOIDENTIDAD (tCodigo, Descripcion, tResumido, lActivo) Values ('01','DNI','DNI',1)
			INSERT INTO TTIPOIDENTIDAD (tCodigo, Descripcion, tResumido, lActivo) Values ('02','RUC','RUC',1)
			INSERT INTO TTIPOIDENTIDAD (tCodigo, Descripcion, tResumido, lActivo) Values ('03','CARNET DE EXTRANJERIA','C.E.',1)
			INSERT INTO TTIPOIDENTIDAD (tCodigo, Descripcion, tResumido, lActivo) Values ('04','PASAPORTE','PAS',1)
			INSERT INTO TTIPOIDENTIDAD (tCodigo, Descripcion, tResumido, lActivo) Values ('05','CEDULA DIPLOMATICA DE IDENTIDAD','CDI',1)
			INSERT INTO TTIPOIDENTIDAD (tCodigo, Descripcion, tResumido, lActivo) Values ('06','VARIOS','VARIOS',1)
		END 
GO
	delete from  tusuario where substring(tcodigousuario,1,1) like '*'
GO
	-- if (select count(*)from TUSUARIO )=0 or (select count(*) from TUSUARIO where tCodigoUsuario='00001')=0 --select * from TUSUARIO
		-- begin
			-- insert into TUSUARIO(tCodigoUsuario,tGrupoUsuario,tDetallado,tResumido,tPassword,lActivo,fRegistro,tBandaMagnetica,tHuella,tUsuarioModifica) 
			-- select '00001','00','ADMIN','ADMIN','!ª)(=',1,GETDATE(),'','',''
		-- end
-- GO
	IF (SELECT COUNT (*) FROM TPARAMETRO) = 0
	BEGIN
		INSERT INTO TPARAMETRO (tIdentificacionTributaria) VALUES  ('11111111111')
	END
GO
	if not exists (select * from TUSUARIO where tDetallado='ADMIN' or tResumido='ADMIN') and (select count(*)from TUSUARIO )>0
		begin
			insert into TUSUARIO (tCodigoUsuario,tGrupoUsuario,tDetallado,tResumido,tPassword,lActivo,fRegistro,tBandaMagnetica,tHuella,tUsuarioModifica)
			select (select RIGHT('00000'+ convert(nvarchar(5),(MAX(tcodigousuario) + 1)),5) from TUSUARIO ),'00','ADMIN','ADMIN','!ª)(=',1,GETDATE(),'','',''
		end
GO
	DELETE FROM TTABLA WHERE TTABLA='MOZO' AND substring(TCODIGO,1,1)='*'
GO
	if (select count(TCODIGO) from TTABLA where  TTABLA='MOZO')=0 
		begin
			insert into TTABLA (TTABLA,TCODIGO,nTamano,tDetallado,tResumido,tIcono,nBoton,tValor,nValor, lActivo,lReplica,tValor2,lAutorizaEnviosMozo)
			select 'MOZO','0001',0,'ADMIN','ADMIN',		'',0,'!ª)(=',1,1,1,'',1
		END
GO
	if (select count(tF1) from TMENSAJE) = 0
	   INSERT INTO TMENSAJE (tF1) VALUES ('')
GO
	--Actualiza Datos
	UPDATE TPRODUCTO SET TUNIDADNEGOCIO='00' where isnull(TUNIDADNEGOCIO,'')=''
	UPDATE DPEDIDO SET TUNIDADNEGOCIO='00' where isnull(TUNIDADNEGOCIO,'')=''
	update tTipoDocumentoImpresora set lImpuesto1 = 0 where isnull(lImpuesto1,0)=0
	update tTipoDocumentoImpresora set lImpuesto2 = 0 where isnull(lImpuesto2,0)=0
	update tTipoDocumentoImpresora set lImpuesto3 = 0 where isnull(lImpuesto3,0)=0
	update tProducto set lImpuesto4 = 0 where isnull(lImpuesto4,0)=0
	update tProducto set lImpuesto5 = 0 where isnull(lImpuesto5,0)=0
	update tProducto set lImpuesto6 = 0 where isnull(lImpuesto6,0)=0
	update tProducto set lImpuesto7 = 0 where isnull(lImpuesto7,0)=0
	update tProducto set lImpuesto8 = 0 where isnull(lImpuesto8,0)=0
	update tProducto set lImpuesto9 = 0 where isnull(lImpuesto9,0)=0
	update tProducto set lImpuesto10 = 0 where isnull(lImpuesto10,0)=0
	update tProducto set lImpuesto11 = 0 where isnull(lImpuesto11,0)=0
	update tProducto set lImpuesto12 = 0 where isnull(lImpuesto12,0)=0
	update tProducto set lImpuesto13 = 0 where isnull(lImpuesto13,0)=0
	update tProducto set lImpuesto14 = 0 where isnull(lImpuesto14,0)=0
	update tProducto set lImpuesto15 = 0 where isnull(lImpuesto15,0)=0
	update tproducto set lDescuento=1 where isnull(lDescuento,1)=1
	update tProducto set lLocal = 1 where isnull(lLocal,1)=1
	update tProducto set lDelivery = 1 where isnull(lDelivery,1)=1
	update tProducto set lLlevar = 1 where isnull(lLlevar,1)=1
	update tProducto set lCanal4 = 1 where isnull(lCanal4,1)=1
	update tProducto set lCanal5 = 1 where isnull(lCanal5,1)=1
	update tProducto set nPrecioDelivery = 0 where isnull(nPrecioDelivery,0)=0
	update tProducto set nPrecioLlevar = 0 where isnull(nPrecioLlevar,0)=0
	update tProducto set nPrecioCanal4 = 0 where isnull(nPrecioCanal4,0)=0
	update tProducto set nPrecioCanal5 = 0 where isnull(nPrecioCanal5,0)=0
	update tproducto set ninsumo=0 where isnull(nInsumo,0)=0
	update tproducto set ninsumo2=0 where isnull(nInsumo2,0)=0
	update tproducto set ninsumo3=0 where isnull(nInsumo3,0)=0
	update tproducto set ninsumo4=0 where isnull(nInsumo4,0)=0
	update tproducto set ninsumo5=0 where isnull(nInsumo5,0)=0
	update tproducto set nManoObra=0 where isnull(nManoObra,0)=0
	update tproducto set nManoObra2=0 where isnull(nManoObra2,0)=0
	update tproducto set nManoObra3=0 where isnull(nManoObra3,0)=0
	update tproducto set nManoObra4=0 where isnull(nManoObra4,0)=0
	update tproducto set nManoObra5=0 where isnull(nManoObra5,0)=0
	update tproducto set nGasto=0 where isnull(nGasto,0)=0
	update tproducto set nGasto2=0 where isnull(nGasto2,0)=0
	update tproducto set nGasto3=0 where isnull(nGasto3,0)=0
	update tproducto set nGasto4=0 where isnull(nGasto4,0)=0
	update tproducto set nGasto5=0 where isnull(nGasto5,0)=0
	update TTABLA set tIcono = 0 where TTABLA = 'TIPODOCUMENTO' and isnull(tIcono,0)=0
	update TTABLA set nTamano = 0 where TTABLA = 'MOZO' and isnull(nTamano,0)=0
	UPDATE TDELIVERY SET lPuntos=1 where isnull(lPuntos,1)=1
	UPDATE TDELIVERY SET nAcumulado=0 where isnull(nAcumulado,0)=0
	UPDATE TDELIVERY SET nUtilizado=0 where isnull(nUtilizado,0)=0
	UPDATE TDELIVERY SET nDisponible=0 where isnull(nDisponible,0)=0
	UPDATE TCAJA SET tTipoPedido = '01' where isnull(tTipoPedido,'')=''
	update TPARAMETRO set nDecimal = 2 where isnull(nDecimal,0)=0
	UPDATE TTABLA SET nTamano = '1' where isnull(nTamano,1)=1 and TTABLA='TIPODOCUMENTO'
	update TSUBGRUPO set nOrden=1 where isnull(nOrden,0)=0
	delete from tproductoarea where tcodigoproducto in (select tcodigoproducto from tproducto where lcombinacion=1)
	update mpedido set fLlegada =fRegistro where isnull(fLlegada,0)=0 and tTipoPedido='02'
	update TOFERTA set lCanal4=1 where isnull(lCanal4,0)=0
	update TOFERTA set lCanal5=1 where isnull(lCanal5,0)=0
	update treserva set tprioridad=1
	UPDATE TUSUARIO set tBandaMagnetica='' where isnull(tBandaMagnetica,'')=''
	UPDATE TPROPIEDAD set tEnlace='' where isnull(tEnlace,'')='' 
	UPDATE TPRODUCTOPROPIEDAD set tEnlace='' where isnull(tEnlace,'')='' 
	UPDATE TCOMBOPROPIEDAD set tEnlace='' where isnull(tEnlace,'')='' 
	UPDATE TMOTIVODESCUENTO set lTopePedido=0 where isnull(lTopePedido,0)=0
	UPDATE TMOTIVODESCUENTO set lBloqueo=0 where isnull(lBloqueo,0)=0
	UPDATE TMOTIVODESCUENTO set lRatio=0 where isnull(lRatio,0)=0
	UPDATE TMOTIVODESCUENTO set lAplicablePedido=0 where isnull(lAplicablePedido,0)=0
	UPDATE TCAJA SET lPagoDirectoCD = 0 where isnull(lPagoDirectoCD,0)=0
	UPDATE TCAJA SET lActivarAvisoRetenciones=0 where isnull(lActivarAvisoRetenciones,0)=0
GO
	update TTABLA set tResumido = substring(tResumido,1,15) where tTABLA='MOZO' and len(ltrim(tResumido))>15
	update TUSUARIO set tResumido = substring(tResumido,1,15) 
GO
	delete from taccesoenvia where substring(tcodigoacceso,4,5)='00000'
GO
	update TLOCAL set tCodigoLocal='0'+LTRIM(RTRIM(tCodigoLocal)) where LEN(LTRIM(tCodigoLocal))=2 
	update TTABLA set tValor='0'+LTRIM(RTRIM(tValor)) where LEN(LTRIM(tValor))=2 and TTABLA='SALON'
GO

--insert into TOPERADOR 
--select tcodigo, tdetallado, tresumido, tvalor, 0, nvalor, 0, nboton, getdate(), 'MASTER', lactivo, 1 from TTABLA where tTabla='OPERADOR'
--delete from TTABLA where tTabla='OPERADOR'

--INSERT INTO TMOTIVODESCUENTO 
--SELECT tCodigo, tDetallado, tResumido, nValor, 0, 1, 1, 1 
--FROM TTABLA where TTABLA='MOTIVODESCUENTO'
--delete ttabla where TTABLA='MOTIVODESCUENTO'

--DELETE FROM TTABLA WHERE TTABLA='GRUPOUSUARIO'
--DELETE FROM TGRUPOUSUARIO
--INSERT INTO TGRUPOUSUARIO (tGrupoUsuario, tDetallado, lModulo01, lModulo02, lModulo03, lOpcion01,lOpcion02,lOpcion03,lOpcion04,lOpcion05,lOpcion06,lOpcion07,lOpcion08,lOpcion09,lOpcion10, lOpcion11, lOpcion12, lOpcion13, lOpcion14, lOpcion15, lActivo) 
--values ('01','SUPERVISOR',1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1)
--INSERT INTO TGRUPOUSUARIO (tGrupoUsuario, tDetallado, lModulo01, lModulo02, lModulo03, lOpcion01,lOpcion02,lOpcion03,lOpcion04,lOpcion05,lOpcion06,lOpcion07,lOpcion08,lOpcion09,lOpcion10, lOpcion11, lOpcion12, lOpcion13, lOpcion14, lOpcion15, lActivo) 
--values ('02','CAJA',1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1)

--Update CPEDIDO set nInsumo = ISNULL(TPRODUCTO.nInsumo,0), nGasto=ISNULL(TPRODUCTO.nGasto,0), nManoObra=ISNULL(TPRODUCTO.nManoObra,0)
--From CPEDIDO, TPRODUCTO
--WHERE CPEDIDO.tProductoCombo=TPRODUCTO.tCodigoProducto and isnull(CPEDIDO.nInsumo,0)=0 


--insert into ttabla (TTABLA,TCODIGO,TDETALLADO,TRESUMIDO,LACTIVO) values('SECTOR','00','SECTOR0','SECTOR0',1)
	update TPARAMETRO  set lImprimeDiaContable =0 where lImprimeDiaContable is null
GO
	if (select count(tCodigolocal) from TLOCAL) = 0 
		BEGIN
			insert into tlocal (tcodigolocal,tdetallado,tresumido,lactivo)
			values('001','PRINCIPAL','PRINCIPAL',1)
		end
GO
	IF(select count(tCodigo) from TTABLA WHERE tCodigo ='00' and TTABLA ='SUCURSAL' )= 0
		BEGIN 
			insert into TTABLA( tTabla, tCodigo, tDetallado, tResumido, nValor, tValor, tIcono, lActivo) values ('SUCURSAL',  '00',  'SIN SUCURSAL',  'SIN SUCURSAL', 0,'','',1) 
		END
GO
	if (select case when ldiacontableautomatico=1 then 2 else 0 end + case when ldiacontablemanual=1 then 2 else 0 end from tparametro) =0
		begin
			update tparametro set ldiacontableautomatico=1 , thoracierrediacontable='06:00'
		end 
GO
	if (SELECT upper(systypes.name)  FROM syscolumns INNER JOIN  sysobjects ON syscolumns.id = sysobjects.id INNER JOIN systypes ON systypes.xtype = syscolumns.xtype WHERE (sysobjects.name = 'DPEDIDO') AND (systypes.name <> 'sysname')  AND   syscolumns.name='fenvio')='SMALLDATETIME'
		BEGIN
			exec RSP_ALTERACOLUMNA 'DPEDIDO','FENVIO','datetime','8','N'
		END
GO
	if (select count(*) from ttabla where TTABLA='ESTADOFRECUENTE') = 0 
		BEGIN
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOFRECUENTE','01','ACTIVO','ACTIVO',0,1)
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOFRECUENTE','02','SUSP CREDITO','SUSP CREDITO',0,1)
			INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOFRECUENTE','03','SUSP DERECHO','SUSP DERECHO',0,1)
		end
GO
	if (select count(tCodigo) from TTABLA where tTabla='TIPODOCUMENTO') > 0 
		BEGIN
			insert into TTIPODOCUMENTO   
			select tCodigo, tDetallado, tResumido, tValor, nValor, tIcono, 0, nTamano, lActivo, 0,'',0,'',0,0,0,0,0,0,'',0,0,0,'',0 from TTABLA where tTABLA='TIPODOCUMENTO'
			DELETE FROM TTABLA where tTABLA='TIPODOCUMENTO'			
		END
GO
	IF (select COUNT(*) from TTIPODOCUMENTO )=0
			begin
				insert into TTIPODOCUMENTO (tCodigoTipoDocumento,tDescripcion,tPrefijo,tCodigoSunat,lPideCliente,nMonto,lTransporte
				,lRegistroVenta,lActivo,lCanjearNotaCredito,lValidaRuc,lImprimeQR,tCodigoFacturacion,lFacturacionE
				,lNotaCredito,lValidaEmail,lImpresionRetencion,lIdentidadRuc,lIdentidadOtros,tImpresionFE,nMontoMaximo
				,lImpObs,nMontoMinimo,tCodigoPlantilla,lValidaUbigeo)
				values 
				('00','CUENTA INTERNA',0,NULL,0,0,0,1,1,0,0,0,'',0,0,0,0,0,0,0,0,NULL,NULL,NULL,NULL)
			end
GO
	if (select count(tCodigoCanalVenta) from TCANALVENTA) = 0 and (select count(*)from ttabla where ttabla='TIPOPEDIDO')>0
		begin
		   INSERT INTO TCANALVENTA(tCodigoCanalVenta,tDetallado,lActivaMozo,lActivaMotorizado,lObligaMesa,lObligaPax,lObligaMozo,lObligaMotorizado,lCanalCentralPedidos,lCanalDelivery,
									lObligaIngresoFechaEntrega,lObligaClienteFrecuente,lActivo,lObligaEntregarA,tEnlaceContable1,tEnlaceContable2,tCodigoSunat)
		   select tCodigo, tDetallado, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, lActivo,0,'','','' from ttabla where ttabla='TIPOPEDIDO'
		end
	else
		Begin
			if (select count(tCodigoCanalVenta) from TCANALVENTA) = 0
				begin
					INSERT INTO TCANALVENTA (tCodigoCanalVenta,tDetallado,lActivaMozo,lActivaMotorizado,lObligaMesa,lObligaPax,lObligaMozo,lObligaMotorizado,lCanalCentralPedidos,
											lCanalDelivery,lObligaIngresoFechaEntrega,lObligaClienteFrecuente, lActivo,lObligaEntregarA,tEnlaceContable1,tEnlaceContable2,tCodigoSunat)
					select '01', 'En el local', 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 1,0,'','',''
				end
		end
GO
	if (select count(tCodigoCanalVenta) from TCANALVENTA) = 1 and (select count(*)from ttabla where ttabla='TIPOPEDIDO')>0
		begin
		   INSERT INTO TCANALVENTA(tCodigoCanalVenta,tDetallado,lActivaMozo,lActivaMotorizado,lObligaMesa,lObligaPax,lObligaMozo,lObligaMotorizado,lCanalCentralPedidos,lCanalDelivery,lObligaIngresoFechaEntrega,lObligaClienteFrecuente,lActivo,lObligaEntregarA,tEnlaceContable1,tEnlaceContable2,tCodigoSunat)
		   select tCodigo, tDetallado, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, lActivo,0,'','','' from ttabla where ttabla='TIPOPEDIDO'
		end
	else
		Begin
			if  (select count(tCodigoCanalVenta) from TCANALVENTA) = 1
			begin
				INSERT INTO TCANALVENTA (tCodigoCanalVenta,tDetallado,lActivaMozo,lActivaMotorizado,lObligaMesa,lObligaPax,lObligaMozo,lObligaMotorizado,lCanalCentralPedidos,lCanalDelivery,lObligaIngresoFechaEntrega,lObligaClienteFrecuente,lActivo,lObligaEntregarA,tEnlaceContable1,tEnlaceContable2,tCodigoSunat)
				select '02', 'Delivery', 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 1,0,'','',''
			end
		end
GO
	if (select count(tCodigoCanalVenta) from TCANALVENTA) = 2
		begin
		   INSERT INTO TCANALVENTA (tCodigoCanalVenta,tDetallado,lActivaMozo,lActivaMotorizado,lObligaMesa,lObligaPax,lObligaMozo,lObligaMotorizado,lCanalCentralPedidos,lCanalDelivery,lObligaIngresoFechaEntrega,lObligaClienteFrecuente,lActivo,lObligaEntregarA,tEnlaceContable1,tEnlaceContable2,tCodigoSunat)
		   select '03', 'Para llevar', 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 1,0,'','',''
		end
GO
	if (select count(tCodigoCanalVenta) from TCANALVENTA) = 3 
		begin
		   INSERT INTO TCANALVENTA (tCodigoCanalVenta,tDetallado,lActivaMozo,lActivaMotorizado,lObligaMesa,lObligaPax,lObligaMozo,lObligaMotorizado,lCanalCentralPedidos,lCanalDelivery,lObligaIngresoFechaEntrega,lObligaClienteFrecuente,lActivo,lObligaEntregarA,tEnlaceContable1,tEnlaceContable2,tCodigoSunat)
		   select '04', '', 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0,0,'','',''
		end
GO
	if (select count(tCodigoCanalVenta) from TCANALVENTA) = 4
		begin
		   INSERT INTO TCANALVENTA (tCodigoCanalVenta,tDetallado,lActivaMozo,lActivaMotorizado,lObligaMesa,lObligaPax,lObligaMozo,lObligaMotorizado,lCanalCentralPedidos,lCanalDelivery,lObligaIngresoFechaEntrega,lObligaClienteFrecuente,lActivo,lObligaEntregarA,tEnlaceContable1,tEnlaceContable2,tCodigoSunat)
		   select '05', '', 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0,0,'','',''
		end
GO
	if (select count(TTABLA) from TTABLA where tTabla='TTIPOPEDIDO') > 0
		Begin
			Update TCANALVENTA set lObligaMesa=TCAJA.lObliga, lObligaPax=TCAJA.lPax, lObligaMozo=TCAJA.lMozo, lObligaMotorizado=TCAJA.lMotorizado
			from TCAJA, TCANALVENTA
			where TCAJA.tTipoPedido=TCANALVENTA.tCodigoCanalVenta 
			delete FROM TTABLA where tTabla='TTIPOPEDIDO'
		End
GO
	if (select avg(len(ltrim(tCodigoDelivery))) FROM TDELIVERY) < 6
		Begin
		   update TDELIVERY set tCodigoDelivery = right('00'+ ltrim(tCodigoDelivery),7)
		End
GO
	IF EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'tPaseCliente')
		BEGIN
			drop table tPaseCliente
		END
	ELSE
		BEGIN
			CREATE TABLE tPaseCliente (tcodigoDelivery nVarChar(10) collate Modern_Spanish_CI_AS, tCodigoCompania nVarChar(10) collate Modern_Spanish_CI_AS)		
		END
GO
	if (select count(tCodigoCliente) FROM TCOMPANIA) > 0
	Begin
		Declare @Correlativo as int
		set @Correlativo = (select ISNULL(max(tCodigoDelivery),0) FROM TDELIVERY)
		insert into  TDELIVERY
		(
		tCodigoDelivery,	tTipoCliente,	tApellido,	tNombre,	tDireccion,	tTelefono,	tReferencia,
		tZona,	tDistrito,	tCodigoCliente,	tCodigoTarjeta,	tNumeroTarjeta,	tFechaTarjeta,	nDescuento,
		fNacimiento,	tEMail,	tObservacion,	lActivo,	lPuntos,	nAcumulado,	nUtilizado,
		nDisponible,	tUsuario,	fRegistro,	fModificacion,	lReplica,	tEstadoFrecuente,
		lExcluyeProductos,	lClienteCtaCte,	nConsumo,	nLinea,	tTipoCtaCte,	tSubTipoCtaCte,
		tTipoIdentidad,	tIdentidad,	iFoto,	tAccionSocio,	nLineaPorCobrar,	nConsumoPorCobrar,
		idLocalMesa,	idClienteMesa,	lEmisionMonedaExtranjera,	nDiasCredito,	lRangoHorarios,
		tHoraInicio,	tHoraFin,	tCanal1,	tCanal2,	tCanal3,	tCanal4,	tCanal5,
		lSoloxDia
		)
	
		select substring('0000000',1,7-len(ltrim(str(@correlativo+tCodigoCliente))))+ltrim(str(@correlativo+tCodigoCliente)), 
		'00', tapecom, tnomsoc, tDireccion, tTelefono1, '','','','','','','',0,
		null,tEmail,'',lActivo,0,0,0,0, tUsuario, fRegistro,getdate(), lReplica,
		'00', 0, 1, nConsumo, nLinea, tTipoCtaCte, tSubTipoCtaCte, '',ISNULL(tIdentidad,''),'','',0,0,'','',0,0,
		 0,	'','',0,0,0,0,0,0		
		From tCompania

		insert tPaseCliente
		select substring('0000000',1,7-len(ltrim(str(@correlativo+tCodigoCliente))))+ltrim(str(@correlativo+tCodigoCliente)), tCodigoCliente
		from TCOMPANIA

		update MPEDIDO set tClienteDelivery= right('00'+tClienteDelivery,7) where len(tClienteDelivery)>0
		update MPEDIDO set tClienteCtaCte= substring('0000000',1,7-len(ltrim(str(@correlativo+tClienteCtaCte))))+ltrim(str(@correlativo+tClienteCtaCte))
		where len(tClienteCtaCte)>0
		delete from tCompania
	End
GO
	IF EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = 'tPaseCliente')
	begin
		if (select count(*) FROM tPaseCliente) > 0
		 begin
				 update MDOCUMENTO
				 set tClientePago =tpasecliente.tcodigodelivery
				 from  MDOCUMENTO inner join tpasecliente on mdocumento.tclientepago=tpasecliente.tcodigocompania
		end
	end
GO
	delete from TTRAMITE
	insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
					values ('00001','Solicitud de Autorizacion', 0,1)
GO
	insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
					values ('00002','Solicitud de Autorizacion por Cambio de Software',1,1)
GO
	insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
					values ('00003','Solicitud de Renovacion de Autorizacion', 1,1)				
GO
	insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
					values ('00004','Baja de la Autorizacion', 1,1)				
GO
	insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
					values ('00005','Inclusion de Puntos de Emision y Tipos de Documentos', 0,1)				
GO
	insert into TTRAMITE (tCodigoTramite,tDescripcion,lSolicitaNAnteriorAutorizacion,lActivo) 
					values ('00006','Exclusion de Puntos de Emision y Tipos de Documentos', 1,1)						
GO
	delete FROM TTABLA where TTABLA='ESTADOSOLICITUD'
GO
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','01','ACTIVO','ACTIVO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','02','APROBADO','APROBADO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','03','EN TRAMITE','EN TRAMITE',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','04','RECHAZADO','RECHAZADO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOSOLICITUD','05','ANULADO','ANULADO',0,1)	
GO
	delete FROM TTABLA where TTABLA='ESTADODETSOLICITUD'
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADODETSOLICITUD','01','APROBADO','APROBADO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADODETSOLICITUD','02','EN TRAMITE','EN TRAMITE',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADODETSOLICITUD','03','RECHAZADO','RECHAZADO',0,1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADODETSOLICITUD','04','DE BAJA','ANULADO',0,1)
GO
	update TPARAMETRO set lMobilePasswordCCaja=0 where lMobilePasswordCCaja is null
GO
	update TPARAMETRO set lMobileUnidadNegocio=0 where lMobileUnidadNegocio is null
GO
	update tcaja set lBloqueaPrecuenta=0 where lBloqueaPrecuenta is null
GO
	update TCAJA set lMultiareasubgrupo=0 where lMultiareasubgrupo is null
GO
	update TCAJA set lMultiareacaja=0 where lMultiareacaja is null
GO
	if (select count(*) from TPARAMETRO where tVersion<'4.94.4111')>0
	begin
			update TCAJA set lmultiareacaja=1 where isnull(tsubalmacen,'')<>''
	end
GO
	update tcaja set tsectorventa='' where tsectorventa is null
GO
	-----CESAR DOCUMENTOS VARIABLES
	DELETE FROM TTABLA where tTabla='FORMULARIO'
	INSERT TTABLA (TTABLA, TCODIGO, nTamano, tDetallado, tResumido, lActivo) values ('FORMULARIO','01',2,'TICKET','TICKET',1)
	INSERT TTABLA (TTABLA, TCODIGO, nTamano, tDetallado, tResumido, lActivo) values ('FORMULARIO','02',2,'VARIABLE','VARIABLE',1)
	INSERT TTABLA (TTABLA, TCODIGO, nTamano, tDetallado, tResumido, lActivo) values ('FORMULARIO','03',2,'TICKETVARIABLE','TICKETVARIABLE',1)
	-------
GO
	----GRUPO USUARIO
	UPDATE TGRUPOUSUARIO SET tNivel = tGrupoUsuario
GO
	---- MOTIVO NOTA CREDITO ------
	DELETE FROM TTABLA WHERE tTabla='MOTIVONOTACREDITO'
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, lActivo) values ('MOTIVONOTACREDITO','01','Anulación de la operación','Anulacion operacion',1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, lActivo) values ('MOTIVONOTACREDITO','02','Anulación por error en el RUC','error RUC',1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, lActivo) values ('MOTIVONOTACREDITO','03','Corrección por error en la descripción','error en la descripción',1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, lActivo) values ('MOTIVONOTACREDITO','04','Descuento global','Descuento global',1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, lActivo) values ('MOTIVONOTACREDITO','05','Descuento por ítem','Descuento por ítem',1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, lActivo) values ('MOTIVONOTACREDITO','06','Devolución total','Devolución total',1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, lActivo) values ('MOTIVONOTACREDITO','07','Devolución por ítem','Devolución por ítem',1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, lActivo) values ('MOTIVONOTACREDITO','08','Bonificación','Bonificación',1)
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, lActivo) values ('MOTIVONOTACREDITO','09','Disminución en el valor','Disminución en el valor',1)
GO
	-----------------------------
	UPDATE TPARAMETRO
	SET lHuellaDigital=1 where lHuellaDigital is null and lHuellaSecugen is null
GO
	if (select count(*) from TPARAMETRO where tVersion<'4.94.5073')>0
	begin
			update ttipodocumento set lidentidadruc=1,lidentidadotros=0 where isnull(lvalidaruc,0)=1
	 end
GO
	if (select count(*) from TPARAMETRO where tVersion<'4.94.5073')>0
	begin
		declare @codigop nvarchar(100)
		select @codigop=tcodigo from vPaisOrigen where lactivo=1
		if (@codigop='000' )
		begin
				update tcliente set tTipoIdentidad='02' where LEN(ltrim(TCLIENTE.tIdentidad))>10
				and isnull(ttipoidentidad,'')=''
		end 
	 end
GO
	update ttipodocumento set nmontomaximo=isnull(nmontomaximo,0)
GO
	UPDATE TPARAMETRO SET lFacturacionE = 0 Where lFacturacionE IS NULL
GO
	UPDATE TCANALVENTA SET LOBLIGAENTREGARA=0 WHERE lobligaentregara is null
GO
	update tproductopropiedad set ncantidad= 1 where ncantidad is null
GO
	update tcombopropiedad set ncantidad= 1 where ncantidad is null
GO
	update TPROPIEDAD set lSolicitaCantidad=0 where lSolicitaCantidad is null
GO
	update tparametro set lLogCajaRapida=0 where llogcajarapida is null
GO
	UPDATE TLOCAL SET tenlacecontable='' where tenlacecontable is null
GO
	update tcaja set lBuscaPedidoNumero=0 where lBuscaPedidoNumero  is null
GO
	update tcaja set lCodigoReciboIngreso=0 where lCodigoReciboIngreso  is null
GO
	UPDATE TCANALVENTA SET tenlacecontable1='' where tenlacecontable1 is null
GO
	UPDATE TCANALVENTA SET tenlacecontable2='' where tenlacecontable2 is null
GO
	UPDATE TPARAMETRO SET lEmailConfirmacion=0 where lemailconfirmacion is null
GO
	UPDATE TPARAMETRO SET lEmailRecordatorio=0 where lemailRecordatorio is null
GO
	UPDATE TPARAMETRO SET lEmailAgradecimiento=0 where lemailAgradecimiento is null
GO
	SET QUOTED_IDENTIFIER OFF 
GO
	update tparametro set nTiempoToleranciaAnf=0 where nTiempoToleranciaAnf is null
GO
	IF  (SELECT TCODIGO FROM vPaisOrigen)='000' 
	BEGIN
		UPDATE  TCLIENTE SET tIdentidad=LTRIM(RTRIM(tIdentidad)) WHERE lActivo=1-- BORRA LOS ESPACION EN BLANCO DELANTE Y DETRAS DEL NUMERO DE IDENTIFICACION PARA PODER BUSCAR LA LONGITUD 
		UPDATE  TCLIENTE SET  lActivo=0 where len(tIdentidad)<>11 AND tTipoIdentidad='02'AND lActivo=1-- PARA TIPO IDENTIFICACION RUC
		UPDATE  TCLIENTE SET  lActivo=0 where len(tIdentidad)<>8 AND tTipoIdentidad='01' AND lActivo=1 -- PARA TIPO IDENTIFICACION DNI
		UPDATE TCLIENTE SET lActivo=0 WHERE ISNULL(tTipoIdentidad,'')=''
		UPDATE TCLIENTE SET tEmpresa=LTRIM(RTRIM(tEmpresa)) WHERE lActivo=1
	END 
GO
--- se configura los correos dentro del sistema 
	delete from TTABLA where TTABLA='EMAIL_PRORROGA' and TCODIGO='01'
GO
	INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('EMAIL_PRORROGA','01','edelacruz@infomatica.pe','edelacruz@infomatica.pe',1,1)
GO
--- configuracion de email
	delete from TTABLA where TTABLA='EMAIL_PRORROGA_CON' 
	INSERT TTABLA (TTABLA, TCODIGO, tResumido, tDetallado, nBoton, lActivo) values ('EMAIL_PRORROGA_CON','C','SMTP','mail.replicadores.com.pe',1,1)
	INSERT TTABLA (TTABLA, TCODIGO, tResumido, tDetallado, nBoton, lActivo) values ('EMAIL_PRORROGA_CON','M1','Asunto','PRORROGA DE PAGO',1,1)
	INSERT TTABLA (TTABLA, TCODIGO, tResumido, tDetallado, nBoton, lActivo) values ('EMAIL_PRORROGA_CON','M2','Cuerpo','Usted debe realizar el pago',1,1)
	INSERT TTABLA (TTABLA, TCODIGO, tResumido, tDetallado, nBoton, lActivo) values ('EMAIL_PRORROGA_CON','P','PORT','25',1,1)
	INSERT TTABLA (TTABLA, TCODIGO, tResumido, tDetallado, nBoton, lActivo) values ('EMAIL_PRORROGA_CON','U','Usuario','dluna@replicadores.com.pe',1,1)
	INSERT TTABLA (TTABLA, TCODIGO, tResumido, tDetallado, nBoton, lActivo) values ('EMAIL_PRORROGA_CON','W','Password','Sistemas1',1,1)
GO
	--Configuracion motivos sunat --genera txt
	delete from ttabla where ttabla ='MOTIVOSUNAT'
	insert into ttabla (ttabla,tcodigo,tDetallado) values ('MOTIVOSUNAT','1','Conexión internet')
	insert into ttabla (ttabla,tcodigo,tDetallado) values ('MOTIVOSUNAT','2','Fallas fluido eléctrico')
	insert into ttabla (ttabla,tcodigo,tDetallado) values ('MOTIVOSUNAT','3','Desastres Naturales')
	insert into ttabla (ttabla,tcodigo,tDetallado) values ('MOTIVOSUNAT','4','Robo') 
	insert into ttabla (ttabla,tcodigo,tDetallado) values ('MOTIVOSUNAT','5','Fallas en el sistema de emisión electrónica')
	insert into ttabla (ttabla,tcodigo,tDetallado) values ('MOTIVOSUNAT','6','ventas por emisores itinerantes')
	insert into ttabla (ttabla,tcodigo,tDetallado) values ('MOTIVOSUNAT','7','otros')
GO
	IF (SELECT COUNT(*) FROM TCAJA)=0
		BEGIN 
			INSERT INTO TCAJA (tCaja,TDESCRIPCION) VALUES('001','CAJA001')
		END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOPAGO' AND TCODIGO='01')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPAGO','01','Efectivo','Efectivo',0,1)
	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOVENTA' )=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOVENTA', '01', 'CONTADO', 'CONTADO', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOVENTA', '02', 'CREDITO', 'CREDITO', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOVENTA', '03', 'CORTESIA', 'CORTESIA', 0, 1)
		
	END
GO

	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='Vehiculo' )=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('Vehiculo', '001', 'CG2721', 'CG2721', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('Vehiculo', '002', 'ABC4545', 'ABC4545', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('Vehiculo', '003', 'WQ8795', 'WQ8795', 0, 1)
		
	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOPEDIDO')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPEDIDO', '01', 'En el Local', 'En el Local', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPEDIDO', '02', '', '', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPEDIDO', '03', 'Para Llevar', 'Para Llevar', 0, 1)
		
	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOCTACTE')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCTACTE', '01', 'Empleados', 'Empleados', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCTACTE', '02', 'Directores', 'Directores', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCTACTE', '03', 'Clientes', 'Clientes', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOCTACTE', '04', 'Empresa', 'Empresa', 0, 1)
	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOATENCION')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOATENCION', '01', 'EN EL MOMENTO', 'EN EL MOMENTO', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOATENCION', '02', 'PROGRAMADO', 'PROGRAMADO', 0, 1)
	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='SUBTIPOCTACTE')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo, tValor) values ('SUBTIPOCTACTE', '0101', 'Funcionarios', 'Funcionarios', 0, 1,'01')
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo, tValor) values ('SUBTIPOCTACTE', '0102', 'Mozos', 'Mozos', 0, 1,'01')
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo, tValor) values ('SUBTIPOCTACTE', '0103', 'Cajeros', 'Cajeros', 0, 1,'01')
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo, tValor) values ('SUBTIPOCTACTE', '0104', 'Playero', 'Playero', 0, 1,'01')
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo, tValor) values ('SUBTIPOCTACTE', '0105', 'Administracion', 'Administracion', 0, 1,'01')
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo, tValor) values ('SUBTIPOCTACTE', '0201', 'Gerencia General', 'Gerencia General', 0, 1,'02')
	END
	ELSE
	BEGIN
		UPDATE TTABLA SET tValor =SUBSTRING(TCODIGO,1,2) where isnull(tValor,'') ='' and  tTabla ='SUBTIPOCTACTE'
	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='ESTADORESERVA')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADORESERVA', '01', 'GENERADO', 'GENERADO', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADORESERVA', '02', 'ATENDIDO', 'ATENDIDO', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADORESERVA', '03', 'ANULADO', 'ANULADO', 0, 1)

	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='ESTADOMESA')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOMESA', '01', 'LISTA', 'LISTA', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOMESA', '02', 'OCUPADA', 'OCUPADA', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOMESA', '03', 'RESERVADA', 'RESERVADA', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOMESA', '04', 'SUCIA', 'SUCIA', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOMESA', '05', 'BLOQUEADA', 'BLOQUEADA', 0, 1)
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('ESTADOMESA', '06', 'FUERA SERVICIO', 'FUERA SERVICIO', 0, 1)
	END
GO


	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOPAGO' AND TCODIGO='02')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPAGO','02','Tarjeta de Credito','Tarjeta de Cred',0,1)
	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOPAGO' AND TCODIGO='03')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPAGO','03','Cheque','Cheque',0,1)
	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOPAGO' AND TCODIGO='04')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPAGO','04','Varios','Varios',0,1)
	END
GO
	IF (SELECT COUNT(*) FROM TTABLA WHERE TTABLA='TIPOPAGO' AND TCODIGO='05')=0
	BEGIN
		INSERT TTABLA (TTABLA, TCODIGO, tDetallado, tResumido, nBoton, lActivo) values ('TIPOPAGO','05','Puntos','Puntos',0,1)
	END
GO
	if (select top 1 TCODIGO from TTABLA where TTABLA='PAISORIGEN' and nValor=1 )='003'
		begin
			DELETE TTABLA WHERE TTABLA='TIPORESPONSABLE'

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
		End

GO
if (select count(*) from TPARAMETRO where tVersion<'4.95.26') > 0
  begin
		-- ACTUALIZACION GENERAL DE ACTIVACIONES PARA TODAS LAS TARJETAS BANCARIAS
	BEGIN TRAN;
	
	-- 1) Insertar registros faltantes (uno por cada tarjeta en TTARJETA)
	INSERT INTO dbo.TVISIBILIDADTARJETACREDITOXCANAL
	    (tcodigotarjeta, lvisibilidad, lcanal1, lcanal2, lcanal3, lcanal4, lcanal5)
	SELECT RTRIM(LTRIM(t.tCodigoTarjeta)), 0, 0, 0, 0, 0, 0
	FROM dbo.TTARJETACREDITO t
	WHERE NOT EXISTS (
	    SELECT 1
	    FROM dbo.TVISIBILIDADTARJETACREDITOXCANAL v
	    WHERE v.tcodigotarjeta = RTRIM(LTRIM(t.tCodigoTarjeta)));
	
	-- 2) Activar visibilidad y los 5 canales para todos los registros
	UPDATE v
	SET
	    v.lvisibilidad = 1,
	    v.lcanal1 = 1,
	    v.lcanal2 = 1,
	    v.lcanal3 = 1,
	    v.lcanal4 = 1,
	    v.lcanal5 = 1
	FROM dbo.TVISIBILIDADTARJETACREDITOXCANAL v
	JOIN dbo.TTARJETACREDITO t ON v.tcodigotarjeta = RTRIM(LTRIM(t.tCodigoTarjeta));
	
	COMMIT;
  end

GO
	IF(SELECT COUNT(*) FROM TPARAMETRO)=0 
		BEGIN
			INSERT INTO TPARAMETRO (tIdentificacionTributaria) values ('11111111111')	
		END
GO
	IF (select count(*) from TTABLA where TTABLA='MONEDA')=0
		BEGIN
			insert into TTABLA (ttabla, tcodigo, tResumido, tDetallado) values ('MONEDA','01','S/.','S/.')
			insert into TTABLA (ttabla, tcodigo, tResumido, tDetallado) values ('MONEDA','02','US$','US$')
			UPDATE TPARAMETRO SET tMonedaN='Soles' ,tMonN='S/.', tMonedaE='Dolares', tMonE='US$'
		END
GO
	IF (select count(*) from TTABLA where TTABLA='INTEGRACION')=0
		BEGIN
			insert into TTABLA (ttabla, tcodigo,tValor) values ('INTEGRACION','01','NIUBIZ')
			insert into TTABLA (ttabla, tcodigo,tValor) values ('INTEGRACION','02','IZIPAY')
		END
GO
--	TRUNCATE TABLE TINTEGRACIONES 
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=0)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (0,'Ninguno','Ninguno','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=1)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (1,'Visa Net QR','Visa Net QR','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=2)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (2,'Rappi','Rappi','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=3)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (3,'Ubert Eats','Ubert Eats','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=4)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (4,'CashDro','CashDro','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=5)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (5,'Mercado Pago','Mercado Pago','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=6)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (6,'Patio de Comida','Patio de Comida','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=7)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (7,'Pago Efectivo','Pago Efectivo','',1)
		END		
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=8)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (8,'Justo','Justo','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=9)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (9,'PedidosYa','PedidosYa','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=10)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (10,'Totem','Totem','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=11)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (11,'FPAY','FPAY','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=12)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (12,'EIS','EIS','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=13)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (13,'JUSTO POS','JUSTO POS','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=14)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (14,'JUSTO PASARELA','JUSTO PASARELA','',1)
		END

	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=15)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (15,'PeYaEfectivo','PeYaEfectivo','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=16)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (16,'QR DINAMICO','QR DINAMICO','',1)
		END

	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=17)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (17,'PINPAD','PINPAD','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=18)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (18,'SERVILEX','SERVILEX','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=19)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (19,'Niubiz Inalambrico','Niubiz Inalambrico','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=20)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo,	Descripcion,	tResumido, tValor, lActivo)
			VALUES (20,'Izipay PinPad','Izipay PinPad','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=21)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo, Descripcion, tResumido, tValor, lActivo)
			VALUES (21,'WHATSAPP EFECTIVO','WHATSAPP EFECTIVO','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=22)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo, Descripcion, tResumido, tValor, lActivo)
			VALUES (22,'WHATSAPP POS','WHATSAPP POS','',1)
		END
	IF (SELECT COUNT(*) FROM TINTEGRACIONES WHERE TCODIGO=23)=0
		BEGIN
			INSERT INTO TINTEGRACIONES (tCodigo, Descripcion, tResumido, tValor, lActivo)
			VALUES (23,'WHATSAPP PASARELA','WHATSAPP PASARELA','',1)
		END
        
	IF (SELECT COUNT(*) FROM DICTIONARY_INFOREST WHERE clave='Mesero')=0
		BEGIN
			INSERT INTO DICTIONARY_INFOREST 
			VALUES('Mesero','Mesero','Mesero','Mesero','Mesero','Mesero','Camarero')
		END
	IF (SELECT COUNT(*) FROM DICTIONARY_INFOREST WHERE clave='Guía')=0
		BEGIN
			INSERT INTO DICTIONARY_INFOREST
			VALUES('Guía','Guía','Guía','Guía','Guía','Guía','Albarán')
		END
GO
	update TPRODUCTO set lImpuesto1=1, lImpuesto4=1, lImpuesto7=1, lImpuesto10=1, lImpuesto13=1
	where isnull(lImpuesto1,0)=1 and isnull(tAplicaExoneracion,0)=0 and isnull(tAplicaInafectacion,0)=0
GO
	update TPRODUCTO set lImpuesto1=1, lImpuesto4=1, lImpuesto7=1, lImpuesto10=1, lImpuesto13=1
	where isnull(lImpuesto4,0)=1 and isnull(tAplicaExoneracion,0)=0 and isnull(tAplicaInafectacion,0)=0
GO
	update TPRODUCTO set lImpuesto1=1, lImpuesto4=1, lImpuesto7=1, lImpuesto10=1, lImpuesto13=1
	where isnull(lImpuesto7,0)=1 and isnull(tAplicaExoneracion,0)=0 and isnull(tAplicaInafectacion,0)=0
GO
	update TPRODUCTO set lImpuesto1=1, lImpuesto4=1, lImpuesto7=1, lImpuesto10=1, lImpuesto13=1
	where isnull(lImpuesto10,0)=1 and isnull(tAplicaExoneracion,0)=0 and isnull(tAplicaInafectacion,0)=0
GO
	update TPRODUCTO set lImpuesto1=1, lImpuesto4=1, lImpuesto7=1, lImpuesto10=1, lImpuesto13=1
	where isnull(lImpuesto13,0)=1 and isnull(tAplicaExoneracion,0)=0 and isnull(tAplicaInafectacion,0)=0
-- UPDATE TPRODUCTO set limpuesto1='1',limpuesto4='1',limpuesto7='1',limpuesto10='1',limpuesto13='1'
GO
	UPDATE TPARAMETRO SET tVersion='4.95.28'
GO



	PRINT 'OK'
