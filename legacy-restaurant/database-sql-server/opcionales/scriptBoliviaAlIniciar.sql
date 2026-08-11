--activacion del parametro bolivia
UPDATE TTABLA SET NVALOR = 0 WHERE TTABLA='PAISORIGEN' AND TCODIGO='000' 
UPDATE TTABLA SET NVALOR = 1 WHERE TTABLA='PAISORIGEN' AND TCODIGO='001'
--actualizacion de tabla parametro
update tparametro set impuesto2=0, timpuesto2='',   --quita el servicio de los parametros
					  impuesto3=0, timpuesto3='',   -- quita el tercer servicio de los parametros
					  nlongitud=10  -- la longitud del identificador tributario
-- actualizacion de los impuestos en tabla producto
update tproducto set limpuesto2=0, limpuesto3=0, limpuesto5=0, limpuesto6=0, limpuesto8=0, limpuesto9=0, limpuesto11=0, limpuesto12=0, limpuesto14=0, limpuesto15=0
--actualizacion sobre multiareas de produccion
update tproducto set lmultiarea=0
update tcaja set tsubalmacen=''
--actualizacion de las cajas: documentos en desglose
 update ttipodocumentoimpresora set limpuesto1=0, limpuesto2=0, limpuesto3=0, lequivadolares=0

--actualizacion de parametro para solo nit de 10
update tparametro set llongitud=1 , nlongitud=10 
