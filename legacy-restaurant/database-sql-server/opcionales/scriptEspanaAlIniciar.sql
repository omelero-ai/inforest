UPDATE TTABLA set nValor=0 where tTabla='PAISORIGEN'
UPDATE TTABLA set nValor=1 where tTabla='PAISORIGEN' and tCodigo='005'

UPDATE TTABLA set tResumido ='€' where tTabla = 'MONEDA' and tCodigo='01'
UPDATE TTABLA set tResumido ='US$' where tTabla = 'MONEDA' and tCodigo='02'

UPDATE TPARAMETRO set tMonedaN='EUROS', tMonN='€', tMonedaE='DOLARES', tMonE='US$'
