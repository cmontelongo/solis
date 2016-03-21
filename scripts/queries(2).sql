DECLARE @XMLCargo XML,
 @ValeHerramienta INT;

SET @XMLCargo = --'<root><row orderid=''323936'' ordernumber=''125924'' cargoready=''01/11/2010 00:00:00'' estpallets=''1'' estweight=''2'' estvolume=''3''/><row orderid=''326695'' ordernumber=''128176'' cargoready=''01/21/2010 00:00:00'' estpallets=''4'' estweight=''5'' estvolume=''6''/></root>';
'<O Nombre="pablo" ><D A="432" C="10"/><D A="1414" C="20"/><D A="1415" C="30"/></O>';
SET @ValeHerramienta = 2;

DECLARE @ValeHerramientaDetalle TABLE (
  CveValeHerramienta INT,
  CveArticulo INT,
  Cantidad TINYINT,
  PendienteEntrega BIT);
INSERT INTO @ValeHerramientaDetalle(CveValeHerramienta, CveArticulo, Cantidad, PendienteEntrega) VALUES (1,432,4,0);
INSERT INTO @ValeHerramientaDetalle(CveValeHerramienta, CveArticulo, Cantidad, PendienteEntrega) VALUES (2,432,1,1);
INSERT INTO @ValeHerramientaDetalle(CveValeHerramienta, CveArticulo, Cantidad, PendienteEntrega) VALUES (2,1414,2,0);
INSERT INTO @ValeHerramientaDetalle(CveValeHerramienta, CveArticulo, Cantidad, PendienteEntrega) VALUES (2,1415,3,1);
INSERT INTO @ValeHerramientaDetalle(CveValeHerramienta, CveArticulo, Cantidad, PendienteEntrega) VALUES (3,432,1,1);
INSERT INTO @ValeHerramientaDetalle(CveValeHerramienta, CveArticulo, Cantidad, PendienteEntrega) VALUES (4,821,10,1);

SELECT * 
  FROM @ValeHerramientaDetalle;


SELECT CveArticulo = T.Item.value('@A', 'INT'),
       Cantidad = T.Item.value('@C', 'TINYINT')
  FROM @XMLCargo.nodes('/O/D') AS T(Item);


UPDATE @ValeHerramientaDetalle 
   SET CveArticulo = T.Item.value('@A', 'INT'),
       Cantidad = T.Item.value('@C', 'TINYINT')
  FROM @XMLCargo.nodes('/O/D') AS T(Item)
 WHERE CveValeHerramienta=@ValeHerramienta
   AND CveArticulo=T.Item.value('@A', 'INT');

SELECT * 
  FROM @ValeHerramientaDetalle;
