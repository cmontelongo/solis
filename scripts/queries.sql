SELECT [CveValeHerramienta]
      ,[Fecha]
      ,[CveValeHerramientaEstatus]
      ,[Nombre]
      ,[Observaciones]
      ,[NombreAutoriza]
  FROM [SICIP].[dbo].[ValeHerramienta]
GO


SELECT [CveValeHerramienta]
      ,[CveArticulo]
      ,[Cantidad]
      ,[PendienteEntrega]
  FROM [SICIP].[dbo].[ValeHerramientaDetalle]
GO


SELECT 'INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES ('+convert(varchar(12),CveValeHerramientaEstatus)+','''+Nombre+''','+convert(varchar(2),Activo)+') GO'
  FROM ValeHerramientaEstatus
GO

INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES (1,'Entregado',1 )
INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES (2,'Devolucion Incomplet1',1 )
INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES (3,'Devuelto',1 )

select * from ValeAlmacenEstatus

SELECT 'Total income is', ((OrderQty * UnitPrice) * (1.0 - UnitPriceDiscount)), ' for ',
p.Name AS ProductName 
FROM Production.Product AS p 
INNER JOIN Sales.SalesOrderDetail AS sod
ON p.ProductID = sod.ProductID 
ORDER BY ProductName ASC;














DECLARE @idoc int, @doc varchar(1000);
SET @doc ='
<ROOT>
<Customer CustomerID="VINET" ContactName="Paul Henriot">
   <Order CustomerID="VINET" EmployeeID="5" OrderDate="1996-07-04T00:00:00">
      <OrderDetail OrderID="10248" ProductID="11" Quantity="12"/>
      <OrderDetail OrderID="10248" ProductID="42" Quantity="10"/>
   </Order>
</Customer>
<Customer CustomerID="LILAS" ContactName="Carlos Gonzlez">
   <Order CustomerID="LILAS" EmployeeID="3" OrderDate="1996-08-16T00:00:00">
      <OrderDetail OrderID="10283" ProductID="72" Quantity="3"/>
   </Order>
</Customer>
</ROOT>';
--Create an internal representation of the XML document.
EXEC sp_xml_preparedocument @idoc OUTPUT, @doc;
-- Execute a SELECT statement that uses the OPENXML rowset provider.
SELECT    *
FROM       OPENXML (@idoc, '/ROOT/Customer',1)
            WITH (CustomerID  varchar(10),
                  ContactName varchar(20));
SELECT    *
FROM       OPENXML (@idoc, '/ROOT/Customer/Order/OrderDetail',1)
            WITH (OrderID  varchar(7),
                  ProductID varchar(3),
                  Quantity varchar(3));









DECLARE @ValeHerramienta int, @hdoc int, @XML varchar(1000);
SET @ValeHerramienta=1002;
SET @XML ='<O Nombre="Cesar Solis" Usuario="MIGUEL" Obs="1234"><D A="398" C="1"/></O>';
	EXEC sp_xml_preparedocument @hdoc OUTPUT, @XML  

--	INSERT INTO ValeHerramientaDetalle (CveValeHerramienta,CveArticulo,Cantidad,PendienteEntrega)  
	SELECT @ValeHerramienta, A.A, A.C, 1
	FROM OPENXML(@hdoc, '/O/D',1) WITH(A int,C tinyint) as A  	




select * from ValeHerramienta

select * from valealmacen
select * from valealmacendetalle




--------------------------------------------------------------------


SELECT A.CveArticulo,A.Nombre + ISNULL('  ('+A.Codigo+')','') Nombre 
FROM Articulo AS A JOIN Familia F ON A.CveFamilia = F.CveFamilia
WHERE A.Activo=1 AND (F.CveFamilia in(9,4) OR F.CveRama = 2) 
/* AND (A.Codigo like '%" & Replace(txtBuscar.Text, " ", "%") & "%' OR A.Nombre like '%" & Replace(txtBuscar.Text, " ", "%") & "%') " & _
     strCondicion & " */
ORDER BY A.Nombre

select * from articulo
select * from Familia


SELECT A.CveArticulo,A.Nombre + ISNULL('  ('+A.Codigo+')','') Nombre 
FROM Articulo AS A JOIN Familia F ON A.CveFamilia = F.CveFamilia
WHERE A.Activo=1 AND (F.CveFamilia not in(9,4) OR F.CveRama = 2) 
/* AND (A.Codigo like '%" & Replace(txtBuscar.Text, " ", "%") & "%' OR A.Nombre like '%" & Replace(txtBuscar.Text, " ", "%") & "%') " & _
     strCondicion & " */
ORDER BY A.Nombre

select --*, CveArticulo,A.Nombre,A.Codigo
A.CveArticulo,
A.Nombre,
A.Codigo
from Articulo A
Where A.CveArticulo = 3113

Fab.
Num Parte
Descripcion
Cant --
Cve Parte


select CveArticulo,A.Nombre,A.Codigo from Articulo A Where A.CveArticulo = 1768

select * from ValeAlmacen
select * from ValeAlmacenDetalle





-------------------------------------------------------------

EXEC ValeHerramienta_PROCESO_ActualizaBeta 
@ValeHerramienta=2,@Fecha='2015-10-29',@Nombre ='pablo',@CveUsuario='',@Observaciones=''
,@XML='<O Nombre="pablo" ><D A="432" C="1"/><D A="1414" C="2"/><D A="1415" C="3"/></O>',@XML2='<O Nombre="pablo" ></O>',@XML3='<O Nombre="pablo" ></O>'

select * from ValeHerramientaDetalle;

SELECT Nombre,CveValeHerramienta FROM ValeHerramienta WHERE CveValeHerramientaEstatus in( 1,2) ORDER BY Nombre,CveValeHerramienta;
select * from ValeHerramienta;
update ValeHerramienta set Nombre='cesar' where CveValeHerramienta=2


select VA.CveValeHerramienta,VA.CveArticulo,A.Nombre,A.Codigo,VA.Cantidad,DV.CantidadRegresada
from ValeHerramientaDetalle VA
  JOIN Articulo A ON A.CveArticulo = VA.CveArticulo
  LEFT JOIN (SELECT VD.CveValeHerramienta,VDD.CveArticulo,SUM(VDD.Cantidad) CantidadRegresada
                FROM DevolucionHerramienta VD
                    JOIN DevolucionHerramientaDetalle VDD ON VDD.CveDevolucionHerramienta = VD.CveDevolucionHerramienta
                group by VD.CveValeHerramienta,VDD.CveArticulo) DV ON VA.CveValeHerramienta = DV.CveValeHerramienta AND DV.CveArticulo = VA.CveArticulo
    WHERE VA.CveValeHerramienta = 6
    ORDER BY A.Nombre;

select * from DevolucionHerramienta;
select * from DevolucionHerramientaDetalle;
select * from ValeHerramienta;
select * from ValeHerramientaDetalle;

select * from Articulo where CveArticulo='564';




select * from DevolucionHerramienta;
select * from DevolucionHerramientaDetalle;
--update DevolucionHerramientaDetalle set Cantidad=0 where CveDevolucionHerramienta in (6,7,8);
SELECT VD.CveDevolucionHerramienta,VD.CveValeHerramienta,VDD.CveArticulo,VDD.Cantidad--SUM(VDD.Cantidad) CantidadRegresada
                FROM DevolucionHerramienta VD
                    JOIN DevolucionHerramientaDetalle VDD 
						ON VDD.CveDevolucionHerramienta = VD.CveDevolucionHerramienta
where CveValeHerramienta=6
                group by VD.CveValeHerramienta,VDD.CveArticulo

select VA.CveValeHerramienta,VA.CveArticulo,A.Nombre,A.Codigo,VA.Cantidad-ISNULL(DV.CantidadRegresada,0) PorRegresar
from ValeHerramientaDetalle VA
  JOIN Articulo A ON A.CveArticulo = VA.CveArticulo
  LEFT JOIN (SELECT VD.CveValeHerramienta,VDD.CveArticulo,SUM(VDD.Cantidad) CantidadRegresada
                FROM DevolucionHerramienta VD
                    JOIN DevolucionHerramientaDetalle VDD 
						ON VDD.CveDevolucionHerramienta = VD.CveDevolucionHerramienta
                group by VD.CveValeHerramienta,VDD.CveArticulo) DV ON VA.CveValeHerramienta = DV.CveValeHerramienta AND DV.CveArticulo = VA.CveArticulo
    WHERE VA.CveValeHerramienta = 7
    ORDER BY A.Nombre;
