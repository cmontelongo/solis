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

INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES (1,'Entregado',1 ) GO
INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES (2,'Devolucion Incomplet1',1 ) GO
INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES (3,'Devuelto',1 ) GO


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


