USE [SICIP]
GO

/****** Object:  StoredProcedure [dbo].[DevolucionHerramienta_PROCESO_ActualizaBeta]    Script Date: 02/19/2016 08:30:22 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[DevolucionHerramienta_PROCESO_ActualizaBeta]
	@ValeHerramienta	int	
    ,@CveUsuario Varchar(20)
    ,@Observaciones varchar(200)
	,@XML varchar(8000)
	,@XML2 varchar(8000)
	,@XML3 varchar(8000)
	 AS

SET NOCOUNT ON  
  
DECLARE @hdoc int
	,@DevolucionHerramienta int

SELECT @DevolucionHerramienta = MAX(CveDevolucionHerramienta) + 1
FROM DevolucionHerramienta

IF @DevolucionHerramienta IS NULL SELECT @DevolucionHerramienta = 1

INSERT INTO DevolucionHerramienta (CveDevolucionHerramienta,Fecha,Observaciones,NombreAutoriza,CveValeHerramienta)
SELECT @DevolucionHerramienta
		,GETDATE()
		,@Observaciones
		,@CveUsuario
		,@ValeHerramienta

EXEC sp_xml_preparedocument @hdoc OUTPUT, @XML  

INSERT INTO DevolucionHerramientaDetalle (CveDevolucionHerramienta,CveArticulo,Cantidad,Observaciones)  
SELECT @DevolucionHerramienta,A.A,A.C,''
FROM OPENXML(@hdoc, '/O/D',1) WITH(A int,C tinyint) as A  	


UPDATE VH
	SET VH.CveValeHerramientaEstatus = CASE WHEN DV.CantidadRegresada = VHD.Cantidad THEN 3  --Entregada al 100%
											WHEN DV.CantidadRegresada < VHD.Cantidad THEN 2  --Entrega parcial
											else VH.CveValeHerramientaEstatus
										END
	--selecT *
FROM ValeHerramienta VH 
	JOIN ValeHerramientaDetalle VHD ON VHD.CveValeHerramienta = VH.CveValeHerramienta
	JOIN (SELECT VD.CveValeHerramienta,VDD.CveArticulo,SUM(VDD.Cantidad) CantidadRegresada
			FROM DevolucionHerramienta VD 
			JOIN DevolucionHerramientaDetalle VDD ON VDD.CveDevolucionHerramienta = VD.CveDevolucionHerramienta
			group by VD.CveValeHerramienta,VDD.CveArticulo) DV ON VH.CveValeHerramienta = DV.CveValeHerramienta AND DV.CveArticulo = VHD.CveArticulo
WHERE VH.CveValeHerramienta = @ValeHerramienta

UPDATE VHD
	SET VHD.PendienteEntrega = CASE WHEN DV.CantidadRegresada = VHD.Cantidad THEN 0  --Entregada al 100%
											WHEN DV.CantidadRegresada < VHD.Cantidad THEN 1  --Entrega parcial
											else VHD.PendienteEntrega
										END
	--selecT *
FROM ValeHerramienta VH 
	JOIN ValeHerramientaDetalle VHD ON VHD.CveValeHerramienta = VH.CveValeHerramienta
	JOIN (SELECT VD.CveValeHerramienta,VDD.CveArticulo,SUM(VDD.Cantidad) CantidadRegresada
			FROM DevolucionHerramienta VD 
			JOIN DevolucionHerramientaDetalle VDD ON VDD.CveDevolucionHerramienta = VD.CveDevolucionHerramienta
			group by VD.CveValeHerramienta,VDD.CveArticulo) DV ON VH.CveValeHerramienta = DV.CveValeHerramienta AND DV.CveArticulo = VHD.CveArticulo
WHERE VH.CveValeHerramienta = @ValeHerramienta
GO


