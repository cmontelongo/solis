USE [SICIP]
GO

/****** Object:  StoredProcedure [dbo].[ValeHerramienta_PROCESO_ActualizaBeta]    Script Date: 02/18/2016 19:18:22 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[ValeHerramienta_PROCESO_ActualizaBeta]
	@ValeHerramienta	int	= NULL
	,@Fecha datetime
    ,@Nombre varchar(50)
    ,@CveUsuario Varchar(20)
    ,@Observaciones varchar(200)
	,@XML varchar(8000)
	,@XML2 varchar(8000)
	,@XML3 varchar(8000)
	 AS

SET NOCOUNT ON  
  
DECLARE @hdoc int,
		@XMLCargo XML;

IF @ValeHerramienta IS NULL
   BEGIN
	SELECT @ValeHerramienta = MAX(CveValeHerramienta) + 1
	FROM ValeHerramienta

	IF @ValeHerramienta IS NULL SELECT @ValeHerramienta = 1
   END

SELECT @ValeHerramienta

IF exists(SELECT CveValeHerramienta
		FROM ValeHerramienta
		WHERE CveValeHerramienta = @ValeHerramienta)
   BEGIN
	UPDATE ValeHerramienta
	SET Nombre = @Nombre
	,Observaciones = @Observaciones
	WHERE CveValeHerramienta = @ValeHerramienta;

   END
ELSE
   BEGIN   

	INSERT INTO ValeHerramienta (CveValeHerramienta,Fecha,CveValeHerramientaEstatus,Nombre,Observaciones,NombreAutoriza)
	SELECT @ValeHerramienta
	,@Fecha
	,1
	,@Nombre
	,@Observaciones
	,@CveUsuario
	
	EXEC sp_xml_preparedocument @hdoc OUTPUT, @XML

	INSERT INTO ValeHerramientaDetalle (CveValeHerramienta,CveArticulo,Cantidad,PendienteEntrega)  
	SELECT @ValeHerramienta,A.A,A.C,1
	FROM OPENXML(@hdoc, '/O/D',1) WITH(A int,C tinyint) as A  	
	
   END

GO


