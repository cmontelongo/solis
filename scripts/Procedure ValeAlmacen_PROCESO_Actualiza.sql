USE [SICIP]
GO
/****** Object:  StoredProcedure [dbo].[ValeHerramienta_PROCESO_ActualizaBeta]    Script Date: 01/26/2016 14:03:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[ValeAlmacen_PROCESO_Actualiza]
	@ValeAlmacen	int	= NULL
	,@Fecha datetime
    ,@Nombre varchar(50)
    ,@CveUsuario Varchar(20)
    ,@Observaciones varchar(200)
	,@MovimientoSalida tinyint
	,@CveObra smallint
	,@CveAlmacen smallint
	,@XML varchar(8000)
	,@XML2 varchar(8000)
	,@XML3 varchar(8000)
	 AS

SET NOCOUNT ON  
  
DECLARE @hdoc int  

IF @ValeAlmacen IS NULL
   BEGIN
	SELECT @ValeAlmacen = MAX(CveValeAlmacen) + 1
	FROM ValeAlmacen

	IF @ValeAlmacen IS NULL SELECT @ValeAlmacen = 1
   END

SELECT @ValeAlmacen

IF exists(SELECT CveValeAlmacen
		FROM ValeAlmacen
		WHERE CveValeAlmacen = @ValeAlmacen)
   BEGIN
	UPDATE ValeAlmacen
	SET Nombre = @Nombre
	,Observaciones = @Observaciones
	,MovimientoSalida = @MovimientoSalida
	,CveObra = @CveObra
	,CveAlmacen = @CveAlmacen
	WHERE CveValeAlmacen = @ValeAlmacen
   END
ELSE
   BEGIN   

	INSERT INTO ValeAlmacen (CveValeAlmacen,Fecha,CveValeAlmacenEstatus,Nombre,Observaciones,NombreAutoriza,MovimientoSalida,CveObra,CveAlmacen)
	SELECT @ValeAlmacen
	,@Fecha
	,1
	,@Nombre
	,@Observaciones
	,@CveUsuario
	,@MovimientoSalida
	,@CveObra
	,@CveAlmacen

	EXEC sp_xml_preparedocument @hdoc OUTPUT, @XML  

	INSERT INTO ValeAlmacenDetalle (CveValeAlmacen, CveArticulo, CveFabrica, NumeroParte, Descripcion, Cantidad, Precio, CveParte, PendienteEntrega)  
	SELECT @ValeAlmacen, 1, A.F, A.N, A.D, A.C, A.P, A.CP, 1
	FROM OPENXML(@hdoc, '/O/D',1) WITH(F varchar(5), N varchar(5), D varchar(100), C int, P varchar(16), CP varchar(5)) as A


   END
