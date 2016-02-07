USE [SICIP]
GO

/*********************************************************************************************/
/****** Object:  Table [dbo].[ValeAlmacen]    Script Date: 01/25/2016 13:35:17 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ValeAlmacen]') AND type in (N'U'))
DROP TABLE [dbo].[ValeAlmacen]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[ValeAlmacen](
	[CveValeAlmacen] [int] NOT NULL,
	[Fecha] [datetime] NOT NULL,
	[CveValeAlmacenEstatus] [tinyint] NOT NULL,
	[Nombre] [varchar](50) NOT NULL,
	[Observaciones] [varchar](200) NULL,
	[NombreAutoriza] [varchar](50) NOT NULL,
	[MovimientoSalida] [tinyint] NOT NULL,
	[CveObra] [smallint] NULL,
	[CveAlmacen] [smallint] NULL,
 CONSTRAINT [PK_ValeAlmacen] PRIMARY KEY CLUSTERED 
(
	[CveValeAlmacen] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


/*********************************************************************************************/
/****** Object:  Table [dbo].[ValeAlmacenDetalle]    Script Date: 01/25/2016 13:35:38 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ValeAlmacenDetalle]') AND type in (N'U'))
DROP TABLE [dbo].[ValeAlmacenDetalle]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ValeAlmacenDetalle](
	[CveValeAlmacen] [int] NOT NULL,
	[CveArticulo] [int] NOT NULL,
	[CveFabrica] [varchar](10) NOT NULL,
	[NumeroParte] [varchar](10) NOT NULL,
	[Descripcion] [varchar](200) NOT NULL,
	[Cantidad] [tinyint] NOT NULL,
	[Precio] [int] NOT NULL,
	[CveParte] [varchar](10) NOT NULL,
	[PendienteEntrega] [bit] NOT NULL,
 CONSTRAINT [PK_ValeAlmacenDetalle] PRIMARY KEY CLUSTERED 
(
	[CveValeAlmacen] ASC,
	[CveArticulo] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


/*********************************************************************************************/
/****** Object:  Table [dbo].[ValeAlmacenEstatus]    Script Date: 01/25/2016 13:36:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[ValeAlmacenEstatus](
	[CveValeAlmacenEstatus] [tinyint] NOT NULL,
	[Nombre] [varchar](30) NOT NULL,
	[Activo] [nchar](10) NOT NULL,
 CONSTRAINT [PK_ValeAlmacenEstatus] PRIMARY KEY CLUSTERED 
(
	[CveValeAlmacenEstatus] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


/*********************************************************************************************/
/****** Object:  View [dbo].[vw_ValeAlmacen]    Script Date: 01/26/2016 13:33:27 ******/
--IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vw_ValeAlmacen]'))
--DROP VIEW [dbo].[vw_ValeAlmacen]
--GO
/*********************************************************************************************/
/****** Object:  View [dbo].[vw_ValeHerramienta]    Script Date: 01/25/2016 14:39:05 ****** /
CREATE VIEW [dbo].[vw_ValeAlmacen] AS
SELECT VH.CveValeAlmacen
	,VH.Fecha,VH.Nombre,VH.Observaciones,VH.NombreAutoriza,VH.CveValeAlmacenEstatus,VHE.Nombre ValeHerramientaEstatus
	,VHD.CveArticulo,VHD.Cantidad
	,A.Codigo,A.Nombre ArticuloNombre

	[CveValeAlmacen] [int] NOT NULL,
	[CveArticulo] [int] NOT NULL,
	[CveFabrica] [varchar](10) NOT NULL,
	[NumeroParte] [varchar](10) NOT NULL,
	[Descripcion] [varchar](200) NOT NULL,
	[Cantidad] [tinyint] NOT NULL,
	[Precio] [int] NOT NULL,
	[CveParte] [varchar](10) NOT NULL,
	[PendienteEntrega] [bit] NOT NULL,

FROM ValeAlmacen VH WITH (NOLOCK)
	JOIN ValeAlmacenEstatus VHE WITH (NOLOCK) ON VHE.CveValeAlmacenEstatus = VH.CveValeAlmacenEstatus
	JOIN ValeAlmacenDetalle VHD WITH (NOLOCK) ON VHD.CveValeAlmacen = VH.CveValeAlmacen
	JOIN Articulo A WITH (NOLOCK) ON A.CveArticulo = VHD.CveArticulo

GO
*/



/******************* /
SELECT 'INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES ('+convert(varchar(12),CveValeHerramientaEstatus)+','''+Nombre+''','+convert(varchar(2),Activo)+') GO'
  FROM ValeHerramientaEstatus
GO

SELECT * FROM ValeAlmacenEstatus
*/

INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES (1,'Entregado',1 )
GO
INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES (2,'Devolucion Incomplet1',1 )
GO
INSERT INTO ValeAlmacenEstatus(CveValeAlmacenEstatus,Nombre,Activo) VALUES (3,'Devuelto',1 )
GO
