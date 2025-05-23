USE [InvestigacionDB]
GO
/****** Object:  Table [dbo].[Consultas]    Script Date: 23/05/2025 15:42:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Consultas](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Prompt] [nvarchar](1000) NOT NULL,
	[Resultado] [nvarchar](max) NOT NULL,
	[FechaConsulta] [datetime] NULL,
PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[Consultas] ADD  DEFAULT (getdate()) FOR [FechaConsulta]
GO
