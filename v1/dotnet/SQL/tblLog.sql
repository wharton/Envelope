USE [Envelope]
GO

/****** Object:  Table [dbo].[tblLog]    Script Date: 05/14/2010 11:35:43 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tblLog]') AND type in (N'U'))
	DROP TABLE [dbo].[tblLog]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tblLog](
	[id] [bigint] IDENTITY(1,1) NOT NULL,
	[logTimestamp] [datetime] NOT NULL,
	[APIKey] [varchar](32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[ImpersonatedUser] [varchar](200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[APIMethod] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[MethodSuccess] [bit] NOT NULL,
	[IPAddress] [varchar](45) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[miscdata] [varchar](max) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblLog_miscdata]  DEFAULT ('{}'),
 CONSTRAINT [PK_tblLog] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
