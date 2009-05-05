/**
	1. Create a database with the name: ajaxedtest
	2. Execute the following script
	3. Run Database tests within ajaxedconsole
**/

USE [ajaxedtest]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[person](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[firstname] [nvarchar](255) NULL,
	[lastname] [nvarchar](255) NULL,
	[age] [int] NULL,
	[cool] [tinyint] NULL,
 CONSTRAINT [PK_person] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

/**
	Data
**/

INSERT INTO [ajaxedtest].[dbo].[person] VALUES ('Michal','Gabrukiewicz','26','0');
INSERT INTO [ajaxedtest].[dbo].[person] VALUES ('cool','and the gang','48','1');