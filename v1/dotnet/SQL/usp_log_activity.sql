USE [Envelope]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_log_activity]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_log_activity]
GO

SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE [dbo].[usp_log_activity]
(
	@logTimestamp		datetime,
	@APIKey				varchar(32),
	@ImpersonatedUser	varchar(200),
	@APIMethod			varchar(100),
	@MethodSuccess		bit,
	@IPAddress			varchar(45),
	@MiscData			varchar(MAX)
)
AS

INSERT INTO tblLog (logTimestamp, APIKey, ImpersonatedUser, APIMethod, MethodSuccess, IPAddress, MiscData)
VALUES (@logTimestamp, @APIKey, @ImpersonatedUser, @APIMethod, @MethodSuccess, @IPAddress, @MiscData)

GO

SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

GRANT  EXECUTE  ON [dbo].[usp_log_activity]  TO [exchangews_user]
GO
