-- EXEC [dbo].[BulkInsertTextFile] @TargetTable ='Stage.VEmployee', @SourceFile = 'C:\Temp\Intermarket\DataSeed\Employee.txt'

CREATE PROCEDURE [dbo].[BulkInsertTextFile]
	@TargetTable nvarchar(100),
	@SourceFile nvarchar(250),
	@FieldTerminator nchar(1) = ','
AS
BEGIN

DECLARE @SQLQuery nvarchar(max)
SET @SQLQuery = 'BULK INSERT '  + @TargetTable + ' FROM ''' + @SourceFile +
''' WITH ( FIELDTERMINATOR = ''' + @FieldTerminator+ ''', ROWTERMINATOR = ''\n'')'
EXEC(@SQLQuery)


RETURN 0
END