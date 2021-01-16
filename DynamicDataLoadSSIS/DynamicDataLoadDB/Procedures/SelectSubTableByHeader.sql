
/*===================================================================================================================================================

	Procedure:			dbo.SelectSubTableByHeader

	Parameters:			
						@SourceTableName	-- Source Table that is creating from LoadXLSFile package	  
						@ColumnName			-- Column header to be found ( first column header name)
						@NumberColumns      -- Number of columns after firts column to be extracted 
						@TargetTableName    -- Target table where fragment of found selection will be copied 
										    
	Description:		

						
	Created by:			Yaroslav Dobryanskyy
	Created on:			2021-01-16
------------------------------------------------------------------------------------------------------------------------------------------------------
	Sample calls:  

		EXEC  dbo.SelectSubTableByHeader 
			@SourceTableName = 'CUSTOMER_SHEET1', @ColumnName = 'CustomerID', @NumberColumns = 4 , @TargetTableName = 'Customer'

		SELECT * FROM Cusromer

=====================================================================================================================================================*/

CREATE PROCEDURE [dbo].[SelectSubTableByHeader]
(
  @SourceTableName nvarchar(100)	
  ,@ColumnName nvarchar(50)		
  ,@NumberColumns int		    
  ,@TargetTableName  nvarchar(100) = NULL	
)
AS
BEGIN
  SET NOCOUNT ON;

  DECLARE @query nvarchar(max) 
  DECLARE @ColumnList NVARCHAR(MAX)
  DECLARE @Tmp TABLE (RowID int, ColumnID int, ColumnAlias NVARCHAR(100))
 
 -- Boundaries for subselect
  DECLARE @ColumnStartID int
  DECLARE @ColumnLastID int
  DECLARE @HeaderRowID int
  DECLARE @LastRowID int
  DECLARE @Index int
 
  -- List of columns available at table
  SELECT @ColumnList = STUFF((SELECT ',' +   QUOTENAME(name)                     
  FROM sys.columns s
  WHERE objecT_id = OBJECT_ID(@SourceTableName)  
	AND [name] like 'COLUMN%'	
  ORDER BY [name]
  FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)') ,1,1,'')
  SELECT @ColumnLastID= MAX(CONVERT(int,value)) FROM STRING_SPLIT(REPLACE(REPLACE( @ColumnList,'[COLUMN',''),']',''),',')

  -- Find ID for row and column by search text = @ColumnName 
  SET @query =  N'SELECT RowID, REPLACE(column_name,''COLUMN'','''') ColumnID' +
				N' FROM ' + @SourceTableName + 
				N' UNPIVOT(  CellValue for column_name IN ('+  @ColumnList + 
				N'))AS tunpiv WHERE CellValue =''' + @ColumnName + N''''
  DELETE FROM @Tmp
  INSERT INTO @Tmp(RowID, ColumnID) EXEC (@query)

  SELECT TOP 1 @HeaderRowID = RowID, @ColumnStartID = ColumnID FROM @Tmp	
  SET @ColumnLastID = IIF( @ColumnStartID + @NumberColumns -1 >  @ColumnLastID, @ColumnLastID,  @ColumnStartID + @NumberColumns -1 )

  -- Find Last Row ID
  SET @query = N'SELECT RowID, 0 ColumnID FROM (  SELECT RowID, [COLUMN' + CONVERT(NVARCHAR(10), @ColumnStartID) + 
				N'] AS FirstColumn , LEAD([COLUMN' + CONVERT(NVARCHAR(10), @ColumnStartID) + 
				N']) OVER(ORDER BY ROWID) AS LeadColumn FROM ' + @SourceTableName + 
				N' WHERE RowID >' +CONVERT(NVARCHAR(10), @HeaderRowID)  + 
				N') A WHERE A.LeadColumn ='''' AND A.FirstColumn !='''' '
   
  DELETE FROM @Tmp
  INSERT INTO @Tmp(RowID, ColumnID)  EXEC (@query) 
  SELECT @LastRowID = RowID FROM @Tmp

  -- New Column list with Aliases from Header Row
  SET @ColumnList = N'[COLUMN'+ CONVERT(NVARCHAR(10),@ColumnStartID) +N']'
  SET @Index = @ColumnStartID+1

  WHILE( @Index <= @ColumnLastID)
  BEGIN
	SET @ColumnList += CONCAT(N',[COLUMN',CONVERT(NVARCHAR(10),@Index),N']')
	SET @Index+=1	
  END

  SET @query =  N'SELECT column_name +'' AS ['' +  CellValue +'']'' FROM ' + @SourceTableName + 
				N' UNPIVOT(  CellValue for column_name IN ('+  @ColumnList + 
				N'))AS tunpiv WHERE RowID =' + CONVERT(NVARCHAR(10),@HeaderRowID)
  DELETE @Tmp
  INSERT INTO @Tmp(ColumnAlias) EXEC (@query)

  SELECT @ColumnList = STUFF((SELECT DISTINCT ',' + ColumnAlias                     
  FROM @Tmp
  FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)') ,1,1,'')
  --SELECT @ColumnList
  
  -- Query return final result 
  IF (@TargetTableName IS NULL)
	 SET @query =  N'SELECT ' + @ColumnList + '  FROM ' + @SourceTableName + 
				N' WHERE RowID >' + CONVERT(NVARCHAR(10),@HeaderRowID) +
				IIF(@LastRowID IS NULL, N'', CONCAT(' AND RowID <=',CONVERT(NVARCHAR(10),@LastRowID)))
  ELSE    
	 SET @query =  N'DROP TABLE IF EXISTS '+ @TargetTableName+ ';SELECT ' + @ColumnList + 'INTO '+ @TargetTableName + ' FROM ' + @SourceTableName + 
				N' WHERE RowID >' + CONVERT(NVARCHAR(10),@HeaderRowID) +
				IIF(@LastRowID IS NULL, N'', CONCAT(' AND RowID <=',CONVERT(NVARCHAR(10),@LastRowID)))

  EXEC (@query)
	
END