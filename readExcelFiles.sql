CREATE PROCEDURE getFiles
    @FolderPath NVARCHAR(500)
AS
BEGIN
    -- Delete existing records in ExcelFiles table
    DELETE FROM ExcelFiles;

    -- Temporary table to store file names
    CREATE TABLE #TempFiles (FileName NVARCHAR(1000));

    -- Command to change directory to the specified folder
    DECLARE @Cmd1 NVARCHAR(1000);
    SET @Cmd1 = 'CD "' + @FolderPath + '"& dir /s /b /a-d';

    -- Insert file names into the temporary table
    INSERT INTO #TempFiles (FileName)
    EXEC xp_cmdshell @Cmd1;

    -- Insert file names into the ExcelFiles table
    INSERT INTO ExcelFiles (FileName)
    SELECT FileName FROM #TempFiles WHERE RIGHT(FileName, 5) IN ('.xls', '.xlsx');

    -- Drop temporary table
    DROP TABLE #TempFiles;

    -- Iterate through each file and get sheet names
    DECLARE @FilePath NVARCHAR(1000);
    DECLARE cur CURSOR FOR SELECT FileName FROM ExcelFiles;
    OPEN cur;
    FETCH NEXT FROM cur INTO @FilePath;

    WHILE @@FETCH_STATUS = 0
    BEGIN
        EXEC getSheets @File_Name = @FilePath ;
        FETCH NEXT FROM cur INTO @FilePath;
    END

    CLOSE cur;
    DEALLOCATE cur;
END;
GO

CREATE PROCEDURE getSheets 
    @File_Name NVARCHAR(1000)
AS
BEGIN
    -- Ensure Result table exists
    IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Result')
    BEGIN
        CREATE TABLE Result (
            FileName NVARCHAR(1000),
            Sheets NVARCHAR(1000)
        );
    END
    -- Set variables
    DECLARE @linkedServerName SYSNAME = 'TempExcelSpreadsheet';
    DECLARE @excelFileUrl NVARCHAR(1000) = @File_Name;

    -- Remove existing linked server if it exists
    IF EXISTS (SELECT NULL FROM sys.servers WHERE name = @linkedServerName)
    BEGIN
        EXEC sp_dropserver @server = @linkedServerName, @droplogins = 'droplogins';
    END

    -- Add the linked server
    EXEC sp_addlinkedserver
        @server = @linkedServerName,
        @srvproduct = 'ACE 12.0',
        @provider = 'Microsoft.ACE.OLEDB.12.0',
        @datasrc = @excelFileUrl,
        @provstr = 'Excel 12.0;HDR=Yes';

    -- Grab the current user to use as a remote login
    DECLARE @suser_sname NVARCHAR(256) = SUSER_SNAME();
	delete from details
    -- Add the current user as a login
    EXEC sp_addlinkedsrvlogin
        @rmtsrvname = @linkedServerName,
        @useself = 'false',
        @locallogin = @suser_sname,
        @rmtuser = NULL,
        @rmtpassword = NULL;
		INSERT INTO details
		exec sp_tables_ex @linkedServerName

		declare @TableName nvarchar(1000)
		declare curs cursor for select TABLE_NAME from details 
		open curs 
		FETCH NEXT FROM curs INTO @TableName;
		while  @@FETCH_STATUS = 0
		Begin
		insert into Result(FileName,Sheets) values (@File_Name,@TableName)
		FETCH NEXT FROM curs INTO @TableName;
		End
		close curs
		deallocate curs
END;
GO



////////////////////////////////////////////////


CREATE PROCEDURE importData
   
AS 
BEGIN

declare @ExcelFilePath NVARCHAR(1000)
declare @SheetName NVARCHAR(1000)
declare curs cursor for select FileName,Sheets from Result
open curs 
Fetch next from curs into @ExcelFilePath,@SheetName
while @@Fetch_Status=0
Begin
IF EXISTS (SELECT 1 FROM sys.tables WHERE name = 'TempTable' or name='JsonTable')
    DROP TABLE TempTable, JsonTable;

DECLARE @SQLRead NVARCHAR(4000) = '
SELECT * 
into TempTable
FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', 
                ''Excel 12.0;Database='+@ExcelFilePath+';HDR=YES'',
                ''SELECT * FROM ['+Trim (@SheetName)+']'')';
EXEC sp_executesql @SQLRead;
DECLARE @TableName NVARCHAR(100) = 'TempTable';
DECLARE @DynamicSQL NVARCHAR(MAX) = '';

SELECT @DynamicSQL += 
    'IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(''' + @TableName + ''') AND name = ''' + Original + ''')
        EXEC sp_rename ''' + @TableName + '.' + Original + ''', ''' + new + ''', ''COLUMN''; '
FROM Dictionary;
EXEC(@DynamicSQL);

DECLARE @json NVARCHAR(MAX);
DECLARE @query NVARCHAR(MAX) =
'
DECLARE @ColumnsToExclude NVARCHAR(1000);
SET @ColumnsToExclude = ''UNIVERSITY,FACULTY,FILE,STUDENTNO,STUDENT_TYPE,ACADEMIC_LEVEL,CURRENT_YEAR,NAME,FATHER,FAMILY,MOTHER,BIRTHDATE,SEX,SOCIALSECURITYNO,COUNTRY,NATIONALITY,CITY,STREET,BUILDING,FLOOR,MOBILE'';
DECLARE @ColumnsQuery NVARCHAR(MAX);
SET @ColumnsQuery = '''';
SELECT @ColumnsQuery = @ColumnsQuery + ''['' + name + ''],''
FROM sys.columns
WHERE object_id = OBJECT_ID(''TempTable'')
AND CHARINDEX('','' + name + '','', '','' + @ColumnsToExclude + '','') = 0;

-- Remove the trailing comma
SET @ColumnsQuery = LEFT(@ColumnsQuery, LEN(@ColumnsQuery) - 1);

DECLARE @DynamicSQL NVARCHAR(MAX);
SET @DynamicSQL = ''SELECT '' + @ColumnsQuery + '' FROM TempTable FOR JSON AUTO;'';

-- Execute dynamic SQL and store result in @json variable
DECLARE @sql NVARCHAR(MAX) = N''SELECT @json = (SELECT '' + @ColumnsQuery + '' FROM TempTable FOR JSON AUTO);'';
EXEC sp_executesql @sql, N''@json NVARCHAR(MAX) OUTPUT'', @json OUTPUT;
';

-- Execute the dynamic query
EXEC sp_executesql @query, N'@json NVARCHAR(MAX) OUTPUT', @json OUTPUT;
Declare @jt nvarchar(1000)='create table JsonTable (JsonValue nvarchar(3000))'
Exec (@jt)
-- Output the result
INSERT INTO JsonTable (JsonValue)
SELECT value 
FROM STRING_SPLIT(@json, '{');
DELETE FROM JsonTable
WHERE JsonValue='['
declare @enumJson nvarchar(100)= 'alter table JsonTable add i INT IDENTITY(1,1)'
Exec (@enumJson)
DECLARE @ColumnsToExclude NVARCHAR(1000);
SET @ColumnsToExclude = 'UNIVERSITY,FACULTY,FILE,STUDENTNO,STUDENT_TYPE,ACADEMIC_LEVEL,CURRENT_YEAR,NAME,FATHER,FAMILY,MOTHER,BIRTHDATE,SEX,SOCIALSECURITYNO,COUNTRY,NATIONALITY,CITY,STREET,BUILDING,FLOOR,MOBILE';
DECLARE @DropColumns NVARCHAR(MAX) = '';

-- Generate ALTER TABLE DROP COLUMN statements
SELECT @DropColumns = @DropColumns + 'ALTER TABLE TempTable DROP COLUMN ' + name + ';'
FROM sys.columns
WHERE object_id = OBJECT_ID('TempTable')
AND CHARINDEX(name, @ColumnsToExclude) = 0;
-- Execute dynamic SQL
EXEC sp_executesql @DropColumns;
declare @enumTemp nvarchar(100)='alter table temptable add iteration INT IDENTITY(1,1)'
Exec (@enumTemp)

DECLARE @columns NVARCHAR(MAX);

-- Get the column names from TempTable excluding 'iteration'
SELECT @columns = STRING_AGG(QUOTENAME(column_name), ', ')
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'TempTable' AND COLUMN_NAME != 'iteration';

-- Get the column names from JsonTable excluding 'i'
SELECT @columns = CONCAT(@columns, ', ', STRING_AGG(QUOTENAME(column_name), ', '))
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'JsonTable' AND COLUMN_NAME != 'i';

-- Construct the dynamic SQL query
DECLARE @sql NVARCHAR(MAX);
SET @sql = '
INSERT INTO allData (' + @columns + ')
SELECT ' + @columns + '
FROM TempTable t
JOIN JsonTable j ON t.iteration = j.i';

-- Execute the dynamic SQL query
EXEC (@sql);
Fetch next from curs into @ExcelFilePath,@SheetName
END
close curs
deallocate curs
END



///////////////////////////////////////////////////

CREATE PROCEDURE getAllData
@folder nvarchar(100)
as
begin
delete from Result
exec getFiles @FolderPath=@folder
exec importData
End

exec getAllData @folder='C:\ExcelFiles'

///////////////////////////////////////////////////



select * from details
select * from ExcelFiles
SELECT * FROM Result
select * from TempTable
select * from JsonTable
SELECT * FROM ALLdata

DELETE FROM ExcelSheets
DELETE FROM Result
delete from details
delete FROM ALLdata

exec importData @ExcelFilePath='C:\ExcelFiles\excel2\test2.xlsx',@SheetName='data'
Exec getFiles @FolderPath='C:\ExcelFiles'
EXEC getSheets @File_Name='C:\ExcelFiles\excel2\test2.xlsx'
