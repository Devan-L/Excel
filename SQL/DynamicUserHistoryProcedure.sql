-- Setup test tables and dummy data
CREATE TABLE Lookups (
	ID INT IDENTITY(1, 1) PRIMARY KEY,
	Scheme VARCHAR(100) NOT NULL,
	TableName VARCHAR(100) NOT NULL,
	ColumnName VARCHAR(100) NOT NULL,
	ActualName VARCHAR(100) NOT NULL
)

CREATE TABLE DataBin (
	ID INT IDENTITY(1, 1) PRIMARY KEY,
	Reference VARCHAR(100) NOT NULL,
	TableName VARCHAR(100) NOT NULL,
	Column1 VARCHAR(100) NULL,
	Column2 VARCHAR(100) NULL,
	Column3 VARCHAR(100) NULL,
	Column4 VARCHAR(100) NULL
)

INSERT INTO Lookups VALUES 
	('A', 'Table1', 'Column1', 'Forename'), ('A', 'Table1', 'Column2', 'Surname'), ('A', 'Table1', 'Column3', 'Title'), ('A', 'Table1', 'Column4', 'FavouriteColour'),
	('A', 'Table2', 'Column1', 'HasALozenge?'), ('A', 'Table2', 'Column2', 'ReleasesTheKraken?'), ('A', 'Table2', 'Column3', 'WhoWantsSandwiches?'),
	('B', 'Table1', 'Column1', 'IsTrue?'), ('B', 'Table1', 'Column2', 'IsFalse?')

INSERT INTO DataBin VALUES 
	('AAA', 'Table1', 'Bob', 'Bobbington', 'Senior', 'Brown'), ('BBB', 'Table1', 'Boblet', 'Bobbinger', 'Lord', 'Pink'),
	('AAA', 'Table2', 'False', 'False', 'False', NULL), ('BBB', 'Table2', 'True', 'True', 'True', NULL),
	('CCC', 'Table1', 'True', 'False', NULL, NULL), ('DDD', 'Table1', 'True', 'False', NULL, NULL)

GO

-- Create dynamic table procedure
CREATE PROCEDURE GetDataBinTable (
	@Scheme VARCHAR(100),
	@TableName VARCHAR(100)
)
AS
BEGIN
	IF OBJECT_ID('tempdb..#TableColumns') IS NOT NULL
	BEGIN
		DROP TABLE #TableColumns
	END

	SELECT * INTO #TableColumns FROM Lookups WHERE TableName = @TableName AND Scheme = @Scheme
	
	IF NOT EXISTS (SELECT 1 FROM #TableColumns)
	BEGIN
		THROW 51000, 'Invalid Scheme and TableName combination', 1
	END

	DECLARE @ColumnsSelect VARCHAR(8000) = ''

	WHILE EXISTS (SELECT 1 FROM #TableColumns)
	BEGIN
		DECLARE @ColumnName VARCHAR(100)
		DECLARE @ActualName VARCHAR(100)	

		SELECT TOP 1 @ColumnName = ColumnName, @ActualName = ActualName FROM #TableColumns
		SET @ColumnsSelect = @ColumnsSelect + IIF(@ColumnsSelect = '', '', ', ') + '[' + @ColumnName + '] AS [' + @ActualName + ']'
		DELETE TOP (1) FROM #TableColumns
	END

	DECLARE @SelectStatement NVARCHAR(MAX) = 'SELECT Reference, ' + @ColumnsSelect + ' FROM DataBin WHERE TableName = ''' + @TableName + ''''
	PRINT @SelectStatement
	EXEC sp_executesql @SelectStatement
END

GO

-- Example usage
EXEC dbo.GetDataBinTable 'B', 'Table1'
