DECLARE @TableName varchar(50)
SET @TableName = 'Sessions'

SET NOCOUNT ON; -- Hide row count so printed output is not affected

DECLARE @DataTypeName varchar(50)
DECLARE @NewLine char
DECLARE @ColumnName varchar(50)
DECLARE @DataType varchar(50)
DECLARE @FieldName varchar(50)

PRINT 'Public Class ' + @TableName;

DECLARE DeclarationCursor CURSOR SCROLL FOR 
    SELECT
        columns.name [ColumnName],
        CASE 
            WHEN columns.system_type_id = 34    THEN 'Byte[]'
            WHEN columns.system_type_id = 35    THEN 'String'
            WHEN columns.system_type_id = 36    THEN 'System.Guid'
            WHEN columns.system_type_id = 48    THEN 'Byte'
            WHEN columns.system_type_id = 52    THEN 'Short'
            WHEN columns.system_type_id = 56    THEN 'Integer'
            WHEN columns.system_type_id = 58    THEN 'System.DateTime'
            WHEN columns.system_type_id = 59    THEN 'float'
            WHEN columns.system_type_id = 60    THEN 'Decimal'
            WHEN columns.system_type_id = 61    THEN 'System.DateTime'
            WHEN columns.system_type_id = 62    THEN 'double'
            WHEN columns.system_type_id = 98    THEN 'Object'
            WHEN columns.system_type_id = 99    THEN 'String'
            WHEN columns.system_type_id = 104   THEN 'Boolean'
            WHEN columns.system_type_id = 106   THEN 'Decimal'
            WHEN columns.system_type_id = 108   THEN 'Decimal'
            WHEN columns.system_type_id = 122   THEN 'Decimal'
            WHEN columns.system_type_id = 127   THEN 'long'
            WHEN columns.system_type_id = 165   THEN 'Byte[]'
            WHEN columns.system_type_id = 167   THEN 'String'
            WHEN columns.system_type_id = 173   THEN 'Byte[]'
            WHEN columns.system_type_id = 175   THEN 'string'
            WHEN columns.system_type_id = 189   THEN 'Long'
            WHEN columns.system_type_id = 231   THEN 'String'
            WHEN columns.system_type_id = 239   THEN 'String'
            WHEN columns.system_type_id = 241   THEN 'String'
            WHEN columns.system_type_id = 241   THEN 'String'
        END [DataType]
FROM              sys.tables tables
    INNER JOIN    sys.schemas schemas ON (tables.schema_id = schemas.schema_id )
    INNER JOIN    sys.columns columns ON (columns.object_id = tables.object_id) 
WHERE
    tables.name = @TableName
ORDER BY 
    columns.object_id ASC;


OPEN DeclarationCursor;

DECLARE @NEWSETTER VARCHAR(MAX) = ''

FETCH FIRST FROM DeclarationCursor 
INTO @ColumnName, @DataType;

WHILE @@FETCH_STATUS = 0
BEGIN
    PRINT '    Public Property ' + @ColumnName + ' As ' + @DataType;
    FETCH NEXT FROM DeclarationCursor 
    INTO @ColumnName, @DataType;
END

CLOSE DeclarationCursor;
DEALLOCATE DeclarationCursor;

PRINT 'End Class';