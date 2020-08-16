CREATE FUNCTION dbo.GetPopulatedAddressLineCount(
    @Ad1 VARCHAR(100),
    @Ad2 VARCHAR(100),
    @Ad3 VARCHAR(100),
    @Ad4 VARCHAR(100),
    @City VARCHAR(100),
    @County VARCHAR(100),
    @Country VARCHAR(100)
)
RETURNS INT
AS
BEGIN
    -- Insert all address lines into a single-column table
    DECLARE @addressLines TABLE ( [Value] VARCHAR(100) )
    INSERT INTO @addressLines VALUES (@Ad1), (@Ad2), (@Ad3), (@Ad4), (@City), (@County), (@Country)

    -- Count populated address lines
    DECLARE @addressLineCount INT
    SELECT @addressLineCount = COUNT(*) FROM @addressLines WHERE [Value] IS NOT NULL

    RETURN @addressLineCount
END

GO

CREATE FUNCTION dbo.GetNewAddress(
    @Ad1 VARCHAR(100),
    @Ad2 VARCHAR(100),
    @Ad3 VARCHAR(100),
    @Ad4 VARCHAR(100),
    @City VARCHAR(100),
    @County VARCHAR(100),
    @Postcode VARCHAR(100),
    @Country VARCHAR(100)
)
RETURNS @address TABLE (
    Ad1 VARCHAR(100),
    Ad2 VARCHAR(100),
    Ad3 VARCHAR(100),
    Ad4 VARCHAR(100),
    Ad5 VARCHAR(100),
    Postcode VARCHAR(100)
)
AS
BEGIN
    -- Insert all address lines into a single-column table
    DECLARE @addressLines TABLE ( [Value] VARCHAR(100) )
    INSERT INTO @addressLines VALUES (@Ad1), (@Ad2), (@Ad3), (@Ad4), (@City), (@County), (@Country)

    -- Create new temp table with only populated address lines
    DECLARE @populatedAddressLines TABLE ( [Line] INT IDENTITY(1, 1), [Value] VARCHAR(100) )
    INSERT INTO @populatedAddressLines
    SELECT * FROM @addressLines WHERE [Value] IS NOT NULL

    -- Select each new address line based on index in the populated address line temp table
    DECLARE @New_Ad1 VARCHAR(100), @New_Ad2 VARCHAR(100), @New_Ad3 VARCHAR(100), @New_Ad4 VARCHAR(100), @New_Ad5 VARCHAR(100)

    SELECT @New_Ad1 = [Value] FROM @populatedAddressLines WHERE [Line] = 1
    SELECT @New_Ad2 = [Value] FROM @populatedAddressLines WHERE [Line] = 2
    SELECT @New_Ad3 = [Value] FROM @populatedAddressLines WHERE [Line] = 3
    SELECT @New_Ad4 = [Value] FROM @populatedAddressLines WHERE [Line] = 4
    SELECT @New_Ad5 = [Value] FROM @populatedAddressLines WHERE [Line] = 5

    -- Add converted address line to results
    INSERT INTO @address VALUES (@New_Ad1, @New_Ad2, @New_Ad3, @New_Ad4, @New_Ad5, @Postcode)

    RETURN
END

GO