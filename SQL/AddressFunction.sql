CREATE FUNCTION GetAddress
(
	@AddressLine1 VARCHAR(500),
	@AddressLine2 VARCHAR(500),
	@AddressLine3 VARCHAR(500),
	@AddressLine4 VARCHAR(500),
	@AddressLine5 VARCHAR(500),
	@AddressLine6 VARCHAR(500)
)
RETURNS VARCHAR(3000)
AS
BEGIN
	IF @AddressLine1 = 'NK#'
		RETURN @AddressLine1

	DECLARE @Address VARCHAR(3000) = @AddressLine1

	IF @AddressLine2 IS NOT NULL
		SET @Address = @Address + '#' + @AddressLine2

	IF @AddressLine3 IS NOT NULL
			SET @Address = @Address + '#' + @AddressLine3

	IF @AddressLine4 IS NOT NULL
			SET @Address = @Address + '#' + @AddressLine4

	IF @AddressLine5 IS NOT NULL
			SET @Address = @Address + '#' + @AddressLine5

	IF @AddressLine6 IS NOT NULL
			SET @Address = @Address + '#' + @AddressLine6

	RETURN @Address
END
GO

SELECT 'dgdfg', 'ghfgh', dbo.GetAddress('beep', 'boop', NULL, NULL, NULL, NULL)
