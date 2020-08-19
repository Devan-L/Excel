CREATE FUNCTION dbo.GetTemporaryNationalInsuranceNumber(
    @DateOfBirth DATE,
    @Gender CHAR(1)
)
RETURNS CHAR(9)
BEGIN
    RETURN 'TN' + CONVERT(VARCHAR(6), @DateOfBirth, 12) + @Gender
END