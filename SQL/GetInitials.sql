CREATE TABLE People (
	Forename VARCHAR(200) NOT NULL
)

INSERT INTO People VALUES ('Adfglk Gdfg'), ('Bsdf Ss F'), ('Chfg O''Fdfg'), ('Dfghgfc'), ('E'), ('Fgdfgfd Shgff Gewfsd'), ('Ggbdf Ffdsf-Hdfgd')

GO

CREATE FUNCTION dbo.GetInitials(@input VARCHAR(MAX))
RETURNS VARCHAR(100)
AS
BEGIN

	DECLARE @initials VARCHAR(100) = ''
	DECLARE @currentCharacter CHAR = ''
	DECLARE @lastCharacter CHAR = ' '
	DECLARE @currentIndex INT = 1

	WHILE @currentIndex <= LEN(@input)
	BEGIN		
		SET @currentCharacter = SUBSTRING(@input, @currentIndex, 1)

		IF @lastCharacter = ' ' AND @currentCharacter <> ' '
		BEGIN
			SET @initials = @initials + @currentCharacter
		END

		SET @lastCharacter = @currentCharacter
		SET @currentIndex = @currentIndex + 1
	END
	
	RETURN @initials
END	

GO

SELECT *, dbo.GetInitials(Forename) AS Initials FROM People
