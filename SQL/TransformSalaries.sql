CREATE DATABASE SalaryConversion
GO

USE SalaryConversion
GO

CREATE TABLE Input (
	MemberRef VARCHAR(100),
	Year INT NOT NULL,
	Value DECIMAL NOT NULL
)

GO

BULK INSERT Input
FROM 'C:\Users\User\OneDrive\Devan\Github\Excel\Salaries.csv'
WITH
(
	FIELDTERMINATOR = ',',
	ROWTERMINATOR = '\n',
	FIRSTROW = 2
)

GO

CREATE VIEW RankedSalaries
AS
SELECT 
	*, 
	ROW_NUMBER() OVER (PARTITION BY MemberRef, Year ORDER BY (SELECT NULL)) AS SalaryRank
FROM Input

GO

CREATE VIEW [dbo].[TransformedSalaries]
AS
SELECT MemberRef, [Year], [1] AS Salary, [2] AS Salary_1, [3] AS Salary_2, [4] AS Salary_3
FROM
(
	SELECT MemberRef, Value, Year, SalaryRank
	 FROM RankedSalaries
) AS source
PIVOT
(
	 SUM(Value)
	 FOR SalaryRank IN (
		 [1] , [2], [3], [4]
	)
) AS PivotTable;

GO

SELECT * FROM TransformedSalaries