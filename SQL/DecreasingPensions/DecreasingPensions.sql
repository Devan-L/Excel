IF OBJECT_ID('tempdb..#MemberPension') IS NOT NULL
BEGIN
	DROP TABLE #MemberPension
END

SELECT *,
	RANK() OVER (PARTITION BY MEMID ORDER BY ItemDate) AS [DateRank],
	RANK() OVER (PARTITION BY MEMID ORDER BY Amount) AS [AmountRank]
INTO #MemberPension FROM RetirementItems 
ORDER BY MEMID, ItemDate

SELECT * FROM #MemberPension
WHERE DateRank <> AmountRank
ORDER BY MEMID
