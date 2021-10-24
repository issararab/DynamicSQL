-- ========================================================================
-- Author      : <Issar Arab>
-- Created     : <date: 26.08.2020,>
-- Modified    : <date: 15.01.2021,>
-- Description : <Dynamic script distributing incentives to Partners,
--                 in both Germany and Western Europe, 
--                 based on their revenue/consumption made on 6 
--                 different programs: Enterprise, OSA, OSU, Azure, CSP, C3
--                 , AGI, SPLAR, and Hosting, Azure Expert>
-- ========================================================================

DECLARE @SQL NVARCHAR(MAX) = '';
DECLARE @FINAL_QUERY NVARCHAR(MAX) = '';
DECLARE @NewLine NCHAR(1) = NCHAR(10);
DECLARE @Temp nvarchar(100) = '';
DECLARE @LeverTempTable nvarchar(1000) = '';
DECLARE @RevenueTempTable nvarchar(1000) = '';
DECLARE @IncentivesTempTable nvarchar(1000) = '';
DECLARE @Division nvarchar(10) = 'GE'; --WE, GE
DECLARE @ProgramName nvarchar(100) = 'BGI'; -- Enterprise, OSA, OSU, Azure, CSP, C3, AGI, SPLAR, and Hosting, AzureExpert, BGI
DECLARE @TablePath nvarchar(500) = 'dbo.GlobalPrograms_'; -- 'dbo.' for WE, and 'dbo.GlobalPrograms_' for GE 


-- Enterprise, OSU, OSA investments distribution for both Germany and Western Europe
IF (@ProgramName = 'OSA' OR @ProgramName = 'OSU')
BEGIN

	--Change Levermapping to grouped version
	--Clean table
	SET @LeverTempTable = '#LeverMapping' + @ProgramName
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @LeverTempTable +''') IS NOT NULL' + @NewLine + 'DROP TABLE '+@LeverTempTable
	         + @NewLine;
	PRINT (@SQL);
	--EXECUTE sp_executesql @SQL
	SET @FINAL_QUERY += @SQL + @NewLine;
	-- EXEC sys.sp_executesql @SQL;  // EXEC(@sql)
	--Retrieve Lever data
	SET @SQL = 'SELECT LeverName, SummaryPricingLevel, SolutionArea, TopReportingProduct, TopReportingProduct_For_Join' ;
	SET @SQL += CASE WHEN @ProgramName = 'Enterprise' THEN @NewLine + 'INTO ' + @LeverTempTable+ @NewLine + 'FROM '+ @TablePath+@Division+'MappingLever lm' + @NewLine + 'WHERE ProgramGroupName= ''Enterprise'''
					 WHEN @ProgramName = 'OSA' THEN @NewLine + 'INTO ' + @LeverTempTable + @NewLine + 'FROM '+ @TablePath+@Division+'MappingLever lm' + @NewLine + 'WHERE ProgramGroupName like ''%OSA%'''
					 WHEN @ProgramName = 'OSU' THEN ', PartnerAttachType'+ @NewLine + 'INTO '+ @LeverTempTable + @NewLine +'FROM '+ @TablePath+@Division+'MappingLever lm' + @NewLine + 'WHERE ProgramGroupName like ''%OSU%'''
					 END

	SET @SQL += CASE WHEN @ProgramName = 'Enterprise' OR @ProgramName = 'OSA' THEN @NewLine +'GROUP BY LeverName,SummaryPricingLevel,SolutionArea,TopReportingProduct,TopReportingProduct_For_Join'
					 ELSE @NewLine +'GROUP BY LeverName,PartnerAttachType,SummaryPricingLevel,SolutionArea,TopReportingProduct,TopReportingProduct_For_Join'
					 END

	SET @SQL += @NewLine;

	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;
	--Update the lever Mapping table column names

	IF (@ProgramName = 'Enterprise')
	BEGIN
		SET @SQL = 'UPDATE ' + @LeverTempTable + @NewLine + 'SET TopReportingProduct_For_Join=''Enterprise Mobility''' + @NewLine +'WHERE TopReportingProduct_For_Join=''Enterprise Mobility E5''' + @NewLine;
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;
		SET @SQL = 'UPDATE ' + @LeverTempTable + @NewLine + 'SET TopReportingProduct_For_Join= NULL, TopReportingProduct=NULL, SolutionArea=''Azure''' + @NewLine +'WHERE TopReportingProduct_For_Join=''Windows Server''' + @NewLine;
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;
		SET @SQL = 'UPDATE ' + @LeverTempTable + @NewLine + 'SET SolutionArea=''Business Applications''' + @NewLine +'WHERE SolutionArea=''Business Application''';
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;
	END

	--Change Incentivedate
	--Clean table
	SET @IncentivesTempTable = '#Incentives' + @ProgramName;
	SET @SQL = @NewLine + 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @IncentivesTempTable +''') IS NOT NULL' + @NewLine + 'DROP TABLE '+ @IncentivesTempTable 
	         + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	--Retrieve Incentives data
	SET @SQL = 'SELECT ' + CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID, Subregion, Subsidiary,' ELSE '' END+ 'PartnerOneID, PartnerOneName, ProgramGroupName, ProgramName, LeverName, FiscalYear, FiscalMonth, EarningsAmtInCD, SuperRevSumDivisionName, TopReportingProduct, SolutionArea, CustomerName'
	SET @SQL += @NewLine + CASE WHEN @ProgramName = 'Enterprise' AND @Division = 'GE' THEN ', EarningsAmtInLC, REPLACE(FiscalMonth,'' '','', '') AS FiscalMonthYearName' 
								WHEN @ProgramName = 'Enterprise' AND @Division = 'WE' THEN ', REPLACE(FiscalMonth,'' '','', '') AS FiscalMonthYearName' 
								ELSE ',DATENAME(MONTH, FiscalMonth) +  '', ''+ DATENAME(YEAR, FiscalMonth) as FiscalMonthYearName' END 
	SET @SQL += @NewLine + 'INTO ' + @IncentivesTempTable + @NewLine + 'FROM '+ @TablePath+@Division+@ProgramName+'Earnings' + @NewLine;

	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	--Change Revenuedate
	--Temporary revenue table
	--Clean table
	SET @SQL = 'IF OBJECT_ID(''tempdb.dbo.#r1'') IS NOT NULL' + @NewLine + 'DROP TABLE #r1' +@NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	--Retrieve Revenue data
	SET @SQL = 'SELECT DISTINCT ' + CASE WHEN @Division = 'WE' THEN + 'r.PartnerOneSubID, r.Subregion, r.Subsidiary,' ELSE '' END+ 'r.PartnerOneID, d.FiscalMonthYearName, r.SolutionArea, r.TopReportingProduct'
	         + CASE WHEN @ProgramName = 'Enterprise' OR @ProgramName = 'OSA' THEN ',r.SummaryPricingLevel' ELSE ',r.PartnerAttachType' END

	SET @Temp = 'Customer';

	SET @SQL += @NewLine + ',CASE WHEN '+@Temp+'Subsegment IS NULL THEN NULL' + 
				@NewLine + 'WHEN ' + @Temp+'Subsegment LIKE ''%Major%'' THEN ''Enterprise''' 
				+@NewLine + 'WHEN ' + @Temp+'Subsegment LIKE ''%Strategic%'' THEN ''Enterprise'''
				+@NewLine + 'WHEN ' + @Temp+'Subsegment LIKE ''%SMB%'' THEN ''SMB'''
				+@NewLine + 'WHEN ' + @Temp+'Subsegment LIKE ''%SM&C%'' AND ' + @Temp+'SubSegment NOT LIKE ''%SMB%'' THEN ''Corporate'''
				+@NewLine + 'ELSE ''Check'' END AS SegmentGroup,'
				+@NewLine + 'Total, TotalType'-- CASE WHEN @ProgramName = 'OSU' THEN ',[Active Usage]' ELSE ',BilledRevenue' END
				+@NewLine + 'INTO #r1' 
				+@NewLine + 'FROM '+ @TablePath+@Division+@ProgramName+'Revenue r' 
				+@NewLine + 'LEFT JOIN OCPIncentives.dbo.Date d'
				+@NewLine + 'ON LEFT(r.fiscalMonth,4)=d.[FiscalYear(YY)] AND RIGHT(r.FiscalMonth,3)=LEFT(d.CalendarMonthName,3)' +@NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	--Delete negative revenue, Group Segment
	SET @SQL = 'DELETE FROM #r1 WHERE Total<=0' +@NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'DELETE FROM #r1 WHERE SegmentGroup =''Check'' OR SegmentGroup IS NULL' +@NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	--Revenues/Usage table
	--Clean table
	SET @RevenueTempTable = '#Revenues' + @ProgramName;
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @RevenueTempTable +''') IS NOT NULL' + @NewLine + 'DROP TABLE '+@RevenueTempTable;
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	--Retrieve Lever data
	SET @SQL = 'SELECT ' + CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID, Subregion, Subsidiary,' ELSE '' END+ 'PartnerOneID, FiscalMonthYearName, SolutionArea, TopReportingProduct,'+CASE WHEN @ProgramName = 'OSU' THEN 'PartnerAttachType' ELSE 'SummaryPricingLevel' END+', SegmentGroup, TotalType, SUM(Total) AS Total'
				+ @NewLine + 'INTO ' + @RevenueTempTable
				+ @NewLine +'FROM #r1'
				+ @NewLine +'GROUP BY'
				+ @NewLine + CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID, Subregion, Subsidiary,' ELSE '' END+'PartnerOneID, FiscalMonthYearName, SolutionArea, TopReportingProduct, '+CASE WHEN @ProgramName = 'OSU' THEN 'PartnerAttachType' ELSE 'SummaryPricingLevel' END+', SegmentGroup, TotalType'

	SET @SQL += @NewLine;

	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	--Product-Level Levers--
	--Sum revenues/Usage to lever-level
	DECLARE @attachORpricing varchar(255) = CASE WHEN @ProgramName = 'OSU' THEN 'PartnerAttachType' ELSE 'SummaryPricingLevel' END
	--Productlevel including pricing level
	SET @SQL = 'SELECT ' +CASE WHEN @Division = 'WE' THEN + 'r.PartnerOneSubID, r.Subregion, r.Subsidiary,' ELSE '' END+'r.PartnerOneID, r.FiscalMonthYearName, r.SegmentGroup, r.SolutionArea, r.TopReportingProduct'+CASE WHEN @ProgramName = 'OSU' THEN ', r.PartnerAttachType, ' ELSE ', ' END + 'lm.LeverName, r.TotalType, SUM(r.Total) AS Total'
				+ @NewLine + 'INTO ' + @RevenueTempTable +'Leverlevel'
				+ @NewLine + 'FROM ' + @RevenueTempTable +' r'
				+ @NewLine + 'INNER JOIN #LeverMapping' + @ProgramName +' lm'
				+ @NewLine + 'ON r.SolutionArea=lm.SolutionArea '+ CASE WHEN @ProgramName = 'OSU' THEN ' AND r.'+@attachORpricing+'=lm.'+ @attachORpricing ELSE '' END + @NewLine 
				+ @NewLine + 'GROUP BY'
				+ @NewLine + CASE WHEN @Division = 'WE' THEN + 'r.PartnerOneSubID, r.Subregion, r.Subsidiary,' ELSE '' END+'r.PartnerOneID, r.FiscalMonthYearName, r.SegmentGroup, r.SolutionArea, r.TopReportingProduct'+CASE WHEN @ProgramName = 'OSU' THEN ', r.PartnerAttachType, ' ELSE ', ' END+'lm.LeverName, r.TotalType'
	
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;
	--Group incentives back to leverlevel
	PRINT '--Group incentives back to leverlevel'
	--PRINT @RevenueTempTable
	--CASE WHEN @ProgramName = 'OSU' THEN 'ActiveUsage' ELSE 'BilledRevenue' END

	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @RevenueTempTable +'LeverlevelGrouped'') IS NOT NULL' + @NewLine + 'DROP TABLE '+@RevenueTempTable + 'LeverlevelGrouped';
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'SELECT '+CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID, Subregion, Subsidiary,' ELSE '' END+'PartnerOneID, FiscalMonthYearName, SegmentGroup, SolutionArea, TopReportingProduct, LeverName, TotalType, SUM(Total) AS Total'
				+ @NewLine + 'INTO ' + @RevenueTempTable +'LeverlevelGrouped'
				+ @NewLine +'FROM '
				+ @NewLine +'(SELECT * FROM ' + @RevenueTempTable + 'Leverlevel) a'
				+ @NewLine +'GROUP BY'
				+ @NewLine +CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID, Subregion, Subsidiary,' ELSE '' END+'PartnerOneID, FiscalMonthYearName, SegmentGroup, SolutionArea, TopReportingProduct, '+CASE WHEN @ProgramName = 'OSU' THEN 'PartnerAttachType,' ELSE '' END+' LeverName, TotalType'

	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'ALTER TABLE ' + @RevenueTempTable +'LeverlevelGrouped' + @NewLine + 'ADD TotalGrouped float, SegmentShare float '--Go';
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	--Add grouped revenue (on PartnerOneID-FM-Lever), use it to calculate the revenue share of each segment
	SET @SQL = 'UPDATE a SET a.TotalGrouped=b.TotalGrouped'
			   + @NewLine + 'FROM ' + @RevenueTempTable +'LeverlevelGrouped a'
			   + @NewLine + 'LEFT JOIN ( SELECT '+CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID, Subregion, Subsidiary,' ELSE '' END+'PartnerOneID, FiscalMonthYearName, LeverName, SUM(Total) AS TotalGrouped' 
			   + @NewLine + 'FROM ' + @RevenueTempTable +'LeverlevelGrouped'
			   + @NewLine + 'GROUP BY '+CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID, Subregion, Subsidiary,' ELSE '' END+'PartnerOneID, FiscalMonthYearName, LeverName) b 
				ON a.'+CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID' ELSE 'PartnerOneID' END+'=b.'+CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID' ELSE 'PartnerOneID' END+' AND a.FiscalMonthYearName=b.FiscalMonthYearName AND a.leverName=b.leverName'
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL ='UPDATE ' + @RevenueTempTable +'LeverlevelGrouped'
				+ @NewLine + 'SET SegmentShare = Total/TotalGrouped'
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.#Incentives' + @ProgramName +'Distributed'') IS NOT NULL' + @NewLine + 'DROP TABLE #Incentives' + @ProgramName +'Distributed';
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'SELECT ' + CASE WHEN @Division = 'WE' THEN + 'i.PartnerOneSubID, i.Subregion, i.Subsidiary,' ELSE '' END+ 'i.PartnerOneID, i.PartnerOneName, i.ProgramGroupName, i.ProgramName, i.LeverName, i.FiscalYear, i.FiscalMonth, i.EarningsAmtInCD, i.SuperRevSumDivisionName, i.FiscalMonthYearName, ' 
	        + CASE WHEN @ProgramName = 'Enterprise' THEN 'r.TopReportingProduct, r.SolutionArea'  ELSE 'i.TopReportingProduct, i.SolutionArea, i.CustomerName' END
	        + CASE WHEN @ProgramName = 'Enterprise' AND @Division = 'GE' THEN ', i.EarningsAmtInLC' ELSE '' END +', r.SegmentGroup,r.Total, r.TotalType, r.SegmentShare'
			+ @NewLine + 'INTO #Incentives' + @ProgramName +'Distributed'
			+ @NewLine + 'FROM ' + @IncentivesTempTable + ' i'
			+ @NewLine + 'LEFT JOIN ' +  @RevenueTempTable +'LeverlevelGrouped r'
			+ @NewLine + 'ON i.'+CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID' ELSE 'PartnerOneID' END+'=r.'+CASE WHEN @Division = 'WE' THEN + 'PartnerOneSubID' ELSE 'PartnerOneID' END+' AND i.FiscalMonthYearName=r.FiscalMonthYearName AND i.LeverName=r.LeverName'
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	--Distribute Incentives by share
	PRINT('--Distribute Incentives by share')
	SET @SQL = 'ALTER TABLE #Incentives' + @ProgramName +'Distributed'
				+ @NewLine + 'ADD EarningsAmtInCDDistributed float' + CASE WHEN @ProgramName = 'Enterprise' THEN ', EarningsAmtInLCDistributed float' ELSE '' END + @NewLine --+ 'GO'
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'UPDATE #Incentives' + @ProgramName +'Distributed'
			 + @NewLine +  CASE WHEN @ProgramName = 'Enterprise' AND @Division = 'GE' THEN 'SET EarningsAmtInCDDistributed=EarningsAmtInCD*SegmentShare, EarningsAmtInLCDistributed=EarningsAmtInLC*SegmentShare' ELSE 'SET EarningsAmtInCDDistributed=EarningsAmtInCD*SegmentShare' END  
	         + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	--Distribute Incentives to Partners with no SegmentGroup
	SET @SQL = 'UPDATE #Incentives' + @ProgramName +'Distributed'
			+ @NewLine + 'Set SegmentShare = 1, SegmentGroup = ''Unknown'', '
			+ @NewLine +  CASE WHEN @ProgramName = 'Enterprise' AND @Division = 'GE' THEN 'EarningsAmtInCDDistributed=EarningsAmtInCD, EarningsAmtInLCDistributed=EarningsAmtInLC' 
								ELSE 'EarningsAmtInCDDistributed=EarningsAmtInCD' END  
			+ @NewLine + 'WHERE SegmentGroup IS NULL'
	        + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;


END
-- CSP investments distribution
ELSE IF (@ProgramName = 'CSP')
BEGIN
	--Prepare incentives table
	SET @IncentivesTempTable = '#Incentives' + @ProgramName;
	SET @SQL = @NewLine + 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @IncentivesTempTable +'Temp'') IS NOT NULL' + @NewLine + 'DROP TABLE '+ @IncentivesTempTable +'Temp'
	         + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	--Retrieve Incentives data
	SET @SQL = 'SELECT *, FiscalYear +  '' - ''+ CONVERT(CHAR(3), DATENAME(MONTH, FiscalMonth)) as FYMonthName' 
	         + @NewLine + 'INTO ' + @IncentivesTempTable +'Temp'
	         + @NewLine + 'FROM '+@TablePath+@Division+@ProgramName+'Earnings'
	         + @NewLine
	         + @NewLine + 'UPDATE ' + @IncentivesTempTable +'Temp'
	         + @NewLine + 'SET TopReportingProduct = ''Microsoft Azure'' , SolutionArea = ''Applications and Infrastructure'''
	         + @NewLine + 'WHERE TopReportingProduct IS NULL AND (SuperRevSumDivisionName = ''Cognitive Services'' OR LeverName like ''%Azure%'') '
	         + @NewLine ;

	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;
	--Group incentives data
	SET @IncentivesTempTable = '#Incentives' + @ProgramName;
	SET @SQL = @NewLine + 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @IncentivesTempTable +''') IS NOT NULL' + @NewLine + 'DROP TABLE '+ @IncentivesTempTable
	         + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	SET @SQL = 'SELECT '+ CASE WHEN @Division = 'WE' THEN 'Subregion, Subsidiary, PartnerOneSubID, ' ELSE '' END +'ProgramGroupName, ProgramName, LeverName , PartnerOneID, PartnerOneName, FiscalYear, FYMonthName ,FiscalMonth, SolutionArea, CustomerName, TopReportingProduct, SUM(EarningsAmtInCD) AS EarningsAmtInCD' 
	         + @NewLine + 'INTO ' + @IncentivesTempTable + @NewLine + 'FROM ' + @IncentivesTempTable +'Temp'
	         + @NewLine + 'GROUP BY '+ CASE WHEN @Division = 'WE' THEN 'Subregion, Subsidiary, PartnerOneSubID, ' ELSE '' END +'ProgramGroupName, ProgramName, LeverName , PartnerOneID, PartnerOneName, FiscalYear, FYMonthName ,FiscalMonth, SolutionArea, CustomerName, TopReportingProduct' 
	         + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	IF OBJECT_ID('tempdb.dbo.#SubprogramNames') IS NOT NULL
	DROP TABLE #SubprogramNames;
	--Loop over all CSP sub-program names
	create table #SubprogramNames (CSPSubprogram varchar(255), id int identity(1,1))
	INSERT #SubprogramNames(CSPSubprogram) VALUES ('IndirectReseller'),('IndirectProvider'),('DirectProvider')
	declare @CSPCounter int = 1;
	while @CSPCounter < (select max(id)+1 from #SubprogramNames)
	begin 

		--Select * from  #SubprogramNames
		--------------------------------
		DECLARE @SubprogramName varchar(255) = (select CSPSubprogram from #SubprogramNames where id = @CSPCounter);

		PRINT CASE WHEN @SubprogramName = 'IndirectProvider' THEN @NewLine + '--CSPIndirectProvider Revenue'
			WHEN @SubprogramName = 'DirectProvider' THEN @NewLine + '--CSPDirectProvider Revenue'
			WHEN @SubprogramName = 'IndirectReseller' THEN @NewLine + '--CSPIndirectReseller Revenue'
		END
	
		SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.#' + @ProgramName + @SubprogramName +'_Base'') IS NOT NULL' + @NewLine + 'DROP TABLE #'+ @ProgramName + @SubprogramName +'_Base';
		SET @SQL += @NewLine;

		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;

		--Retrieve Revenue data
		SET @SQL = 'Select *'+ ',CASE WHEN CustomerSubsegment IS NULL THEN NULL' + 
			@NewLine + 'WHEN CustomerSubsegment LIKE ''%Major%'' THEN ''Enterprise''' 
			+@NewLine + 'WHEN CustomerSubsegment LIKE ''%Strategic%'' THEN ''Enterprise'''
			+@NewLine + 'WHEN CustomerSubsegment LIKE ''%SMB%'' THEN ''SMB'''
			+@NewLine + 'WHEN CustomerSubsegment LIKE ''%SM&C%'' AND CustomerSubsegment NOT LIKE ''%SMB%'' THEN ''Corporate'''
			+@NewLine + 'ELSE ''Check'' END AS SegmentGroup'
			+@NewLine + 'INTO #' + @ProgramName + @SubprogramName +'_Base'
			+@NewLine + 'FROM '+ @TablePath+ @Division+ @ProgramName + @SubprogramName +'Revenue'
			+@NewLine + 'WHERE CustomerSubSegment <> ''Unknown'''
			+@NewLine + 'AND Total > 0' + @NewLine;
				
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;


		SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.#' + @ProgramName + @SubprogramName +'_Rev'') IS NOT NULL' + @NewLine + 'DROP TABLE #'+ @ProgramName + @SubprogramName +'_Rev';
		SET @SQL += @NewLine;
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;

		SET @SQL = 'SELECT '+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+', FiscalMonth, SolutionArea, SegmentGroup, TotalType, SUM(Total) AS Rev'
			+@NewLine + 'INTO #' + @ProgramName + @SubprogramName +'_Rev'
			+@NewLine + 'FROM #' + @ProgramName + @SubprogramName +'_Base'
			+@NewLine + 'GROUP BY '+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+', FiscalMonth, SolutionArea, SegmentGroup, TotalType' + @NewLine;
				
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;

		SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.#' + @ProgramName + @SubprogramName +'_3level_SUMs'') IS NOT NULL' + @NewLine + 'DROP TABLE #'+ @ProgramName + @SubprogramName +'_3level_SUMs';
		SET @SQL += @NewLine;
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;

		SET @SQL = 'SELECT '+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+', FiscalMonth, SolutionArea, TotalType,SUM(Rev) AS Rev_2'
			+@NewLine + 'INTO #' + @ProgramName + @SubprogramName +'_3level_SUMs'
			+@NewLine + 'FROM #' + @ProgramName + @SubprogramName +'_Rev'
			+@NewLine + 'GROUP BY '+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+', FiscalMonth, SolutionArea, TotalType' + @NewLine;
				
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;

		SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.#' + @ProgramName + @SubprogramName +'_Shares'') IS NOT NULL' + @NewLine + 'DROP TABLE #'+ @ProgramName + @SubprogramName +'_Shares';
		SET @SQL += @NewLine;
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;

		SET @SQL = 'SELECT r.*, s.Rev_2, r.Rev / NULLIF(s.Rev_2,0)  AS SegmentShare'
			+@NewLine + 'INTO #' + @ProgramName + @SubprogramName +'_Shares'
			+@NewLine + 'FROM #' + @ProgramName + @SubprogramName +'_Rev r'
			+@NewLine + 'LEFT JOIN #' + @ProgramName + @SubprogramName +'_3level_SUMs s ON r.'+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+' = s.'+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+' AND r.FiscalMonth = s.FiscalMonth AND r.SolutionArea = s.SolutionArea';
				
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine;

		------------------
		--select leverlevel from #product_level_levers where id = @counter
		set @CSPCounter = @CSPCounter + 1
	END

	Print('--Union Rev')
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.#' + @ProgramName + '_Rev'') IS NOT NULL' + @NewLine + 'DROP TABLE #'+ @ProgramName + '_Rev';
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	SET @SQL = 'SELECT x.*'
		+@NewLine + 'INTO #' + @ProgramName +'_Rev'
		+@NewLine + 'FROM ('
		+@NewLine + 'SELECT ''CSP Indirect Reseller'' AS ProgramName,* FROM #' + @ProgramName +'IndirectReseller_Shares'
		+@NewLine + 'UNION'
		+@NewLine + 'SELECT ''CSP Indirect Provider'' AS ProgramName,* FROM #' + @ProgramName +'IndirectProvider_Shares'
		+@NewLine + 'UNION'
		+@NewLine + 'SELECT ''CSP Direct Provider'' AS ProgramName,* FROM #' + @ProgramName +'DirectProvider_Shares'
		+@NewLine + ') x';
			
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	-- Distribute data
	Print('--Union Rev')
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.#Incentives' + @ProgramName +'Distributed'') IS NOT NULL' + @NewLine + 'DROP TABLE #Incentives' + @ProgramName +'Distributed'
	         + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	SET @SQL = 'SELECT i.*, r.SegmentGroup, r.SegmentShare, i.EarningsAmtInCD * r.SegmentShare AS EarningsAmtInCDDistributed'
		+@NewLine + 'INTO #Incentives' + @ProgramName +'Distributed'
		+@NewLine + 'FROM ' + @IncentivesTempTable + ' i'
		+@NewLine + 'LEFT JOIN #' + @ProgramName +'_Rev r ON i.ProgramName = r.ProgramName AND i.'+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+' = r.'+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+' AND i.FYMonthName = r.FiscalMonth AND i.SolutionArea = r.SolutionArea'  + @NewLine;
		
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	SET @SQL = 'UPDATE #Incentives' + @ProgramName +'Distributed'
		+@NewLine + 'SET EarningsAmtInCDDistributed = EarningsAmtInCD, SegmentShare = 1'
		+@NewLine + 'WHERE EarningsAmtInCDDistributed IS NULL AND SegmentGroup IS NOT NULL'  + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	SET @SQL = 'UPDATE #Incentives' + @ProgramName +'Distributed'
		+@NewLine + 'SET EarningsAmtInCDDistributed = EarningsAmtInCD, SegmentShare = 1 , SegmentGroup = ''Unknown'''
		+@NewLine + 'WHERE EarningsAmtInCDDistributed IS NULL'  + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;

	SET @SQL = 'UPDATE #Incentives' + @ProgramName +'Distributed'
		+@NewLine + 'SET SolutionArea = ''Applications and Infrastructure'''
		+@NewLine + 'WHERE SolutionArea IS NULL AND LeverName LIKE ''%Azure%'''  + @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine;


	--SET @SQL = 'UPDATE #Incentives' + @ProgramName +'Distributed'
	--	+@NewLine + 'SET SolutionArea = ''Unknown'''
	--	+@NewLine + 'WHERE SolutionArea IS NULL' + @NewLine;
	--PRINT (@SQL);
	--SET @FINAL_QUERY += @SQL + @NewLine;

END
-- C3 investments distribution for Germany only
ELSE IF (@ProgramName = 'C3' AND @Division = 'GE')
BEGIN 
	--PRINT('-- Adding the C3 LastRefreshDate')
	--SET @SQL = 'ALTER TABLE [OCPIncentives].[dbo].[GEC3_Data]'  
	--		+ @NewLine + 'ADD LastRefreshDate date' + @NewLine
	--PRINT (@SQL);
	--SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'UPDATE [OCPIncentives].[dbo].[GEC3_Data]'  
			+ @NewLine + 'SET LastRefreshDate = CAST(GETDATE() AS DATE)' + @NewLine
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'DROP TABLE [OCPIncentives].[dbo].[C3_Data_Monthly]'+ @NewLine
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'SELECT DISTINCT a.*, b.[Fiscal Month] AS CalendarMonth'  
			+ @NewLine + ', LEFT(a.Time, 4) +''-''+b.[Fiscal Month] AS FiscalMonth' 
			+ @NewLine + ', LEFT(a.Time, 4) AS FiscalYear' 
			+ @NewLine + ', [Investment Per SA] / 3 AS MonthlyInvesmentPerSA' 
			+ @NewLine + 'INTO [OCPIncentives].[dbo].[C3_Data_Monthly]' 
			+ @NewLine + 'FROM [OCPIncentives].[dbo].[GEC3_Data] a' 
			+ @NewLine + 'LEFT JOIN [OCPIncentives].[dbo].[C3_Time] b ON a.Time = b.Time' + @NewLine
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'UPDATE [OCPIncentives].[dbo].[C3_Data_Monthly]'  
			+ @NewLine + 'SET [Solution Areas] = CASE WHEN [Solution Areas] = ''Azure(final)'' THEN ''Azure''' 
			+ @NewLine + 'WHEN [Solution Areas] = ''BizApps(final)'' THEN ''Business Applications''' 
			+ @NewLine + 'WHEN [Solution Areas] = ''ModernWorkplace(final)'' THEN ''Modern Work & Security''' 
			+ @NewLine + 'ELSE [Solution Areas] END ' + @NewLine
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

END
-- Azure, AGI, Hosting, SPLAR investments distribution 
ELSE IF (@ProgramName = 'Azure' OR @ProgramName = 'AGI' OR @ProgramName = 'Hosting' OR @ProgramName = 'SPLAR')
BEGIN
	--Get Azure incentives data
	PRINT('--Get Azure incentives data')
	SET @IncentivesTempTable = '#Incentives' + @ProgramName
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @IncentivesTempTable +'_Temp'') IS NOT NULL' + @NewLine + 'DROP TABLE '+ @IncentivesTempTable +'_Temp';
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine
	SET @SQL = 'SELECT p.*,'+CASE WHEN @Division ='WE' THEN ' DATENAME(MONTH, p.FiscalMonth) +  '', ''+ DATENAME(YEAR, p.FiscalMonth) as FiscalMonthYearName' ELSE ' pm.PartnerOneID AS PartnerOneID_pm, pm.PartnerOneName AS PartnerOneName_pm,d.FiscalMonthYearName' END 
			+ @NewLine + 'INTO ' + @IncentivesTempTable +'_Temp'
			+ @NewLine + 'FROM '+@TablePath+@Division+@ProgramName+'Earnings p'
			+ CASE WHEN @Division ='GE' THEN @NewLine + 'LEFT JOIN OCPIncentives.dbo.GEPartnerMaster pm ON p.PartnerOneID = pm.SourceID' ELSE '' END
			+ CASE WHEN @Division ='GE' THEN @NewLine +'LEFT JOIN OCPIncentives.dbo.Date d ON cast(p.FiscalMonth AS DATE) =d.Date' ELSE '' END
			+ @NewLine + CASE WHEN @ProgramName ='Azure' THEN 'WHERE ProgramGroupName = ''Azure''' ELSE '' END

	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine
	
	IF (@Division = 'GE')
	BEGIN
		SET @SQL = @NewLine + 'UPDATE ' + @IncentivesTempTable +'_Temp'
				+ @NewLine + 'SET ' + @IncentivesTempTable +'_Temp.PartnerOneID_pm = PartnerOneID'
				+ @NewLine + ', ' + @IncentivesTempTable +'_Temp.PartnerOneName_pm = PartnerOneName'
				+ @NewLine + 'WHERE ' + @IncentivesTempTable +'_Temp.PartnerOneID_pm IS NULL' + @NewLine
		PRINT (@SQL);
		SET @FINAL_QUERY += @SQL + @NewLine
	END
	
	--Group incentives
	PRINT('--Group incentives')
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @IncentivesTempTable +''') IS NOT NULL' + @NewLine + 'DROP TABLE '+ @IncentivesTempTable;
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'SELECT PartnerOneID,'+CASE WHEN @Division ='WE' THEN ' Subregion, Subsidiary, PartnerOneSubID,' ELSE '' END +' LeverName AS Lever, ProgramName AS ProgramName, ProgramGroupName AS ProgramGroupName, PartnerOneName AS PartnerOneName, FiscalYear AS FiscalYear, FiscalMonth AS FiscalMonth, FiscalMonthYearName, SolutionArea, CustomerName, TopReportingProduct, SUM(EarningsAmtInCD) AS EarningsAmtInCD' 
			+ @NewLine + 'INTO ' + @IncentivesTempTable
			+ @NewLine + 'FROM ' + @IncentivesTempTable +'_Temp'
			+ @NewLine + 'GROUP BY PartnerOneID ,'+CASE WHEN @Division ='WE' THEN ' Subregion, Subsidiary, PartnerOneSubID,' ELSE '' END +' ProgramName, ProgramGroupName, PartnerOneName, LeverName, FiscalYear, FiscalMonthYearName, FiscalMonth, SolutionArea, CustomerName, TopReportingProduct'+ @NewLine
		
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine
	PRINT('--Get Azure Consumed Revenue')
	
	--Change Revenuedate
	--Temporary revenue table
	--Clean table
	SET @SQL = ''
	SET @SQL = 'IF OBJECT_ID(''tempdb.dbo.#r1'') IS NOT NULL' + @NewLine + 'DROP TABLE #r1' +@NewLine
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	--Retrieve Revenue data
	SET @SQL = ''
	SET @Temp = 'Customer';
	
	SET @SQL = 'SELECT r.PartnerOneID, '+CASE WHEN @Division ='WE' THEN ' r.Subregion, r.Subsidiary, r.PartnerOneSubID,' ELSE '' END+ CASE WHEN @ProgramName = 'Azure' THEN 'REPLACE(r.FiscalMonth, '' '', '', '') AS FiscalMonthYearName,' ELSE 'd.FiscalMonthYearName,' END +' r.SolutionArea, r.TopReportingProduct'
			+ @NewLine + ',CASE WHEN '+@Temp+'Subsegment IS NULL THEN NULL'
			+@NewLine + 'WHEN ' + @Temp+'Subsegment LIKE ''%Major%'' THEN ''Enterprise''' 
			+@NewLine + 'WHEN ' + @Temp+'Subsegment LIKE ''%Strategic%'' THEN ''Enterprise'''
			+@NewLine + 'WHEN ' + @Temp+'Subsegment LIKE ''%SMB%'' THEN ''SMB'''
			+@NewLine + 'WHEN ' + @Temp+'Subsegment LIKE ''%SM&C%'' AND ' + @Temp+'Subsegment NOT LIKE ''%SMB%'' THEN ''Corporate'''
			+@NewLine + 'ELSE ''Check'' END AS SegmentGroup'
			+@NewLine + ', r.Total, r.TotalType'
			+@NewLine + 'INTO #r1' 
			+@NewLine + 'FROM '+@TablePath+@Division+CASE WHEN @ProgramName = 'Hosting' OR @ProgramName = 'SPLAR' THEN 'HostingSPLAR' ELSE @ProgramName END+'Revenue'+' r' 
			+@NewLine + CASE WHEN @ProgramName = 'AGI' OR @ProgramName = 'Hosting' OR @ProgramName = 'SPLAR' THEN 'LEFT JOIN '+CASE WHEN @Division ='WE' THEN @TablePath ELSE 'OCPIncentives.dbo.' END +'date d' ELSE '' END
			+@NewLine + CASE WHEN @ProgramName = 'AGI' OR @ProgramName = 'Hosting' OR @ProgramName = 'SPLAR' THEN 'ON r.FiscalYear=d.[FiscalYear(YY)] AND '+CASE WHEN @ProgramName = 'AGI' OR @ProgramName = 'Hosting' OR @ProgramName = 'SPLAR' THEN 'RIGHT' ELSE 'LEFT' END+'(r.FiscalMonth,3)=LEFT(d.CalendarMonthName,3)'  ELSE '' END
			+@NewLine + CASE WHEN @ProgramName = 'Azure' THEN 'WHERE r.SummaryPricingLevel NOT IN (''suites'', ''CSP'', ''other vl programs'')' ELSE '' END
			+@NewLine + CASE WHEN @ProgramName = 'Azure' THEN 'AND r.PartnerAttachType = ''Partner of Record''' ELSE '' END+@NewLine
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine


	--Delete negative revenue, Group Segment
	SET @SQL = ''
	SET @SQL = 'DELETE FROM #r1 WHERE Total<=0' +@NewLine
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = ''
	SET @SQL = 'DELETE FROM #r1 WHERE SegmentGroup =''Check'' OR SegmentGroup IS NULL' +@NewLine
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	--Revenues/Usage table
	--Clean table
	SET @RevenueTempTable = '#Revenues' + @ProgramName
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @RevenueTempTable +'_Temp' +''') IS NOT NULL' + @NewLine + 'DROP TABLE '+@RevenueTempTable+ '_Temp' ;
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	--Retrieve Lever data
	SET @SQL = ''
	SET @SQL = 'SELECT PartnerOneID, FiscalMonthYearName,'+CASE WHEN @Division ='WE' THEN ' Subregion, Subsidiary, PartnerOneSubID, ' ELSE '' END+' SolutionArea, TopReportingProduct, SegmentGroup, SUM(Total) AS GroupedTotal, TotalType'
				+ @NewLine + 'INTO ' + @RevenueTempTable +'_Temp' 
				+ @NewLine +'FROM #r1'
				+ @NewLine +'GROUP BY'
				+ @NewLine +'PartnerOneID, '+CASE WHEN @Division ='WE' THEN ' Subregion, Subsidiary, PartnerOneSubID, ' ELSE '' END+' FiscalMonthYearName, SolutionArea, TopReportingProduct, SegmentGroup, TotalType'

	SET @SQL += @NewLine;

	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	--Compute shares of revenue per Segment group
	PRINT('--Compute the shares for each PartnerID')
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.' + @RevenueTempTable +''') IS NOT NULL' + @NewLine + 'DROP TABLE '+@RevenueTempTable;
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	SET @SQL = 'SELECT Prt.*,CASE WHEN Prt.GroupedTotal = 0 OR Grp.GroupedTotal = 0 THEN 0 ELSE Prt.GroupedTotal/Grp.GroupedTotal END AS SegmentShare' + 
			@NewLine + 'INTO ' + @RevenueTempTable
			+@NewLine + 'FROM ' + @RevenueTempTable +'_Temp Prt'
			+@NewLine + 'LEFT JOIN (SELECT '+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+', FiscalMonthYearName, SUM(GroupedTotal) AS GroupedTotal'
			+@NewLine + 'FROM ' + @RevenueTempTable +'_Temp'
			+@NewLine + 'GROUP BY '+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+', FiscalMonthYearName) Grp'
			+@NewLine + 'ON Grp.'+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+' = Prt.'+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+' AND Grp.FiscalMonthYearName = Prt.FiscalMonthYearName' +@NewLine
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	--Distribute incentives
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.#Incentives' + @ProgramName +'Distributed'') IS NOT NULL' + @NewLine + 'DROP TABLE #Incentives' + @ProgramName +'Distributed';
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	PRINT('--Distribute incentives')
	SET @SQL = 'SELECT DISTINCT i.PartnerOneID, '+CASE WHEN @Division ='WE' THEN ' i.Subregion, i.Subsidiary, i.PartnerOneSubID, ' ELSE '' END+'i.ProgramName, i.ProgramGroupName,i.PartnerOneName,i.Lever,i.FiscalMonthYearName,i.FiscalMonth,i.FiscalYear, i.SolutionArea, i.CustomerName, i.TopReportingProduct, r.SegmentGroup, r.TotalType,i.EarningsAmtInCD, r.SegmentShare, i.EarningsAmtInCD * r.SegmentShare AS EarningsAmtInCDDistributed'
			 + @NewLine + 'INTO ' + '#Incentives' + @ProgramName +'Distributed'
			 + @NewLine + 'FROM ' + @IncentivesTempTable + ' i'
			 + @NewLine + 'LEFT JOIN ' +  @RevenueTempTable +' r'
			 + @NewLine + 'ON i.'+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+'=r.'+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+' AND i.FiscalMonthYearName=r.FiscalMonthYearName'
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine
	-- Distribute the remaining incentives to the PartenerOneIDs without revenues
	PRINT('--- Distribute the remaining incentives to the PartenerOneIDs without revenues')
	SET @SQL = 'UPDATE ' + '#Incentives' + @ProgramName +'Distributed'
			 + @NewLine + 'SET SegmentGroup = ''Unknown'', SegmentShare = 1, EarningsAmtInCDDistributed = EarningsAmtInCD' 
			 + @NewLine + 'WHERE EarningsAmtInCDDistributed IS NULL AND '+CASE WHEN @Division ='WE' THEN 'PartnerOneSubID' ELSE 'PartnerOneID' END+' IS NOT NULL'
	PRINT (@SQL);
	
	SET @FINAL_QUERY += @SQL + @NewLine
END
-- Azure Expert and BGI
ELSE IF (@ProgramName = 'AzureExpert' or @ProgramName = 'BGI' or @ProgramName = 'Enterprise')
BEGIN
	--Distribute incentives
	SET @SQL = 'IF OBJECT_ID(''' + 'tempdb.dbo.#Incentives' + @ProgramName +'Distributed'') IS NOT NULL' + @NewLine + 'DROP TABLE #Incentives' + @ProgramName +'Distributed';
	SET @SQL += @NewLine;
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	PRINT('--Distribute incentives')
	SET @SQL = 'SELECT DISTINCT PartnerOneID, ProgramName, ProgramGroupName,PartnerOneName,LeverName AS Lever,'+CASE WHEN @ProgramName = 'Enterprise' THEN ' (SUBSTRING(FiscalMonth,1,CHARINDEX('' '',FiscalMonth)-1)) +  '', ''+ SUBSTRING(FiscalMonth,CHARINDEX('' '',FiscalMonth)+1,len(FiscalMonth))' ELSE ' DATENAME(MONTH, FiscalMonth) +  '', ''+ DATENAME(YEAR, FiscalMonth)' END+' AS FiscalMonthYearName,'+CASE WHEN @ProgramName = 'Enterprise' THEN 'TRIM(STR(MONTH(SUBSTRING(FiscalMonth,1,CHARINDEX('' '',FiscalMonth)-1) + '' 1 2020''))) + ''/'' +''1/''+SUBSTRING(FiscalMonth,CHARINDEX('' '',FiscalMonth)+1,len(FiscalMonth))  As' ELSE '' END+' FiscalMonth,FiscalYear, SolutionArea, CustomerName, TopReportingProduct, '+CASE WHEN @ProgramName ='BGI' THEN '''Unknown'' AS ' ELSE '' END+'SegmentGroup, EarningsAmtInCD, 1 As SegmentShare, EarningsAmtInCD * 1 AS EarningsAmtInCDDistributed'
			 + @NewLine + 'INTO ' + '#Incentives' + @ProgramName +'Distributed'
			 + @NewLine + 'FROM '+@TablePath+@Division+@ProgramName+'Earnings'+CASE WHEN @ProgramName = 'Enterprise' THEN 'New' ELSE '' END
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine
	-- Distribute the remaining incentives to the PartenerOneIDs without revenues
	PRINT('--- Distribute the remaining incentives to the PartenerOneIDs without revenues')
	SET @SQL = 'UPDATE ' + '#Incentives' + @ProgramName +'Distributed'
			 + @NewLine + 'SET SegmentGroup = ''Unknown'', SegmentShare = 1, EarningsAmtInCDDistributed = EarningsAmtInCD' 
			 + @NewLine + 'WHERE EarningsAmtInCDDistributed IS NULL AND PartnerOneID IS NOT NULL'
	PRINT (@SQL);
	
	SET @FINAL_QUERY += @SQL + @NewLine
	SET @FINAL_QUERY += @SQL + @NewLine
END

IF (@ProgramName = 'Enterprise' OR @ProgramName = 'OSA' OR @ProgramName = 'OSU' OR @ProgramName = 'CSP' OR @ProgramName = 'Azure' OR @ProgramName = 'AGI' OR @ProgramName = 'Hosting' OR @ProgramName = 'SPLAR' OR @ProgramName = 'AzureExpert' OR @ProgramName = 'BGI')
BEGIN
	--Backup old table if there hasn't been any update yet, drop old final table and overwrite with new data

	-- If backup table is missing create one (First time)
	SET @SQL = @NewLine + 'IF OBJECT_ID(''' +@TablePath+@Division+@ProgramName+'DistributedIncentivesBackupTable'') IS NULL'
			 + @NewLine + 'SELECT *, CAST(GETDATE() AS DATE) AS BackupDate '
			 + @NewLine + 'INTO ' + @TablePath+@Division+@ProgramName+'DistributedIncentivesBackupTable'
			 + @NewLine + 'FROM #Incentives' + @ProgramName +'Distributed' +@NewLine 
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	-- Check if data is backed up, if not back it up
	SET @SQL = 'IF ( (SELECT CAST(MAX(BackupDate) AS DATE) FROM ' +@TablePath+@Division+@ProgramName+'DistributedIncentivesBackupTable)'
			 + @NewLine + '<'
			 + @NewLine + '(SELECT CAST(GETDATE() AS DATE)))'
			 + @NewLine + 'INSERT INTO '+@TablePath+@Division+@ProgramName+'DistributedIncentivesBackupTable'
			 + @NewLine + 'SELECT *, CAST(GETDATE() AS DATE) AS BackupDate'
			 + @NewLine + 'FROM '+@TablePath+@Division+@ProgramName+'DistributedIncentives' +@NewLine 
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine

	-- Insert new data
	SET @SQL = 'IF OBJECT_ID(''' +@TablePath+@Division+@ProgramName+'DistributedIncentives'') IS NOT NULL'
			 + @NewLine +'DROP TABLE ' +@TablePath+@Division+@ProgramName+'DistributedIncentives'
			 + @NewLine +'SELECT * '
			 + @NewLine +'INTO ' +@TablePath+@Division+@ProgramName+'DistributedIncentives'
			 + @NewLine + 'FROM #Incentives' + @ProgramName +'Distributed' +@NewLine 
	PRINT (@SQL);
	SET @FINAL_QUERY += @SQL + @NewLine
END

EXECUTE sp_executesql @FINAL_QUERY
