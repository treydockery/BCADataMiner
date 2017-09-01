Imports System.IO
Imports System.Net
Imports System.Text
Imports BCADataMiner.RootObject
Imports Newtonsoft.Json


Public Class Main

    Dim strMarketPrice As String = Nothing
    Dim bitFullyLoaded As Boolean = False
    Public dtMainSymbolList As New DataTable
    Public dtWeeklySymbolList As New DataTable
    Public dtROSSymbolList As New DataTable
    Public dtIndustryPEInfo As New DataTable

    Dim bitLoggedIntoCBOE As Boolean = False

    Public Structure SECTOR_SCORES
        Dim strSectorName As String
        Dim intVeryStrong As Integer
        Dim intStrong As Integer
        Dim intNeutral As Integer
        Dim intWeak As Integer
        Dim intVeryWeak As Integer
        Dim decTotalScore As Decimal
    End Structure

    Public Structure OPTION_MONTHS
        Dim strOptionMonth As String
        Dim strOptionValue As String
        Dim intDTE As Integer
    End Structure

    Public Structure PERPETUAL_INCOME
        Dim strStockSymbol As String
        Dim strMarketPrice As String
        Dim strDiffInStrikes As String
        Dim strPutStrikeLevel As String
        Dim strPutMid As String
        Dim strPutOpenInterest As String
        Dim strPutVolume As String
        Dim strPutPercentOfMarketPrice As String
        Dim strCallStrikeLevel As String
        Dim strCallMid As String
        Dim strCallOpenInterest As String
        Dim strCallVolume As String
        Dim strCallPercentOfMarketPrice As String
        Dim dtProcessed As DateTime
        Dim dCorrectExpirationDate As Date
    End Structure

    Public strResponseString As String = String.Empty

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' NOT USED ANYMORE
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    End Sub

    Private Sub btnGetLegacyScores_Click(sender As Object, e As EventArgs) Handles btnGetLegacyScores.Click
        btnGetLegacyScores.Enabled = False

        txtActivityLog.AppendText(Now & " Started Data Mining Program..." & vbCrLf)
        txtActivityLog.AppendText(Now & " Getting Symbol List." & vbCrLf)

        dtMainSymbolList = GetSymbolList()

        If cbOnlyRunRelativeStrength.Checked = True Then
            GoTo RUN_RS
        End If

        If cbOnlyRunPerpetualIncome.Checked = True Then
            GoTo RUN_PI
        End If

        If cbRuleOfSix.Checked = True Then
            GoTo RUN_ROS
        End If

        dtIndustryPEInfo = GetIndustryPEInfo()

        If dtMainSymbolList.Rows.Count = 0 Then
            txtActivityLog.AppendText(Now & " No rows were retreived.  Exiting routine." & vbCrLf)
            Exit Sub
        End If

        txtActivityLog.AppendText(Now & " Getting Daily Data." & vbCrLf)
        Dim intCounter As Integer = 0
        For Each myDataRow In dtMainSymbolList.Rows
            Application.DoEvents()
            intCounter += 1
            Dim strSymbol As String = myDataRow("strStockSymbol")
            strMarketPrice = String.Empty

            txtActivityLog.AppendText(Now & " Getting Data for " & strSymbol & ". # " & intCounter & " of " & dtMainSymbolList.Rows.Count & vbCrLf)

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Company Revenue Items #1 and #2
            txtActivityLog.AppendText(Now & " Getting Company Revenue." & vbCrLf)
            Dim bitCompanyRevenueSucessful As Boolean = GetCompanyRevenueAndEPS(strSymbol)
            If bitCompanyRevenueSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Revenue for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Return on Equity #3
            txtActivityLog.AppendText(Now & " Getting Return on Equity." & vbCrLf)
            Dim bitROESucessful As Boolean = GetROE(strSymbol)
            If bitROESucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Return on Equity for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' All this used to be done from Yahoo.   But they changed the site.
            ' Zacks is better anyway.
            ' Analyst Recommendations, Earnings Growth
            ' Items #4, 7
            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Analyst Recommendations, Earnings Growth
            ' Items #8, and forward PE
            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #9.  Industry Earnings
            txtActivityLog.AppendText(Now & " Getting Full Company Report From Zacks." & vbCrLf)
            Dim bitFullCompanyReportFromZacksSucessful As Boolean = GetFullCompanyReportFromZacks(strSymbol)
            If bitFullCompanyReportFromZacksSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Full Company Report From Zacks for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Earnings Surprises #5
            txtActivityLog.AppendText(Now & " Getting Earnings Surprises." & vbCrLf)
            Dim bitEarningsSurprisesSucessful As Boolean = GetEarningsSurprises(strSymbol)
            If bitEarningsSurprisesSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Earnings Surprises for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Earnings Forecast #6
            txtActivityLog.AppendText(Now & " Getting Earnings Forecast." & vbCrLf)
            Dim bitEarningsForecastSucessful As Boolean = GetEarningsForecast(strSymbol)
            If bitEarningsSurprisesSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Earnings Forecast for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #11. Insider Trading
            txtActivityLog.AppendText(Now & " Getting Insider Trading" & vbCrLf)
            Dim bitInsideTrading As Boolean = GetInsideTrading(strSymbol)
            If bitInsideTrading = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Insider Trading for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #12. Weighted Alpha
            txtActivityLog.AppendText(Now & " Getting  Weighted Alpha" & vbCrLf)
            Dim bitWeightedAlpha As Boolean = GetWeightedAlpha(strSymbol)
            If bitWeightedAlpha = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting  Weighted Alpha for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Calulate the score and put into the table
            txtActivityLog.AppendText(Now & " Calculating Overall Score" & vbCrLf)
            Dim bitCalculateOverallScore As Boolean = CalculateOverallScore(strSymbol)
            If bitCalculateOverallScore = False Then
                txtActivityLog.AppendText(Now & " There was an issue Calculating Overall Score for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Calulate the target price
            'txtLegacyScoreActivityLog.AppendText(Now & " Calculating Target Price" & vbCrLf)
            'Dim bitCalculateTargetPrice As Boolean = CalculateTargetPrice(strSymbol)
            'If bitCalculateTargetPrice = False Then
            '    txtLegacyScoreActivityLog.AppendText(Now & " There was an issue Calculating Target Price for " & strSymbol & "." & vbCrLf)
            'End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Get the Next earnings date
            txtActivityLog.AppendText(Now & " Getting the Expected Earnings Date" & vbCrLf)
            Dim bitExpEarningsDate As Boolean = GetExpEarningsDate(strSymbol)
            If bitExpEarningsDate = False Then
                txtActivityLog.AppendText(Now & " There was an issue Getting the Expected Earnings Date." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Get the PE
            txtActivityLog.AppendText(Now & " Getting the PE From Zacks" & vbCrLf)
            Dim bitPEFromZacks As Boolean = GetPEFromZacks(strSymbol)
            If bitPEFromZacks = False Then
                txtActivityLog.AppendText(Now & " There was an issue Getting the PE From Zacks." & vbCrLf)
            End If

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Get the IV

            txtActivityLog.AppendText(Now & " Getting the IV From CBOE" & vbCrLf)
            Dim bitIVFromCBOE As Boolean = GetIVFromCBOE(strSymbol)
            If bitIVFromCBOE = False Then
                txtActivityLog.AppendText(Now & " There was an issue Getting the IV From CBOE." & vbCrLf)
            End If

        Next
        '++++++++++++++++++++++++++++++++++++++++++++++
        ' #Calulate the Sector Scores and put into the table
        txtActivityLog.AppendText(Now & " Calculating Overall Scores" & vbCrLf)
        Dim bitCalculateSectorScore As Boolean = CalculateSectorScores()
        If bitCalculateSectorScore = False Then
            txtActivityLog.AppendText(Now & " There was an issue Calculating Sector Scores." & vbCrLf)
        End If
        Application.DoEvents()

        '++++++++++++++++++++++++++++++++++++++++++++++
        ' Now we are going to get and calc the Relative Strength
        ' only do it once a week on monday or later
        ' (Weekly prices are posted Sunday night)

        Dim dsSPResults As DataSet = RunSP("dbo.spGetRSLastTimeRun")
        Dim dLastTimeRun As Date = Nothing
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                For Each myRow In dsSPResults.Tables(0).Rows
                    dLastTimeRun = myRow.Item("dLastTimeRun")
                Next
            End If
        End If
        Application.DoEvents()

        '****************************************************************************
        'Dim bitRunRSPercentile = True 'FOR TESTING. BE SURE AND TAKE OUT
        '****************************************************************************
        Dim bitRunRSPercentile = False
        If cbForceRS.Checked = True Then
            bitRunRSPercentile = True
        Else
            bitRunRSPercentile = DoIRunRSPercentile(dLastTimeRun)
        End If

        If bitRunRSPercentile = True Then
            intCounter = 0
RUN_RS:
            For Each myDataRow In dtMainSymbolList.Rows
                intCounter += 1
                Dim strSymbol As String = myDataRow("strStockSymbol")

                txtActivityLog.AppendText(Now & " Calculating Relative Strength for " & strSymbol & ". # " & intCounter & " of " & dtMainSymbolList.Rows.Count & vbCrLf)
                Dim bitRelativeStrengthSucessful As Boolean = CalculateRelativeStrength(strSymbol)
                If bitRelativeStrengthSucessful = False Then
                    txtActivityLog.AppendText(Now & " There was an issue getting Relative Strength for " & strSymbol & "." & vbCrLf)
                End If
                Application.DoEvents()

            Next
            '++++++++++++++++++++++++++++++++++++++++++++++
            ' And Finally the Relative Strength Percentiles
            txtActivityLog.AppendText(Now & " Getting Relative Strength Percentile." & vbCrLf)
            Dim bitRelativeStrengthPercentileSucessful As Boolean = GetRelativeStrengthPercentile()
            If bitRelativeStrengthPercentileSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Relative Strength Percentile." & vbCrLf)
            End If

            Application.DoEvents()

            txtActivityLog.AppendText(Now & " Getting Sector Relative Strength Percentile." & vbCrLf)
            Dim bitSectorRelativeStrengthPercentileSucessful As Boolean = GetSectorRelativeStrengthPercentile()
            If bitSectorRelativeStrengthPercentileSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Sector Relative Strength Percentile." & vbCrLf)
            End If
        Else
            txtActivityLog.AppendText(Now & " Not Running the Relative Strength Routines." & vbCrLf)
        End If
        Application.DoEvents()


RUN_PI:
        dtWeeklySymbolList = GetWeeklySymbolList()
        txtActivityLog.AppendText(Now & " Getting Perpetual Income Data." & vbCrLf)
        intCounter = 0
        For Each myDataRow In dtWeeklySymbolList.Rows
            Dim myExpDate As Date
            Application.DoEvents()
            intCounter += 1
            Dim strSymbol As String = myDataRow("strStockSymbol")
            txtActivityLog.AppendText(Now & " Getting Data for " & strSymbol & ". # " & intCounter & " of " & dtWeeklySymbolList.Rows.Count & vbCrLf)

            Dim strURI = "https://query1.finance.yahoo.com/v7/finance/options/" & strSymbol & "?formatted=true&crumb=ytyxjZiBVhF&lang=en-US&region=US&corsDomain=finance.yahoo.com"
            Dim strResponseString As String = GetYahooAPIData(strURI)
            If strResponseString = "False" Then
                Continue For
            End If

            'this is just to get the right contract
            Dim myOptionData As New YahooOptionChainData
            myOptionData = JsonConvert.DeserializeObject(Of YahooOptionChainData)(strResponseString)

            'walk thru the contracts and get the correct one.
            'lets get the first contract that is 4 or more days out
            Dim strCorrectExpirationDate As String = String.Empty
            Try
                For Each myContract In myOptionData.optionChain.result(0).expirationDates
                    myExpDate = FromUNIXTimeStamps(myContract.ToString)
                    If DateDiff(DateInterval.Day, Today, myExpDate) >= 4 Then
                        strCorrectExpirationDate = myContract
                        Exit For
                    End If
                Next

            Catch ex As Exception
                Continue For
            End Try

            If strCorrectExpirationDate = String.Empty Then
                Continue For
            End If

            'now get the correct contract
            strURI = strURI & "&date=" & strCorrectExpirationDate
            strResponseString = GetYahooAPIData(strURI)
            If strResponseString = "False" Then
                Continue For
            End If

            Dim myCorrectContractOptionData As New YahooOptionChainData
            myCorrectContractOptionData = JsonConvert.DeserializeObject(Of YahooOptionChainData)(strResponseString)

            Dim myPerpetualIncome As New PERPETUAL_INCOME

            '5. get price from the ATM put And call
            '	For Each option chain -> result -> 0 -> options -> 0 -> (calls Or puts)
            '		look for inTheMoney (T/F)
            '	then the first one OTM, get (ask+bid)/2
            Dim strPriorStrikeLevel As String = String.Empty
            For Each myCallOptionContract In myCorrectContractOptionData.optionChain.result(0).options(0).calls
                If myCallOptionContract.inTheMoney = False Then
                    Try
                        myPerpetualIncome.strCallStrikeLevel = myCallOptionContract.strike.fmt.ToString
                    Catch ex As Exception
                    End Try
                    Try
                        myPerpetualIncome.strCallMid = (myCallOptionContract.ask.raw + myCallOptionContract.bid.raw) / 2
                    Catch ex As Exception
                    End Try
                    Try
                        myPerpetualIncome.strCallOpenInterest = myCallOptionContract.openInterest.raw
                    Catch ex As Exception
                    End Try
                    Try
                        myPerpetualIncome.strCallVolume = myCallOptionContract.volume.raw
                    Catch ex As Exception
                    End Try
                    Try
                        myPerpetualIncome.dCorrectExpirationDate = myCallOptionContract.expiration.fmt
                    Catch ex As Exception
                    End Try
                    Exit For
                Else
                    strPriorStrikeLevel = myCallOptionContract.strike.fmt.ToString
                End If
            Next



            'gotta step backwards through this one. have to use for not each
            For intPutCounter As Integer = (myCorrectContractOptionData.optionChain.result(0).options(0).puts.Count - 1) To 0 Step -1
                Dim myPutOptionContract = myCorrectContractOptionData.optionChain.result(0).options(0).puts(intPutCounter)
                If myPutOptionContract.inTheMoney = False Then
                    Try
                        myPerpetualIncome.strPutStrikeLevel = myPutOptionContract.strike.fmt.ToString
                    Catch ex As Exception
                    End Try
                    Try
                        myPerpetualIncome.strPutMid = (myPutOptionContract.ask.raw + myPutOptionContract.bid.raw) / 2
                    Catch ex As Exception
                    End Try
                    Try
                        myPerpetualIncome.strPutOpenInterest = myPutOptionContract.openInterest.raw
                    Catch ex As Exception
                    End Try
                    Try
                        myPerpetualIncome.strPutVolume = myPutOptionContract.volume.raw
                    Catch ex As Exception
                    End Try

                    If myPerpetualIncome.dCorrectExpirationDate = #1/1/0001 12:00:00 AM# Then
                        Try
                            myPerpetualIncome.dCorrectExpirationDate = myPutOptionContract.expiration.fmt
                        Catch ex As Exception
                        End Try
                    End If

                    Exit For
                End If
            Next

            '6. Also gather the:
            '	b.stock symbol
            '   c.market price
            '	a.diff in strikes
            '   d.DT processed
            '   e.correct expiration date
            myPerpetualIncome.strStockSymbol = strSymbol
            Try
                myPerpetualIncome.strMarketPrice = myCorrectContractOptionData.optionChain.result(0).quote.regularMarketPrice.ToString
            Catch ex As Exception
            End Try
            Try
                myPerpetualIncome.strDiffInStrikes = myPerpetualIncome.strCallStrikeLevel - strPriorStrikeLevel
            Catch ex As Exception
            End Try
            myPerpetualIncome.dtProcessed = Now
            Try
                myPerpetualIncome.dCorrectExpirationDate = myExpDate
            Catch ex As Exception
            End Try

            '7. Determine if this Is 3/4% of the market price?

            Dim intDaysToExpiration As Integer = DateDiff(DateInterval.Day, Today, myPerpetualIncome.dCorrectExpirationDate)

            Try
                Dim decPercentOfMarketPrice As Decimal = ((((myPerpetualIncome.strCallMid) / myPerpetualIncome.strMarketPrice) / intDaysToExpiration) * 365) * 100
                myPerpetualIncome.strCallPercentOfMarketPrice = FormatNumber(decPercentOfMarketPrice, 2)
            Catch ex As Exception
            End Try

            Try
                Dim decPercentOfMarketPrice As Decimal = ((((myPerpetualIncome.strPutMid) / myPerpetualIncome.strMarketPrice) / intDaysToExpiration) * 365) * 100
                myPerpetualIncome.strPutPercentOfMarketPrice = FormatNumber(decPercentOfMarketPrice, 2)
            Catch ex As Exception
            End Try

            'Run SP merge statement for the insert
            Dim params(14) As SqlClient.SqlParameter
            params(0) = New SqlClient.SqlParameter("@strStockSymbol", SqlDbType.VarChar)
            params(0).Value = myPerpetualIncome.strStockSymbol
            params(1) = New SqlClient.SqlParameter("@strMarketPrice", SqlDbType.VarChar)
            params(1).Value = myPerpetualIncome.strMarketPrice
            params(2) = New SqlClient.SqlParameter("@dCorrectExpirationDate", SqlDbType.Date)
            params(2).Value = myPerpetualIncome.dCorrectExpirationDate
            params(3) = New SqlClient.SqlParameter("@strDiffInStrikes", SqlDbType.VarChar)
            params(3).Value = myPerpetualIncome.strDiffInStrikes
            params(4) = New SqlClient.SqlParameter("@strPutStrikeLevel", SqlDbType.VarChar)
            params(4).Value = myPerpetualIncome.strPutStrikeLevel
            params(5) = New SqlClient.SqlParameter("@strPutMid", SqlDbType.VarChar)
            params(5).Value = myPerpetualIncome.strPutMid
            params(6) = New SqlClient.SqlParameter("@strPutOpenInterest", SqlDbType.VarChar)
            params(6).Value = myPerpetualIncome.strPutOpenInterest
            params(7) = New SqlClient.SqlParameter("@strPutVolume", SqlDbType.VarChar)
            params(7).Value = myPerpetualIncome.strPutVolume
            params(8) = New SqlClient.SqlParameter("@strPutPercentOfMarketPrice", SqlDbType.VarChar)
            params(8).Value = myPerpetualIncome.strPutPercentOfMarketPrice
            params(9) = New SqlClient.SqlParameter("@strCallStrikeLevel", SqlDbType.VarChar)
            params(9).Value = myPerpetualIncome.strCallStrikeLevel
            params(10) = New SqlClient.SqlParameter("@strCallMid", SqlDbType.VarChar)
            params(10).Value = myPerpetualIncome.strCallMid
            params(11) = New SqlClient.SqlParameter("@strCallOpenInterest", SqlDbType.VarChar)
            params(11).Value = myPerpetualIncome.strCallOpenInterest
            params(12) = New SqlClient.SqlParameter("@strCallVolume", SqlDbType.VarChar)
            params(12).Value = myPerpetualIncome.strCallVolume
            params(13) = New SqlClient.SqlParameter("@strCallPercentOfMarketPrice", SqlDbType.VarChar)
            params(13).Value = myPerpetualIncome.strCallPercentOfMarketPrice
            params(14) = New SqlClient.SqlParameter("@dtProcessed", SqlDbType.DateTime2)
            params(14).Value = myPerpetualIncome.dtProcessed

            Dim dsSPPIResults As DataSet = RunSP("dbo.spUpdatePerpetualIncome", params)
            'Dim x = 1
        Next

        GoTo DONE

RUN_ROS:
        dtROSSymbolList = GetROSSymbolList()
        txtActivityLog.AppendText(Now & " Getting Rule of 6 Data." & vbCrLf)
        intCounter = 0
        For Each myDataRow In dtROSSymbolList.Rows
            intCounter += 1
            Dim strSymbol As String = myDataRow("strSectorSymbol")
            txtActivityLog.AppendText(Now & " Getting Data for " & strSymbol & ". # " & intCounter & " of " & dtROSSymbolList.Rows.Count & vbCrLf)

            Dim dPeriod2 As Date = Today
            Dim strPeriod2 As String = GetUNIXTimeStamps(dPeriod2)

            Dim strURI = "https://query2.finance.yahoo.com/v8/finance/chart/" & strSymbol & "?formatted=true&region=US&period1=345448800&period2=" & strPeriod2 & "&interval=1mo&events=div%7Csplit&corsDomain=finance.yahoo.com"
            Dim strResponseString As String = GetYahooAPIData(strURI)
            If strResponseString = "False" Then
                Continue For
            End If

            Try
                strResponseString = strResponseString.Replace("result", "PriceResult")
                strResponseString = strResponseString.Replace("quote", "Pricequote")
                Dim myCorrectContractOptionData = JsonConvert.DeserializeObject(Of YahooHistoricalPrices)(strResponseString)


                'this works.
                'create a dictionary with the key being the date and the other being the price
                'then look at the json and walk backwards down the prices and grab the date from the other place
                'then put them into the dictionary

                'then put in a loop and write a merge statemetn to get it all in the database.

                Dim intNumberOfMonths As Integer = 12
                If myCorrectContractOptionData.chart.Priceresult(0).indicators.Pricequote(0).close.Count - 1 > 12 Then
                    intNumberOfMonths = (myCorrectContractOptionData.chart.Priceresult(0).indicators.Pricequote(0).close.Count - 1) - 12
                Else
                    'less than 12 months
                    intNumberOfMonths = 0
                End If

                For i As Integer = myCorrectContractOptionData.chart.Priceresult(0).indicators.Pricequote(0).close.Count - 1 To intNumberOfMonths Step -1
                    'Last 12 Months  ^^^
                    'For i As Integer = myCorrectContractOptionData.chart.Priceresult(0).indicators.Pricequote(0).close.Count - 1 To 0 Step -1
                    ' All monthly prices ^^^
                    Try
                        Dim myDate As Date = FromUNIXTimeStamps(myCorrectContractOptionData.chart.Priceresult(0).timestamp(i))
                        Dim strYearMonth As String = Year(myDate).ToString & Month(myDate).ToString().PadLeft(2, "0")

                        Dim strOpen As String = "None"
                        Try
                            strOpen = FormatNumber(myCorrectContractOptionData.chart.Priceresult(0).indicators.Pricequote(0).open(i))
                        Catch ex As Exception
                        End Try

                        Dim strHigh As String = "None"
                        Try
                            strHigh = FormatNumber(myCorrectContractOptionData.chart.Priceresult(0).indicators.Pricequote(0).high(i))
                        Catch ex As Exception
                        End Try

                        Dim strLow As String = "None"
                        Try
                            strLow = FormatNumber(myCorrectContractOptionData.chart.Priceresult(0).indicators.Pricequote(0).low(i))
                        Catch ex As Exception
                        End Try

                        Dim strClose As String = "None"
                        Try
                            strClose = FormatNumber(myCorrectContractOptionData.chart.Priceresult(0).indicators.Pricequote(0).close(i))
                        Catch ex As Exception
                        End Try

                        'Run SP merge statement for the insert
                        Dim params(5) As SqlClient.SqlParameter
                        params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
                        params(0).Value = strSymbol
                        params(1) = New SqlClient.SqlParameter("@strYearMonth", SqlDbType.VarChar)
                        params(1).Value = strYearMonth
                        params(2) = New SqlClient.SqlParameter("@strOpen", SqlDbType.VarChar)
                        params(2).Value = strOpen
                        params(3) = New SqlClient.SqlParameter("@strHigh", SqlDbType.VarChar)
                        params(3).Value = strHigh
                        params(4) = New SqlClient.SqlParameter("@strLow", SqlDbType.VarChar)
                        params(4).Value = strLow
                        params(5) = New SqlClient.SqlParameter("@strClose", SqlDbType.VarChar)
                        params(5).Value = strClose

                        Dim dsSPPIResults As DataSet = RunSP("dbo.spUpdateROSMonthlyPrice", params)

                    Catch ex As Exception
                        txtActivityLog.AppendText(Now & " There was an issue getting Monthly Prices for " & strSymbol & vbCrLf)
                    End Try

                Next
            Catch ex As Exception
                txtActivityLog.AppendText(Now & " There was an issue getting Monthly Prices for " & strSymbol & vbCrLf)
            End Try
        Next


        GoTo DO_YEARLY_CALC 'not going to do quarterlys
        '########################################################################################################################
        '#
        '# START OF QUARTERLY CALCULATIONS
        '#
        '########################################################################################################################
        'ok... now for the calculations
        '        txtActivityLog.AppendText(Now & " Starting ROS calculations " & vbCrLf)
        '        Dim strCurrentQuarter = GetCurentQuarter(Today)

        '        If strCurrentQuarter <> "None" Then
        '            Dim strPreviousQuarter As String = String.Empty

        '            'get the previous quarter and that is where we will start
        '            Dim params(0) As SqlClient.SqlParameter
        '            params(0) = New SqlClient.SqlParameter("@strCurrentQuarter", SqlDbType.VarChar)
        '            params(0).Value = strCurrentQuarter

        '            Dim dsSPQuarterResults As DataSet = RunSP("dbo.spGetPreviousQtr", params)
        '            If dsSPQuarterResults.Tables.Count > 0 Then
        '                If dsSPQuarterResults.Tables(0).Rows.Count > 0 Then
        '                    For Each myRow In dsSPQuarterResults.Tables(0).Rows
        '                        Try
        '                            strPreviousQuarter = myRow.Item("strPreviousQuarter")
        '                        Catch ex As Exception
        '                            GoTo ROSError
        '                        End Try
        '                    Next
        '                End If
        '            End If
        '            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '            ' ALL SYMBOLS OF PERCENT CHANGE DATA
        '            Dim lstAllPercentChangeData As New List(Of PercentChangeSymbolList)

        '            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '            ' START OF SPY CALCULATIONS
        '            txtActivityLog.AppendText(Now & " Starting SPY Calculations for ROS." & vbCrLf)

        '            'we have to create a table that will have percent inc or dec
        '            'get all the monthly prices of the "mother" symbol
        '            'to start with, that will be spy
        '            Dim SpyMonthlyPrices As DataTable = GetMonthlyPrices("SPY", GetQuarterStartYearMonth(strCurrentQuarter))

        '            ' get the fisrt quarter of the table, which is current quarter and it returns previous
        '            '   which is the starting place.
        '            Dim strFirstQuarterName As String = GetPreviousQuarter(SpyMonthlyPrices.Rows(0).Item(0))
        '            Dim strLastQuarterName As String = GetCurentQuarterYearMonth(SpyMonthlyPrices.Rows(SpyMonthlyPrices.Rows.Count - 1).Item(0))

        '            ' we have a linked list that we will walk down until we get to the oldest quarter of this table
        '            Dim SPYPercentChangeSymbols As New PercentChangeSymbolList
        '            SPYPercentChangeSymbols.strSymbol = "SPY"

        '            Dim strQuarterName As String = String.Empty
        '            Dim dtQuarterEnd As String = String.Empty
        '            Dim strNextQuarterToLookUp As String = strFirstQuarterName

        '            Do
        '                Dim params2(0) As SqlClient.SqlParameter
        '                params2(0) = New SqlClient.SqlParameter("@strYearMonth", SqlDbType.VarChar)
        '                params2(0).Value = strNextQuarterToLookUp

        '                Dim dsSPResults2 As DataSet = RunSP("dbo.spGetQuarterInformation", params2)
        '                If dsSPResults2.Tables.Count > 0 Then
        '                    If dsSPResults2.Tables(0).Rows.Count > 0 Then
        '                        For Each myRow As DataRow In dsSPResults2.Tables(0).Rows
        '                            strQuarterName = myRow.Item("strQuarterName")
        '                            dtQuarterEnd = myRow.Item("dtQuarterEnd")
        '                            strNextQuarterToLookUp = myRow.Item("strPreviousQuarter")
        '                        Next
        '                    End If
        '                End If

        '                Dim SPYPercentDiff As New PercentChangeData
        '                SPYPercentDiff.strQuarterName = strQuarterName
        '                SPYPercentDiff.strQuarterEndDate = dtQuarterEnd
        '                SPYPercentDiff.strQuarterPreviousName = strNextQuarterToLookUp

        '                SPYPercentChangeSymbols.lstPercentChangeData.Add(SPYPercentDiff)
        '            Loop Until strLastQuarterName = strNextQuarterToLookUp

        '            'now calculate percent diffs for the mother stock
        '            For Each mySpyQuarter In SPYPercentChangeSymbols.lstPercentChangeData

        '                Dim strCurrentPrice As String = String.Empty
        '                Dim strPreviousPrice As String = String.Empty

        '                Dim myCurrentYearMonth = GetYearMonth(mySpyQuarter.strQuarterEndDate)
        '                Dim myCurrentPriceRow = SpyMonthlyPrices.Select("strYearMonth = '" & myCurrentYearMonth & "'")
        '                For Each myRow In myCurrentPriceRow
        '                    strCurrentPrice = myRow.Item("strClose")
        '                Next

        '                Dim strPreviousQuarterEndDate As String = GetQuarterEndYearMonth(mySpyQuarter.strQuarterPreviousName)
        '                Dim myPreviousPriceRow = SpyMonthlyPrices.Select("strYearMonth = '" & strPreviousQuarterEndDate & "'")
        '                For Each myRow In myPreviousPriceRow
        '                    strPreviousPrice = myRow.Item("strClose")
        '                Next

        '                Dim decPercentIncrease As Decimal = -9999
        '                Try
        '                    decPercentIncrease = FormatNumber(((CDec(strCurrentPrice) - CDec(strPreviousPrice)) / CDec(strPreviousPrice)) * 100)
        '                    mySpyQuarter.strQuarterPercentInc = decPercentIncrease.ToString
        '                Catch ex As Exception
        '                End Try
        '            Next

        '            'add to the overall list
        '            lstAllPercentChangeData.Add(SPYPercentChangeSymbols)

        '            ' END OF SPY CALCULATIONS
        '            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        '            'we can hardcode the first one
        '            Dim myChildSymbols = GetChildrenSymbols("SPY")

        '            For Each myChildSymbolRow In myChildSymbols.Rows
        '                Dim ChildPercentChangeSymbols As New PercentChangeSymbolList

        '                ChildPercentChangeSymbols.strSymbol = myChildSymbolRow("strSectorSymbol")
        '                ChildPercentChangeSymbols.strParentSymbol = "SPY"

        '                lstAllPercentChangeData.Add(ChildPercentChangeSymbols)
        '                txtActivityLog.AppendText(Now & " Adding Child Symbol " & ChildPercentChangeSymbols.strSymbol & "." & vbCrLf)

        '            Next

        '            'after all this, do for each sector
        '            '******************************************************************************************
        '            ' SECTOR CALCULATIONS
        '            For Each myChildSector In lstAllPercentChangeData
        '                If lstAllPercentChangeData.IndexOf(myChildSector) = 0 Then
        '                    Continue For
        '                End If
        '                txtActivityLog.AppendText(Now & " Starting Sector Calculation for " & myChildSector.strSymbol & "." & vbCrLf)

        '                Dim ChildMonthlyPrices As DataTable = GetMonthlyPrices(myChildSector.strSymbol, GetQuarterStartYearMonth(strCurrentQuarter))

        '                ' get the fisrt quarter of the table, which is current quarter and it returns previous
        '                '   which is the starting place.
        '                Dim strChildFirstQuarterName As String = GetPreviousQuarter(ChildMonthlyPrices.Rows(0).Item(0))
        '                Dim strChildLastQuarterName As String = GetCurentQuarterYearMonth(ChildMonthlyPrices.Rows(ChildMonthlyPrices.Rows.Count - 1).Item(0))

        '                ' we have a linked list that we will walk down until we get to the oldest quarter of this table
        '                ' Dim ChildPercentChangeSymbols As New PercentChangeSymbolList
        '                'ChildPercentChangeSymbols.strSymbol = myChildSector.strSymbol

        '                Dim strChildQuarterName As String = String.Empty
        '                Dim dtChildQuarterEnd As String = String.Empty
        '                Dim strChildNextQuarterToLookUp As String = strChildFirstQuarterName

        '                Do
        '                    Dim params2(0) As SqlClient.SqlParameter
        '                    params2(0) = New SqlClient.SqlParameter("@strYearMonth", SqlDbType.VarChar)
        '                    params2(0).Value = strChildNextQuarterToLookUp

        '                    Dim dsSPResults2 As DataSet = RunSP("dbo.spGetQuarterInformation", params2)
        '                    If dsSPResults2.Tables.Count > 0 Then
        '                        If dsSPResults2.Tables(0).Rows.Count > 0 Then
        '                            For Each myRow As DataRow In dsSPResults2.Tables(0).Rows
        '                                strChildQuarterName = myRow.Item("strQuarterName")
        '                                dtChildQuarterEnd = myRow.Item("dtQuarterEnd")
        '                                strChildNextQuarterToLookUp = myRow.Item("strPreviousQuarter")
        '                            Next
        '                        End If
        '                    End If

        '                    Dim ChildPercentDiff As New PercentChangeData
        '                    ChildPercentDiff.strQuarterName = strChildQuarterName
        '                    ChildPercentDiff.strQuarterEndDate = dtChildQuarterEnd
        '                    ChildPercentDiff.strQuarterPreviousName = strChildNextQuarterToLookUp

        '                    myChildSector.lstPercentChangeData.Add(ChildPercentDiff)
        '                Loop Until strChildLastQuarterName = strChildNextQuarterToLookUp

        '                'now calculate percent diffs for the mother stock
        '                txtActivityLog.AppendText(Now & " Starting Quarter Calculations for " & myChildSector.strSymbol & "." & vbCrLf)
        '                For Each myChildQuarter In myChildSector.lstPercentChangeData

        '                    Dim strCurrentPrice As String = String.Empty
        '                    Dim strPreviousPrice As String = String.Empty

        '                    Dim myCurrentYearMonth = GetYearMonth(myChildQuarter.strQuarterEndDate)
        '                    Dim myCurrentPriceRow = ChildMonthlyPrices.Select("strYearMonth = '" & myCurrentYearMonth & "'")
        '                    For Each myRow In myCurrentPriceRow
        '                        strCurrentPrice = myRow.Item("strClose")
        '                    Next

        '                    Dim strPreviousQuarterEndDate As String = GetQuarterEndYearMonth(myChildQuarter.strQuarterPreviousName)
        '                    Dim myPreviousPriceRow = ChildMonthlyPrices.Select("strYearMonth = '" & strPreviousQuarterEndDate & "'")
        '                    For Each myRow In myPreviousPriceRow
        '                        strPreviousPrice = myRow.Item("strClose")
        '                    Next

        '                    Dim decPercentIncrease As Decimal = -9999
        '                    Try
        '                        decPercentIncrease = FormatNumber(((CDec(strCurrentPrice) - CDec(strPreviousPrice)) / CDec(strPreviousPrice)) * 100)
        '                        myChildQuarter.strQuarterPercentInc = decPercentIncrease.ToString
        '                    Catch ex As Exception
        '                        myChildQuarter.strQuarterPercentInc = "N/A"
        '                    End Try
        '                Next
        '            Next


        '            'now is the tricky part, compare it against the SPY
        '            'maybe create a data table to search from the spy list?

        '            'get the spyTable
        '            Dim lstSpyDataTable As New List(Of PercentChangeData)
        '            lstSpyDataTable = lstAllPercentChangeData(0).lstPercentChangeData

        '            For Each mySector In lstAllPercentChangeData
        '                If lstAllPercentChangeData.IndexOf(mySector) = 0 Then
        '                    'first one is parent. skip it
        '                    Continue For
        '                End If
        '                txtActivityLog.AppendText(Now & " Starting Sector Comparison Against Parent for " & mySector.strSymbol & "." & vbCrLf)

        '                For Each myQuarter In mySector.lstPercentChangeData
        '                    'search the parent for the match
        '                    Try
        '                        Dim FirstSpyMatch As New PercentChangeData
        '                        FirstSpyMatch = lstSpyDataTable.Find(Function(p) p.strQuarterName = myQuarter.strQuarterName)

        '                        If IsNothing(FirstSpyMatch) = False Then
        '                            'got a match. else will remain false.
        '                            If CDec(myQuarter.strQuarterPercentInc) > CDec(FirstSpyMatch.strQuarterPercentInc) Then
        '                                myQuarter.bitBeatParentSymbol = True
        '                            End If
        '                        End If

        '                    Catch ex As Exception
        '                    End Try
        '                Next
        '            Next


        '            'put it into the database
        '            For Each mySector In lstAllPercentChangeData
        '                If lstAllPercentChangeData.IndexOf(mySector) = 0 Then
        '                    'first one is parent. skip it
        '                    Continue For
        '                End If
        '                txtActivityLog.AppendText(Now & " Updating Database items for Sector " & mySector.strSymbol & "." & vbCrLf)
        '                For Each myQuarter In mySector.lstPercentChangeData
        '                    'only do the last 6 quarters.
        '                    'If mySector.lstPercentChangeData.IndexOf(myQuarter) >= 6 Then
        '                    ' Exit For
        '                    ' End If
        '                    Try
        '                        Dim paramsQtr(2) As SqlClient.SqlParameter
        '                        paramsQtr(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        '                        paramsQtr(0).Value = mySector.strSymbol
        '                        paramsQtr(1) = New SqlClient.SqlParameter("@strQuarterName", SqlDbType.VarChar)
        '                        paramsQtr(1).Value = myQuarter.strQuarterName
        '                        paramsQtr(2) = New SqlClient.SqlParameter("@bitBeatParentSymbol", SqlDbType.Bit)
        '                        paramsQtr(2).Value = myQuarter.bitBeatParentSymbol

        '                        Dim dsSPResultsQtr As DataSet = RunSP("dbo.spUpdateROSQuarterlyResult", paramsQtr)
        '                    Catch ex As Exception
        '                    End Try
        '                Next
        '            Next
        '            '########################################################################################################################
        '            '#
        '            '# END QUARTERLY CALCULATIONS
        '            '#
        '            '########################################################################################################################

        '        Else
        'ROSError:
        '            txtActivityLog.AppendText(Now & " There was an issue with ROS calculations." & vbCrLf)
        '        End If
DO_YEARLY_CALC:
        '########################################################################################################################
        '#
        '# START OF YEARLY CALCULATIONS
        '#
        '########################################################################################################################

        Dim intCurrentYear As Integer = Year(Today)
        Dim intPreviousYear As Integer = Year(DateAdd(DateInterval.Year, -1, Today))

        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' ALL SYMBOLS OF PERCENT CHANGE DATA
        Dim lstAllPercentChangeDataYearly As New List(Of PercentChangeSymbolListYearly)

        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' START OF SPY CALCULATIONS
        txtActivityLog.AppendText(Now & " Starting SPY Calculations for ROS." & vbCrLf)

        'we have to create a table that will have percent inc or dec
        'get all the monthly prices of the "mother" symbol
        'to start with, that will be spy
        'Dim SpyMonthlyPricesEachYear As DataTable = GetMonthlyPrices("SPY", intCurrentYear & "0101")
        Dim SpyMonthlyPricesEachYear As DataTable = GetMonthlyPrices("SPY", intCurrentYear & "1231")

        ' get the fisrt year of the table, which is current year and it returns previous
        '   which is the starting placeSpyMonthlyPricesEachYear
        Dim strFirstYearName As String = intPreviousYear
        Dim strLastYearName As String = Strings.Left(SpyMonthlyPricesEachYear.Rows(SpyMonthlyPricesEachYear.Rows.Count - 1).Item(0), 4)

        ' we have a linked list that we will walk down until we get to the oldest year of this table
        Dim SPYPercentChangeYearly As New PercentChangeSymbolListYearly
        SPYPercentChangeYearly.strSymbol = "SPY"

        Dim strYearName As String = intCurrentYear
        'Dim dtYearEnd As String = "12/01/" & intCurrentYear
        Dim dtYearEnd As String = Today.Month.ToString & "/01/" & intCurrentYear
        Dim strNextYearToLookUp As String = intPreviousYear

        'Dim strYearName As String = intPreviousYear
        'Dim dtYearEnd As String = "12/01/" & intPreviousYear
        'Dim strNextYearToLookUp As String = intPreviousYear - 1

        Do Until strLastYearName = strYearName
            Dim SPYPercentDiff As New PercentChangeDataYearly
            SPYPercentDiff.strYearName = strYearName
            SPYPercentDiff.strYearEndDate = dtYearEnd
            SPYPercentDiff.strYearPreviousName = strNextYearToLookUp

            SPYPercentChangeYearly.lstPercentChangeData.Add(SPYPercentDiff)

            strYearName = strNextYearToLookUp
            dtYearEnd = "12/01/" & strYearName
            Try
                strNextYearToLookUp = CInt(strYearName) - 1
            Catch ex As Exception
                Exit Do
            End Try
        Loop

        'now calculate percent diffs for the mother stock
        For Each mySpyYear In SPYPercentChangeYearly.lstPercentChangeData

            Dim strCurrentPrice As String = String.Empty
            Dim strPreviousPrice As String = String.Empty

            Dim myCurrentYearMonth = GetYearMonth(mySpyYear.strYearEndDate)
            Dim myCurrentPriceRow = SpyMonthlyPricesEachYear.Select("strYearMonth = '" & myCurrentYearMonth & "'")
            For Each myRow In myCurrentPriceRow
                strCurrentPrice = myRow.Item("strClose")
            Next

            Dim strPreviousYearEndDate As String = mySpyYear.strYearPreviousName & "12"
            Dim myPreviousPriceRow = SpyMonthlyPricesEachYear.Select("strYearMonth = '" & strPreviousYearEndDate & "'")
            For Each myRow In myPreviousPriceRow
                strPreviousPrice = myRow.Item("strClose")
            Next

            Dim decPercentIncrease As Decimal = -9999
            Try
                decPercentIncrease = FormatNumber(((CDec(strCurrentPrice) - CDec(strPreviousPrice)) / CDec(strPreviousPrice)) * 100)
                mySpyYear.strYearPercentInc = decPercentIncrease.ToString
            Catch ex As Exception
            End Try
        Next

        'add to the overall list
        lstAllPercentChangeDataYearly.Add(SPYPercentChangeYearly)

        ' END OF SPY CALCULATIONS
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'we can hardcode the first one
        Dim myChildSymbols = GetChildrenSymbolsAll()

        For Each myChildSymbolRow In myChildSymbols.Rows
            Dim ChildPercentChangeSymbols As New PercentChangeSymbolListYearly

            ChildPercentChangeSymbols.strSymbol = myChildSymbolRow("strSectorSymbol")
            ChildPercentChangeSymbols.strParentSymbol = "SPY"

            lstAllPercentChangeDataYearly.Add(ChildPercentChangeSymbols)
            txtActivityLog.AppendText(Now & " Adding Child Symbol " & ChildPercentChangeSymbols.strSymbol & "." & vbCrLf)

        Next

        '******************************************************************************************
        ' SECTOR CALCULATIONS
        For Each myChildSector In lstAllPercentChangeDataYearly
            If lstAllPercentChangeDataYearly.IndexOf(myChildSector) = 0 Then
                Continue For
            End If
            txtActivityLog.AppendText(Now & " Starting Sector Calculation for " & myChildSector.strSymbol & "." & vbCrLf)

            'Dim ChildMonthlyPrices As DataTable = GetMonthlyPrices(myChildSector.strSymbol, intCurrentYear & "0101")
            Dim ChildMonthlyPrices As DataTable = GetMonthlyPrices(myChildSector.strSymbol, intCurrentYear & "1231")

            ' get the fisrt quarter of the table, which is current quarter and it returns previous
            '   which is the starting place.
            ' Dim strChildFirstYearName As String = intPreviousYear
            Dim strChildFirstYearName As String = intCurrentYear
            Dim strChildLastYearName As String = Strings.Left(ChildMonthlyPrices.Rows(ChildMonthlyPrices.Rows.Count - 1).Item(0), 4)

            ' we have a linked list that we will walk down until we get to the oldest quarter of this table
            ' Dim ChildPercentChangeSymbols As New PercentChangeSymbolList
            'ChildPercentChangeSymbols.strSymbol = myChildSector.strSymbol

            'Dim strChildYearName As String = intPreviousYear
            'Dim dtChildYearEnd As String = "12/01/" & intPreviousYear
            'Dim strChildNextYearToLookUp As String = intPreviousYear - 1

            Dim strChildYearName As String = intCurrentYear
            ' Dim dtChildYearEnd As String = "12/01/" & intCurrentYear
            Dim dtChildYearEnd As String = Today.Month.ToString & "/01/" & intCurrentYear
            Dim strChildNextYearToLookUp As String = intPreviousYear

            Do Until strChildLastYearName = strChildYearName

                Dim ChildPercentDiff As New PercentChangeDataYearly
                ChildPercentDiff.strYearName = strChildYearName
                ChildPercentDiff.strYearEndDate = dtChildYearEnd
                ChildPercentDiff.strYearPreviousName = strChildNextYearToLookUp

                myChildSector.lstPercentChangeData.Add(ChildPercentDiff)

                strChildYearName = strChildNextYearToLookUp
                dtChildYearEnd = "12/01/" & strChildYearName
                Try
                    strChildNextYearToLookUp = CInt(strChildYearName) - 1
                Catch ex As Exception
                    Exit Do
                End Try
            Loop
            'now calculate percent diffs for the child stock
            txtActivityLog.AppendText(Now & " Starting Yearly Calculations for " & myChildSector.strSymbol & "." & vbCrLf)
            For Each myChildYear In myChildSector.lstPercentChangeData
                Dim strCurrentPrice As String = String.Empty
                Dim strPreviousPrice As String = String.Empty

                Dim myCurrentYearMonth = GetYearMonth(myChildYear.strYearEndDate)
                Dim myCurrentPriceRow = ChildMonthlyPrices.Select("strYearMonth = '" & myCurrentYearMonth & "'")
                For Each myRow In myCurrentPriceRow
                    strCurrentPrice = myRow.Item("strClose")
                Next

                Dim strPreviousYearEndDate As String = myChildYear.strYearPreviousName & "12"
                Dim myPreviousPriceRow = ChildMonthlyPrices.Select("strYearMonth = '" & strPreviousYearEndDate & "'")
                For Each myRow In myPreviousPriceRow
                    strPreviousPrice = myRow.Item("strClose")
                Next

                Dim decPercentIncrease As Decimal = -9999
                Try
                    decPercentIncrease = FormatNumber(((CDec(strCurrentPrice) - CDec(strPreviousPrice)) / CDec(strPreviousPrice)) * 100)
                    myChildYear.strYearPercentInc = decPercentIncrease.ToString
                Catch ex As Exception
                    myChildYear.strYearPercentInc = "N/A"
                End Try
            Next
        Next

        'now is the tricky part, compare it against the SPY
        'maybe create a data table to search from the spy list?

        'get the spyTable
        Dim lstSpyDataTable As New List(Of PercentChangeDataYearly)
        lstSpyDataTable = lstAllPercentChangeDataYearly(0).lstPercentChangeData

        For Each mySymbol In lstAllPercentChangeDataYearly
            If lstAllPercentChangeDataYearly.IndexOf(mySymbol) = 0 Then
                'first one is parent. skip it
                Continue For
            End If
            txtActivityLog.AppendText(Now & " Starting Sector Comparison Against Parent for " & mySymbol.strSymbol & "." & vbCrLf)

            For Each mySymbolYear In mySymbol.lstPercentChangeData
                'search the parent for the match
                Try
                    Dim FirstSpyMatch As New PercentChangeDataYearly
                    FirstSpyMatch = lstSpyDataTable.Find(Function(p) p.strYearName = mySymbolYear.strYearName)

                    If IsNothing(FirstSpyMatch) = False Then
                        'got a match. else will remain false.
                        If CDec(mySymbolYear.strYearPercentInc) > CDec(FirstSpyMatch.strYearPercentInc) Then
                            mySymbolYear.bitBeatParentSymbol = True
                        End If
                    End If

                Catch ex As Exception
                End Try
            Next
        Next
        'put it into the database
        For Each mySymbol In lstAllPercentChangeDataYearly
            If lstAllPercentChangeDataYearly.IndexOf(mySymbol) = 0 Then
                'first one is parent. skip it
                Continue For
            End If
            txtActivityLog.AppendText(Now & " Updating Database items for Sector " & mySymbol.strSymbol & "." & vbCrLf)
            For Each mySymbolYear In mySymbol.lstPercentChangeData
                Try
                    Dim paramsQtr(2) As SqlClient.SqlParameter
                    paramsQtr(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
                    paramsQtr(0).Value = mySymbol.strSymbol
                    paramsQtr(1) = New SqlClient.SqlParameter("@strYearName", SqlDbType.VarChar)
                    paramsQtr(1).Value = mySymbolYear.strYearName
                    paramsQtr(2) = New SqlClient.SqlParameter("@bitBeatParentSymbol", SqlDbType.Bit)
                    paramsQtr(2).Value = mySymbolYear.bitBeatParentSymbol

                    Dim dsSPResultsQtr As DataSet = RunSP("dbo.spUpdateROSYearlyResult", paramsQtr)
                Catch ex As Exception
                End Try
            Next
        Next

        Dim x = 1
        '########################################################################################################################
        '#
        '# END OF YEARLY CALCULATIONS
        '#
        '########################################################################################################################

DONE:
        txtActivityLog.AppendText(Now & " Done Getting Scores." & vbCrLf)
        btnGetLegacyScores.Enabled = True
    End Sub

    Public Function GetCompanyRevenueAndEPS(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim stWorkString As String = String.Empty
        Dim i As Integer = 1
        Dim j As Integer = 0
        Dim intQtrNumber As Integer = 0

        Dim strCurrentFYRevenue(4) As String
        Dim strLastFYRevenue(4) As String
        Dim strCurrentFYEPS(4) As String
        Dim strLastFYEPS(4) As String

        Dim strRevenueScore As String = "FAIL"
        Dim strEPSScore As String = "FAIL"

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "%27")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "%27")

        Dim URI As String = "http://fundamentals.nasdaq.com/redpage.asp?selected=" & strURLEncodedSymbol

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0

        Dim intRetryNumber As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        If strResponseString = "" Then
            txtActivityLog.AppendText(Now & " RECEIVED AND EMPTY PAGE. RETRYING. " & strSymbol & vbCrLf)
            intRetryNumber += 1
            If intRetryNumber > 5 Then
                Return False
            End If
            GoTo RETRY
        End If
        If strResponseString.Contains(">THIS FEATURE CURRENTLY IS UNAVAILABLE FOR") Then
            txtActivityLog.AppendText(Now & " THIS FEATURE CURRENTLY IS UNAVAILABLE FOR " & strSymbol & vbCrLf)

            Return False
        End If

        ' walk down the page and put the variables into the array
        ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' REVENUE
        i = 1
        Dim strRowString As String = String.Empty
        Try

            i = InStr(i, strResponseString, "REVENUE</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYRevenue(0) = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYRevenue(0) = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i, strResponseString, "REVENUE</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYRevenue(1) = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYRevenue(1) = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i, strResponseString, "REVENUE</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYRevenue(2) = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYRevenue(2) = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i, strResponseString, "REVENUE</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYRevenue(3) = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYRevenue(3) = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i, strResponseString, "REVENUE</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYRevenue(4) = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYRevenue(4) = ExtractNumbers(strRowString)

            For intCounter As Integer = 1 To 4
                Dim strCurrentFYRevenueTemp As String = Nothing
                Dim strLastFYRevenueTemp As String = Nothing

                If strCurrentFYRevenue(intCounter) Is Nothing Then
                    If strCurrentFYRevenue(intCounter - 1) > strLastFYRevenue(intCounter - 1) Then
                        strRevenueScore = "PASS"
                    End If
                    strCurrentFYRevenueTemp = strCurrentFYRevenue(intCounter - 1)
                    strLastFYRevenueTemp = strLastFYRevenue(intCounter - 1)

                    Dim params(4) As SqlClient.SqlParameter
                    params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
                    params(0).Value = strSymbol
                    params(1) = New SqlClient.SqlParameter("@strCurrentFYRevenue", SqlDbType.VarChar)
                    params(1).Value = strCurrentFYRevenueTemp
                    params(2) = New SqlClient.SqlParameter("@strLastFYRevenue", SqlDbType.VarChar)
                    params(2).Value = strLastFYRevenueTemp
                    params(3) = New SqlClient.SqlParameter("@strRevenueScore", SqlDbType.VarChar)
                    params(3).Value = strRevenueScore
                    params(4) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
                    params(4).Value = Today.ToShortDateString

                    Dim dsSPResults As DataSet = RunSP("dbo.spUpdateRevenue", params)

                    Exit For
                ElseIf intCounter = 4 Then
                    If strCurrentFYRevenue(intCounter) > strLastFYRevenue(intCounter) Then
                        strRevenueScore = "PASS"
                    End If
                    strCurrentFYRevenueTemp = strCurrentFYRevenue(intCounter)
                    strLastFYRevenueTemp = strLastFYRevenue(intCounter)

                    Dim params(4) As SqlClient.SqlParameter
                    params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
                    params(0).Value = strSymbol
                    params(1) = New SqlClient.SqlParameter("@strCurrentFYRevenue", SqlDbType.VarChar)
                    params(1).Value = strCurrentFYRevenueTemp
                    params(2) = New SqlClient.SqlParameter("@strLastFYRevenue", SqlDbType.VarChar)
                    params(2).Value = strLastFYRevenueTemp
                    params(3) = New SqlClient.SqlParameter("@strRevenueScore", SqlDbType.VarChar)
                    params(3).Value = strRevenueScore
                    params(4) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
                    params(4).Value = Today.ToShortDateString

                    Dim dsSPResults As DataSet = RunSP("dbo.spUpdateRevenue", params)

                    Exit For
                End If

            Next
        Catch ex As Exception

        End Try
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' EPS
        i = 1
        Try

            strRowString = String.Empty

            i = InStr(i, strResponseString, "EPS</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYEPS(0) = ExtractNumbersEPS(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYEPS(0) = ExtractNumbersEPS(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i, strResponseString, "EPS</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYEPS(1) = ExtractNumbersEPS(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYEPS(1) = ExtractNumbersEPS(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i, strResponseString, "EPS</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYEPS(2) = ExtractNumbersEPS(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYEPS(2) = ExtractNumbersEPS(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i, strResponseString, "EPS</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYEPS(3) = ExtractNumbersEPS(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYEPS(3) = ExtractNumbersEPS(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i, strResponseString, "EPS</TD>")  'Start of section to grab
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentFYEPS(4) = ExtractNumbersEPS(strRowString)

            i = j + 1
            strRowString = String.Empty
            i = InStr(i + 12, strResponseString, ">") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strLastFYEPS(4) = ExtractNumbersEPS(strRowString)

            For intCounter As Integer = 1 To 4
                Dim strCurrentFYEPSTemp As String = Nothing
                Dim strLastFYEPSTemp As String = Nothing

                If strCurrentFYEPS(intCounter) Is Nothing Then
                    If strCurrentFYEPS(intCounter - 1) > strLastFYEPS(intCounter - 1) Then
                        strEPSScore = "PASS"
                    End If
                    strCurrentFYEPSTemp = strCurrentFYEPS(intCounter - 1)
                    strLastFYEPSTemp = strLastFYEPS(intCounter - 1)

                    Dim params(4) As SqlClient.SqlParameter
                    params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
                    params(0).Value = strSymbol
                    params(1) = New SqlClient.SqlParameter("@strCurrentFYEPS", SqlDbType.VarChar)
                    params(1).Value = strCurrentFYEPSTemp
                    params(2) = New SqlClient.SqlParameter("@strLastFYEPS", SqlDbType.VarChar)
                    params(2).Value = strLastFYEPSTemp
                    params(3) = New SqlClient.SqlParameter("@strEPSScore", SqlDbType.VarChar)
                    params(3).Value = strEPSScore
                    params(4) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
                    params(4).Value = Today.ToShortDateString

                    Dim dsSPResults As DataSet = RunSP("dbo.spUpdateEPS", params)

                    Exit For
                ElseIf intCounter = 4 Then
                    If strCurrentFYEPS(intCounter) > strLastFYEPS(intCounter) Then
                        strEPSScore = "PASS"
                    End If
                    strCurrentFYEPSTemp = strCurrentFYEPS(intCounter)
                    strLastFYEPSTemp = strLastFYEPS(intCounter)

                    Dim params(4) As SqlClient.SqlParameter
                    params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
                    params(0).Value = strSymbol
                    params(1) = New SqlClient.SqlParameter("@strCurrentFYEPS", SqlDbType.VarChar)
                    params(1).Value = strCurrentFYEPSTemp
                    params(2) = New SqlClient.SqlParameter("@strLastFYEPS", SqlDbType.VarChar)
                    params(2).Value = strLastFYEPSTemp
                    params(3) = New SqlClient.SqlParameter("@strEPSScore", SqlDbType.VarChar)
                    params(3).Value = strEPSScore
                    params(4) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
                    params(4).Value = Today.ToShortDateString

                    Dim dsSPResults As DataSet = RunSP("dbo.spUpdateEPS", params)

                    Exit For
                End If
            Next
        Catch ex As Exception

        End Try
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        Return bitSucessful
    End Function

    Public Function GetROE(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim stWorkString As String = String.Empty
        Dim i As Integer = 1
        Dim j As Integer = 0
        Dim intQtrNumber As Integer = 0

        Dim strCurrentROE As String = String.Empty
        Dim strPastROE1yrAgo As String = String.Empty
        Dim strPastROE2yrsAgo As String = String.Empty

        Dim strROEScore As String = "FAIL"
        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "-")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "-")

        Dim URI As String = "http://www.nasdaq.com/symbol/" & strURLEncodedSymbol.ToLower & "/financials?query=ratios"
        'symbol has to be lower case
        'http://www.nasdaq.com/symbol/aa/financials?query=ratios

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        If strResponseString = String.Empty Then
            txtActivityLog.AppendText(Now & " NO DATA WAS RETURNED FOR " & strSymbol & vbCrLf)

            Return False
        End If
        If strResponseString.Contains(">THIS IS AN UNKNOWN SYMBOL") Then
            txtActivityLog.AppendText(Now & " THIS FEATURE CURRENTLY IS UNAVAILABLE FOR " & strSymbol & vbCrLf)

            Return False
        End If
        If strResponseString.Contains("THERE IS CURRENTLY NO DATA FOR THIS SYMBOL.") Then
            txtActivityLog.AppendText(Now & " THERE IS CURRENTLY NO DATA FOR " & strSymbol & vbCrLf)

            Return False
        End If

        If strResponseString.Contains("THIS SYMBOL CHANGED.") Then
            txtActivityLog.AppendText(Now & " THIS SYMBOL CHANGED " & strSymbol & vbCrLf)

            Return False
        End If


        ' walk down the page and put the variables into the array
        ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' ROE #3
        Try

            i = 1

            i = InStr(i, strResponseString, ">AFTER TAX ROE</TH>")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")

            Dim strRowString As String = Mid(strResponseString, (i + 1), (j - i - 1))
            strCurrentROE = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty

            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")
            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strPastROE1yrAgo = ExtractNumbers(strRowString)

            i = j + 1
            strRowString = String.Empty

            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")

            strRowString = Mid(strResponseString, (i + 1), (j - i - 1))
            strPastROE2yrsAgo = ExtractNumbers(strRowString)

            If IsNumeric(strCurrentROE) Then
                If strCurrentROE > strPastROE1yrAgo AndAlso strPastROE1yrAgo > strPastROE2yrsAgo Then
                    strROEScore = "PASS"
                End If

                Dim params(5) As SqlClient.SqlParameter
                params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
                params(0).Value = strSymbol
                params(1) = New SqlClient.SqlParameter("@strCurrentROE", SqlDbType.VarChar)
                params(1).Value = strCurrentROE
                params(2) = New SqlClient.SqlParameter("@strPastROE1yrAgo", SqlDbType.VarChar)
                params(2).Value = strPastROE1yrAgo
                params(3) = New SqlClient.SqlParameter("@strPastROE2yrsAgo", SqlDbType.VarChar)
                params(3).Value = strPastROE2yrsAgo
                params(4) = New SqlClient.SqlParameter("@strROEScore", SqlDbType.VarChar)
                params(4).Value = strROEScore
                params(5) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
                params(5).Value = Today.ToShortDateString

                Dim dsSPResults As DataSet = RunSP("dbo.spUpdateROE", params)
            End If
        Catch ex As Exception

        End Try
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        Return bitSucessful
    End Function

    Public Function GetEarningsSurprises(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim stWorkString As String = String.Empty
        Dim i As Integer = 1
        Dim j As Integer = 0
        Dim intQtrNumber As Integer = 0

        Dim strCurrentEarningsSurprise As String = String.Empty
        Dim strPastEarningsSurprise1Ago As String = String.Empty
        Dim strPastEarningsSurprise2Ago As String = String.Empty
        Dim strPastEarningsSurprise3Ago As String = String.Empty

        Dim strEarningsSurpriseScore As String = "FAIL"

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "-")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "-")

        Dim URI As String = "http://www.nasdaq.com/symbol/" & strURLEncodedSymbol.ToLower & "/earnings-surprise"
        'symbol has to be lower case
        'http://www.nasdaq.com/symbol/aapl/earnings-surprise

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        If strResponseString = String.Empty Then
            txtActivityLog.AppendText(Now & " NO DATA WAS RETURNED FOR " & strSymbol & vbCrLf)

            Return False
        End If
        If strResponseString.Contains("INSUFFICIENT INFORMATION TO DISPLAY THE GRAPH FOR THIS SYMBOL.") Then
            txtActivityLog.AppendText(Now & " INSUFFICIENT INFORMATION TO DISPLAY THE GRAPH FOR THIS SYMBOL. FOR " & strSymbol & vbCrLf)

            Return False
        End If
        If strResponseString.Contains("THERE IS CURRENTLY NO DATA FOR THIS SYMBOL.") Then
            txtActivityLog.AppendText(Now & " THERE IS CURRENTLY NO DATA FOR " & strSymbol & vbCrLf)

            Return False
        End If

        If strResponseString.Contains("THIS SYMBOL CHANGED.") Then
            txtActivityLog.AppendText(Now & " THIS SYMBOL CHANGED " & strSymbol & vbCrLf)

            Return False
        End If


        ' walk down the page and put the variables into the array
        ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' EarningsSurprise current
        Try

            i = 1

            i = InStr(i, strResponseString, "<TH>% SURPRISE</TH>")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")
            strCurrentEarningsSurprise = Mid(strResponseString, (i + 1), (j - i - 1))

            i = j + 1
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")
            strPastEarningsSurprise1Ago = Mid(strResponseString, (i + 1), (j - i - 1))

            i = j + 1
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")
            strPastEarningsSurprise2Ago = Mid(strResponseString, (i + 1), (j - i - 1))

            i = j + 1
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")
            strPastEarningsSurprise3Ago = Mid(strResponseString, (i + 1), (j - i - 1))

            If IsNumeric(strCurrentEarningsSurprise) Then
                Try
                    If CDec(strCurrentEarningsSurprise) > 0 _
                    AndAlso CDec(strPastEarningsSurprise1Ago) > 0 _
                    AndAlso CDec(strPastEarningsSurprise2Ago) > 0 _
                    AndAlso CDec(strPastEarningsSurprise3Ago) > 0 Then
                        strEarningsSurpriseScore = "PASS"
                    End If
                Catch ex As Exception
                End Try

                Dim params(6) As SqlClient.SqlParameter
                params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
                params(0).Value = strSymbol
                params(1) = New SqlClient.SqlParameter("@strCurrentEarningsSurprise", SqlDbType.VarChar)
                params(1).Value = strCurrentEarningsSurprise
                params(2) = New SqlClient.SqlParameter("@strPastEarningsSurprise1Ago", SqlDbType.VarChar)
                If IsNumeric(strPastEarningsSurprise1Ago) Then
                    params(2).Value = strPastEarningsSurprise1Ago
                Else
                    params(2).Value = DBNull.Value
                End If
                params(3) = New SqlClient.SqlParameter("@strPastEarningsSurprise2Ago", SqlDbType.VarChar)
                If IsNumeric(strPastEarningsSurprise2Ago) Then
                    params(3).Value = strPastEarningsSurprise2Ago
                Else
                    params(3).Value = DBNull.Value
                End If
                params(4) = New SqlClient.SqlParameter("@strPastEarningsSurprise3Ago", SqlDbType.VarChar)
                If IsNumeric(strPastEarningsSurprise3Ago) Then
                    params(4).Value = strPastEarningsSurprise3Ago
                Else
                    params(4).Value = DBNull.Value
                End If
                params(5) = New SqlClient.SqlParameter("@strEarningsSurpriseScore", SqlDbType.VarChar)
                params(5).Value = strEarningsSurpriseScore
                params(6) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
                params(6).Value = Today.ToShortDateString

                Dim dsSPResults As DataSet = RunSP("dbo.spUpdateEarningsSurprise", params)
            End If
        Catch ex As Exception

        End Try
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        Return bitSucessful
    End Function

    Public Function GetEarningsForecast(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim stWorkString As String = String.Empty
        Dim i As Integer = 1
        Dim j As Integer = 0
        Dim intQtrNumber As Integer = 0

        Dim strCurrentEarningsForecast As String = String.Empty
        Dim strPastEarningsForecast1Ago As String = String.Empty
        Dim strPastEarningsForecast2Ago As String = String.Empty

        Dim strEarningsForecastScore As String = "FAIL"

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "-")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "-")

        Dim URI As String = "http://www.nasdaq.com/symbol/" & strURLEncodedSymbol.ToLower & "/earnings-forecast"
        'symbol has to be lower case
        'http://www.nasdaq.com/symbol/aapl/earnings-forecast

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        If strResponseString = String.Empty Then
            txtActivityLog.AppendText(Now & " NO DATA WAS RETURNED FOR " & strSymbol & vbCrLf)

            Return False
        End If
        '
        If strResponseString.Contains("NO DATA AVAIABLE.") Then
            txtActivityLog.AppendText(Now & " THIS FEATURE CURRENTLY IS UNAVAILABLE FOR " & strSymbol & vbCrLf)

            Return False
        End If
        If strResponseString.Contains(">THIS IS AN UNKNOWN SYMBOL") Then
            txtActivityLog.AppendText(Now & " THIS FEATURE CURRENTLY IS UNAVAILABLE FOR " & strSymbol & vbCrLf)

            Return False
        End If
        If strResponseString.Contains("THERE IS CURRENTLY NO DATA FOR THIS SYMBOL.") Then
            txtActivityLog.AppendText(Now & " THERE IS CURRENTLY NO DATA FOR " & strSymbol & vbCrLf)

            Return False
        End If

        If strResponseString.Contains("THIS SYMBOL CHANGED.") Then
            txtActivityLog.AppendText(Now & " THIS SYMBOL CHANGED " & strSymbol & vbCrLf)

            Return False
        End If


        ' walk down the page and put the variables into the array
        ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' EarningsForecast current
        Try

            i = 1

            i = InStr(i, strResponseString, "<H2>YEARLY EARNINGS FORECASTS</H2>")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")
            strCurrentEarningsForecast = Mid(strResponseString, (i + 1), (j - i - 1))

            i = j + 1
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")
            strPastEarningsForecast1Ago = Mid(strResponseString, (i + 1), (j - i - 1))

            i = j + 1
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, "<TD>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")
            strPastEarningsForecast2Ago = Mid(strResponseString, (i + 1), (j - i - 1))

        Catch ex As Exception
        End Try

        If IsNumeric(strCurrentEarningsForecast) Then
            Try
                If CDec(strCurrentEarningsForecast) > 0 _
                AndAlso CDec(strPastEarningsForecast1Ago) > 0 _
                AndAlso CDec(strPastEarningsForecast2Ago) > 0 Then
                    strEarningsForecastScore = "PASS"
                End If
            Catch ex As Exception
            End Try

            Dim params(5) As SqlClient.SqlParameter
            params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            params(0).Value = strSymbol
            params(1) = New SqlClient.SqlParameter("@strCurrentEarningsForecast", SqlDbType.VarChar)
            params(1).Value = strCurrentEarningsForecast
            params(2) = New SqlClient.SqlParameter("@strPastEarningsForecast1Ago", SqlDbType.VarChar)
            If IsNumeric(strPastEarningsForecast1Ago) Then
                params(2).Value = strPastEarningsForecast1Ago
            Else
                params(2).Value = DBNull.Value
            End If
            params(3) = New SqlClient.SqlParameter("@strPastEarningsForecast2Ago", SqlDbType.VarChar)
            If IsNumeric(strPastEarningsForecast2Ago) Then
                params(3).Value = strPastEarningsForecast2Ago
            Else
                params(3).Value = DBNull.Value
            End If
            params(4) = New SqlClient.SqlParameter("@strEarningsForecastScore", SqlDbType.VarChar)
            params(4).Value = strEarningsForecastScore
            params(5) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            params(5).Value = Today.ToShortDateString

            Dim dsSPResults As DataSet = RunSP("dbo.spUpdateEarningsForecast", params)
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        Return bitSucessful
    End Function

    Public Function GetWeightedAlpha(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 1
        Dim j As Integer = 0

        Dim strWeightedAlpha As String = String.Empty
        Dim strWeightedAlphaScore As String = "FAIL"

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", ".")

        'Dim URI As String = "https://www.barchart.com/stocks/quotes/" & strURLEncodedSymbol.ToLower
        Dim URI As String = "http://old.barchart.com/quotes/stocks/" & strURLEncodedSymbol.ToLower
        'symbol has to be lower case
        'http://www.barchart.com/quotes/stocks/a

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)

                        'Throw ex
                End Select
            End If
        End Try
        If strResponseString.Contains("YOUR SEARCH CRITERIA DID NOT RETURN") Then
            txtActivityLog.AppendText(Now & " THIS FEATURE CURRENTLY IS UNAVAILABLE FOR " & strSymbol & vbCrLf)

            Return False
        End If
        ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' #12 Weighted Alpha current
        Try

            i = 1

            i = InStr(i, strResponseString, "ALPHA</TD>")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "<STRONG>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, vbTab)
            strWeightedAlpha = Mid(strResponseString, (i + 1), (j - i - 1))

        Catch ex As Exception
        End Try

        If IsNumeric(strWeightedAlpha) Then
            Try
                If CDec(strWeightedAlpha) > 0 Then
                    strWeightedAlphaScore = "PASS"
                End If
            Catch ex As Exception
            End Try

            Dim params(3) As SqlClient.SqlParameter
            params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            params(0).Value = strSymbol
            params(1) = New SqlClient.SqlParameter("@strWeightedAlpha", SqlDbType.VarChar)
            params(1).Value = strWeightedAlpha
            params(2) = New SqlClient.SqlParameter("@strWeightedAlphaScore", SqlDbType.VarChar)
            params(2).Value = strWeightedAlphaScore
            params(3) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            params(3).Value = Today.ToShortDateString

            Dim dsSPResults As DataSet = RunSP("dbo.spUpdateWeightedAlpha", params)
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        Return bitSucessful
    End Function

    Public Function GetInsideTrading(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 1
        Dim j As Integer = 0

        Dim strInsiderTrading As String = String.Empty
        Dim strInsiderTradingScore As String = "FAIL"

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "-")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "-")

        Dim URI As String = "http://www.nasdaq.com/symbol/" & strURLEncodedSymbol.ToLower & "/insider-trades"
        'symbol has to be lower case
        'http://www.nasdaq.com/symbol/aapl/insider-trades

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        If strResponseString = String.Empty Then
            txtActivityLog.AppendText(Now & " NO DATA WAS RETURNED FOR " & strSymbol & vbCrLf)

            Return False
        End If
        '
        If strResponseString.Contains(">THERE ARE NO INSIDERS FOR THIS SECURITY<") Then
            txtActivityLog.AppendText(Now & " THIS FEATURE CURRENTLY IS UNAVAILABLE FOR " & strSymbol & vbCrLf)

            Return False
        End If
        If strResponseString.Contains(">THIS IS AN UNKNOWN SYMBOL") Then
            txtActivityLog.AppendText(Now & " THIS FEATURE CURRENTLY IS UNAVAILABLE FOR " & strSymbol & vbCrLf)

            Return False
        End If
        If strResponseString.Contains("THERE IS CURRENTLY NO DATA FOR THIS SYMBOL.") Then
            txtActivityLog.AppendText(Now & " THERE IS CURRENTLY NO DATA FOR " & strSymbol & vbCrLf)

            Return False
        End If

        If strResponseString.Contains("THIS SYMBOL CHANGED.") Then
            txtActivityLog.AppendText(Now & " THIS SYMBOL CHANGED " & strSymbol & vbCrLf)

            Return False
        End If

        ' walk down the page and put the variables into the array
        ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' InsideTrading current
        Try
            i = 1

            i = InStr(i, strResponseString, "<TH>NET ACTIVITY</TH>")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "<TD")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TD")
            strInsiderTrading = Mid(strResponseString, (i + 1), (j - i - 1))
        Catch ex As Exception
        End Try

        If IsNumeric(strInsiderTrading) Then
            Try
                If CDec(strInsiderTrading) > 0 Then
                    strInsiderTradingScore = "PASS"
                End If
            Catch ex As Exception
            End Try

            Dim params(3) As SqlClient.SqlParameter
            params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            params(0).Value = strSymbol
            params(1) = New SqlClient.SqlParameter("@strInsiderTrading", SqlDbType.VarChar)
            params(1).Value = strInsiderTrading
            params(2) = New SqlClient.SqlParameter("@strInsiderTradingScore", SqlDbType.VarChar)
            params(2).Value = strInsiderTradingScore
            params(3) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            params(3).Value = Today.ToShortDateString

            Dim dsSPResults As DataSet = RunSP("dbo.spUpdateInsiderTrading", params)
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        Return bitSucessful
    End Function

    Public Function CalculateOverallScore(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True

        Dim intCompanyID As Integer = Nothing
        Dim strAnalystRecommendationScore As String = Nothing
        Dim strDaysToCoverScore As String = Nothing
        Dim strEarningsForecastScore As String = Nothing
        Dim strEarningsGrowthScore As String = Nothing
        Dim strEarningsSurpriseScore As String = Nothing
        Dim strEPSScore As String = Nothing
        Dim strIndustryEarningsScore As String = Nothing
        Dim strInsiderTradingScore As String = Nothing
        Dim strPEGRatioScore As String = Nothing
        Dim strRevenueScore As String = Nothing
        Dim strROEScore As String = Nothing
        Dim strWeightedAlphaScore As String = Nothing

        Dim intNumberOfPassingScores As Integer = 0

        Dim strOverallScore As String = Nothing

        '1. run get single score
        Dim params(0) As SqlClient.SqlParameter
        params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        params(0).Value = strSymbol

        Dim dsSPResults As DataSet = RunSP("dbo.spGetSingleScore", params)
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                For Each myRow As DataRow In dsSPResults.Tables(0).Rows
                    intCompanyID = myRow.Item("intCompanyID")

                    For Each myItem In myRow.ItemArray
                        If myItem.ToString = "PASS" Then
                            intNumberOfPassingScores += 1
                        End If
                    Next

                    strAnalystRecommendationScore = myRow.Item("strAnalystRecommendationScore").ToString
                    strDaysToCoverScore = myRow.Item("strDaysToCoverScore").ToString
                    strEarningsForecastScore = myRow.Item("strEarningsForecastScore").ToString
                    strEarningsGrowthScore = myRow.Item("strEarningsGrowthScore").ToString
                    strEarningsSurpriseScore = myRow.Item("strEarningsSurpriseScore").ToString
                    strEPSScore = myRow.Item("strEPSScore").ToString
                    strIndustryEarningsScore = myRow.Item("strIndustryEarningsScore").ToString
                    strInsiderTradingScore = myRow.Item("strInsiderTradingScore").ToString
                    strPEGRatioScore = myRow.Item("strPEGRatioScore").ToString
                    strRevenueScore = myRow.Item("strRevenueScore").ToString
                    strROEScore = myRow.Item("strROEScore").ToString
                    strWeightedAlphaScore = myRow.Item("strWeightedAlphaScore").ToString
                Next
            End If
        End If

        '2. create the score
        Dim strOverallScoreNumber As String = Nothing
        Dim decOverallScoreNumber As Decimal = Nothing
        Try

            decOverallScoreNumber = FormatNumber((intNumberOfPassingScores / 12) * 10, 2)
            If decOverallScoreNumber >= 8 Then
                strOverallScore = "VERY STRONG"
            ElseIf decOverallScoreNumber >= 7 Then
                strOverallScore = "STRONG"
            ElseIf decOverallScoreNumber >= 5 Then
                strOverallScore = "NEUTRAL"
            ElseIf decOverallScoreNumber >= 3 Then
                strOverallScore = "WEAK"
            ElseIf decOverallScoreNumber >= 0 Then
                strOverallScore = "VERY WEAK"
            Else
                strOverallScore = "N/A"
            End If
            strOverallScoreNumber = decOverallScoreNumber.ToString
        Catch ex As Exception
            strOverallScore = "N/A"
        End Try


        '3. update the table.
        Try
            Dim paramsUpdate(15) As SqlClient.SqlParameter
            paramsUpdate(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            paramsUpdate(0).Value = strSymbol
            paramsUpdate(1) = New SqlClient.SqlParameter("@strAnalystRecommendationScore", SqlDbType.VarChar)
            paramsUpdate(1).Value = strAnalystRecommendationScore
            paramsUpdate(2) = New SqlClient.SqlParameter("@strDaysToCoverScore", SqlDbType.VarChar)
            paramsUpdate(2).Value = strDaysToCoverScore
            paramsUpdate(3) = New SqlClient.SqlParameter("@strEarningsForecastScore", SqlDbType.VarChar)
            paramsUpdate(3).Value = strEarningsForecastScore
            paramsUpdate(4) = New SqlClient.SqlParameter("@strEarningsGrowthScore", SqlDbType.VarChar)
            paramsUpdate(4).Value = strEarningsGrowthScore
            paramsUpdate(5) = New SqlClient.SqlParameter("@strEarningsSurpriseScore", SqlDbType.VarChar)
            paramsUpdate(5).Value = strEarningsSurpriseScore
            paramsUpdate(6) = New SqlClient.SqlParameter("@strEPSScore", SqlDbType.VarChar)
            paramsUpdate(6).Value = strEPSScore
            paramsUpdate(7) = New SqlClient.SqlParameter("@strIndustryEarningsScore", SqlDbType.VarChar)
            paramsUpdate(7).Value = strIndustryEarningsScore
            paramsUpdate(8) = New SqlClient.SqlParameter("@strInsiderTradingScore", SqlDbType.VarChar)
            paramsUpdate(8).Value = strInsiderTradingScore
            paramsUpdate(9) = New SqlClient.SqlParameter("@strPEGRatioScore", SqlDbType.VarChar)
            paramsUpdate(9).Value = strPEGRatioScore
            paramsUpdate(10) = New SqlClient.SqlParameter("@strRevenueScore", SqlDbType.VarChar)
            paramsUpdate(10).Value = strRevenueScore
            paramsUpdate(11) = New SqlClient.SqlParameter("@strROEScore", SqlDbType.VarChar)
            paramsUpdate(11).Value = strROEScore
            paramsUpdate(12) = New SqlClient.SqlParameter("@strWeightedAlphaScore", SqlDbType.VarChar)
            paramsUpdate(12).Value = strWeightedAlphaScore
            paramsUpdate(13) = New SqlClient.SqlParameter("@strOverallScore", SqlDbType.VarChar)
            paramsUpdate(13).Value = strOverallScore
            paramsUpdate(14) = New SqlClient.SqlParameter("@strOverallScoreNumber", SqlDbType.VarChar)
            paramsUpdate(14).Value = strOverallScoreNumber
            paramsUpdate(15) = New SqlClient.SqlParameter("@dScored", SqlDbType.Date)
            paramsUpdate(15).Value = Today.ToShortDateString

            Dim dsSPResultsUpdate As DataSet = RunSP("dbo.spUpdateCompanyScore", paramsUpdate)
        Catch ex As Exception

        End Try

        Return bitSucessful
    End Function

    Public Function CalculateTargetPrice(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim strCurrentPE As String = String.Empty
        Dim strForwardPE As String = String.Empty

        Dim lstOptionMonths As New List(Of OPTION_MONTHS)

        Dim strFundamentalTarget As String = "99999"
        Dim str30DayTargetUp As String = "99999"
        Dim str30DayTargetDown As String = "0"
        Dim str60DayTargetUp As String = "99999"
        Dim str60DayTargetDown As String = "0"
        Dim str90DayTargetUp As String = "99999"
        Dim str90DayTargetDown As String = "0"

        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'get the current and forward PE from the table
        Dim params(0) As SqlClient.SqlParameter
        params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        params(0).Value = strSymbol

        Dim dsSPResults As DataSet = RunSP("dbo.spGetCurrentAndForwardPE", params)
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                For Each myRow As DataRow In dsSPResults.Tables(0).Rows
                    Try
                        strCurrentPE = myRow.Item("strCurrentPE")
                        strForwardPE = myRow.Item("strForwardPE")
                    Catch ex As Exception
                        strCurrentPE = String.Empty
                        strForwardPE = String.Empty
                    End Try
                Next
            End If
        End If

        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'fundamental target price
        'Price x ((current P/E) / (forward P/E)) = future price (or price target)
        '109.61 x (11.66/10.95) = 116.71

        If IsNumeric(strMarketPrice) AndAlso strMarketPrice <> 0 AndAlso IsNumeric(strCurrentPE) AndAlso IsNumeric(strForwardPE) Then
            Try
                strFundamentalTarget = Math.Round(CDec(strMarketPrice) * (CDec(strCurrentPE) / CDec(strForwardPE)), 2)

                If strFundamentalTarget < strMarketPrice Then
                    strFundamentalTarget = "N/A"
                End If
            Catch ex As Exception
                strFundamentalTarget = "N/A"
            End Try
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '30/60/90 expected moves
        '1. go to the page and get the list of option contracts

        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'call web browser form

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "-")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "-")

        strResponseString = String.Empty ' result will be here.
        Dim URI As String = "https://beta.finance.yahoo.com/quote/" & strURLEncodedSymbol & "/options"

        If strResponseString = "FAILED_TO_LOAD" Then
            GoTo FAILED_TO_LOAD
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '2.put the dates and values into an array for 30/60/90
        Try
            i = 1
            i = InStr(i + 1, strResponseString, "OPTION-CONTRACT-CONTROL DROP-DOWN-SELECTOR")
            If i <> 0 Then ' found the list
                Dim k As Integer = InStr(i + 1, strResponseString, "</SELECT>") 'end of the list

                Do
                    i = InStr(i + 1, strResponseString, "<OPTION ")
                    If i <> 0 AndAlso i < k Then
                        Dim myOptionContractListItem As New OPTION_MONTHS

                        i = InStr(i + 1, strResponseString, "VALUE=")
                        i += 6
                        j = InStr(i + 8, strResponseString, """")
                        myOptionContractListItem.strOptionValue = Mid(strResponseString, (i + 1), (j - i - 1))

                        i = InStr(j + 1, strResponseString, ">")
                        j = InStr(i, strResponseString, "</")
                        myOptionContractListItem.strOptionMonth = Mid(strResponseString, (i + 1), (j - i - 1))

                        Try
                            myOptionContractListItem.intDTE = DateDiff(DateInterval.Day, Today, CDate(myOptionContractListItem.strOptionMonth))
                            lstOptionMonths.Add(myOptionContractListItem)
                        Catch ex As Exception
                        End Try

                    End If
                Loop While i <> 0 AndAlso i < k
            End If
        Catch ex As Exception
        End Try

        'GET THE MARKET PRICE
        'strMarketPrice = GetMarketPriceFromYahooResponseString(strResponseString)

        '3. go to each page and get the IV for the ATM contract
        '3a. get both put and call to the closest and average them
        Dim str30DTEValue As String = String.Empty
        Dim str60DTEValue As String = String.Empty
        Dim str90DTEValue As String = String.Empty

        Dim int30DayDTE As Integer = Nothing
        Dim int60DayDTE As Integer = Nothing
        Dim int90DayDTE As Integer = Nothing

        For Each myOptionContract As OPTION_MONTHS In lstOptionMonths
            If myOptionContract.intDTE < 45 Then
                str30DTEValue = myOptionContract.strOptionValue
                int30DayDTE = myOptionContract.intDTE
            End If
            If myOptionContract.intDTE < 75 Then
                str60DTEValue = myOptionContract.strOptionValue
                int60DayDTE = myOptionContract.intDTE
            End If
            If myOptionContract.intDTE < 105 Then
                str90DTEValue = myOptionContract.strOptionValue
                int90DayDTE = myOptionContract.intDTE
            End If
        Next

        Dim str30DTETargetUpDown = GetCompanyTargetUpDown(strSymbol, str30DTEValue, int30DayDTE)
        If str30DTETargetUpDown <> "N/A" Then
            Try
                str30DayTargetUp = CDec(strMarketPrice) + CDec(str30DTETargetUpDown)
                str30DayTargetDown = CDec(strMarketPrice) - CDec(str30DTETargetUpDown)
            Catch ex As Exception
            End Try
        End If

        Dim str60DTETargetUpDown = GetCompanyTargetUpDown(strSymbol, str60DTEValue, int60DayDTE)
        If str60DTETargetUpDown <> "N/A" Then
            Try
                str60DayTargetUp = CDec(strMarketPrice) + CDec(str60DTETargetUpDown)
                str60DayTargetDown = CDec(strMarketPrice) - CDec(str60DTETargetUpDown)
            Catch ex As Exception
            End Try
        End If

        Dim str90DTETargetUpDown = GetCompanyTargetUpDown(strSymbol, str90DTEValue, int90DayDTE)
        If str90DTETargetUpDown <> "N/A" Then
            Try
                str90DayTargetUp = CDec(strMarketPrice) + CDec(str90DTETargetUpDown)
                str90DayTargetDown = CDec(strMarketPrice) - CDec(str90DTETargetUpDown)
            Catch ex As Exception
            End Try
        End If

        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'insert/update

FAILED_TO_LOAD:
        Dim paramsCompanyTargets(8) As SqlClient.SqlParameter
        paramsCompanyTargets(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        paramsCompanyTargets(0).Value = strSymbol
        paramsCompanyTargets(1) = New SqlClient.SqlParameter("@strFundamentalTarget", SqlDbType.VarChar)
        paramsCompanyTargets(1).Value = strFundamentalTarget
        paramsCompanyTargets(2) = New SqlClient.SqlParameter("@str30DayTargetUp", SqlDbType.VarChar)
        paramsCompanyTargets(2).Value = str30DayTargetUp
        paramsCompanyTargets(3) = New SqlClient.SqlParameter("@str30DayTargetDown", SqlDbType.VarChar)
        paramsCompanyTargets(3).Value = str30DayTargetDown
        paramsCompanyTargets(4) = New SqlClient.SqlParameter("@str60DayTargetUp", SqlDbType.VarChar)
        paramsCompanyTargets(4).Value = str60DayTargetUp
        paramsCompanyTargets(5) = New SqlClient.SqlParameter("@str60DayTargetDown", SqlDbType.VarChar)
        paramsCompanyTargets(5).Value = str60DayTargetDown
        paramsCompanyTargets(6) = New SqlClient.SqlParameter("@str90DayTargetUp", SqlDbType.VarChar)
        paramsCompanyTargets(6).Value = str90DayTargetUp
        paramsCompanyTargets(7) = New SqlClient.SqlParameter("@str90DayTargetDown", SqlDbType.VarChar)
        paramsCompanyTargets(7).Value = str90DayTargetDown
        paramsCompanyTargets(8) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
        paramsCompanyTargets(8).Value = Today.ToShortDateString
        Dim dsSPResultsCompanyTargets As DataSet = RunSP("dbo.spUpdateCompanyTargets", paramsCompanyTargets)

        Return bitSucessful
    End Function

    Public Function GetExpEarningsDate(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim intRetryNumber As Integer = 0

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", ".")

        Dim URI As String = "http://www.zacks.com/stock/quote/" & strURLEncodedSymbol.ToUpper

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        If strResponseString = "" Then
            txtActivityLog.AppendText(Now & " RECEIVED AND EMPTY PAGE. RETRYING. " & strSymbol & vbCrLf)
            intRetryNumber += 1
            If intRetryNumber > 5 Then
                Return False
            End If
            GoTo RETRY
        End If
        'page error
        'If responseString.Contains(">THIS FEATURE CURRENTLY IS UNAVAILABLE FOR") Then
        '    txtLegacyScoreActivityLog.AppendText(Now & " THIS FEATURE CURRENTLY IS UNAVAILABLE FOR " & strSymbol & vbCrLf)

        '    Return False
        'End If


        i = 1
        Dim strExpEarningsDate As String = String.Empty
        Try
            i = InStr(i, strResponseString, ">EXP EARNINGS DATE<")  'Start of section to grab
            If i = 0 Then
                txtActivityLog.AppendText(Now & " UNKNOWN SYMBOL. " & strSymbol & vbCrLf)
                Return False

            End If
            i = InStr(i + 12, strResponseString, "<TD>") 'Start of section to grab
            j = InStr(i, strResponseString, "</TD")

            strExpEarningsDate = Mid(strResponseString, (i + 1), (j - i - 1))
            If strExpEarningsDate.Contains("<SUP") Then
                i = InStr(i + 1, strResponseString, "</SUP>") 'Start of section to grab
                i = InStr(i + 1, strResponseString, ">") 'Start of section to grab
                strExpEarningsDate = Mid(strResponseString, (i + 1), (j - i - 1))
            End If

            If strExpEarningsDate.Contains("<A HREF") Then
                i = InStr(1, strExpEarningsDate, ">") 'Start of section to grab
                j = InStr(i + 1, strExpEarningsDate, "<") 'Start of section to grab
                strExpEarningsDate = Mid(strExpEarningsDate, (i + 1), (j - i - 1))
            End If

        Catch ex As Exception
        End Try

        Dim paramsCompanyTargets(2) As SqlClient.SqlParameter
        paramsCompanyTargets(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        paramsCompanyTargets(0).Value = strSymbol
        paramsCompanyTargets(1) = New SqlClient.SqlParameter("@strExpEarningsDate", SqlDbType.VarChar)
        paramsCompanyTargets(1).Value = strExpEarningsDate
        paramsCompanyTargets(2) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
        paramsCompanyTargets(2).Value = Today
        Dim dsSPResultsCompanyTargets As DataSet = RunSP("dbo.spUpdateExpEarningsDate", paramsCompanyTargets)

        Return bitSucessful
    End Function

    Public Function GetPEFromZacks(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim intRetryNumber As Integer = 0

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", ".")

        Dim URI As String = "http://www.zacks.com/stock/quote/" & strURLEncodedSymbol.ToUpper & "/financial-overview"

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        If strResponseString = "" Then
            txtActivityLog.AppendText(Now & " RECEIVED AND EMPTY PAGE. RETRYING. " & strSymbol & vbCrLf)
            intRetryNumber += 1
            If intRetryNumber > 5 Then
                Return False
            End If
            GoTo RETRY
        End If


        i = 1
        Dim strTrailingPE As String = String.Empty
        '">TRAILING P/E</
        i = InStr(i + 1, strResponseString, ">TRAILING 12 MONTHS</")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strTrailingPE = Mid(strResponseString, (i + 1), (j - i - 1))

        If IsNumeric(strTrailingPE) Then
            Dim paramsPERatio(2) As SqlClient.SqlParameter
            paramsPERatio(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            paramsPERatio(0).Value = strSymbol
            paramsPERatio(1) = New SqlClient.SqlParameter("@strPERatio", SqlDbType.VarChar)
            paramsPERatio(1).Value = strTrailingPE
            paramsPERatio(2) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            paramsPERatio(2).Value = Today.ToShortDateString

            Dim dsSPResultsPERatio As DataSet = RunSP("dbo.spUpdatePERatio", paramsPERatio)
        End If
        Return bitSucessful
    End Function

    Public Function GetIVFromCBOE(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim intRetryNumber As Integer = 0
        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", ".")

        If bitLoggedIntoCBOE = False Then
            'AM I Logged in
            'you only have to do this once.
            'http://www.cboe.com/tradtool/ivolservice8.aspx
            GetResponseStringFromURL("http://www.cboe.com/tradtool/ivolservice8.aspx", "</HTML>")

            bitLoggedIntoCBOE = True
        End If

RETRY:
        strResponseString = ""
        'Load the page and get the response
        Dim URI As String = "http://cboe.ivolatility.com/options.j?ticker=" & strURLEncodedSymbol.ToLower & "&R=0&x=0&y=0"
        strResponseString = GetResponseStringFromURL(URI, "IVOLATILITY.COM  ALL RIGHTS RESERVED.  IVOLATILITY <")

        If strResponseString = "" Then
            txtActivityLog.AppendText(Now & " RECEIVED AND EMPTY PAGE. RETRYING. " & strSymbol & vbCrLf)
            intRetryNumber += 1
            If intRetryNumber > 5 Then
                Return False
            End If
            GoTo RETRY
        End If

        'we are going to do this so that it works no matter if they change things.
        '1. find the data point NAME down in the code
        i = 1
        '>IV&NBSP;INDEX MEAN
        i = InStr(i + 1, strResponseString, ">IV&NBSP;INDEX MEAN")
        If i = 0 Then
            Return False 'there is no page for the symbol
        End If
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, "WR(D[")
        j = InStr(i, strResponseString, "]")
        Dim strIVXMeanDataPointName As String = Mid(strResponseString, (i + 3), (j - i - 2))

        i = j + 1
        '>52 week High
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, "WR(D[")
        j = InStr(i, strResponseString, "]")
        Dim strIVXHighDataPointName As String = Mid(strResponseString, (i + 3), (j - i - 2))

        i = j + 1
        '>52 week Low
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, "WR(D[")
        j = InStr(i, strResponseString, "]")
        Dim strIVXLowDataPointName As String = Mid(strResponseString, (i + 3), (j - i - 2))

        '2. look for the names in the array

        i = 1
        i = InStr(i + 1, strResponseString, strIVXMeanDataPointName & "=")
        i = InStr(i + 1, strResponseString, "'")
        j = InStr(i + 1, strResponseString, "'")
        Dim strIVXMeanDataPoint As String = Mid(strResponseString, (i + 1), (j - i - 1))

        i = 1
        i = InStr(i + 1, strResponseString, strIVXHighDataPointName & "=")
        i = InStr(i + 1, strResponseString, "'")
        j = InStr(i + 1, strResponseString, "'")
        Dim strIVXHighDataPoint As String = Mid(strResponseString, (i + 1), (j - i - 1))

        i = 1
        i = InStr(i + 1, strResponseString, strIVXLowDataPointName & "=")
        i = InStr(i + 1, strResponseString, "'")
        j = InStr(i + 1, strResponseString, "'")
        Dim strIVXLowDataPoint As String = Mid(strResponseString, (i + 1), (j - i - 1))

        'Calculate the IV Percentile

        'here I think we are going to subtact the 52week low from the high and the current
        'then get a percentage. that will give us a percentage within the range

        Dim strIVXResultDataPoint As String = 0

        Try
            Dim decIVXMeanDataPoint As Decimal = CDec(Replace(strIVXMeanDataPoint, "%", ""))
            Dim decIVXHighDataPoint As Decimal = CDec(Replace(strIVXHighDataPoint, "%", ""))
            Dim decIVXLowDataPoint As Decimal = CDec(Replace(strIVXLowDataPoint, "%", ""))

            Dim decCurrentIVXDataPoint = decIVXMeanDataPoint - decIVXLowDataPoint
            Dim decTempIVXHighDataPoint = decIVXHighDataPoint - decIVXLowDataPoint

            strIVXResultDataPoint = FormatNumber((decCurrentIVXDataPoint / decTempIVXHighDataPoint) * 100)

        Catch ex As Exception
            Return False
        End Try

        Dim params(5) As SqlClient.SqlParameter
        params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        params(0).Value = strSymbol
        params(1) = New SqlClient.SqlParameter("@strImpliedVolatility", SqlDbType.VarChar)
        params(1).Value = strIVXMeanDataPoint
        params(2) = New SqlClient.SqlParameter("@strImpliedVolatility52WkHigh", SqlDbType.VarChar)
        params(2).Value = strIVXHighDataPoint
        params(3) = New SqlClient.SqlParameter("@strImpliedVolatility52WkLow", SqlDbType.VarChar)
        params(3).Value = strIVXLowDataPoint
        params(4) = New SqlClient.SqlParameter("@strImpliedVolatilityScore", SqlDbType.VarChar)
        params(4).Value = strIVXResultDataPoint
        params(5) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
        params(5).Value = Today.ToShortDateString

        Dim spUpdateImpliedVolatility As DataSet = RunSP("dbo.spUpdateImpliedVolatility", params)
        Return bitSucessful
    End Function

    Public Function GetFullCompanyReportFromZacks(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim intRetryNumber As Integer = 0

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", ".")

        Dim URI As String = "http://www.zacks.com/stock/quote/" & strURLEncodedSymbol.ToUpper & "/financial-overview"

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        If strResponseString = "" Then
            txtActivityLog.AppendText(Now & " RECEIVED AND EMPTY PAGE. RETRYING. " & strSymbol & vbCrLf)
            intRetryNumber += 1
            If intRetryNumber > 5 Then
                Return False
            End If
            GoTo RETRY
        End If


        i = 1
        Dim strTrailingPE As String = String.Empty
        '">TRAILING P/E</
        i = InStr(i + 1, strResponseString, ">TRAILING 12 MONTHS</")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strTrailingPE = Mid(strResponseString, (i + 1), (j - i - 1))

        If IsNumeric(strTrailingPE) Then
            Dim paramsPERatio(2) As SqlClient.SqlParameter
            paramsPERatio(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            paramsPERatio(0).Value = strSymbol
            paramsPERatio(1) = New SqlClient.SqlParameter("@strPERatio", SqlDbType.VarChar)
            paramsPERatio(1).Value = strTrailingPE
            paramsPERatio(2) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            paramsPERatio(2).Value = Today.ToShortDateString

            Dim dsSPResultsPERatio As DataSet = RunSP("dbo.spUpdatePERatio", paramsPERatio)
        End If
        Return bitSucessful
    End Function

    Public Function CalculateSectorScores() As Boolean
        Dim bitSucessful As Boolean = True
        '1. Get the sector list
        Dim dtSectorList As DataTable = GetSectorList()

        '2. loop thru and get the companies that comprise it.
        For Each mySectorRow As DataRow In dtSectorList.Rows
            Dim intTotalNumerofStocks As Integer = 0

            Dim myCurrentSectorScores As New SECTOR_SCORES
            Dim strSector As String = mySectorRow.Item(0).ToString

            Dim dtGetCompanyOveralScores As DataTable = GetCompanyOveralScoresBySectorName(strSector)

            For Each myCompanyScoreRow As DataRow In dtGetCompanyOveralScores.Rows
                Select Case myCompanyScoreRow.ItemArray(16)
                    Case "VERY STRONG"
                        myCurrentSectorScores.intVeryStrong += 1
                    Case "STRONG"
                        myCurrentSectorScores.intStrong += 1
                    Case "NEUTRAL"
                        myCurrentSectorScores.intNeutral += 1
                    Case "WEAK"
                        myCurrentSectorScores.intWeak += 1
                    Case "VERY WEAK"
                        myCurrentSectorScores.intVeryWeak += 1
                End Select
                intTotalNumerofStocks += 1
            Next
            myCurrentSectorScores.strSectorName = strSector
            Try
                myCurrentSectorScores.decTotalScore = ((myCurrentSectorScores.intVeryStrong * 10) +
                (myCurrentSectorScores.intStrong * 8) +
                (myCurrentSectorScores.intNeutral * 6) +
                (myCurrentSectorScores.intWeak * 4) +
                (myCurrentSectorScores.intVeryWeak * 2)) / intTotalNumerofStocks

            Catch ex As Exception
                myCurrentSectorScores.decTotalScore = 0
            End Try

            Dim params(7) As SqlClient.SqlParameter
            params(0) = New SqlClient.SqlParameter("@strSector", SqlDbType.VarChar)
            params(0).Value = myCurrentSectorScores.strSectorName
            params(1) = New SqlClient.SqlParameter("@intVeryStrong", SqlDbType.VarChar)
            params(1).Value = myCurrentSectorScores.intVeryStrong
            params(2) = New SqlClient.SqlParameter("@intStrong", SqlDbType.VarChar)
            params(2).Value = myCurrentSectorScores.intStrong
            params(3) = New SqlClient.SqlParameter("@intNeutral", SqlDbType.VarChar)
            params(3).Value = myCurrentSectorScores.intNeutral
            params(4) = New SqlClient.SqlParameter("@intWeak", SqlDbType.VarChar)
            params(4).Value = myCurrentSectorScores.intWeak
            params(5) = New SqlClient.SqlParameter("@intVeryWeak", SqlDbType.VarChar)
            params(5).Value = myCurrentSectorScores.intVeryWeak
            params(6) = New SqlClient.SqlParameter("@decTotalScore", SqlDbType.Decimal)
            params(6).Value = myCurrentSectorScores.decTotalScore
            params(7) = New SqlClient.SqlParameter("@dScored", SqlDbType.Date)
            params(7).Value = Today.ToShortDateString

            Dim dsSPResults As DataSet = RunSP("dbo.spUpdateSectorScores", params)
        Next

        Return bitSucessful
    End Function

    Public Function GetCompanyTargetUpDown(ByRef strSymbol As String, ByRef strValue As String, ByRef intDTE As Integer) As String
        Dim SW2 As New Stopwatch
        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim strStrikeLevel As String = String.Empty
        Dim strStrikeLevelIV As String = String.Empty

        Dim strTargetUpDown As String = "N/A"

        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'call web browser form

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "-")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "-")

        strResponseString = String.Empty ' result will be here.
        Dim URI As String = "https://beta.finance.yahoo.com/quote/" & strURLEncodedSymbol & "/options?date=" & strValue

        If strResponseString = "FAILED_TO_LOAD" Then
            GoTo FAILED_TO_LOAD
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'get the values and do the calcs
        '1. get the market price and walk thru the table to get the IV
        Try
            Dim strStrikeLevelTemp As String
            i = 1
            i = InStr(i + 1, strResponseString, ">PUTS</")
            If i <> 0 Then ' found the list
                Do

                    i = InStr(i + 1, strResponseString, "$STRIKE.1")
                    i = InStr(i + 1, strResponseString, ">")
                    j = InStr(i + 1, strResponseString, "</")
                    strStrikeLevelTemp = Mid(strResponseString, (i + 1), (j - i - 1))
                    Try
                        If CDec(strStrikeLevelTemp) < CDec(strMarketPrice) Then
                            strStrikeLevel = strStrikeLevelTemp
                            i = InStr(j + 1, strResponseString, ".$IMPLIEDVOLATILITY")
                            i = InStr(i + 1, strResponseString, ">")
                            j = InStr(i, strResponseString, "</")
                            strStrikeLevelIV = Mid(strResponseString, (i + 1), (j - i - 1))
                            strStrikeLevelIV = strStrikeLevelIV.Replace("%", "")
                        End If

                    Catch ex As Exception

                    End Try
                Loop While CDec(strStrikeLevelTemp) < CDec(strMarketPrice)
            End If
        Catch ex As Exception
        End Try

        '2. run the calc and return the number
        Try
            Dim decIVStandardDev As Decimal = strMarketPrice * (strStrikeLevelIV / 100) * (Math.Sqrt(intDTE / 365))
            strTargetUpDown = Math.Round(decIVStandardDev, 2)
        Catch ex As Exception

        End Try
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
FAILED_TO_LOAD:
        Return strTargetUpDown
    End Function

    Public Function CalculateRelativeStrength(ByRef strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim stWorkString As String = String.Empty
        Dim i As Integer = 1
        Dim j As Integer = 0
        Dim SW2 As New Stopwatch

        'get UNIX time stamps :-)
        Dim dPeriod1 As Date = DateAdd(DateInterval.Day, -182, Today) '26 weeks ago
        Dim dPeriod2 As Date = Today
        Dim strPeriod1 As String = GetUNIXTimeStamps(dPeriod1)
        Dim strPeriod2 As String = GetUNIXTimeStamps(dPeriod2)

        Dim decSMA26 As Decimal = 0
        Dim decRelativeStrength As Decimal = 0

        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'Start webclient for google
        Dim URI As String = "https://www.google.com/finance/historical?q=" & strSymbol & "&start=0&num=200"

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        'end webclient for google
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '1. get 26 weeks of prices
        '1a. Get the TBODY and put into string
        Dim strTbody As String = String.Empty
        Try
            i = 1
            i = InStr(i + 1, strResponseString, "GF-TABLE HISTORICAL_PRICE")
            i = InStr(i + 1, strResponseString, "<TR>")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TABLE>")
            strTbody = Mid(strResponseString, (i + 1), (j - i - 1)).Trim
        Catch ex As Exception
            strTbody = "0"
        End Try
        '1b. do a split on the rows
        '1c. for each row get data if  not a dividend row

        Dim strTableRows() As String = strTbody.Split(New String() {"<TR>"}, StringSplitOptions.RemoveEmptyEntries)

        '2. from the same page get the close of last week
        '2a. the first row is this item

        Dim strLatestClosingPrice As String = String.Empty
        Dim decClosingPriceTotal As Decimal = 0
        Dim decClosingPriceTemp As Decimal = 0
        Dim intNumberOfWeeks As Integer = 0
        For intCounter As Integer = 1 To strTableRows.Length - 1
            decClosingPriceTemp = 0
            '1. find the price
            If strTableRows(intCounter).Contains(">DIVIDEND<") = False Then
                Try
                    i = 1
                    i = InStr(i + 1, strTableRows(intCounter), ">")
                    i = InStr(i + 1, strTableRows(intCounter), ">")
                    i = InStr(i + 1, strTableRows(intCounter), ">")
                    i = InStr(i + 1, strTableRows(intCounter), ">")
                    i = InStr(i + 1, strTableRows(intCounter), ">")
                    j = InStr(i, strTableRows(intCounter), "<")
                    decClosingPriceTemp = CDec(Mid(strTableRows(intCounter), (i + 1), (j - i - 1)).Trim)
                    intNumberOfWeeks += 1
                Catch ex As Exception
                End Try

                If intCounter = 1 Then
                    'latest price
                    Try
                        strLatestClosingPrice = decClosingPriceTemp.ToString
                    Catch ex As Exception
                    End Try
                End If

                'add to total
                decClosingPriceTotal += decClosingPriceTemp
            End If
        Next

        '3. divide P/SMA26
        Try
            decSMA26 = FormatNumber(decClosingPriceTotal / intNumberOfWeeks)
            decRelativeStrength = FormatNumber(CDec(strLatestClosingPrice) / decSMA26, 4)
        Catch ex As Exception
        End Try

FAILED_TO_LOAD:
        If strResponseString = "FAILED_TO_LOAD" Then
            decRelativeStrength = 0
        End If

        Dim params(2) As SqlClient.SqlParameter
        params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        params(0).Value = strSymbol
        params(1) = New SqlClient.SqlParameter("@strRelativeStrength", SqlDbType.VarChar)
        params(1).Value = decRelativeStrength.ToString
        params(2) = New SqlClient.SqlParameter("@dScored", SqlDbType.Date)
        params(2).Value = Today.ToShortDateString
        Dim dsSPResults As DataSet = RunSP("dbo.spUpdateRelativeStrength", params)

        Return bitSucessful
    End Function

    Public Function GetRelativeStrengthPercentile() As Boolean
        Dim bitSucessful As Boolean = True

        Dim dtRelativeStrengthList As DataTable = GetRelativeStrengthList()
        Dim intNumberOfRows = dtRelativeStrengthList.Rows.Count

        Dim intCounter As Integer = 0
        For Each myDataRow In dtRelativeStrengthList.Rows
            intCounter += 1
            txtActivityLog.AppendText(Now & " Getting Relative Strength for # " & intCounter & " of " & dtRelativeStrengthList.Rows.Count & vbCrLf)
            Try
                Dim intCompanyID As Integer = myDataRow("intCompanyID")
                Dim strRelativeStrength As String = myDataRow("strRelativeStrength")
                Dim strDateScored As String = myDataRow("dScored")

                Dim SameRowsResult() As DataRow = dtRelativeStrengthList.Select("strRelativeStrength = " & strRelativeStrength)
                Dim intNumberOfSameRows As Integer = 0
                Try
                    intNumberOfSameRows = SameRowsResult.Length
                Catch ex As Exception
                End Try

                Dim RowsBelowResult() As DataRow = dtRelativeStrengthList.Select("strRelativeStrength < " & strRelativeStrength)
                Dim intNumberOfRowsBelow As Integer = 0
                Try
                    intNumberOfRowsBelow = RowsBelowResult.Length
                Catch ex As Exception
                End Try

                Dim decRSPercentile As Decimal = FormatNumber((((intNumberOfRowsBelow + (0.5 * intNumberOfSameRows)) / intNumberOfRows) * 100))

                If IsNumeric(decRSPercentile) Then
                    Dim params(2) As SqlClient.SqlParameter
                    params(0) = New SqlClient.SqlParameter("@intCompanyID", SqlDbType.Int)
                    params(0).Value = intCompanyID
                    params(1) = New SqlClient.SqlParameter("@strRSPercentile", SqlDbType.VarChar)
                    params(1).Value = decRSPercentile.ToString
                    params(2) = New SqlClient.SqlParameter("@dScored", SqlDbType.Date)
                    params(2).Value = strDateScored

                    Dim dsSPResults As DataSet = RunSP("dbo.spUpdateRelativeStrengthPercentile", params)
                End If
            Catch ex As Exception
                Continue For
                'Return False
            End Try
        Next

        Return bitSucessful
    End Function

    Public Function GetSectorRelativeStrengthPercentile() As Boolean
        Dim bitSucessful As Boolean = True
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'put a loop here that goes thru all the sectors

        Dim mySectorList As DataTable = GetSectorList()

        For Each mySectorRow As DataRow In mySectorList.Rows

            Dim strSector As String = mySectorRow.Item(0).ToString

            Dim dtSectorRelativeStrengthList As DataTable = GetSectorRelativeStrengthList(strSector)
            Dim intNumberOfRows = dtSectorRelativeStrengthList.Rows.Count

            Dim intCounter As Integer = 0
            For Each myDataRow In dtSectorRelativeStrengthList.Rows
                intCounter += 1
                txtActivityLog.AppendText(Now & " Getting Sector Relative Strength for # " & intCounter & " of " & dtSectorRelativeStrengthList.Rows.Count & vbCrLf)
                Try
                    Dim intCompanyID As Integer = myDataRow("intCompanyID")
                    Dim strRelativeStrength As String = myDataRow("strRelativeStrength")
                    Dim strDateScored As String = myDataRow("dScored")

                    Dim SameRowsResult() As DataRow = dtSectorRelativeStrengthList.Select("strRelativeStrength = " & strRelativeStrength)
                    Dim intNumberOfSameRows As Integer = 0
                    Try
                        intNumberOfSameRows = SameRowsResult.Length
                    Catch ex As Exception
                    End Try

                    Dim RowsBelowResult() As DataRow = dtSectorRelativeStrengthList.Select("strRelativeStrength < " & strRelativeStrength)
                    Dim intNumberOfRowsBelow As Integer = 0
                    Try
                        intNumberOfRowsBelow = RowsBelowResult.Length
                    Catch ex As Exception
                    End Try

                    Dim decRSPercentile As Decimal = FormatNumber((((intNumberOfRowsBelow + (0.5 * intNumberOfSameRows)) / intNumberOfRows) * 100))

                    If IsNumeric(decRSPercentile) Then
                        Dim params(2) As SqlClient.SqlParameter
                        params(0) = New SqlClient.SqlParameter("@intCompanyID", SqlDbType.Int)
                        params(0).Value = intCompanyID
                        params(1) = New SqlClient.SqlParameter("@strRSPercentileSector", SqlDbType.VarChar)
                        params(1).Value = decRSPercentile.ToString
                        params(2) = New SqlClient.SqlParameter("@dScored", SqlDbType.Date)
                        params(2).Value = strDateScored

                        Dim dsSPResults As DataSet = RunSP("dbo.spUpdateSectorRelativeStrengthPercentile", params)
                    End If
                Catch ex As Exception
                    Continue For
                    'Return False
                End Try
            Next
        Next
        'end sector list
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Return bitSucessful
    End Function

#Region "Misc Functions"
    Public Function GetIndustryPEInfo() As DataTable
        Dim strWorkString As String = String.Empty
        Dim i As Integer = 1
        Dim j As Integer = 0
        Dim z As Integer = 0 'end of table

        txtActivityLog.AppendText(Now & " Getting Industry Information." & vbCrLf)

        Dim myIndustryInfo As New DataTable
        Dim IndustryNameColumn As DataColumn = New DataColumn("strIndustryName", GetType(String))
        myIndustryInfo.Columns.Add(IndustryNameColumn)
        myIndustryInfo.Columns.Add(New DataColumn("strIndustryPE", GetType(String)))

        Dim KeyColumn(0) As DataColumn
        KeyColumn(0) = IndustryNameColumn
        myIndustryInfo.PrimaryKey = KeyColumn

        'table.Rows.Add(25, "Indocin", "David", DateTime.Now) example

        Dim URI As String = "https://biz.yahoo.com/p/sum_conameu.html"

        Dim webClient As New WebClient()
        strResponseString = String.Empty

        Dim intServerRetry As Integer = 0
RETRY:
        Try
            strResponseString = webClient.DownloadString(URI).ToUpper
        Catch ex As WebException
            If TypeOf ex.Response Is HttpWebResponse Then
                Select Case DirectCast(ex.Response, HttpWebResponse).StatusCode
                    Case HttpStatusCode.NotFound
                        txtActivityLog.AppendText(Now & HttpStatusCode.NotFound & vbCrLf)

                    Case HttpStatusCode.InternalServerError
                        If intServerRetry = 0 Then
                            txtActivityLog.AppendText(Now & ex.Message & " Trying one more time." & vbCrLf)
                            intServerRetry += 1
                            Dim SW2 As New Stopwatch
                            SW2.Restart()
                            Do

                            Loop Until SW2.ElapsedMilliseconds >= 2000

                            GoTo RETRY
                        End If

                    Case Else
                        txtActivityLog.AppendText(Now & ex.Message & vbCrLf)
                        'Throw ex
                End Select
            End If
        End Try

        ' walk down the page and put the variables into the datatable
        ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' Industry Info
        Try
            i = 1

            i = InStr(i, strResponseString, ">DOWNLOAD SPREADSHEET<")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "<TR>")

            'inside the table here.  find the end
            z = InStr(i + 1, strResponseString, "</TABLE>")

            i = InStr(i + 1, strResponseString, "<TR>")
            i = InStr(i + 1, strResponseString, "<TR>")

            Do While i < z
                Dim strIndustryName As String = String.Empty
                Dim strIndustryPE As String = String.Empty

                'start of row
                i = InStr(i + 1, strResponseString, ">")
                i = InStr(i + 1, strResponseString, ">")
                i = InStr(i + 1, strResponseString, ">")
                i = InStr(i + 1, strResponseString, ">")
                j = InStr(i, strResponseString, "</")

                strWorkString = Mid(strResponseString, (i + 1), (j - i - 1))
                strIndustryName = strWorkString.Replace(vbLf, " ")

                i = j
                j = 0
                i = InStr(i + 1, strResponseString, "<TD")
                i = InStr(i + 1, strResponseString, "<TD")
                i = InStr(i + 1, strResponseString, "<TD")
                i = InStr(i + 1, strResponseString, ">")
                i = InStr(i + 1, strResponseString, ">")
                j = InStr(i, strResponseString, "</")
                strWorkString = Mid(strResponseString, (i + 1), (j - i - 1))
                strIndustryPE = strWorkString

                myIndustryInfo.Rows.Add(strIndustryName, strIndustryPE)

                i = j
                j = 0

                'find the start of the next row
                i = InStr(i + 1, strResponseString, "<TR>")

                If i = 0 Then
                    Exit Do
                End If

            Loop
        Catch ex As Exception
        End Try
        Return myIndustryInfo
    End Function

    Public Function GetSymbolList() As DataTable
        Dim mySymbolList As New DataTable
        'get the list
        Dim dsSPResults As DataSet = RunSP("dbo.spGetMasterListOfSymbols")
        'txtLegacyScoreActivityLog.AppendText(Now & " getSymbolList." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables.Count & " tables were retreived." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables(0).Rows.Count & " rows were retreived." & vbCrLf)
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                mySymbolList = dsSPResults.Tables(0)
            End If
        End If

        Return mySymbolList
    End Function

    Public Function GetWeeklySymbolList() As DataTable
        Dim mySymbolList As New DataTable
        'get the list
        Dim dsSPResults As DataSet = RunSP("dbo.spGetCompaniesWithWeeklyOptions")
        'txtLegacyScoreActivityLog.AppendText(Now & " getSymbolList." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables.Count & " tables were retreived." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables(0).Rows.Count & " rows were retreived." & vbCrLf)
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                mySymbolList = dsSPResults.Tables(0)
            End If
        End If

        Return mySymbolList
    End Function

    Public Function GetROSSymbolList() As DataTable
        Dim mySymbolList As New DataTable
        'get the list
        Dim dsSPResults As DataSet = RunSP("dbo.spGetROSMasterListOfSymbols")
        'txtLegacyScoreActivityLog.AppendText(Now & " getSymbolList." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables.Count & " tables were retreived." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables(0).Rows.Count & " rows were retreived." & vbCrLf)
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                mySymbolList = dsSPResults.Tables(0)
            End If
        End If

        Return mySymbolList
    End Function

    Public Function GetSectorList() As DataTable
        Dim mySectorList As New DataTable
        'get the list
        Dim dsSPResults As DataSet = RunSP("dbo.spGetSectorList")
        'txtLegacyScoreActivityLog.AppendText(Now & " getSymbolList." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables.Count & " tables were retreived." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables(0).Rows.Count & " rows were retreived." & vbCrLf)
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                mySectorList = dsSPResults.Tables(0)
            End If
        End If

        Return mySectorList
    End Function

    Public Function GetCompanyOveralScoresBySectorName(ByRef strSector As String) As DataTable
        Dim mySectorList As New DataTable
        'get the list

        Dim params(0) As SqlClient.SqlParameter
        params(0) = New SqlClient.SqlParameter("@strSector", SqlDbType.VarChar)
        params(0).Value = strSector
        Dim dsSPResults As DataSet = RunSP("dbo.spGetCompanyOveralScoresBySectorName", params)

        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables.Count & " tables were retreived." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables(0).Rows.Count & " rows were retreived." & vbCrLf)
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                mySectorList = dsSPResults.Tables(0)
            End If
        End If

        Return mySectorList
    End Function

    Public Function GetRelativeStrengthList() As DataTable
        Dim mySymbolList As New DataTable
        'get the list
        Dim dsSPResults As DataSet = RunSP("dbo.spGetRelativeStrengthScores")
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables.Count & " tables were retreived." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables(0).Rows.Count & " rows were retreived." & vbCrLf)
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                mySymbolList = dsSPResults.Tables(0)
            End If
        End If

        Return mySymbolList
    End Function

    Public Function GetSectorRelativeStrengthList(strSector As String) As DataTable
        Dim mySymbolList As New DataTable
        'get the list

        Dim params(0) As SqlClient.SqlParameter
        params(0) = New SqlClient.SqlParameter("@strSectorName", SqlDbType.VarChar)
        params(0).Value = strSector

        Dim dsSPResults As DataSet = RunSP("dbo.spGetSectorRelativeStrengthScores", params)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables.Count & " tables were retreived." & vbCrLf)
        txtActivityLog.AppendText(Now & " " & dsSPResults.Tables(0).Rows.Count & " rows were retreived." & vbCrLf)
        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                mySymbolList = dsSPResults.Tables(0)
            End If
        End If

        Return mySymbolList
    End Function

    Public Function ExtractNumbers(strOriginalString As String) As String
        Dim ExtractedNumbers As String = Nothing

        Dim myChars() As Char = strOriginalString.ToCharArray()
        For Each ch As Char In myChars
            If Char.IsDigit(ch) Then
                ExtractedNumbers = ExtractedNumbers & ch
            End If
        Next


        Return ExtractedNumbers
    End Function

    Public Function ExtractNumbersEPS(strOriginalString As String) As String
        Dim ExtractedNumbers As String = Nothing

        Dim myChars() As String = strOriginalString.Split("&NBSP;")

        If myChars(0) <> "" Then
            ExtractedNumbers = myChars(0)
        End If
        Return ExtractedNumbers
    End Function

    Public Function GetUNIXTimeStamps(ByRef dMyDate As Date) As String
        Dim strMyUNIXTimeStamp As String = String.Empty
        Try
            Dim uTime As Integer = (dMyDate - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds
            strMyUNIXTimeStamp = uTime.ToString
        Catch ex As Exception
        End Try
        Return strMyUNIXTimeStamp
    End Function

    Public Function FromUNIXTimeStamps(ByRef dMyUNIXDate As String) As Date
        Dim strMyStartDateTimeStamp As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        Dim strMyDateTimeStamp As New DateTime

        Try
            strMyDateTimeStamp = strMyStartDateTimeStamp.AddSeconds(dMyUNIXDate)
            strMyDateTimeStamp = strMyDateTimeStamp.ToShortDateString
        Catch ex As Exception
        End Try
        Return strMyDateTimeStamp
    End Function

    Public Function DoIRunRSPercentile(ByRef dScored As Date) As Boolean
        Dim bitRunRoutine As Boolean = False

        'get the sunday of the current week
        Dim dSundaysDate As Date = Today.AddDays(0 - Today.DayOfWeek)
        'has it been run this week?
        If dScored < dSundaysDate Then
            'This will return TRUE if it is past midnight Saturday
            'ie. Sunday 12:01 AM
            'That is ok.
            Return True
        End If
        Return bitRunRoutine
    End Function

    Public Function GetQuarterStartYearMonth(strQuarterName As String) As String
        If strQuarterName.Contains("Q1") Then
            Return Strings.Right(strQuarterName, 4) & "01"
        ElseIf strQuarterName.Contains("Q2") Then
            Return Strings.Right(strQuarterName, 4) & "04"
        ElseIf strQuarterName.Contains("Q3") Then
            Return Strings.Right(strQuarterName, 4) & "07"
        ElseIf strQuarterName.Contains("Q4") Then
            Return Strings.Right(strQuarterName, 4) & "10"
        Else
            Return "None"
        End If

    End Function

    Public Function GetQuarterEndYearMonth(strQuarterName As String) As String
        If strQuarterName.Contains("Q1") Then
            Return Strings.Right(strQuarterName, 4) & "03"
        ElseIf strQuarterName.Contains("Q2") Then
            Return Strings.Right(strQuarterName, 4) & "06"
        ElseIf strQuarterName.Contains("Q3") Then
            Return Strings.Right(strQuarterName, 4) & "09"
        ElseIf strQuarterName.Contains("Q4") Then
            Return Strings.Right(strQuarterName, 4) & "12"
        Else
            Return "None"
        End If

    End Function

    Public Function GetYearMonth(dtQuarterEnd) As String
        If IsDate(dtQuarterEnd) = False Then
            Return "None"
        End If

        Dim strYearMonth As String = Year(dtQuarterEnd).ToString & Month(dtQuarterEnd).ToString.PadLeft(2, "0")
        Return strYearMonth
    End Function

    Public Function GetCurentQuarter(dtCurrentDate As Date) As String
        If dtCurrentDate.Month <= 3 Then
            Return "Q1" & Year(Today)
        ElseIf dtCurrentDate.Month <= 6 Then
            Return "Q2" & Year(Today)
        ElseIf dtCurrentDate.Month <= 9 Then
            Return "Q3" & Year(Today)
        ElseIf dtCurrentDate.Month <= 12 Then
            Return "Q4" & Year(Today)
        Else
            Return "None"
        End If

    End Function

    Public Function GetCurentQuarterYearMonth(strYearMonth As String) As String

        If IsNumeric(strYearMonth) = False Then
            Return "None"
        End If

        Try
            Dim intYear As String = CInt(Strings.Left(strYearMonth, 4))
            Dim intMonth As String = CInt(Strings.Right(strYearMonth, 2))

            If intMonth <= 3 Then
                Return "Q1" & intYear.ToString
            ElseIf intMonth <= 6 Then
                Return "Q2" & intYear.ToString
            ElseIf intMonth <= 9 Then
                Return "Q3" & intYear.ToString
            ElseIf intMonth <= 12 Then
                Return "Q4" & (intYear - 1).ToString
            Else
                Return "None"
            End If

        Catch ex As Exception
            Return "None"
        End Try

    End Function

    Public Function GetPreviousQuarter(strYearMonth As String) As String

        If IsNumeric(strYearMonth) = False Then
            Return "None"
        End If

        Try
            Dim intYear As String = CInt(Strings.Left(strYearMonth, 4))
            Dim intMonth As String = CInt(Strings.Right(strYearMonth, 2))

            If intMonth <= 3 Then
                Return "Q4" & (intYear - 1).ToString
            ElseIf intMonth <= 6 Then
                Return "Q1" & intYear.ToString
            ElseIf intMonth <= 9 Then
                Return "Q2" & intYear.ToString
            ElseIf intMonth <= 12 Then
                Return "Q3" & intYear.ToString
            Else
                Return "None"
            End If

        Catch ex As Exception
            Return "None"
        End Try

    End Function

    Public Function GetMonthlyPrices(strSymbol As String, strYearMonth As String) As DataTable
        Dim myMonthlyPrice As New DataTable

        Dim params(1) As SqlClient.SqlParameter
        params(0) = New SqlClient.SqlParameter("@strYearMonth", SqlDbType.VarChar)
        params(0).Value = strYearMonth
        params(1) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        params(1).Value = strSymbol

        Dim dsSPResults As DataSet = RunSP("dbo.spGetMonthlyPrices", params)

        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                myMonthlyPrice = dsSPResults.Tables(0)
            End If
        End If

        Return myMonthlyPrice
    End Function

    Public Function GetChildrenSymbols(strSymbol As String) As DataTable
        Dim myChildSymbols As New DataTable

        Dim params(0) As SqlClient.SqlParameter
        params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        params(0).Value = strSymbol

        Dim dsSPResults As DataSet = RunSP("dbo.spGetChildrenSymbols", params)

        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                myChildSymbols = dsSPResults.Tables(0)
            End If
        End If

        Return myChildSymbols
    End Function

    Public Function GetChildrenSymbolsAll() As DataTable
        Dim myChildSymbols As New DataTable

        'Dim dsSPResults As DataSet = RunSP("dbo.spGetChildrenSymbolsAll")
        Dim dsSPResults As DataSet = RunSP("dbo.spGetChildrenSymbolsAllWithSectors")

        If dsSPResults.Tables.Count > 0 Then
            If dsSPResults.Tables(0).Rows.Count > 0 Then
                myChildSymbols = dsSPResults.Tables(0)
            End If
        End If

        Return myChildSymbols
    End Function

#End Region

#Region "WebPageBrowserStuff"

    'VVVVV Calling Internal Web Browser ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    Public Function GetResponseStringFromURL(strURL As String, strWhatToLookFor As String) As String
        Dim intReloadCounter As Integer = 0
        Dim intSecondaryReloadCounter As Integer = 0
        Dim SW2 As New Stopwatch
        Dim bitFailedToLoad As Boolean = False
        Dim strTempResponseString As String = String.Empty

RELOAD:
        MainWebBrowser.Navigate(strURL)

KEEP_LOADING:
        Try
            Do
                Application.DoEvents()
            Loop Until MainWebBrowser.IsBusy = False And (MainWebBrowser.ReadyState = WebBrowserReadyState.Complete OrElse MainWebBrowser.ReadyState = WebBrowserReadyState.Interactive)

            Application.DoEvents()
            If MainWebBrowser.Document.Body Is Nothing = True OrElse MainWebBrowser.DocumentText.ToUpper = String.Empty Then
                SW2.Restart()
                Do
                    Application.DoEvents()
                Loop Until SW2.ElapsedMilliseconds >= 500
                intReloadCounter += 1
                If intReloadCounter >= 21 Then
                    bitFailedToLoad = True
                    txtActivityLog.AppendText(Now & " bitFailedToLoad = True" & vbCrLf)
                ElseIf intReloadCounter = 11 Then
                    txtActivityLog.AppendText(Now & " GoTo RELOAD" & vbCrLf)
                    GoTo RELOAD
                Else
                    txtActivityLog.AppendText(Now & " GoTo KEEP_LOADING" & vbCrLf)
                    GoTo KEEP_LOADING
                End If
            End If
        Catch ex As Exception
            Return False
        End Try

        strTempResponseString = MainWebBrowser.DocumentText.ToUpper

        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'LOOK FOR AN ERROR ON THE PAGE

        If strTempResponseString.Contains("WILL BE RIGHT BACK") = True Then
            intReloadCounter += 1
            txtActivityLog.AppendText(Now & " WILL BE RIGHT BACK - GoTo RELOAD" & vbCrLf)
            GoTo RELOAD
        End If

        If strTempResponseString.Contains(">NOT FOUND</P>") = True Then
            txtActivityLog.AppendText(Now & " >NOT FOUND</P> - Return False" & vbCrLf)
            Return False
        End If

        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        'this is at the bottom of page, if it is there then you can go forward
        If strTempResponseString.Contains(strWhatToLookFor) = False Then
            txtActivityLog.AppendText(Now & " strTempResponseString.Contains(strWhatToLookFor) = False" & vbCrLf)
            intSecondaryReloadCounter += 1
            'lets give it a couple seconds here
            SW2.Restart()
            Do

            Loop Until SW2.ElapsedMilliseconds >= 500

            If intSecondaryReloadCounter = 10 Then
                MainWebBrowser.Navigate(strURL)
                txtActivityLog.AppendText(Now & " intSecondaryReloadCounter = 10 - GoTo RELOAD" & vbCrLf)
                GoTo RELOAD
            End If

            If intSecondaryReloadCounter < 20 Then
                txtActivityLog.AppendText(Now & " intSecondaryReloadCounter < 20 - GoTo KEEP_LOADING" & vbCrLf)
                GoTo KEEP_LOADING
            End If

            If intSecondaryReloadCounter >= 21 Then
                txtActivityLog.AppendText(Now & " intSecondaryReloadCounter >= 21 - bitFailedToLoad = True" & vbCrLf)
                bitFailedToLoad = True
            End If

        End If
        'should be loaded now.

        If bitFailedToLoad = True Then
            strTempResponseString = "FAILED_TO_LOAD"
            txtActivityLog.AppendText(Now & " strTempResponseString = FAILED_TO_LOAD" & vbCrLf)
        End If

        Return strTempResponseString
    End Function

#End Region

#Region "Stored Proc"
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Example of how to run stored procedure
    'Private Sub btnClickMe_Click(sender As Object, e As EventArgs) Handles btnClickMe.Click
    '    Dim strTextForBox As String = String.Empty
    '    Dim params(0) As SqlClient.SqlParameter
    '    params(0) = New SqlClient.SqlParameter("@myWatchlistID", SqlDbType.Int)
    '    params(0).Value = 11

    '    Dim dsSPResults As DataSet = RunSP("dbo.spGetWatchListItems", params)
    '    If dsSPResults.Tables.Count > 0 Then
    '        If dsSPResults.Tables(0).Rows.Count > 0 Then
    '            For Each myRow In dsSPResults.Tables(0).Rows
    '                strTextForBox = strTextForBox & myRow.Item("strSymbol") & " " & myRow.Item("strCompanyName") & vbCrLf
    '            Next
    '        End If
    '    End If
    '    txtActivityLog.Text = strTextForBox
    'End Sub
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    Public Function RunSP(ByVal p_SP As String, Optional ByVal params() As SqlClient.SqlParameter = Nothing) As DataSet
        Dim l_DataSet As DataSet = Nothing
        Dim intRetry As Integer = 0
        Dim SW2 As New Stopwatch
TRY_AGAIN:

        Try
            l_DataSet = New DataSet()
            'txtLegacyScoreActivityLog.AppendText(Now & " inside runSP." & vbCrLf)
            ' Run the stored procedure
            Dim cmd As New SqlClient.SqlCommand()
            cmd.Connection = New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("DefaultConnection").ConnectionString)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = p_SP

            If Not params Is Nothing Then
                For i As Integer = 0 To params.Count - 1
                    cmd.Parameters.Add(params(i))
                Next i
            End If
            ' Capture the returned dataset to return
            Dim l_Adapter As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter()
            l_Adapter.SelectCommand = cmd
            l_Adapter.Fill(l_DataSet)
        Catch ex As Exception
            'ok...  we are getting network related error messages.  I am going to attempt to 
            'wait a minute and try a few times.  it is usually cleared up in 5  minutes or so.

            'Also, I will need to check the status and not do this for non network related stuff
            If Err.Description.Contains("network") Then
                intRetry += 1
                If intRetry < 5 Then
                    SW2.Restart()
                    Do

                    Loop Until SW2.ElapsedMilliseconds >= 60000

                    GoTo TRY_AGAIN
                End If
            End If
            txtActivityLog.AppendText(Now & " Proc: ExecuteSQL ErrorTrap: " & Err.Number.ToString() & " " & Err.Description & vbCrLf & vbCrLf)
        End Try

        Return l_DataSet
    End Function

#End Region

#Region "Yahoo API"
    Private Function GetYahooAPIData(strURI As String) As String
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader

        Dim strResponseResults As String = Nothing

        Try
            ' Create the web request  
            request = DirectCast(WebRequest.Create(strURI), HttpWebRequest)

            ' Get response  
            response = DirectCast(request.GetResponse(), HttpWebResponse)

            ' Get the response stream into a reader  
            reader = New StreamReader(response.GetResponseStream())

            ' Console application output  
            'Console.WriteLine(reader.ReadToEnd())

            strResponseResults = reader.ReadToEnd()
        Catch ex As Exception
            strResponseResults = "False"
        Finally
            If Not response Is Nothing Then response.Close()
        End Try
        Return strResponseResults
    End Function

    Function ConvertTimestamp(ByVal timestamp As Double) As DateTime
        Return New DateTime(1970, 1, 1, 0, 0, 0).AddSeconds(timestamp).ToLocalTime()
    End Function

#End Region

End Class
