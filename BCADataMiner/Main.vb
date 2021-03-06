Imports System.IO
Imports System.Net
Imports System.Net.Security
Imports System.Text
Imports System.Collections.Specialized
Imports System.Web
Imports Newtonsoft.Json
Imports System.Security.Cryptography.X509Certificates

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
        'This is used for automatic running from Task Manager
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        Dim myArgs() = Command().Split(" ")

        If myArgs(0) = "Auto=TRUE" Then
            Me.Show()
            btnGetLegacyScores_Click(Me, EventArgs.Empty)
        End If

    End Sub

    Private Sub btnGetLegacyScores_Click(sender As Object, e As EventArgs) Handles btnGetLegacyScores.Click
        btnGetLegacyScores.Enabled = False

        txtActivityLog.AppendText(Now & " Started Data Mining Program..." & vbCrLf)
        txtActivityLog.AppendText(Now & " Getting Symbol List." & vbCrLf)

        dtMainSymbolList = GetSymbolList()

        'dtIndustryPEInfo = GetIndustryPEInfo()

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
            'Debug
            'strSymbol = "AAPL"
            strMarketPrice = String.Empty

            txtActivityLog.AppendText(Now & " Getting Data for " & strSymbol & ". # " & intCounter & " of " & dtMainSymbolList.Rows.Count & vbCrLf)

            'for debug only
            'GoTo DEBUG
            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Company Revenue Items #1 and #2
            'Working
            txtActivityLog.AppendText(Now & " Getting Company Revenue." & vbCrLf)
            Dim bitCompanyRevenueSucessful As Boolean = GetCompanyRevenueAndEPS(strSymbol)
            If bitCompanyRevenueSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Revenue for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Return on Equity #3

            ' NOT WORKING !!!! Moved to another Section

            'txtActivityLog.AppendText(Now & " Getting Return on Equity." & vbCrLf)
            'Dim bitROESucessful As Boolean = GetROE(strSymbol)
            'If bitROESucessful = False Then
            '    txtActivityLog.AppendText(Now & " There was an issue getting Return on Equity for " & strSymbol & "." & vbCrLf)
            'End If
            'Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Days to Cover. See Document
            ' Working
            txtActivityLog.AppendText(Now & " Getting Days To Cover / Short Interest." & vbCrLf)
            Dim bitDaysToCoverSucessful As Boolean = GetDaysToCover(strSymbol)
            If bitDaysToCoverSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Days To Cover / Short Interest for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()
            '
            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Analyst Recommendations
            ' Working
            txtActivityLog.AppendText(Now & " Getting Zacks Analyst Recommendations." & vbCrLf)
            Dim bitZacksAnalystRecommendations As Boolean = GetZacksAnalystRecommendationsAndIndustryComparison(strSymbol)
            If bitZacksAnalystRecommendations = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Zacks Analyst Recommendations for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' forward PE
            ' working
            txtActivityLog.AppendText(Now & " Getting Full Company Report From Zacks." & vbCrLf)
            Dim bitFullCompanyReportFromZacksSucessful As Boolean = GetFullCompanyReportFromZacks(strSymbol)
            If bitFullCompanyReportFromZacksSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Full Company Report From Zacks for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Earnings Growth
            ' Working
            txtActivityLog.AppendText(Now & " Getting Earnings Growth Report From Zacks." & vbCrLf)
            Dim bitEarningsGrowthFromZacksSucessful As Boolean = GetEarningsGrowthFromZacks(strSymbol)
            If bitEarningsGrowthFromZacksSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Earnings Growth Report From Zacks for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Earnings Surprises #5
            ' Working!!!
            txtActivityLog.AppendText(Now & " Getting Earnings Surprises." & vbCrLf)
            Dim bitEarningsSurprisesSucessful As Boolean = GetEarningsSurprises(strSymbol)
            If bitEarningsSurprisesSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Earnings Surprises for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' Earnings Forecast #6
            'Working!!
            txtActivityLog.AppendText(Now & " Getting Earnings Forecast." & vbCrLf)
            Dim bitEarningsForecastSucessful As Boolean = GetEarningsForecast(strSymbol)
            If bitEarningsSurprisesSucessful = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Earnings Forecast for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #11. Insider Trading
            'Working!!!
            txtActivityLog.AppendText(Now & " Getting Insider Trading" & vbCrLf)
            Dim bitInsideTrading As Boolean = GetInsideTrading(strSymbol)
            If bitInsideTrading = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting Insider Trading for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #12. Weighted Alpha
            ' Working

            'DEBUG:

            txtActivityLog.AppendText(Now & " Getting  Weighted Alpha" & vbCrLf)
            Dim bitWeightedAlpha As Boolean = GetWeightedAlpha(strSymbol)
            If bitWeightedAlpha = False Then
                txtActivityLog.AppendText(Now & " There was an issue getting  Weighted Alpha for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            'For Debug only
            ' Continue For
            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Calulate the score and put into the table
            ' Working
            txtActivityLog.AppendText(Now & " Calculating Overall Score" & vbCrLf)
            Dim bitCalculateOverallScore As Boolean = CalculateOverallScore(strSymbol)
            If bitCalculateOverallScore = False Then
                txtActivityLog.AppendText(Now & " There was an issue Calculating Overall Score for " & strSymbol & "." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Get Analyst price Targets
            ' Working
            txtActivityLog.AppendText(Now & " Getting Analyst Price Targets" & vbCrLf)
            Dim bitGetTargetPrice As Boolean = GetBCATargetPrice(strSymbol)
            'If bitCalculateTargetPrice = False Then
            '    txtLegacyScoreActivityLog.AppendText(Now & " There was an issue Calculating Target Price for " & strSymbol & "." & vbCrLf)
            'End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Get the Next earnings date
            ' Working
            txtActivityLog.AppendText(Now & " Getting the Expected Earnings Date" & vbCrLf)
            Dim bitExpEarningsDate As Boolean = GetExpEarningsDate(strSymbol)
            If bitExpEarningsDate = False Then
                txtActivityLog.AppendText(Now & " There was an issue Getting the Expected Earnings Date." & vbCrLf)
            End If
            Application.DoEvents()

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Get the PE and ROE
            ' Working
            txtActivityLog.AppendText(Now & " Getting the PE From Zacks" & vbCrLf)
            Dim bitPEFromZacks As Boolean = GetPEFromZacks(strSymbol)
            If bitPEFromZacks = False Then
                txtActivityLog.AppendText(Now & " There was an issue Getting the PE From Zacks." & vbCrLf)
            End If

            '++++++++++++++++++++++++++++++++++++++++++++++
            ' #Get the IV

            txtActivityLog.AppendText(Now & " Getting the IV From CBOE" & vbCrLf)
            'Dim bitIVFromCBOE As Boolean = GetIVFromCBOE(strSymbol)
            'If bitIVFromCBOE = False Then
            '    txtActivityLog.AppendText(Now & " There was an issue Getting the IV From CBOE." & vbCrLf)
            'End If

        Next
        '++++++++++++++++++++++++++++++++++++++++++++++
        ' #Calulate the Sector Scores and put into the table
        txtActivityLog.AppendText(Now & " Calculating Overall Scores" & vbCrLf)
        Dim bitCalculateSectorScore As Boolean = CalculateSectorScores()
        If bitCalculateSectorScore = False Then
            txtActivityLog.AppendText(Now & " There was an issue Calculating Sector Scores." & vbCrLf)
        End If
        Application.DoEvents()

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

        Dim URI As String = "https://fundamentals.nasdaq.com/redpage.asp?selected=" & strURLEncodedSymbol

        strResponseString = String.Empty

        Dim webClient As New WebClient()

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidateCertificate)

        webClient.Headers.Add(HttpRequestHeader.Host, "www.nasdaq.com")
        webClient.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Mobile Safari/537.36")
        webClient.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8")
        'these didnt work ugh....
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36
        'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Mobile Safari/537.36
        webClient.Headers.Add(HttpRequestHeader.AcceptEncoding, "deflate")
        webClient.Headers.Add(HttpRequestHeader.AcceptLanguage, "en-US,en;q=0.9")

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

        Dim URI As String = "https://www.nasdaq.com/symbol/" & strURLEncodedSymbol.ToLower & "/financials?query=ratios"
        'symbol has to be lower case
        'http://www.nasdaq.com/symbol/aa/financials?query=ratios

        strResponseString = String.Empty

        Dim webClient As New WebClient()

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        webClient.Headers.Add(HttpRequestHeader.Host, "www.nasdaq.com")
        webClient.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Mobile Safari/537.36")
        webClient.Headers.Add(HttpRequestHeader.Accept, "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8")
        'these didnt work ugh....
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36
        'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Mobile Safari/537.36
        webClient.Headers.Add(HttpRequestHeader.AcceptEncoding, "deflate")
        webClient.Headers.Add(HttpRequestHeader.AcceptLanguage, "en-US,en;q=0.9")

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

        If strResponseString.Contains(">DATA IS CURRENTLY NOT AVAILABLE</DIV>") Then
            txtActivityLog.AppendText(Now & " >DATA IS CURRENTLY NOT AVAILABLE " & strSymbol & vbCrLf)

            Return False
        End If
        '
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

    Public Function GetDaysToCover(strSymbol) As Boolean
        Dim bitSucessful As Boolean = True

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "-")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "-")
        'https://finance.yahoo.com/quote/AAPL/key-statistics?p=AAPL
        Dim URI As String = "https://finance.yahoo.com/quote/" & strURLEncodedSymbol & "/key-statistics?p=" & strURLEncodedSymbol

        Dim webClient As New WebClient()

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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
            Return False
        End If
        'first find the start of the data

        Dim i As Integer = 1
        Dim j As Integer = -1

        Dim strSharesShort As String = String.Empty
        Dim strDaysToCover As String
        Dim strDaysToCoverScore As String = "FAIL"

        Try
            i = InStr(i, strResponseString, ">SHARES SHORT (")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "<TD")
            'i = InStr(i + 1, strResponseString, "<SPAN")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</")
            strSharesShort = Mid(strResponseString, (i + 1), (j - i - 1))


            'If strSharesShort <> "N/A" Then
            '    i = 1
            '    i = InStr(i, strResponseString, ">SHARES SHORT</")  'Start of section to grab
            '    i = InStr(i + 1, strResponseString, "<TD")
            '    i = InStr(i + 1, strResponseString, ">")
            '    j = InStr(i, strResponseString, "</")
            '    strSharesShort = Mid(strResponseString, (i + 1), (j - i - 1))
            'End If
            i = 1

            i = InStr(i, strResponseString, ">SHORT RATIO (")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "<TD")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</")
            strDaysToCover = Mid(strResponseString, (i + 1), (j - i - 1))

            If strSharesShort = "N/A" Then
                strDaysToCover = "99"
            ElseIf IsNumeric(strDaysToCover) = False Then
                strDaysToCover = "99"
            ElseIf strDaysToCover < 2 Then
                strDaysToCoverScore = "PASS"
            End If

            Dim params(3) As SqlClient.SqlParameter
            params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            params(0).Value = strSymbol
            params(1) = New SqlClient.SqlParameter("@strDaysToCover", SqlDbType.VarChar)
            params(1).Value = strDaysToCover
            params(2) = New SqlClient.SqlParameter("@strDaysToCoverScore", SqlDbType.VarChar)
            params(2).Value = strDaysToCoverScore
            params(3) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            params(3).Value = Today.ToShortDateString

            Dim dsSPResults As DataSet = RunSP("dbo.spUpdateDaysToCover", params)
        Catch ex As Exception
            Return False
        End Try


        'Get PEG From here also

        '>PEG RATIO
        Dim strPEGRatio As String
        Try
            i = 1
            i = InStr(i, strResponseString, ">PEG RATIO")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "<TD")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</")
            strPEGRatio = Mid(strResponseString, (i + 1), (j - i - 1))

            If IsNumeric(strPEGRatio) = False Then
                strPEGRatio = "99"
            End If

            Dim strPEGRatioScore As String = "FAIL"
            If strPEGRatio > 0 AndAlso strPEGRatio < 1.0 Then
                strPEGRatioScore = "PASS"
            End If

            Dim paramsPEG(3) As SqlClient.SqlParameter
            paramsPEG(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            paramsPEG(0).Value = strSymbol
            paramsPEG(1) = New SqlClient.SqlParameter("@strPEGRatio", SqlDbType.VarChar)
            paramsPEG(1).Value = strPEGRatio
            paramsPEG(2) = New SqlClient.SqlParameter("@strPEGRatioScore", SqlDbType.VarChar)
            paramsPEG(2).Value = strPEGRatioScore
            paramsPEG(3) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            paramsPEG(3).Value = Today.ToShortDateString

            Dim dsSPResultsPEG As DataSet = RunSP("dbo.spUpdatePEGRatio", paramsPEG)

        Catch ex As Exception
            Return False
        End Try
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
        ServicePointManager.SecurityProtocol = (SecurityProtocolType.Tls Or (SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12))

        Try
            ''https://api.nasdaq.com/api/company/AAPL/earnings-surprise
            Dim httpGetData = CType(WebRequest.Create("https://api.nasdaq.com/api/company/" & strURLEncodedSymbol.ToLower & "/earnings-surprise"), HttpWebRequest)
            httpGetData.Host = "api.nasdaq.com"
            httpGetData.UserAgent = "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Mobile Safari/537.36"
            httpGetData.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
            httpGetData.Headers.Add(HttpRequestHeader.AcceptEncoding, "deflate")
            httpGetData.Headers.Add(HttpRequestHeader.AcceptLanguage, "en-US,en;q=0.9")

            Dim httpDataResponse As HttpWebResponse = CType(httpGetData.GetResponse, HttpWebResponse)

            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
            ' Pipes the response stream to a higher level stream reader with the required encoding format. 
            Dim receiveStream As Stream = httpDataResponse.GetResponseStream()
            Dim readStream As New StreamReader(receiveStream, encode)
            Dim strTheRawJSON As String = readStream.ReadToEnd
            readStream.Close()
            'Parse the Raw Json and put into database
            Dim myData = JsonConvert.DeserializeObject(strTheRawJSON)

            Try
                strCurrentEarningsSurprise = myData.GetValue("data").Item("earningsSurpriseTable").Item("rows")(0).Item("percentageSurprise").ToString.ToUpper
            Catch ex As Exception
                strCurrentEarningsSurprise = "N/A"
            End Try

            Try
                strPastEarningsSurprise1Ago = myData.GetValue("data").Item("earningsSurpriseTable").Item("rows")(1).Item("percentageSurprise").ToString.ToUpper
            Catch ex As Exception
                strPastEarningsSurprise1Ago = "N/A"
            End Try

            Try
                strPastEarningsSurprise2Ago = myData.GetValue("data").Item("earningsSurpriseTable").Item("rows")(2).Item("percentageSurprise").ToString.ToUpper
            Catch ex As Exception
                strPastEarningsSurprise2Ago = "N/A"
            End Try

            Try
                strPastEarningsSurprise3Ago = myData.GetValue("data").Item("earningsSurpriseTable").Item("rows")(3).Item("percentageSurprise").ToString.ToUpper
            Catch ex As Exception
                strPastEarningsSurprise3Ago = "N/A"
            End Try

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
        Catch e As Exception
            Console.WriteLine(e.Message)
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

        strResponseString = String.Empty

        ServicePointManager.Expect100Continue = True
        'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        System.Net.ServicePointManager.SecurityProtocol = (SecurityProtocolType.Tls Or (SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12))

        Try

            'https://api.nasdaq.com/api/company/AAPL/earnings-surprise
            Dim httpGetData = CType(WebRequest.Create("https://api.nasdaq.com/api/analyst/" & strURLEncodedSymbol.ToLower & "/earnings-forecast"), HttpWebRequest)
            httpGetData.Host = "api.nasdaq.com"
            httpGetData.UserAgent = "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Mobile Safari/537.36"
            httpGetData.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
            httpGetData.Headers.Add(HttpRequestHeader.AcceptEncoding, "deflate")
            httpGetData.Headers.Add(HttpRequestHeader.AcceptLanguage, "en-US,en;q=0.9")
            Dim httpDataResponse As HttpWebResponse = CType(httpGetData.GetResponse, HttpWebResponse)

            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
            ' Pipes the response stream to a higher level stream reader with the required encoding format. 
            Dim receiveStream As Stream = httpDataResponse.GetResponseStream()
            Dim readStream As New StreamReader(receiveStream, encode)
            Dim strTheRawJSON As String = readStream.ReadToEnd
            readStream.Close()
            'Parse the Raw Json and put into database
            Dim myData = JsonConvert.DeserializeObject(strTheRawJSON)

            Try
                strCurrentEarningsForecast = myData.GetValue("data").Item("quarterlyForecast").Item("rows")(0).Item("consensusEPSForecast").ToString.ToUpper
            Catch ex As Exception
                strCurrentEarningsForecast = "N/A"
            End Try

            Try
                strPastEarningsForecast1Ago = myData.GetValue("data").Item("quarterlyForecast").Item("rows")(1).Item("consensusEPSForecast").ToString.ToUpper
            Catch ex As Exception
                strPastEarningsForecast1Ago = "N/A"
            End Try

            Try
                strPastEarningsForecast2Ago = myData.GetValue("data").Item("quarterlyForecast").Item("rows")(2).Item("consensusEPSForecast").ToString.ToUpper
            Catch ex As Exception
                strPastEarningsForecast2Ago = "N/A"
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
        Catch ex As Exception
            Dim x = 1
        End Try

        Return bitSucessful
    End Function

    Public Function GetWeightedAlpha(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 1
        Dim j As Integer = 0

        Dim strWeightedAlpha As String = String.Empty
        Dim strWeightedAlphaScore As String = "FAIL"

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", ".")

        Dim URI As String = "https://www.barchart.com/stocks/quotes/" & strURLEncodedSymbol.ToLower
        'Dim URI As String = "http://old.barchart.com/quotes/stocks/" & strURLEncodedSymbol.ToLower
        'symbol has to be lower case
        'http://www.barchart.com/quotes/stocks/a

        Dim webClient As New WebClient()
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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

            i = InStr(i, strResponseString, "WEIGHTEDALPHA&QUOT;")  'Start of section to grab
            i = InStr(i + 1, strResponseString, "WEIGHTEDALPHA&QUOT;")  'Start of section to grab
            i = InStr(i + 1, strResponseString, ";")
            i = InStr(i + 1, strResponseString, ";")
            j = InStr(i, strResponseString, "&")
            strWeightedAlpha = Mid(strResponseString, (i + 1), (j - i - 1))

            If strWeightedAlpha(0) = "+" Then
                strWeightedAlpha = Strings.Right(strWeightedAlpha, Len(strWeightedAlpha) - 1)
                ' ElseIf strWeightedAlpha(0) = "-" Then
                '     strWeightedAlpha = Strings.Right(strWeightedAlpha, Len(strWeightedAlpha) - 1)
            End If

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
        txtActivityLog.AppendText(Now & " The Weighted Alpha for  " & strSymbol & " is " & strWeightedAlpha & vbCrLf)
        txtActivityLog.AppendText(Now & " The Weighted Alpha SCORE for  " & strSymbol & " is " & strWeightedAlphaScore & vbCrLf)
        Return bitSucessful
    End Function

    Public Function GetInsideTrading(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 1
        Dim j As Integer = 0

        Dim strSharesBought As String = String.Empty
        Dim strSharesSold As String = String.Empty
        Dim strInsiderTradingScore As String = "FAIL"

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "-")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "-")
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        Try

            'https://api.nasdaq.com/api/company/AAPL/insider-trades?limit=99999&type=ALL
            Dim httpGetData = CType(WebRequest.Create("https://api.nasdaq.com/api/company/" & strURLEncodedSymbol.ToLower & "/insider-trades?limit=99999&type=ALL"), HttpWebRequest)
            httpGetData.Host = "api.nasdaq.com"
            httpGetData.UserAgent = "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Mobile Safari/537.36"
            httpGetData.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
            httpGetData.Headers.Add(HttpRequestHeader.AcceptEncoding, "deflate")
            httpGetData.Headers.Add(HttpRequestHeader.AcceptLanguage, "en-US,en;q=0.9")

            Dim httpDataResponse As HttpWebResponse = CType(httpGetData.GetResponse, HttpWebResponse)

            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
            ' Pipes the response stream to a higher level stream reader with the required encoding format. 
            Dim receiveStream As Stream = httpDataResponse.GetResponseStream()
            Dim readStream As New StreamReader(receiveStream, encode)
            Dim strTheRawJSON As String = readStream.ReadToEnd
            readStream.Close()
            'Parse the Raw Json and put into database
            Dim myData = JsonConvert.DeserializeObject(strTheRawJSON)

            Try
                strSharesBought = myData.GetValue("data").Item("numberOfSharesTraded").Item("rows")(0).Item("months3").ToString.ToUpper
            Catch ex As Exception
                strSharesBought = "N/A"
            End Try
            Try
                strSharesSold = myData.GetValue("data").Item("numberOfSharesTraded").Item("rows")(1).Item("months3").ToString.ToUpper
            Catch ex As Exception
                strSharesSold = "N/A"
            End Try

            If IsNumeric(strSharesBought) AndAlso IsNumeric(strSharesSold) Then
                Dim intNetResult As Integer = CInt(strSharesBought) - CInt(strSharesSold)
                Try
                    If intNetResult > 0 Then
                        strInsiderTradingScore = "PASS"
                    End If
                Catch ex As Exception
                End Try

                Dim params(3) As SqlClient.SqlParameter
                params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
                params(0).Value = strSymbol
                params(1) = New SqlClient.SqlParameter("@strInsiderTrading", SqlDbType.VarChar)
                Try
                    params(1).Value = CStr(intNetResult)

                Catch ex As Exception
                    params(1).Value = "N/A"
                End Try
                params(2) = New SqlClient.SqlParameter("@strInsiderTradingScore", SqlDbType.VarChar)
                params(2).Value = strInsiderTradingScore
                params(3) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
                params(3).Value = Today.ToShortDateString

                Dim dsSPResults As DataSet = RunSP("dbo.spUpdateInsiderTrading", params)
            End If
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try

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

    Public Function GetBCATargetPrice(strSymbol As String)
        Dim bitSucessful As Boolean = True
        Dim intInitialStartingPlace As Integer = 0
        Dim intCurrentPlace As Integer = 0
        Dim intStartingPlace As Integer = 1
        Dim intEndingPlace As Integer = 0

        Dim strTargetLowPrice As String = "N/A"
        Dim strTargetMeanPrice As String = "N/A"
        Dim strTargetHighPrice As String = "N/A"

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", "-")
        strURLEncodedSymbol = strURLEncodedSymbol.Replace(".", "-")

        Dim URI As String = "https://finance.yahoo.com/quote/" & strURLEncodedSymbol & "/analysts?p=" & strURLEncodedSymbol

        Dim webClient As New WebClient()

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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

        'first find the start of the data

        intCurrentPlace = InStr(1, strResponseString, "ROOT.APP.MAIN =")

        If intCurrentPlace <> 0 Then
            intInitialStartingPlace = intCurrentPlace
            intStartingPlace = intCurrentPlace
            intEndingPlace = intCurrentPlace

            '"TARGETLOWPRICE":{"RAW":129.58,"FMT":"129.58"}
            intCurrentPlace = InStr(intStartingPlace, strResponseString, "TARGETLOWPRICE")
            If intCurrentPlace <> 0 Then
                intStartingPlace = intCurrentPlace + 23
                intEndingPlace = InStr(intStartingPlace, strResponseString, ",")

                Dim strTempPrice = Mid(strResponseString, intStartingPlace, (intEndingPlace - intStartingPlace))
                If IsNumeric(strTempPrice) = True Then
                    strTargetLowPrice = FormatNumber(strTempPrice)
                End If
            End If

            '"TARGETMEANPRICE":{"RAW":187.74,"FMT":"187.74"}
            intCurrentPlace = intInitialStartingPlace
            intStartingPlace = intCurrentPlace
            intEndingPlace = intCurrentPlace

            intCurrentPlace = InStr(intStartingPlace, strResponseString, "TARGETMEANPRICE")
            If intCurrentPlace <> 0 Then
                intStartingPlace = intCurrentPlace + 24
                intEndingPlace = InStr(intStartingPlace, strResponseString, ",")

                Dim strTempPrice = Mid(strResponseString, intStartingPlace, (intEndingPlace - intStartingPlace))
                If IsNumeric(strTempPrice) = True Then
                    strTargetMeanPrice = FormatNumber(strTempPrice)
                End If
            End If

            '"TARGETHIGHPRICE":{"RAW":235,"FMT":"235.00"}
            intCurrentPlace = intInitialStartingPlace
            intStartingPlace = intCurrentPlace
            intEndingPlace = intCurrentPlace

            intCurrentPlace = InStr(intStartingPlace, strResponseString, "TARGETHIGHPRICE")
            If intCurrentPlace <> 0 Then
                intStartingPlace = intCurrentPlace + 24
                intEndingPlace = InStr(intStartingPlace, strResponseString, ",")

                Dim strTempPrice = Mid(strResponseString, intStartingPlace, (intEndingPlace - intStartingPlace))
                If IsNumeric(strTempPrice) = True Then
                    strTargetHighPrice = FormatNumber(strTempPrice)
                End If
            End If
        End If

        Dim paramsCompanyTargets(4) As SqlClient.SqlParameter
        paramsCompanyTargets(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        paramsCompanyTargets(0).Value = strSymbol
        paramsCompanyTargets(1) = New SqlClient.SqlParameter("@strTargetLowPrice", SqlDbType.VarChar)
        paramsCompanyTargets(1).Value = strTargetLowPrice
        paramsCompanyTargets(2) = New SqlClient.SqlParameter("@strTargetMeanPrice", SqlDbType.VarChar)
        paramsCompanyTargets(2).Value = strTargetMeanPrice
        paramsCompanyTargets(3) = New SqlClient.SqlParameter("@strTargetHighPrice", SqlDbType.VarChar)
        paramsCompanyTargets(3).Value = strTargetHighPrice
        paramsCompanyTargets(4) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
        paramsCompanyTargets(4).Value = Today
        Dim dsSPResultsCompanyTargets As DataSet = RunSP("dbo.spBCAUpdateTargetPrices", paramsCompanyTargets)

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

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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

        'This is quarterly ROE
        i = 1
        Dim strCurrentROE As String = String.Empty
        '">ROE</
        i = InStr(i + 1, strResponseString, ">ROE</")
        i = InStr(i + 1, strResponseString, "</TD")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strCurrentROE = Mid(strResponseString, (i + 1), (j - i - 1))

        Dim strPastROE1yrAgo As String = String.Empty
        '">ROE</
        i = InStr(i + 1, strResponseString, "<TR>")
        i = InStr(i + 1, strResponseString, "</TD")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strPastROE1yrAgo = Mid(strResponseString, (i + 1), (j - i - 1))

        Dim strPastROE2yrsAgo As String = String.Empty
        '">ROE</
        i = InStr(i + 1, strResponseString, "<TR>")
        i = InStr(i + 1, strResponseString, "</TD")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strPastROE2yrsAgo = Mid(strResponseString, (i + 1), (j - i - 1))

        Try
            If IsNumeric(strCurrentROE) Then
                Dim strROEScore As String = "FAIL"
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

        Return bitSucessful
    End Function

    Public Function GetIVFromCBOE(strSymbol As String) As Boolean
        Dim SW2 As New Stopwatch
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
            SW2.Restart()
            Do
                Application.DoEvents()
            Loop Until SW2.ElapsedMilliseconds >= 20000

            Application.DoEvents()

        End If

RETRY:
        strResponseString = ""
        'Load the page and get the response
        Dim URI As String = "http://cboe.ivolatility.com/options.j?ticker=" & strURLEncodedSymbol.ToLower & "&R=0&x=0&y=0"
        strResponseString = GetResponseStringFromURL(URI, ">PRIVACY STATEMENT</A>")

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
        'i = InStr(i + 1, strResponseString, "WR(D[")
        i = InStr(i + 4, strResponseString, ">")
        'j = InStr(i, strResponseString, "]")
        j = InStr(i, strResponseString, "</")
        'Dim strIVXMeanDataPointName As String = Mid(strResponseString, (i + 1), (j - i - 2))
        Dim strIVXMeanDataPoint As String = Mid(strResponseString, (i + 1), (j - i - 2))

        i = j + 1
        '>52 week High
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, "<TD")
        'i = InStr(i + 1, strResponseString, "WR(D[")
        i = InStr(i + 4, strResponseString, ">")
        'j = InStr(i, strResponseString, "]")
        j = InStr(i, strResponseString, "%")
        'Dim strIVXHighDataPointName As String = Mid(strResponseString, (i + 1), (j - i - 1))
        Dim strIVXHighDataPoint As String = Mid(strResponseString, (i + 1), (j - i - 1))

        i = j + 1
        '>52 week Low
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 4, strResponseString, ">")
        'j = InStr(i, strResponseString, "]")
        j = InStr(i, strResponseString, "%")
        'Dim strIVXLowDataPointName As String = Mid(strResponseString, (i + 1), (j - i - 1))
        Dim strIVXLowDataPoint As String = Mid(strResponseString, (i + 1), (j - i - 1))

        '2. look for the names in the array

        'i = 1
        'i = InStr(i + 1, strResponseString, strIVXMeanDataPointName & "=")
        'i = InStr(i + 1, strResponseString, "'")
        'j = InStr(i + 1, strResponseString, "'")
        'Dim strIVXMeanDataPoint As String = Mid(strResponseString, (i + 1), (j - i - 1))

        'i = 1
        'i = InStr(i + 1, strResponseString, strIVXHighDataPointName & "=")
        'i = InStr(i + 1, strResponseString, "'")
        'j = InStr(i + 1, strResponseString, "'")
        'Dim strIVXHighDataPoint As String = Mid(strResponseString, (i + 1), (j - i - 1))

        'i = 1
        'i = InStr(i + 1, strResponseString, strIVXLowDataPointName & "=")
        'i = InStr(i + 1, strResponseString, "'")
        'j = InStr(i + 1, strResponseString, "'")
        'Dim strIVXLowDataPoint As String = Mid(strResponseString, (i + 1), (j - i - 1))

        'Calculate the IV Percentile

        'here I think we are going to subtact the 52week low from the high and the current
        'then get a percentage. that will give us a percentage within the range

        Dim strIVXResultDataPoint As String = 0

        Try
            'Dim decIVXMeanDataPoint As Decimal = CDec(Replace(strIVXMeanDataPoint, "%", ""))
            'Dim decIVXHighDataPoint As Decimal = CDec(Replace(strIVXHighDataPoint, "%", ""))
            'Dim decIVXLowDataPoint As Decimal = CDec(Replace(strIVXLowDataPoint, "%", ""))
            Dim decIVXMeanDataPoint As Decimal = CDec(strIVXMeanDataPoint)
            Dim decIVXHighDataPoint As Decimal = CDec(strIVXHighDataPoint)
            Dim decIVXLowDataPoint As Decimal = CDec(strIVXLowDataPoint)

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

    Public Function GetZacksAnalystRecommendationsAndIndustryComparison(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim intRetryNumber As Integer = 0

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", ".")

        Dim URI As String = "http://www.zacks.com/stock/research/" & strURLEncodedSymbol.ToUpper & "/industry-comparison"

        Dim webClient As New WebClient()

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        strResponseString = String.Empty
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12


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
        Dim strAnalystRecommendation As String = String.Empty
        Dim strIndustryAnalystRecommendation As String = String.Empty
        Dim strAnalystRecommendationScore As String = String.Empty
        Dim strIndustryAnalystRecommendationScore As String = String.Empty

        '>AVERAGE RECOMMENDATION (1=BUY, 5=SELL)</
        i = InStr(i + 1, strResponseString, ">AVERAGE RECOMMENDATION (1=BUY, 5=SELL)</")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strAnalystRecommendation = Mid(strResponseString, (i + 1), (j - i - 1))

        'industry analyst recs
        i = InStr(j + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strIndustryAnalystRecommendation = Mid(strResponseString, (i + 1), (j - i - 1))

        Try
            If strAnalystRecommendation < 2.25 Then
                strAnalystRecommendationScore = "PASS"
            Else
                strAnalystRecommendationScore = "FAIL"
            End If
        Catch ex As Exception
            strAnalystRecommendationScore = "FAIL"
        End Try

        If IsNumeric(strAnalystRecommendation) Then
            Dim params(3) As SqlClient.SqlParameter
            params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            params(0).Value = strSymbol
            params(1) = New SqlClient.SqlParameter("@strAnalystRecommendation", SqlDbType.VarChar)
            params(1).Value = strAnalystRecommendation
            params(2) = New SqlClient.SqlParameter("@strAnalystRecommendationScore", SqlDbType.VarChar)
            params(2).Value = strAnalystRecommendationScore
            params(3) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            params(3).Value = Today.ToShortDateString

            Dim dsSPResultsPERatio As DataSet = RunSP("dbo.spUpdateAnalystRecommendation", params)
        End If

        '#############################################################################
        ' HERE WE ARE GOING TO MAKE A FEW COMPARISONS TO GET THE INDUSTY SCORE.

        Dim strLast5Years As String = String.Empty
        Dim strNext5Years As String = String.Empty
        Dim strNetProfitMargin As String = String.Empty
        Dim strReturnOnEquity As String = String.Empty

        Dim strIndustryLast5Years As String = String.Empty
        Dim strIndustryNext5Years As String = String.Empty
        Dim strIndustryNetProfitMargin As String = String.Empty
        Dim strIndustryReturnOnEquity As String = String.Empty

        '>LAST 5 YEARS</
        i = 1
        i = InStr(i + 1, strResponseString, ">LAST 5 YEARS</")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strLast5Years = Mid(strResponseString, (i + 1), (j - i - 1))

        i = InStr(j + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strIndustryLast5Years = Mid(strResponseString, (i + 1), (j - i - 1))

        '>NEXT 5 YEARS</
        i = 1
        i = InStr(i + 1, strResponseString, ">NEXT 5 YEARS</")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strNext5Years = Mid(strResponseString, (i + 1), (j - i - 1))

        i = InStr(j + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strIndustryNext5Years = Mid(strResponseString, (i + 1), (j - i - 1))

        '>NET PROFIT MARGIN (TTM)</
        i = 1
        i = InStr(i + 1, strResponseString, ">NET PROFIT MARGIN (TTM)</")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strNetProfitMargin = Mid(strResponseString, (i + 1), (j - i - 1))

        i = InStr(j + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strIndustryNetProfitMargin = Mid(strResponseString, (i + 1), (j - i - 1))


        '>RETURN ON EQUITY (TTM)</
        i = 1
        i = InStr(i + 1, strResponseString, ">RETURN ON EQUITY (TTM)</")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</")
        strReturnOnEquity = Mid(strResponseString, (i + 1), (j - i - 1))

        i = InStr(j + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</")
        strIndustryReturnOnEquity = Mid(strResponseString, (i + 1), (j - i - 1))


        'So now we are going to compare all this stuff agains its Industry numbers
        'For now, we are going to say 4 to pass.  I might make it 3.  we will see
        Dim intNumberOfPassingGrades As Integer = 0 'Start out failing.
        Dim intNumberOfTotalGrades As Integer = 0 'Start out failing.

        If IsNumeric(strAnalystRecommendation) = True AndAlso IsNumeric(strIndustryAnalystRecommendation) = True Then
            intNumberOfTotalGrades += 1
            Try
                If strAnalystRecommendation < strIndustryAnalystRecommendation Then
                    intNumberOfPassingGrades += 1
                End If
            Catch ex As Exception
            End Try
        End If

        If IsNumeric(strLast5Years) = True AndAlso IsNumeric(strIndustryLast5Years) = True Then
            intNumberOfTotalGrades += 1
            Try
                If strLast5Years > strIndustryLast5Years Then
                    intNumberOfPassingGrades += 1
                End If
            Catch ex As Exception
            End Try
        End If

        If IsNumeric(strNext5Years) = True AndAlso IsNumeric(strIndustryNext5Years) = True Then
            intNumberOfTotalGrades += 1
            Try
                If strNext5Years > strIndustryNext5Years Then
                    intNumberOfPassingGrades += 1
                End If
            Catch ex As Exception
            End Try
        End If

        strNetProfitMargin = strNetProfitMargin.Remove(Len(strNetProfitMargin) - 1)
        strIndustryNetProfitMargin = strIndustryNetProfitMargin.Remove(Len(strIndustryNetProfitMargin) - 1)
        If IsNumeric(strNetProfitMargin) = True AndAlso IsNumeric(strIndustryNetProfitMargin) = True Then
            intNumberOfTotalGrades += 1
            Try
                If strNetProfitMargin > strIndustryNetProfitMargin Then
                    intNumberOfPassingGrades += 1
                End If
            Catch ex As Exception
            End Try
        End If

        strReturnOnEquity = strReturnOnEquity.Remove(Len(strReturnOnEquity) - 1)
        strIndustryReturnOnEquity = strIndustryReturnOnEquity.Remove(Len(strIndustryReturnOnEquity) - 1)
        If IsNumeric(strReturnOnEquity) = True AndAlso IsNumeric(strIndustryReturnOnEquity) = True Then
            intNumberOfTotalGrades += 1
            Try
                If strReturnOnEquity > strIndustryReturnOnEquity Then
                    intNumberOfPassingGrades += 1
                End If
            Catch ex As Exception
            End Try
        End If

        Dim strIndustryEarningsScore As String = "FAIL"
        Try
            If intNumberOfPassingGrades / intNumberOfTotalGrades >= 0.5 Then
                strIndustryEarningsScore = "PASS"
            End If
        Catch ex As Exception
            'it is already fail
        End Try
        'If intNumberOfPassingGrades >= 4 Then
        '    strIndustryEarningsScore = "PASS"
        'End If

        Dim paramsIndustry(3) As SqlClient.SqlParameter
        paramsIndustry(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
        paramsIndustry(0).Value = strSymbol
        paramsIndustry(1) = New SqlClient.SqlParameter("@strIndustryEarningsScore", SqlDbType.VarChar)
        paramsIndustry(1).Value = strIndustryEarningsScore
        paramsIndustry(2) = New SqlClient.SqlParameter("@intNumberOfPassingGrades", SqlDbType.VarChar)
        paramsIndustry(2).Value = intNumberOfPassingGrades
        paramsIndustry(3) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
        paramsIndustry(3).Value = Today.ToShortDateString

        Dim dsSPResultsparamsIndustry As DataSet = RunSP("dbo.spUpdateIndustryEarnings", paramsIndustry)

        Return bitSucessful
    End Function

    Public Function GetEarningsGrowthFromZacks(strSymbol As String) As Boolean
        Dim bitSucessful As Boolean = True
        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim intRetryNumber As Integer = 0

        Dim strURLEncodedSymbol As String = strSymbol.Replace("/", ".")

        Dim URI As String = "https://www.zacks.com/stock/quote/" & strURLEncodedSymbol.ToUpper & "?q=" & strURLEncodedSymbol.ToUpper
        'https://www.zacks.com/stock/quote/AAPL?q=AAPL 
        Dim webClient As New WebClient()

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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


        'Earnings Growth
        i = 1
        Dim strEarningsGrowth As String = String.Empty
        Dim strEarningsGrowthScore As String = String.Empty
        '>VS. PREVIOUS YEAR</
        i = InStr(i + 1, strResponseString, ">EXP EPS GROWTH")
        i = InStr(i + 1, strResponseString, "<P ")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</")
        strEarningsGrowth = Mid(strResponseString, (i + 1), (j - i - 1))
        'This is a percentage string ie. 13.56%
        strEarningsGrowth = strEarningsGrowth.Remove(strEarningsGrowth.Length - 1)
        Try
            If strEarningsGrowth >= 8 Then
                strEarningsGrowthScore = "PASS"
            Else
                strEarningsGrowthScore = "FAIL"
            End If
        Catch ex As Exception
            strEarningsGrowthScore = "FAIL"
        End Try

        If IsNumeric(strEarningsGrowth) Then
            Dim paramsEarnings(3) As SqlClient.SqlParameter
            paramsEarnings(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            paramsEarnings(0).Value = strSymbol
            paramsEarnings(1) = New SqlClient.SqlParameter("@strEarningsGrowth", SqlDbType.VarChar)
            paramsEarnings(1).Value = strEarningsGrowth
            paramsEarnings(2) = New SqlClient.SqlParameter("@strEarningsGrowthScore", SqlDbType.VarChar)
            paramsEarnings(2).Value = strEarningsGrowthScore
            paramsEarnings(3) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            paramsEarnings(3).Value = Today.ToShortDateString

            Dim dsSPResultsparamsEarnings As DataSet = RunSP("dbo.spUpdateEarningsGrowth", paramsEarnings)
        End If
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

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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

        'forward PE
        i = 1
        Dim strForwardPE As String = String.Empty
        '">P/E (F1)</
        i = InStr(i + 1, strResponseString, ">P/E (F1)</")
        i = InStr(i + 1, strResponseString, "<TD")
        i = InStr(i + 1, strResponseString, ">")
        j = InStr(i, strResponseString, "</TD")
        strForwardPE = Mid(strResponseString, (i + 1), (j - i - 1))

        If IsNumeric(strForwardPE) Then
            Dim params(2) As SqlClient.SqlParameter
            params(0) = New SqlClient.SqlParameter("@strSymbol", SqlDbType.VarChar)
            params(0).Value = strSymbol
            params(1) = New SqlClient.SqlParameter("@strForwardPE", SqlDbType.VarChar)
            params(1).Value = strForwardPE
            params(2) = New SqlClient.SqlParameter("@dProcessed", SqlDbType.Date)
            params(2).Value = Today.ToShortDateString

            Dim dsSPResultsPERatio As DataSet = RunSP("dbo.spUpdateForwardPE", params)
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
        'Dim URI As String = "https://www.google.com/finance/historical?q=" & strSymbol & "&start=0&num=200"
        Dim URI As String = "https://finance.yahoo.com/quote/" & strSymbol & "/history?period1=1496880000&period2=1512604800&interval=1wk&filter=history&frequency=1wk"
        'https://finance.yahoo.com/quote/AAPL/history?period1=1496880000&period2=1512604800&interval=1wk&filter=history&frequency=1wk


        Dim webClient As New WebClient()

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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
            i = InStr(i + 1, strResponseString, "HISTORICAL-PRICES")
            i = InStr(i + 1, strResponseString, "<TBODY ")
            i = InStr(i + 1, strResponseString, ">")
            j = InStr(i, strResponseString, "</TBODY>")
            strTbody = Mid(strResponseString, (i + 1), (j - i - 1)).Trim
        Catch ex As Exception
            strTbody = "0"
        End Try
        '1b. do a split on the rows
        '1c. for each row get data if  not a dividend row

        Dim strTableRows() As String = strTbody.Split(New String() {"<TR "}, StringSplitOptions.RemoveEmptyEntries)

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
                    i = InStr(i + 1, strTableRows(intCounter), "<TD")
                    i = InStr(i + 1, strTableRows(intCounter), "<TD")
                    i = InStr(i + 1, strTableRows(intCounter), "<TD")
                    i = InStr(i + 1, strTableRows(intCounter), "<TD")
                    i = InStr(i + 1, strTableRows(intCounter), "<TD")
                    i = InStr(i + 1, strTableRows(intCounter), "<TD")
                    i = InStr(i + 1, strTableRows(intCounter), "<") 'span
                    i = InStr(i + 1, strTableRows(intCounter), ">")
                    j = InStr(i + 1, strTableRows(intCounter), "<")
                    decClosingPriceTemp = CDec(Mid(strTableRows(intCounter), (i + 1), (j - i - 1)).Trim)
                    intNumberOfWeeks += 1
                Catch ex As Exception
                    Dim x = 1
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

    Public Function CalculateRelativeStrengthJSON(ByRef strSymbol As String) As Boolean
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

        Dim strLatestClosingPrice As String = String.Empty
        Dim decClosingPriceTotal As Decimal = 0
        Dim decClosingPriceTemp As Decimal = 0
        Dim intNumberOfWeeks As Integer = 26

        Dim strURI = "https://query1.finance.yahoo.com/v7/finance/spark?symbols=" & strSymbol & "&range=26wk&interval=1wk&indicators=close&includeTimestamps=false&includePrePost=false&corsDomain=finance.yahoo.com&.tsrc=finance"
        Dim strResponseString As String = GetYahooAPIData(strURI)
        If strResponseString = "False" Then
            'Continue For
        End If

        Try
            Dim myCorrectContractOptionData As QuickType.HistoricalYahooClosePrices = QuickType.HistoricalYahooClosePrices.FromJson(strResponseString)
            Dim y = 1

            If myCorrectContractOptionData.Spark.Error = Nothing Then
                For Each myPrice In myCorrectContractOptionData.Spark.Result(0).Response(0).Indicators.Quote(0).Close
                    If IsNumeric(myPrice) = True Then
                        decClosingPriceTotal += myPrice

                        strLatestClosingPrice = myPrice
                    End If
                Next
            End If

        Catch ex As Exception
            Dim x = 1
        End Try

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
    Private Shared Function ValidateCertificate(ByVal sender As Object, ByVal certificate As X509Certificate, ByVal chain As X509Chain, ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function
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

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

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
