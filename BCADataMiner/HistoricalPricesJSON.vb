
Imports System
Imports System.Net
Imports System.Collections.Generic
Imports Newtonsoft.Json
Imports System.Runtime.CompilerServices

Namespace QuickType

    Partial Public Class HistoricalYahooClosePrices

        <JsonProperty("spark")>
        Public Property Spark As Spark
    End Class

    Partial Public Class Spark

        <JsonProperty("error")>
        Public Property [Error] As Object

        <JsonProperty("result")>
        Public Property Result As Result()
    End Class

    Partial Public Class Result

        <JsonProperty("response")>
        Public Property Response As Response()

        <JsonProperty("symbol")>
        Public Property Symbol As String
    End Class

    Partial Public Class Response

        <JsonProperty("indicators")>
        Public Property Indicators As Indicators

        <JsonProperty("meta")>
        Public Property Meta As Meta

        <JsonProperty("timestamp")>
        Public Property Timestamp As Long()
    End Class

    Partial Public Class Meta

        <JsonProperty("chartPreviousClose")>
        Public Property ChartPreviousClose As Double

        <JsonProperty("currency")>
        Public Property Currency As String

        <JsonProperty("currentTradingPeriod")>
        Public Property CurrentTradingPeriod As CurrentTradingPeriod

        <JsonProperty("dataGranularity")>
        Public Property DataGranularity As String

        <JsonProperty("exchangeName")>
        Public Property ExchangeName As String

        <JsonProperty("exchangeTimezoneName")>
        Public Property ExchangeTimezoneName As String

        <JsonProperty("firstTradeDate")>
        Public Property FirstTradeDate As Long

        <JsonProperty("gmtoffset")>
        Public Property Gmtoffset As Long

        <JsonProperty("instrumentType")>
        Public Property InstrumentType As String

        <JsonProperty("symbol")>
        Public Property Symbol As String

        <JsonProperty("timezone")>
        Public Property Timezone As String

        <JsonProperty("validRanges")>
        Public Property ValidRanges As String()
    End Class

    Partial Public Class CurrentTradingPeriod

        <JsonProperty("post")>
        Public Property Post As Pre

        <JsonProperty("pre")>
        Public Property Pre As Pre

        <JsonProperty("regular")>
        Public Property Regular As Pre
    End Class

    Partial Public Class Pre

        <JsonProperty("end")>
        Public Property [End] As Long

        <JsonProperty("gmtoffset")>
        Public Property Gmtoffset As Long

        <JsonProperty("start")>
        Public Property Start As Long

        <JsonProperty("timezone")>
        Public Property Timezone As String
    End Class

    Partial Public Class Indicators

        <JsonProperty("adjclose")>
        Public Property Adjclose As Adjclose()

        <JsonProperty("quote")>
        Public Property Quote As Quote()

        <JsonProperty("unadjclose")>
        Public Property Unadjclose As Unadjclose()
    End Class

    Partial Public Class Unadjclose

        <JsonProperty("unadjclose")>
        Public Property PurpleUnadjclose As Double()
    End Class

    Partial Public Class Quote

        <JsonProperty("close")>
        Public Property Close As Double()
    End Class

    Partial Public Class Adjclose

        <JsonProperty("adjclose")>
        Public Property PurpleAdjclose As Double()
    End Class

    Partial Public Class HistoricalYahooClosePrices

        Public Shared Function FromJson(ByVal json As String) As HistoricalYahooClosePrices
            Return JsonConvert.DeserializeObject(Of HistoricalYahooClosePrices)(json, Converter.Settings)
        End Function
    End Class

    Module Serialize

        <Extension()>
        Function ToJson(ByVal self As HistoricalYahooClosePrices) As String
            Return JsonConvert.SerializeObject(self, Converter.Settings)
        End Function
    End Module

    Public Class Converter

        Public Shared ReadOnly Settings As JsonSerializerSettings = New JsonSerializerSettings With {.MetadataPropertyHandling = MetadataPropertyHandling.Ignore, .DateParseHandling = DateParseHandling.None}
    End Class
End Namespace