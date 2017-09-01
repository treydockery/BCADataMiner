Public Class PercentChangeData
    Public Property strQuarterName() As String
    Public Property strQuarterPreviousName() As String
    Public Property strQuarterEndDate() As String
    Public Property strQuarterPercentInc() As String
    Public Property bitBeatParentSymbol() As Boolean
End Class

Public Class PercentChangeDataYearly
    Public Property strYearName() As String
    Public Property strYearPreviousName() As String
    Public Property strYearEndDate() As String
    Public Property strYearPercentInc() As String
    Public Property bitBeatParentSymbol() As Boolean
End Class

Public Class PercentChangeSymbolList
    Public Property strSymbol() As String
    Public Property strParentSymbol() As String
    Public lstPercentChangeData As New List(Of PercentChangeData)
End Class

Public Class PercentChangeSymbolListYearly
    Public Property strSymbol() As String
    Public Property strParentSymbol() As String
    Public lstPercentChangeData As New List(Of PercentChangeDataYearly)
End Class


