VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Returns a new date offset from the given date.
' The "Offset" argument must start with a number (positive for future dates, negative for
' past dates) and end with one of "Y" (year), "Q" (quarter), "M" (month) or "W" (week).
Public Function Offset(ByVal DateTime As Date, ByVal OffsetVal As Variant) As Date
    
    If VBA.IsNumeric(OffsetVal) Then
        Let Offset = DateTime + CDbl(OffsetVal)
        Exit Function
    End If
    
    Dim offsetUnit As String
    offsetUnit = VBA.UCase$(VBA.Right$(CStr(OffsetVal), 1))
    
    Dim offsetValue As Double
    offsetValue = CDbl(VBA.Replace(CStr(OffsetVal), offsetUnit, vbNullString))
    
    Select Case offsetUnit
        Case "Y"
            Let Offset = DateTime + (360 * offsetValue)
        Case "Q"
            Let Offset = QuarterEndOffset(DateTime, offsetValue)
        Case "M"
            Let Offset = DateTime + (30 * offsetValue)
        Case "W"
            Let Offset = DateTime + (7 * offsetValue)
        Case Else
            Throw RuntimeError, "Dates.Offset", "Invalid date offset unit"
    End Select
    
End Function

' Returns the Fiscal Year as an Integer.
'
' Can be given an offset (# of years). If no date is given then "now" is assumed.
'
' @param Date AsOfDate [Optional]
' @param Long YearOffset [Optional]
' @return Integer
Public Function Year(Optional ByVal AsOfDate As Date, Optional YearOffset As Long) As Integer

    If VBA.IsEmpty(AsOfDate) Then AsOfDate = VBA.Now
    AsOfDate = VBA.DateAdd("yyyy", YearOffset, AsOfDate)
    Year = VBA.Year(AsOfDate)
End Function

' Returns the number of the current fiscal quarter (1-4) for a given date as an Integer.
'
' If no date is given then "now" is assumed.
'
' @param Date AsOfDate [Optional]
' @return Integer
Public Function QuarterNum(Optional ByVal AsOfDate As Date) As Integer

    If VBA.IsEmpty(AsOfDate) Then AsOfDate = VBA.Now
    Select Case VBA.Month(AsOfDate) Mod 12
        Case 1, 2, 3
            QuarterNum = 1
        Case 4, 5, 6
            QuarterNum = 2
        Case 7, 8, 9
            QuarterNum = 3
        Case 10, 11, 12, 0 'Zero case handles 12 mod 12
            QuarterNum = 4
    End Select
End Function

' Returns the last completed quarter number (1-4) as an Integer based on a given date.
'
' If no date is given then "now" is assumed.
'
' @param Date AsOfDate [Optional]
' @return Integer
Public Function LastCompleteQuarterNum(Optional ByVal AsOfDate As Date) As Integer

    If VBA.IsEmpty(AsOfDate) Then AsOfDate = Now
    Select Case VBA.Month(AsOfDate)
        Case 10, 11, 12
            LastCompleteQuarterNum = 3
        Case 1, 2, 3
            LastCompleteQuarterNum = 4
        Case 4, 5, 6
            LastCompleteQuarterNum = 1
        Case 7, 8, 9
            LastCompleteQuarterNum = 2
    End Select
End Function

' Returns the end date of the quarter in which the given date falls.
'
' If no date is given then "Now" is assumed.
'
' @param Date AsOfDate [Optional]
' @return Date
Public Function QuarterEnd(Optional ByVal AsOfDate As Variant) As Variant

    If VBA.IsMissing(AsOfDate) Or VBA.IsEmpty(AsOfDate) Then AsOfDate = Now
    Select Case QuarterNum(AsOfDate)
        Case 1
            QuarterEnd = VBA.DateValue("3/31/" & VBA.Year(AsOfDate))
        Case 2
            QuarterEnd = VBA.DateValue("6/30/" & VBA.Year(AsOfDate))
        Case 3
            QuarterEnd = VBA.DateValue("9/30/" & VBA.Year(AsOfDate))
        Case 4
            QuarterEnd = VBA.DateValue("12/31/" & VBA.Year(AsOfDate))
        Case Else
            QuarterEnd = CVErr(xlErrNA)
    End Select
End Function

' Returns the end date of the last complete quarter as of the given date.
'
' If no date is given then "Now" is assumed.
'
' @param Date AsOfDate [Optional]
' @return Date
Public Function LastCompleteQuarterEnd(Optional ByVal AsOfDate As Date) As Variant
    
    If VBA.IsEmpty(AsOfDate) Then AsOfDate = Now
    Select Case LastCompleteQuarterNum(AsOfDate)
        Case 1
            LastCompleteQuarterEnd = VBA.DateValue("3/31/" & VBA.Year(AsOfDate))
        Case 2
            LastCompleteQuarterEnd = VBA.DateValue("6/30/" & VBA.Year(AsOfDate))
        Case 3
            LastCompleteQuarterEnd = VBA.DateValue("9/30/" & VBA.Year(AsOfDate))
        Case 4
            LastCompleteQuarterEnd = VBA.DateValue("12/31/" & VBA.Year(AsOfDate) - 1)
        Case Else
            LastCompleteQuarterEnd = CVErr(xlErrNA)
    End Select
End Function

Function QuarterEndOffset(AsOfDate As Date, OffsetNumQuarters As Double) As Variant
    QuarterEndOffset = LastCompleteQuarterEnd(AsOfDate + (OffsetNumQuarters * 90))
End Function
