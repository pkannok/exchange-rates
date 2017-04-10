Option Explicit

Sub GetExchangeRates()
'    Dim DataSheet As Worksheet
    Dim DataBook As Workbook
    Dim rateSheet As String
    Dim http As Object, apiResponse As String, allRates() As String, oneRate() As String, i As Integer, rowNum As Integer
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
'    Application.Calculation = xlCalculationManual

    'Delete old rate sheet and create new blank sheet as last sheet
    Set DataBook = ActiveWorkbook
    rateSheet = "xRates"

    On Error Resume Next
    Sheets(rateSheet).Delete
    On Error GoTo 0

    Sheets.Add.Name = rateSheet
    Sheets(rateSheet).Move after:=Worksheets(Worksheets.Count)

    Set http = CreateObject("MSXML2.XMLHTTP") 'Tools > References: Add Microsoft XML
    http.Open "GET", "http://api.fixer.io/latest?base=USD&symbols=AUD,CAD,CNY,EUR,JPY,KRW,SGD", False 'http://fixer.io/
    http.Send
    
    apiResponse = Replace(Replace(Replace(Replace(http.responseText, Chr(34), ""), "{", ""), "}", ""), "rates:", "")
    allRates = Split(apiResponse, ",")
    
    Sheets(rAtesheet).Cells.Clear
    Sheets(rAtesheet).Range("A1").Value = "Current daily exchange rates from the European Central Bank via fixer.io"
    Sheets(rAtesheet).Range("C3").Value = "Local to USD"
    
    For i = LBound(allRates) To UBound(allRates)
        rowNum = i + 3
        oneRate = Split(allRates(i), ":")
        Sheets(rAtesheet).Range("A" & rowNum).Value = oneRate(0)
        Sheets(rAtesheet).Range("B" & rowNum).Value = oneRate(1)
        If rowNum > 4 Then
            With Sheets(rAtesheet).Range("C" & rowNum)
                .Value = 1 / oneRate(1)
                .Name = LCase(Sheets(rAtesheet).Range("A" & rowNum).Value) & "ToUsd"
            End With
        End If
    Next

    Application.DisplayAlerts = True

End Sub

