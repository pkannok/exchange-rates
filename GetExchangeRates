Option Explicit

Sub GetExchangeRates()
'    Dim DataSheet As Worksheet
    Dim DataBook As Workbook
    Dim sDate As String
    Dim eDate As String
    Dim str As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
'    Application.Calculation = xlCalculationManual

    Sheets.Add.Name = "xRates"

'    Set DataSheet = ActiveSheet
    Set DataBook = ActiveWorkbook

    ' Rolling average daily rate for past 5 weeks: start date = 5 weeks before yesterday; end date = yesterday
    sDate = Format(Now - (5 * 7) - 1, "yyyy-mm-dd")
    eDate = Format(Now - 1, "yyyy-mm-dd")

    ' http://www.oanda.com/currency/historical-rates/download?quote_currency=USD&end_date=2016-05-19&start_date=2016-04-14&period=daily&display=absolute&rate=0&data_range=c&price=mid&view=table&base_currency_0=CAD&base_currency_1=EUR&base_currency_2=AUD&base_currency_3=KRW&base_currency_4=JPY&base_currency_5=SGD&download=csv
    str = "http://www.oanda.com/currency/historical-rates/download?quote_currency=USD&end_date=" _
        & eDate & "&start_date=" & sDate _
        & "&period=daily&display=absolute&rate=0&data_range=c&price=mid&view=table&base_currency_0=CAD&base_currency_1=EUR&base_currency_2=AUD&base_currency_3=KRW&base_currency_4=JPY&base_currency_5=SGD&download=csv"

QueryQuote:
    With Sheets("xRates").QueryTables.Add(Connection:="URL;" & str, Destination:=Sheets("xRates").Range("A1"))
        .BackgroundQuery = True
        .TablesOnlyFromHTML = False
        .Refresh BackgroundQuery:=False
        .SaveData = True
    End With

    Sheets("xRates").Range("A5:A22").CurrentRegion.TextToColumns Destination:=Sheets("xRates").Range("A5:A22"), DataType:=xlDelimited, _
                                                           TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                                                           Semicolon:=False, Comma:=True, Space:=False, other:=True, OtherChar:=",", FieldInfo:=Array(1, 2)

    With Sheets("xRates")
    ' Inverse the average daily rate and apply names
        With .Range("B4")
            .Value = 1 / Application.Average(Sheets("xRates").Range("B6:B41"))
            .Name = "cadToUsd"
        End With
        With .Range("C4")
            .Value = 1 / Application.Average(Sheets("xRates").Range("C6:C41"))
            .Name = "eurToUsd"
        End With
        With .Range("D4")
            .Value = 1 / Application.Average(Sheets("xRates").Range("D6:D41"))
            .Name = "audToUsd"
        End With
        With .Range("E4")
            .Value = 1 / Application.Average(Sheets("xRates").Range("E6:E41"))
            .Name = "krwToUsd"
        End With
        With .Range("F4")
            .Value = 1 / Application.Average(Sheets("xRates").Range("F6:F41"))
            .Name = "jpyToUsd"
        End With
        With .Range("G4")
            .Value = 1 / Application.Average(Sheets("xRates").Range("G6:G41"))
            .Name = "sgdToUsd"
        End With

        ' Label rates on sheet
        .Range("B3").Value = "CAD/USD"
        .Range("C3").Value = "EUR/USD"
        .Range("D3").Value = "AUD/USD"
        .Range("E3").Value = "KRW/USD"
        .Range("F3").Value = "JPY/USD"
        .Range("G3").Value = "SGD/USD"
    End With


    Application.DisplayAlerts = True
    
    Call KillConnections

End Sub

Sub KillConnections()
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Connections.Count
    If ActiveWorkbook.Connections.Count = 0 Then Exit Sub
    ActiveWorkbook.Connections.Item(i).Delete
    i = i - 1
    Next i
End Sub
