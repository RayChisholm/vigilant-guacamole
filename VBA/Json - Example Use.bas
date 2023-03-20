''
' This code utilizes the JsonConverter library in this same repo.
' It pulls the JSON data from the pastebin link included
' And then parses it into a range on a designated sheet
' This should be used as an example/framework for intended uses
''


Option Explicit

Sub Test()

    Dim sJSONString As String
    Dim vJSON
    Dim sState As String
    Dim aData()
    Dim aHeader()

    ' Retrieve JSON content
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", "https://pastebin.com/raw/hA2UEDXy", True
        .send
        Do Until .readyState = 4: DoEvents: Loop
        sJSONString = .responseText
    End With
    ' Parse JSON sample
    JSON.Parse sJSONString, vJSON, sState
    If sState = "Error" Then MsgBox "Invalid JSON": End
    ' Convert JSON to 2D Array
    JSON.ToArray vJSON("AppointmentList"), aData, aHeader
    ' Output to worksheet #1
    Output aHeader, aData, ThisWorkbook.Sheets(1)
    MsgBox "Completed"

End Sub

Sub Output(aHeader, aData, oDestWorksheet As Worksheet)

    With oDestWorksheet
        .Activate
        .Cells.Delete
        With .Cells(1, 1)
            .Resize(1, UBound(aHeader) - LBound(aHeader) + 1).Value = aHeader
            .Offset(1, 0).Resize( _
                    UBound(aData, 1) - LBound(aData, 1) + 1, _
                    UBound(aData, 2) - LBound(aData, 2) + 1 _
                ).Value = aData
        End With
        .Columns.AutoFit
    End With

End Sub
