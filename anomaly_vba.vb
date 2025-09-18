Sub Open_Photos() ' Script to open multiple URLs in Chrome browser
    Dim chromeFileLocation As String
    Dim cellContent As String
    Dim urls() As String
    Dim i As Integer
    
    chromeFileLocation = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
    cellContent = ActiveSheet.Range("W8").Value 
    urls = Split(cellContent, ", ")
    
    ' Loop through each URL and open it in Chrome
    For i = LBound(urls) To UBound(urls)
        Shell (chromeFileLocation & " -url " & urls(i))
    Next i
    
End Sub

Sub Increase_Index() 'Script to increment an index
    Dim max As Integer
    
    Worksheets("Photos").Activate
    max = Range("X4")
    
    ' Check if current index (H2) has reached maximum
    If Range("H2") = max Then
        MsgBox "Last index reached, no anomaly exists with index greater than " & max
        Exit Sub
    End If
    
    Range("H2") = Range("H2") + 1
    
End Sub

Sub Decrease_Index() 'Script to decrement an index
    Dim min As Integer
    
    Worksheets("Photos").Activate
    min = Range("W4")
    
    ' Check if current index (H2) has reached minimum
    If Range("H2") = min Then
        MsgBox "First index reached, no anomaly exists with index less than " & min
        Exit Sub
    End If
    
    Range("H2") = Range("H2") - 1
    
End Sub