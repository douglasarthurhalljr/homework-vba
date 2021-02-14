VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub moderateoption()

'set dimensions
Dim total As Double
Dim i As Long
Dim change As Single
Dim j As Integer
Dim start As Long
Dim rowcount As Long
Dim percentagechange As Single
Dim ws As Worksheet

For Each ws In Worksheets

    'set values for each worksheet
    j = 0
    total = 0
    change = 0
    start = 2
    
    'set title row
ws.Range("I1").Value = "ticker"
ws.Range("j1").Value = "yearly change"
ws.Range("k1").Value = "percent change"
ws.Range("l1").Value = "total stock volume"

'set intitial values
j = 0
total = 0
change = 0
start = 2

'get the row number with the last row with data
rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowcount

    ' if the ticker changes then print results
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
      'store results in variables
      total = total + ws.Cells(i, 7).Value
      
      'handle zero total volume
      If total = 0 Then
        'print results
        ws.Range("i" & 2 + j).Value = Cells(i, 1).Value
        ws.Range("j" & 2 + j).Value = 0
        ws.Range("k" & 2 + j).Value = "%" & 0
        ws.Range("l" & 2 + j).Value = 0
    Else
    
        'find first non zero starting value
        If ws.Cells(start, 3) = 0 Then
            For find_value = start To i
                If ws.Cells(find_value, 3).Value <> 0 Then
                    start = find_value
               Exit For
             End If
            Next find_value
           End If
           
           'calculate change
           change = (ws.Cells(i, 6) - ws.Cells(start, 3))
           percentchange = Round((change / ws.Cells(start, 3) * 100), 2)
           
           'start of the next stock ticker
           start = i + 1
           
           'print the results
           ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
           ws.Range("j" & 2 + j).Value = Round(change, 2)
           ws.Range("k" & 2 + j).Value = "%" & percentchange
           ws.Range("l" & 2 + j).Value = total
           
           'colors positive green and negatives red
            Select Case change
            Case Is > 0
                ws.Range("j" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                ws.Range("j" & 2 + j).Interior.ColorIndex = 3
            Case Else
                ws.Range("j" & 2 + j).Interior.ColorIndex = 0
            End Select
        End If
        
        'reset variables for new stock ticker
        total = 0
        change = 0
        j = j + 1
        
    'if ticker is still the same add results
    Else
        total = total + ws.Cells(i, 7).Value
    End If
    
    Next i
    
   Next ws
   
        
End Sub


