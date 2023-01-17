Private Sub CommandButton1_Click()
Dim CTrk As Worksheet, CLog As Worksheet

Set CTrk = Sheet1
Set CLog = Sheet2



Dim x As Range, xx As Range, xxx As Range, xxxx As Range
Dim xxxxx As Range, xxxxxx As Range, xxxxxxx As Range

Set x = CTrk.Range("D8")
Set xx = CTrk.Range("G8")
Set xxx = CTrk.Range("D11")
Set xxxx = CTrk.Range("G11")
Set xxxxx = CTrk.Range("J8")
Set xxxxxx = CTrk.Range("J11")
Set xxxxxxx = CTrk.Range("G14")

Dim xxxxxxxx As Range

If CLog.Range("A2") = "" Then
    Set xxxxxxxx = CLog.Range("A2")
Else
End If


If x = "" Then
    MsgBox "what do u want"
    Exit Sub
End If

x.Copy xxxxxxxx
xxxxx.Copy xxxxxxxx.Offset(0, 1)
xxxxxx.Copy xxxxxxxx.Offset(0, 2)
xx.Copy xxxxxxxx.Offset(0, 3)
xxx.Copy xxxxxxxx.Offset(0, 4)
xxxx.Copy xxxxxxxx.Offset(0, 5)
xxxxxxx.Copy xxxxxxxx.Offset(0, 6)


x.ClearContents
xx.ClearContents
xxx.ClearContents
xxxx.ClearContents
xxxxx.ClearContents
xxxxxx.ClearContents
xxxxxxx.ClearContents
End Sub
