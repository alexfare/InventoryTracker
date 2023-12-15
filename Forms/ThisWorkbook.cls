VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet
    Dim WorkBookCount As Integer
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws

    WorkBookCount = ws.Range("B47")
    WorkBookPlusOne = WorkBookCount + 1
    ws.Range("B47") = WorkBookPlusOne
    ws.Range("B52") = "" 'Clears logged in user.
    ws.Range("B55") = "1" 'Clears Persistent Login From Last Session
    ThisWorkbook.Save
End Sub

