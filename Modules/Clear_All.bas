Attribute VB_Name = "Clear_All"
Sub Clear_Run()
    Clear_Admin
    Clear_Credentials
    Clear_AuditLog
    Delete_Rows
    Clear_completed
End Sub

Sub Delete_Rows()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim List_Select
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Long
    ws.AutoFilterMode = False
    
    For i = 999 To 3 Step -1
        ws.Rows(i).EntireRow.Delete
    Next i
    MsgBox "Cleared Rows."
End Sub

Sub Clear_Admin()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim List_Select
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Integer
    
    For i = 2 To 999
        ws.Range("B" & i).ClearContents
    Next i
    MsgBox "Cleared Admin Settings."
End Sub

Sub Clear_Credentials()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim List_Select
    List_Select = "Credentials"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Integer
    
    For i = 999 To 3 Step -1
        ws.Rows(i).EntireRow.Delete
    Next i
    MsgBox "Cleared Credentials."
End Sub

Sub Clear_AuditLog()
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim ws As Worksheet
Dim List_Select
    List_Select = "Audit"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim i As Integer
    
    For i = 2 To 999
        ws.Range("A" & i).ClearContents
    Next i
    MsgBox "Cleared Audit Logs."
End Sub

Sub Clear_completed()
    MsgBox "Formatting has completed."
End Sub
