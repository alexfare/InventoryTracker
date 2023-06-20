VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "InventoryTracker - Created By Alex Fare"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9720.001
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Gage Tracker
' Created By: Alex Fare

Dim r As Long        ' variable used for storing row number
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean        ' to store update enable flag after search
Dim GN_Verify
Dim currrentUser As String

'/Start up script /'
Private Sub UserForm_Initialize()
'/Code Confirm for production use only/'
    Dim CodeCompare As Integer
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim LoginCount  As Integer
    Dim ws          As Worksheet
    Dim List_Select
    
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    CodeCompare = ws.Range("B56")
    If CodeCompare = "1" Then
        Unload Menu
        CodeConfirm.Show
    End If
'/ End code confirm /'

'/Prevent Issues in the future, Call back the main page/'
    List_Select = "CreatedByAlexFare"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
End Sub
Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Button_Click
        Gage_Number.SetFocus
    End If
End Sub

'/ Search Button
Public Sub Search_Button_Click()
    Dim ws          As Worksheet
    Dim DateEdit 'Update Last searched
    Dim Gage_Number_Save
    
    ' clear previous data from form, except "Gage Number"
    ' --------------------------------------------------------
    Gage_Number_Save = Gage_Number
    Clear_Form
    Gage_Number = Gage_Number_Save
    ' ---------------------------------------------------------
    
    List_Select = "CreatedByAlexFare"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)) Then
        Update_Button_Enable = False
        ErrMsg
    Else
        r = Application.Match(IIf(IsNumeric(Gage_Number), Val(Gage_Number), Gage_Number), ws.Columns(1), 0)
        GN_Verify = Gage_Number
        Descriptiontxt = ws.Cells(r, "B")
        inventoryTxt = ws.Cells(r, "C")
        onOrder = ws.Cells(r, "D")
        
        '/Receive Page/'
        txtProduct = Gage_Number
        txtCurrentQty = inventoryTxt
        
        '/Order Page/'
        OrderProductTxt = Gage_Number
        currentInventory = inventoryTxt
        currentOnOrdertxt = onOrder
        
        Update_Button_Enable = True
        Option4_Custom = True
        
        '/ Audit Log
        lblDateEdit = ws.Cells(r, "AL")
        lblSearchedDate = DateEdit 'Update Last searched
        lastUser = ws.Cells(r, "AN")
                
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Searching..."
        Status
        
    End If
    'Gage_Number.SetFocus
End Sub

'/ Update Button
Private Sub Update_Button_Click()
    If Update_Button_Enable = True Then
        If GN_Verify = Gage_Number Then
            Update_Worksheet
        Else
            MSG_Verify_Update
        End If
    Else
        MsgBox ("Must search For entry before updating"), , "Nothing To Update"
    End If
End Sub

Sub ErrMsg()
    MsgBox ("Gage Number Not Found"), , "Not Found"
    Gage_Number.SetFocus
End Sub

Sub ErrMsg_Duplicate()
    MsgBox ("Gage number already in use"), , "Duplicate"
    Gage_Number.SetFocus
End Sub

Private Sub Clear_Form()
    Gage_Number = ""
    Descriptiontxt = ""
    inventoryTxt = ""
    onOrder = ""
    txtProduct = ""
    txtCurrentQty = ""
End Sub

Private Sub Update_Worksheet()
    If Update_Button_Enable = True Then
        Dim gnString As String
        Set ws = Worksheet_Set
        If IsNumeric(Gage_Number) Then
            gnString = Val(Gage_Number.Value)
        Else
            gnString = Gage_Number
        End If
    Dim userInputr As String
    Dim userInputo As String
    Dim userInputi As String
    Dim ConvertOrderQty As Double
    Dim ConvertCurrentOrder As Double
    Dim convertedNumberi As Double
    Dim inventoryIntToTxt As String
    
    userInputr = receiveInput.Value
    ConvertOrderQty = Val(userInputr)
    userInputo = onOrder.Value
    ConvertCurrentOrder = Val(userInputo)
    userInputi = inventoryTxt.Value
    convertedNumberi = Val(userInputi)
    
    ConvertCurrentOrder = ConvertCurrentOrder - ConvertOrderQty
    convertedNumberi = convertedNumberi + ConvertOrderQty
    
    receiveInput = ConvertOrderQty
    inventoryTxt = convertedNumberi
    onOrder = ConvertCurrentOrder
    
     ' If OnOrder is negative, set it to 0
    If onOrder < 0 Then
        onOrder = 0
    End If
    
    'MsgBox receiveInput
    'MsgBox onOrder
    'MsgBox inventoryTxt

    
        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = Descriptiontxt
        ws.Cells(r, "C") = inventoryTxt
        ws.Cells(r, "D") = onOrder
        
        '/ Audit Log
        currrentUser = Application.userName
        lastUser = currrentUser
        ws.Cells(r, "AN") = lastUser
        ws.Cells(r, "AL") = lblDateEdit

    'Gage_Number.SetFocus 'Clear_Form 'Clear form after update
    
    '/Audit Log/'
    Dim UpdateCount As Integer
    
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    UpdateCount = ws.Range("B50")
    UpdateCountPlusOne = UpdateCount + 1
    ws.Range("B50") = UpdateCountPlusOne
    
    '/Prevent Issues in the future, Call back the main page/'
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    '/ End Audit Log /'
    
    '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Updating..."
        Status
        
    '/Update Menu
    Search_Button_Click
Else
    MsgBox ("Must search For entry before updating"), , "Nothing To Update"
End If

'Update_Button_Enable = False 'Remove comment if you want to require searching again after an update.

End Sub

Sub MSG_Verify_Update()
    
    MSG1 = MsgBox("Are you sure you want To change the Gage ID?", vbYesNo, "Verify")
    
    If MSG1 = vbYes Then
        Update_Worksheet
    Else
        Gage_Number = GN_Verify
    End If
    
End Sub

'/ Clear Button
Private Sub btnClear_Click()
    Update_Button_Enable = False
    Clear_Form
    Gage_Number.SetFocus
End Sub

Private Sub btnSave_click()
    ThisWorkbook.Save
    
    '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Saving..."
        Status
End Sub

Private Sub btnLogout_Click()
    Unload Menu
    Worksheets("Login").Activate
    LoginForm.Show
    ThisWorkbook.Save
End Sub

'/Admin Panel - Bring up admin menu to edit audit dates/'
Private Sub btnAdmin_click()
    '/Add to the login count /'
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim LoginCount  As Integer
    
    Dim ws          As Worksheet
    Dim List_Select
    Dim TempLogin   As Integer
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Persistent_Login = ws.Range("B55")
    
    If Persistent_Login = "1" Then
        Unload Menu
        LoginForm.Show
    End If
    
    If Persistent_Login = "2" Then
        Sheets("CreatedByAlexFare").Activate
        Unload Menu
        AdminForm.Show
    End If
End Sub

'/Report Issue Panel /'
Private Sub btnReportIssue_click()
    Unload Menu
    ReportIssue.Show
End Sub

'/Label Printing /'
Private Sub btnLabel_Click()
    Label.Show
End Sub

'/Gage R&R /'
Private Sub btnGageRR_Click()
    'MsgBox "NOTE: This is a WIP preview. Calculation formula is not displaying correctly!"
    GageRnR.Show
End Sub

'/Display Status at the bottom & Freeze/'
Private Sub Status()
    Dim startTime As Date
    Dim elapsedTime As Long
    Dim waitTimeInSeconds As Long
    
    waitTimeInSeconds = 2 'change this to the desired wait time in seconds
    
    startTime = Now
    Do While elapsedTime < waitTimeInSeconds
        DoEvents 'allow the program to process any pending events
        Application.Wait (Now + TimeValue("0:00:02"))
        elapsedTime = DateDiff("s", startTime, Now)
    Loop
        statusLabel.Caption = ""
        statusLabelLog.Caption = ""
End Sub
Private Sub AutoSave()
    ThisWorkbook.Save
End Sub

Private Sub ReceiveInSub()
End Sub

Private Sub OnOrder_Button_Click()
    If Update_Button_Enable = True Then
        If GN_Verify = Gage_Number Then
            OnOrderSub
        Else
            MSG_Verify_Update
        End If
    Else
        MsgBox ("Must search For entry before updating"), , "Nothing To Update"
    End If
End Sub

Private Sub OnOrderSub()
If Update_Button_Enable = True Then
        Dim gnString As String
        Set ws = Worksheet_Set
        If IsNumeric(Gage_Number) Then
            gnString = Val(Gage_Number.Value)
        Else
            gnString = Gage_Number
        End If
    Dim InputOrderQty As String
    Dim userInputo As String
    Dim userInputi As String
    Dim ConvertOrderQty As Double
    Dim ConvertCurrentOrder As Double
    Dim convertedNumberi As Double
    Dim inventoryIntToTxt As String
    
    InputOrderQty = orderQty.Value
    ConvertOrderQty = Val(InputOrderQty)
    userInputo = onOrder.Value
    ConvertCurrentOrder = Val(userInputo)
    userInputi = inventoryTxt.Value
    convertedNumberi = Val(userInputi)
    
    ConvertCurrentOrder = ConvertCurrentOrder + ConvertOrderQty
    
    orderQty = ConvertOrderQty
    onOrder = ConvertCurrentOrder

        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = Descriptiontxt
        ws.Cells(r, "C") = inventoryTxt
        ws.Cells(r, "D") = onOrder
        
        '/ Audit Log
        currrentUser = Application.userName
        lastUser = currrentUser
        ws.Cells(r, "AN") = lastUser
        ws.Cells(r, "AL") = lblDateEdit

    'Gage_Number.SetFocus 'Clear_Form 'Clear form after update
    
    '/Audit Log/'
    Dim UpdateCount As Integer
    
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    UpdateCount = ws.Range("B50")
    UpdateCountPlusOne = UpdateCount + 1
    ws.Range("B50") = UpdateCountPlusOne
    
    '/Prevent Issues in the future, Call back the main page/'
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    '/ End Audit Log /'
    
    '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Updating..."
        Status
        
    '/Update Menu
    Search_Button_Click
Else
    MsgBox ("Must search For entry before updating"), , "Nothing To Update"
End If
End Sub
