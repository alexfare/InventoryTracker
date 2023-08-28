VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "InventoryTracker - Created By Alex Fare"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9420.001
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Created By: Alex Fare

Dim r As Long        ' variable used for storing row number
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean        ' to store update enable flag after search
Dim GN_Verify
Dim currrentUser As String
Dim ActionLog As String
Dim AuditTime As String
Dim AuditUser As String

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'

    Dim Worksheet_Set
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "CreatedByAlexFare"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    vDisplay = ws.Range("D1")
End Sub

Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Button_Click
    End If
End Sub

'/------- Search Button -------/'
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
        
        '/Usage Page/'
        UsageProduct = Gage_Number
        UsageCurrent = inventoryTxt
        UsageOnOrder = onOrder
        
        '/ Audit Log
        DateEdit = ws.Cells(r, "AM") 'Update Last searched
        ws.Cells(r, "AM") = Now        'Update Last searched
        lblSearchedDate = DateEdit 'Update Last searched
        lblDateEdit = ws.Cells(r, "AL")
        lastUser = ws.Cells(r, "AN")
        lblReceivedIn = ws.Cells(r, "T")
        lblOrderEntry = ws.Cells(r, "S")
        lblUsageReport = ws.Cells(r, "R")
        
        ActionLog = "Searched"
        AuditTime = Now
        AuditUser = Application.userName
        auditLog
                
        '/Status/'
        statusLabel_fix.Caption = "Status:"
        statusLabelLog.Caption = "Searching..."
        Status
        
        '/Enables Edit/'
        Update_Button_Enable = True
    End If
End Sub

'/ ------- Error Handles ------- /'
Sub ErrMsg()
    MsgBox ("Gage Number Not Found"), , "Not Found"
End Sub

Sub ErrMsg_Duplicate()
    MsgBox ("Gage number already in use"), , "Duplicate"
End Sub

'/ ------- Clear Button ------- /'
Private Sub Clear_Form()
    Gage_Number = ""
    Descriptiontxt = ""
    inventoryTxt = ""
    onOrder = ""
    
    '/ReceiveIn/'
    txtProduct = ""
    txtCurrentQty = ""
    receiveInput = ""
    
    '/OnOrder/'
    OrderProductTxt = ""
    currentInventory = ""
    currentOnOrdertxt = ""
    orderQty = ""
    
    '/Usage/'
    UsageProduct = ""
    UsageCurrent = ""
    UsageOnOrder = ""
    txtUse = ""
    
    '/Audit Log/'
    lastUser = ""
    lblDateEdit = ""
    lblSearchedDate = ""
    lblUsageReport = ""
    lblOrderEntry = ""
    lblReceivedIn = ""
End Sub

'/------- Receive In Button -------/'
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
        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = Descriptiontxt
        ws.Cells(r, "C") = inventoryTxt
        ws.Cells(r, "D") = onOrder
        
        '/ Audit Log
        currrentUser = Application.userName
        lastUser = currrentUser
        ws.Cells(r, "AN") = lastUser
        ws.Cells(r, "AL") = lblDateEdit
        lblReceivedIn = Now
        ws.Cells(r, "T") = lblReceivedIn
    
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
    
    ActionLog = "Received In"
    AuditTime = Now
    AuditUser = Application.userName
    auditLog
    '/ End Audit Log /'
    
    '/Status/'
        statusLabel_fix.Caption = "Status:"
        statusLabelLog.Caption = "Receiving In " + receiveInput + " " + gnString
        Status
        AutoSave
        
    '/Update Menu
    receiveInput = ""
    Search_Button_Click
Else
    MsgBox ("Must search For entry before updating"), , "Nothing To Update"
End If
End Sub

Sub MSG_Verify_Update()
        MSG1 = MsgBox("Are you sure you want To change the Product ID?", vbYesNo, "Verify")
    If MSG1 = vbYes Then
        Update_Worksheet
    Else
        Gage_Number = GN_Verify
    End If
End Sub

'/ Clear Button
Private Sub btnClear_Click()
    Clear_Form
End Sub

Private Sub btnSave_click()
    ThisWorkbook.Save
    
    '/Status/'
        statusLabel_fix.Caption = "Status:"
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

'/------- Report Issue Panel ------- /'
Private Sub btnReportIssue_click()
    Unload Menu
    ReportIssue.Show
End Sub

'/------- Display Status -------/'
Private Sub Status()
    Dim startTime As Date
    Dim elapsedTime As Long
    Dim waitTimeInSeconds As Long
        
    waitTimeInSeconds = 2 'change this to the desired wait time in seconds
    
    startTime = Now
    Do While elapsedTime < waitTimeInSeconds
        DoEvents 'allow the program to process any pending events
        elapsedTime = DateDiff("s", startTime, Now)
    Loop
        statusLabel_fix.Caption = ""
        statusLabelLog.Caption = ""
End Sub

Private Sub AutoSave()
    ThisWorkbook.Save
    statusLabel_fix.Caption = "Status:"
    statusLabelLog.Caption = "Auto-Saving..."
    Status
End Sub

'/ ------- On-Order Tab ------- /'
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
        lblOrderEntry = Now
        ws.Cells(r, "S") = lblOrderEntry
    
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
    
    ActionLog = "Order Entry"
    AuditTime = Now
    AuditUser = Application.userName
    auditLog
    '/ End Audit Log /'
    
    '/Status/'
        statusLabel_fix.Caption = "Status:"
        statusLabelLog.Caption = "" + orderQty + " " + gnString + " Added to On-Order!"
        Status
        AutoSave
        
    '/Update Menu
    orderQty = ""
    Search_Button_Click
Else
    MsgBox ("Must search For entry before updating"), , "Nothing To Update"
End If
End Sub

'/ ------- Usage Tab ------- /'
Private Sub Usage_Button_Click()
    If Update_Button_Enable = True Then
        If GN_Verify = Gage_Number Then
            UsageSub
        Else
            MSG_Verify_Update
        End If
    Else
        MsgBox ("Must search For entry before updating"), , "Nothing To Update"
    End If
End Sub

Private Sub UsageSub()
If Update_Button_Enable = True Then
        Dim gnString As String
        Set ws = Worksheet_Set
        If IsNumeric(Gage_Number) Then
            gnString = Val(Gage_Number.Value)
        Else
            gnString = Gage_Number
        End If
    Dim InputUsageQty As String
    Dim userInputu As String
    Dim userInputi As String
    Dim ConvertUsageQty As Double
    Dim ConvertCurrentUsage As Double
    Dim convertedNumberi As Double
    Dim inventoryIntToTxt As String
    
    InputUsageQty = txtUse.Value
    ConvertUsageQty = Val(InputUsageQty)
    userInputu = txtUse.Value
    ConvertCurrentUsage = Val(userInputu)
    userInputi = inventoryTxt.Value
    convertedNumberi = Val(userInputi)
    
    inventoryTxt = convertedNumberi - ConvertCurrentUsage

        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = Descriptiontxt
        ws.Cells(r, "C") = inventoryTxt
        
        '/ Audit Log
        currrentUser = Application.userName
        lastUser = currrentUser
        ws.Cells(r, "AN") = lastUser
        ws.Cells(r, "AL") = lblDateEdit
        lblUsageReport = Now
        ws.Cells(r, "R") = lblUsageReport
    
    '/Audit Log/'
    Dim UpdateCount As Integer
    
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    UpdateCount = ws.Range("B50")
    UpdateCountPlusOne = UpdateCount + 1
    ws.Range("B50") = UpdateCountPlusOne
    
    '/Status/'
        statusLabel_fix.Caption = "Status:"
        statusLabelLog.Caption = "" + txtUse + " " + gnString + " Has been consumed.."
        Status
        
    '/Prevent Issues in the future, Call back the main page/'
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    ActionLog = "Usage Report"
    AuditTime = Now
    AuditUser = Application.userName
    'AuditStatus
    auditLog
    '/ End Audit Log /'
    
    '/Status/'
        statusLabel_fix.Caption = "Status:"
        statusLabelLog.Caption = "" + txtUse + " " + gnString + " Has been consumed.."
        Status
        AutoSave
        
    '/Update Menu
    txtUse = ""
    Search_Button_Click
Else
    MsgBox ("Must search For entry before updating"), , "Nothing To Update"
End If
End Sub

'/ ------- Audit Log ------- /'
Private Sub auditLog()
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim ws          As Worksheet
    Dim List_Select
    List_Select = "Audit"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    Dim auditLog As String
    Dim AuditAdd As String
    Dim AuditDate As String
    
    auditLog = ws.Range("A2")
    AuditDate = Now
    AuditAdd = " ||| Date: " + AuditDate + " User: " + AuditUser + " Action: " + ActionLog + " ||| "
    auditLog = auditLog + " " + AuditAdd
    ws.Range("A2") = auditLog
End Sub

Private Sub auditBTN_Click()
    Unload Menu
    Worksheets("Audit").Activate
End Sub
