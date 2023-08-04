VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminForm 
   Caption         =   "Admin Panel  - Created By Alex Fare"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "AdminForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long ' variable used for storing row number
Dim Worksheet_Set ' variable used for selecting and storing the active worksheet
Dim Update_Button_Enable As Boolean ' to store update enable flag after search
Dim GN_Verify
Dim currrentUser    As String
Dim rlStatus As Integer

Private Sub UserForm_Initialize()
    '/ Display Admin Audit Log/'
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim ws          As Worksheet
    Dim List_Select
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    txtWorkbookOpened = ws.Range("B47")
    txtLogins = ws.Range("B48")
    txtUserCounts = ws.Range("B51")
    lblLoggedUser = ws.Range("B52")
    
    '/Prevent Issues in the future, Call back the main page/'
    List_Select = "CreatedByAlexFare"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
End Sub

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

'/ Pressing Enter will instantly search /'
Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Button_Click
    End If
End Sub

Public Sub Search_Button_Click()
    ' clear previous data from form, except "Gage Number"
    ' --------------------------------------------------------
    Dim Gage_Number_Save
    Gage_Number_Save = Gage_Number
    Clear_Form
    Gage_Number = Gage_Number_Save
    ' ---------------------------------------------------------
    
    Dim ws          As Worksheet
    
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
        
        '/ Audit Log
        DateEdit = ws.Cells(r, "AM") 'Update Last searched
        ws.Cells(r, "AM") = Now        'Update Last searched
        lblSearchedDate = DateEdit 'Update Last searched
        lblDateEdit = ws.Cells(r, "AL")
        lastUser = ws.Cells(r, "AN")
        lblReceivedIn = ws.Cells(r, "T")
        lblOrderEntry = ws.Cells(r, "S")
        lblUsageReport = ws.Cells(r, "R")
        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Searching"
        Status

        Update_Button_Enable = True
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
        '/ Audit
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
        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Updated"
        Status
        
    Else
        MsgBox ("Must search For entry before updating"), , "Nothing To Update"
        
    End If
    
    'Update_Button_Enable = False 'Remove ' if you want to require searching again after an update.
    
End Sub

Sub MSG_Verify_Update()
    
    MSG1 = MsgBox("Are you sure you want To change the Gage ID?", vbYesNo, "Verify")
    
    If MSG1 = vbYes Then
        Update_Worksheet
    Else
        Gage_Number = GN_Verify
    End If
    
End Sub

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

Private Sub btnClear_Click()
    Update_Button_Enable = False
    Clear_Form
End Sub

Sub CheckForUpdate_Click()
    Dim url         As String
    url = "https://github.com/alexfare/GageCalibrationTracker"
    ActiveWorkbook.FollowHyperlink url
End Sub

Private Sub btnClose_Click()
    Unload AdminForm
    
    '/Save Logged In User For The Session /'
    List_Select = "Admin"        ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    ws.Range("B55") = "2"       ' 1 = Required | 2 = Not Required
End Sub

Private Sub btnCreateAccount_click()
    Unload AdminForm
    CreateAccount.Show
End Sub

Private Sub btnUpdateUser_click()
    Unload AdminForm
    ChangePassword.Show
End Sub

Private Sub btnDevMode_click()
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
End Sub

Private Sub btnAbout_Click()
    MsgBox "Code protection password Is GageTracker2022"
End Sub

Private Sub auditBTN_Click()
    Audit.Show
End Sub

Private Sub btnFormat_Click()
    Unload AdminForm
    Worksheets("Format").Activate
End Sub

Private Sub btnLogout_Click()
        List_Select = "Admin"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        ws.Range("B55") = "1"
        Unload AdminForm
End Sub

