VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangePassword 
   Caption         =   "Change Password"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3645
   OleObjectBlob   =   "ChangePassword.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Long        ' variable used for storing row number
Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
Dim btnUpdate_Enable As Boolean        ' to store update enable flag after search
Dim GN_Verify

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub inputPass_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnChangePassword_Click
    End If
End Sub

Public Sub Search_Button_Click()
    Dim ws As Worksheet
    
    List_Select = "Credentials"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(inputUser), Val(inputUser), inputUser), ws.Columns(1), 0)) Then
        btnUpdate_Enable = False
        ErrMsg
    Else
        r = Application.Match(IIf(IsNumeric(inputUser), Val(inputUser), inputUser), ws.Columns(1), 0)
        GN_Verify = inputUser
        btnUpdate_Enable = True
        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Searching..."
        Status
        btnUpdate
    End If
End Sub

Private Sub btnChangePassword_Click()
    If inputUser <> "" And inputPass <> "" Then
        Search_Button_Click
    Else
        MsgBox "Please provide Username & Password.", vbExclamation
    End If
End Sub

Private Sub btnUpdate()
    If btnUpdate_Enable = True Then
        If GN_Verify = inputUser Then
            Update_Worksheet
        Else
            MSG_Verify_Update
        End If
    Else
        ErrMsg_Search
    End If
End Sub

Private Sub Clear_Form()
    inputUser = ""
    inputPass = ""
End Sub

Private Sub Update_Worksheet()
    If btnUpdate_Enable = True Then
        Dim gnString As String
        Set ws = Worksheet_Set
        If IsNumeric(inputUser) Then
            gnString = Val(inputUser.Value)
        Else
            gnString = inputUser
        End If
        
        '/ Hash /'
        s = inputPass
        
        Dim sIn As String, sOut As String, b64 As Boolean
        Dim sH As String, sSecret As String
        
        'Password to be converted
        sIn = s
        sSecret = ""
        
        b64 = True
        
        sH = SHA512(sIn, b64)
        
        savePass = sH
        '/ Hash /'
        
        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = savePass
        
        '/Status/'
        statusLabel.Caption = "Status:"
        statusLabelLog.Caption = "Password Updated!"
        Status
        Clear_Form
    Else
        ErrMsg_Search
    End If
End Sub

Sub MSG_Verify_Update()
    MSG1 = MsgBox("Are you sure you want To change the Username?", vbYesNo, "Verify")
    If MSG1 = vbYes Then
        Update_Worksheet
    Else
        inputUser = GN_Verify
    End If
End Sub

Private Sub btnBack_click()
    Unload ChangePassword
    AdminForm.Show
End Sub

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
        statusLabel.Caption = ""
        statusLabelLog.Caption = ""
End Sub

'/ ------- Error Handles ------- /'
Sub ErrMsg()
    MsgBox ("Username Not Found."), , "Not Found"
End Sub

Sub ErrMsg_Duplicate()
    MsgBox ("Username already in use."), , "Duplicate"
End Sub

Sub ErrMsg_Search()
    MsgBox ("Must search for entry before updating."), vbInformation, "Error"
    Clear_Form
End Sub
