VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateAccount 
   Caption         =   "Create Account"
   ClientHeight    =   8460.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   OleObjectBlob   =   "CreateAccount.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim r As Long ' variable used for storing row number
Dim Worksheet_Set ' variable used for selecting and storing the active worksheet
Dim PassMatch As Boolean

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
inputUser.SetFocus
End Sub

Private Sub inputPass_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnSubmit_Click
    End If
End Sub

Private Sub btnSubmit_Click()
    If inputUser <> "" And inputPass <> "" And inputPassx2 <> "" Then
        PasswordMatch
    Else
        MsgBox "Username And Password Inputs Cannot Be Blank.", vbExclamation
    End If
End Sub

Private Sub PasswordMatch()
    If inputPass = inputPassx2 Then
        btnCreate_Click
    Else
        MsgBox "Passwords Do Not Match.", vbInformation, "Error"
    End If
End Sub

Private Sub btnCreate_Click()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Credentials" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    If IsError(Application.Match(IIf(IsNumeric(inputUser), Val(inputUser), inputUser), ws.Columns(1), 0)) Then
        
        Dim lLastRow As Long ' lLastRow = variable to store the result of the row count calculation
        lLastRow = ws.ListObjects.Item(1).ListRows.Count
        r = lLastRow + 2 ' Add number for every header tab created
        Dim gnString As String
        If IsNumeric(inputUser) Then
            gnString = inputUser.Value
        Else
            gnString = inputUser
        End If
        
        '/ Hash /'
        s = inputPass
        
        Dim sIn As String, sOut As String, b64 As Boolean
        Dim sH As String, sSecret As String
        
        'Password to be converted
        sIn = s
        sSecret = "" 'secret key for StrToSHA512Salt only
        
        b64 = True 'output base-64
        
        sH = SHA512(sIn, b64)
        
        Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
        savePass = sH
        '/ Hash /'
        
        ws.Cells(r, "A") = gnString
        ws.Cells(r, "B") = savePass
        ws.Cells(r, "C") = userName
        ws.Cells(r, "D") = userPhone
        ws.Cells(r, "E") = userAddress
        ws.Cells(r, "F") = userPosition
        ws.Cells(r, "G") = userEmail
        
        '/Add to Users count/'
        Dim AddUser As Integer
        
        List_Select = "Admin"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        
        AddUser = ws.Range("B51")
        AddUserPlusOne = AddUser + 1
        ws.Range("B51") = AddUserPlusOne
        
        '/Prevent Issues in the future, Call back the Credentials page/'
        List_Select = "Credentials"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        
        '/Status/'
        statusLabel_fix.Caption = "Status:"
        statusLabelLog.Caption = "Creating Account..."
        Status
        MsgBox ("Account Created."), , "Duplicate"
        Clear_Form
    Else
        ErrMsg_Duplicate
    End If
End Sub

Sub ErrMsg()
    MsgBox ("Username Not Found."), , "Not Found"
End Sub

Sub ErrMsg_Duplicate()
    MsgBox ("Username Taken."), , "Duplicate"
End Sub

Private Sub Clear_Form()
    inputUser = ""
    inputPass = ""
    inputPassx2 = ""
    userName = ""
    userPhone = ""
    userAddress = ""
    userPosition = ""
    userEmail = ""
End Sub

Private Sub btnBack_click()
    Unload CreateAccount
    AdminForm.Show
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
