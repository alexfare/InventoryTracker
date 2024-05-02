VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CompanyProfile 
   Caption         =   "Company Profile"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4215
   OleObjectBlob   =   "CompanyProfile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CompanyProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Worksheet_Set       ' variable used for selecting and storing the active worksheet

'/Positioning /'
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    inputCName = ws.Range("B2")
    inputCPhone = ws.Range("B3")
    inputCAddress = ws.Range("B4")
    inputCWebsite = ws.Range("B5")
End Sub

Private Sub UserForm_Activate()
'/Positioning /'
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
'/End Positioning /'
End Sub

Private Sub btnBack_click()
    Unload Me
End Sub

Private Sub btnSubmit_Click()
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "Admin" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    
    ws.Range("B2") = inputCName
    ws.Range("B3") = inputCPhone
    ws.Range("B4") = inputCAddress
    ws.Range("B5") = inputCWebsite
    
    '/Status/'
    statusLabel_fix.Caption = "Status:"
    statusLabelLog.Caption = "Updating..."
    Status
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

