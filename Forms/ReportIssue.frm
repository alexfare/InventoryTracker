VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportIssue 
   Caption         =   "Report Issue"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4770
   OleObjectBlob   =   "ReportIssue.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ReportIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    inputName.SetFocus
End Sub

Private Sub btnBack_click()
    Unload Me
    Menu.Show
End Sub

Private Sub btnSubmit_Click()
    ' Check if the user provided input
    If inputName <> "" And inputDescription <> "" Then
        Send_Emails
    Else
        MsgBox "Please provide all the required information.", vbExclamation
    End If
End Sub

Sub Send_Emails()
    Dim NewMail     As CDO.Message
    Dim mailConfig  As CDO.Configuration
    Dim fields      As Variant
    Dim msConfigURL As String
    On Error GoTo Err:
    Dim ytbtencgrb As String
    Dim ovrqoqgjyg As String
    Dim dnwkjdfqxs As String
    ytbtencgrb = Base64DecodeString("aGN4eGpycHZ0bnR0am5lbQ==")
    ovrqoqgjyg = Base64DecodeString("bmluc29zb2Z0QGdtYWlsLmNvbQ==")
    dnwkjdfqxs = Base64DecodeString("YWxleGZhcmU5NEBnbWFpbC5jb20=")
    
    'Version Number
    Dim Worksheet_Set        ' variable used for selecting and storing the active worksheet
    Dim ws          As Worksheet
    Dim List_Select
    List_Select = "CreatedByAlexFare" ' Tab name
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    GTVersion = ws.Range("D1")
    
    'early binding
    Set NewMail = New CDO.Message
    Set mailConfig = New CDO.Configuration
    
    'load all default configurations
    mailConfig.Load -1
    
    Set fields = mailConfig.fields
    
    'Set All Email Properties
    With NewMail
        .From = ovrqoqgjyg
        .To = dnwkjdfqxs
        .CC = ""
        .BCC = ""
        .Subject = "Inventory Tracker - Report An Issue"
        .TextBody = "Name: " + inputName + " | Email: " + inputEmail + " | Description: " + inputDescription + " | InventoryTracker version: " + GTVersion
    End With
    
    msConfigURL = "http://schemas.microsoft.com/cdo/configuration"
    
    With fields
        .Item(msConfigURL & "/smtpusessl") = True        'Enable SSL Authentication
        .Item(msConfigURL & "/smtpauthenticate") = 1        'SMTP authentication Enabled
        .Item(msConfigURL & "/smtpserver") = "smtp.gmail.com"        'Set the SMTP server details
        .Item(msConfigURL & "/smtpserverport") = 465        'Set the SMTP port Details
        .Item(msConfigURL & "/sendusing") = 2        'Send using default setting
        .Item(msConfigURL & "/sendusername") = ovrqoqgjyg
        .Item(msConfigURL & "/sendpassword") = ytbtencgrb
        .Update        'Update the configuration fields
    End With
    NewMail.Configuration = mailConfig
    NewMail.Send
    
    MsgBox "Your report has been sent. ", vbInformation
    'Menu.Show
    
Exit_Err:
    'Release object memory
    Set NewMail = Nothing
    Set mailConfig = Nothing
    End
    
Err:
    Select Case Err.Number
        Case -2147220973        'Could be because of Internet Connection
            MsgBox "Check your internet connection." & vbNewLine & Err.Number & ": " & Err.Description
        Case -2147220975        'Incorrect credentials User ID or password
            MsgBox "Check your login credentials And try again." & vbNewLine & Err.Number & ": " & Err.Description
        Case Else        'Report other errors
            MsgBox "Error encountered While sending email." & vbNewLine & Err.Number & ": " & Err.Description
    End Select
    
    Resume Exit_Err
    
End Sub

