VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminForm 
   Caption         =   "Admin Panel  - Created By Alex Fare"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6375
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

    Dim Worksheet_Set
    Dim ws As Worksheet
    Dim List_Select
    List_Select = "CreatedByAlexFare"
    Set ws = Sheets(List_Select)
    Set Worksheet_Set = ws
    vDisplay = ws.Range("D1")
End Sub

'/ Pressing Enter will instantly search /'
Private Sub Gage_Number_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Search_Button_Click
    End If
End Sub

Sub CheckForUpdate_Click()
    Dim url         As String
    url = "https://github.com/alexfare/InventoryTracker"
    ActiveWorkbook.FollowHyperlink url
End Sub

Private Sub btnCreateAccount_click()
    Unload AdminForm
    CreateAccount.Show
End Sub

Private Sub btnUpdateUser_click()
    Unload AdminForm
    ChangePassword.Show
End Sub

Private Sub btnAbout_Click()
    MsgBox "Simple Inventory Tracker For Excel."
End Sub

Private Sub auditBTN_Click()
    Unload Menu
    Worksheets("Audit").Activate
End Sub

Private Sub btnFormat_Click()
    Format_Form.Show
End Sub

Private Sub btnLogout_Click()
        List_Select = "Admin"        ' Tab name
        Set ws = Sheets(List_Select)
        Set Worksheet_Set = ws
        ws.Range("B55") = "1"
        Unload AdminForm
End Sub

Private Sub btnCompanyProfile_Click()
    CompanyProfile.Show
End Sub

