VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCreateNewStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Employee"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7710
   HelpContextID   =   20
   Icon            =   "frmCreateNewStaff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtDateBirth 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   141361155
      CurrentDate     =   44707
   End
   Begin MSMask.MaskEdBox txtDNI 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "DNI of the new employee"
      Top             =   2160
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#########"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtPhoneNumber 
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Phone number of the new employee"
      Top             =   2160
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "####-####"
      PromptChar      =   "_"
   End
   Begin MSWLess.WLCommand btnCreate 
      Height          =   975
      Left            =   960
      TabIndex        =   8
      Top             =   4920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "&Create"
      ForeColor       =   -1
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWLess.WLCommand btnReset 
      Height          =   975
      Left            =   4200
      TabIndex        =   9
      Top             =   4920
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "&Reset"
      ForeColor       =   -1
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWLess.WLText txtPassword 
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "Password to enter the application"
      Top             =   4080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      PasswordChar    =   "*"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWLess.WLText txtUsername 
      Height          =   495
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Username to enter the application"
      Top             =   4080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWLess.WLCombo cmbRole 
      Height          =   390
      Left            =   4200
      TabIndex        =   5
      ToolTipText     =   "Role of the new employee"
      Top             =   3120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   -300
      Text            =   "cmbRole"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      List            =   "frmCreateNewStaff.frx":3AFA
   End
   Begin VB.Label Label5 
      Caption         =   "Date of birth"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblPhoneNumber 
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin MSWLess.WLText txtLastName 
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      ToolTipText     =   "Last name of the new employee"
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWLess.WLText txtName 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Name of the new employee"
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Staff"
      BeginProperty Font 
         Name            =   "Javanese Text"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblLastName 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblDNI 
      Caption         =   "DNI"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblRole 
      Caption         =   "Role"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "frmCreateNewStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PreviousDNI As String, PreviousUsername As String
Dim AlreadyMarked As Boolean

Private Sub btnCreate_Click()
    On Error GoTo ErrHandler
        If IsInformationValid Then
            If btnCreate.Caption = "&Create" Then
                ExecuteSQL "Select * from Staff where DNI = '" & Replace(txtDNI.Text, "-", "") & "'"
                ExecuteSQL2 "Select * from Staff where Username = '" & txtUsername & "'"
                If Not rs.EOF Then
                    MsgBox "DNI already exists", vbCritical, "Error"
                    Exit Sub
                ElseIf Not rs2.EOF Then
                    MsgBox "Username already exists", vbCritical, "Error"
                    Exit Sub
                End If
                
                rs.AddNew
                
                MsgBox "New user added successfully!", vbOKOnly, "Information"
            ElseIf btnCreate.Caption = "&Update" Then
                ExecuteSQL "Select * from Staff where DNI = '" & PreviousDNI & "'"
                If UserExists(Replace(txtDNI.Text, "-", ""), txtUsername) Then Exit Sub
                
                MsgBox "User updated successfully!", vbOKOnly, "Information"
            End If
        Else
            MsgBox "Fill all the required spaces in the form", vbInformation, "Information"
            CheckEmptyText
            AlreadyMarked = True
            Exit Sub
        End If
        
        rs!Staff_Name = txtName
        rs!Staff_LastName = txtLastName
        rs!DNI = Replace(txtDNI.Text, "-", "")
        rs!Phone_Number = Replace(txtPhoneNumber.Text, "-", "")
        rs!DateBirth = dtDateBirth
        rs!Role_ID = GetRoleID
        rs!UserName = txtUsername
        rs!Password = txtPassword
        rs!Available = True
        rs!PreviousRole_ID = GetRoleID
        rs.Update
        
        AlreadyMarked = False
        RemoveMark Me
        ClearForm
        Exit Sub
ErrHandler:
    MsgBox "There was an error during the operation", vbCritical, "Error"
    Exit Sub
End Sub

Function UserExists(UserDNI As String, UserName As String) As Boolean
    Dim result As Boolean
    ExecuteSQL2 "Select * from Staff where DNI = '" & UserDNI & "'"
    ExecuteSQL3 "Select * from Staff where Username = '" & UserName & "'"
    If UserDNI = PreviousDNI And UserName = PreviousUsername Then
        result = False
    ElseIf Not rs2.EOF And UserDNI <> PreviousDNI Then
        MsgBox "DNI already exists!", vbCritical, "Error"
        result = True
    ElseIf Not rs3.EOF And UserName <> PreviousUsername Then
        MsgBox "Username already exists!", vbCritical, "Error"
        result = True
    End If
    UserExists = result
End Function

Sub CheckEmptyText()
    AddRequiredMark lblName, vbRed, txtName
    AddRequiredMark lblLastName, vbRed, txtLastName
    AddRequiredMark lblDNI, vbRed, txtDNI
    AddRequiredMark lblPhoneNumber, vbRed, txtPhoneNumber
    AddRequiredMark lblRole, vbRed, cmbRole
    AddRequiredMark lblUsername, vbRed, txtUsername
    AddRequiredMark lblPassword, vbRed, txtPassword
End Sub

Function GetRoleID() As Integer
    ExecuteSQL2 "Select * from Role where Role = '" & cmbRole.Text & "'"
    GetRoleID = rs2.Fields!ID
End Function

Private Sub btnReset_Click()
    ClearForm
End Sub

Sub ClearForm()
    RemoveMark Me
    AlreadyMarked = False
    txtName.Text = ""
    txtLastName.Text = ""
    txtDNI.Text = ""
    txtPhoneNumber.Text = ""
    dtDateBirth.value = DateTime.Date
    cmbRole.ListIndex = -1
    txtUsername.Text = ""
    txtPassword.Text = ""
End Sub

Function IsInformationValid() As Boolean
    If txtName.Text <> "" And txtLastName.Text <> "" And txtDNI.Text <> "" And _
        txtPhoneNumber.Text <> "" And txtUsername.Text <> "" And _
        txtPassword.Text <> "" And cmbRole.ListIndex <> -1 Then
        IsInformationValid = True
    Else
        IsInformationValid = False
    End If
End Function

Private Sub Form_Load()
    dtDateBirth.value = DateTime.Date
    LoadRoles
    VerifyCurrentRole
End Sub

Sub VerifyCurrentRole()
    If frmMain.CurrentUserRoleID = 1 Then
        MsgBox "The current user does not have permission to add or modify employees information", vbInformation, "Information"
        btnCreate.Enabled = False
    End If
End Sub

Sub LoadRoles()
    ExecuteSQL "Select * from Role Where ID <> 4"
    cmbRole.Clear
    While Not rs.EOF
        cmbRole.AddItem (rs.Fields!Role)
        rs.MoveNext
    Wend
End Sub

Private Sub txtDNI_GotFocus()
    If txtDNI.Text <> "" Then txtDNI.Text = Replace(txtDNI.Text, "-", "")
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    VerifyChar KeyAscii
    If Len(txtDNI) = 9 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Sub VerifyChar(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
        MsgBox "Enter only numeric characters!", vbInformation, "Information"
    End If
End Sub

Public Sub txtDNI_LostFocus()
    txtDNI.Text = Format(txtDNI.Text, "#-####-####")
End Sub

Private Sub txtPhoneNumber_GotFocus()
    If txtPhoneNumber.Text <> "" Then txtPhoneNumber.Text = Replace(txtPhoneNumber.Text, "-", "")
End Sub

Private Sub txtPhoneNumber_KeyPress(KeyAscii As Integer)
    If Len(txtPhoneNumber) = 8 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Public Sub txtPhoneNumber_LostFocus()
    txtPhoneNumber.Text = Format(txtPhoneNumber.Text, "####-####")
End Sub
