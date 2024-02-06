VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   HelpContextID   =   10
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsername 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If VerifyUser Then
        LoginSucceeded = True
        SetMessage
        Me.Hide
        frmMain.Show
    Else
        MsgBox "Invalid Username or Password, try again!", vbCritical, "Login Error"
        txtPassword.SetFocus
    End If
End Sub

Sub SetMessage()
    Dim FullName As String, Role As String
    ExecuteSQL "Select * from Staff where Username = '" & txtUsername & "'"
    
    FullName = rs.Fields!Staff_Name & " " & rs.Fields!Staff_LastName
    Role = GetRoleName(rs.Fields!Role_ID)
    
    frmMain.lblUser.Caption = "Welcome " & FullName & "!"
    frmMain.lblRole.Caption = "Role: " & Role
        
    frmMain.CurrentUserRoleID = rs.Fields!Role_ID
End Sub

Function GetRoleName(RoleID As Integer) As String
    ExecuteSQL2 "Select * from Role where ID = " & RoleID
    GetRoleName = rs2.Fields!Role
End Function

Function VerifyUser() As Boolean
    ExecuteSQL "Select * from Staff where Username = '" & txtUsername & "' and Password = '" & txtPassword & "'"
    
    If rs.EOF Then
        VerifyUser = False
    Else
        If rs!Role_ID = 4 Then
            VerifyUser = False
            MsgBox "The user can not access the application anymore.", vbError, "Error"
            Exit Function
        End If
        VerifyUser = True
    End If
End Function

