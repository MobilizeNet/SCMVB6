VERSION 5.00
Begin VB.Form frmChangeStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Stock"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4620
   Icon            =   "frmChangeStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameStock 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox chkOutOfStock 
         Caption         =   "Out of Stock?"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Change &Stock"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   5
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtNewStock 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtCurrentStock 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblNewStock 
         Caption         =   "New Stock:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblCurrent 
         Caption         =   "Current Stock:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmChangeStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkOutOfStock_Click()
    If chkOutOfStock.value <> 0 Then txtNewStock.Enabled = False
End Sub

Private Sub cmdAccept_Click()
    If txtNewStock = "" And chkOutOfStock.value = 0 Then
        MsgBox "Enter a valid amount", vbExclamation, "Information"
        Exit Sub
    End If
    If MsgBox("Do you really want to change the Stock of this model?", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        If chkOutOfStock.value <> 0 Then
            rs3.Fields!Quantity = 0
        Else
            rs3.Fields!Quantity = FormatNumber(txtNewStock, 0)
        End If
        rs3.Update
        MsgBox "Stock changed successfully!", vbInformation, "Information"
        Unload Me
    End If
End Sub

Public Sub LoadQuery(query As String)
    ExecuteSQL3 query
    txtCurrentStock = rs3.Fields!Quantity
    txtCurrentStock.Enabled = False
End Sub

Private Sub txtCurrentStock_LostFocus()
    txtCurrentStock = Format(txtCurrentStock, "#,###")
End Sub

Private Sub txtNewStock_GotFocus()
    If txtNewStock <> "" Then txtNewStock.Text = FormatNumber(txtNewStock, 0)
End Sub

Private Sub txtNewStock_KeyPress(KeyAscii As Integer)
    VerifyChar KeyAscii
End Sub

Private Sub txtNewStock_LostFocus()
    If Len(txtNewStock) Mod 3 = 0 Then txtNewStock.MaxLength = 7
    txtNewStock = Format(txtNewStock, "#,###")
End Sub

Sub VerifyChar(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
        MsgBox "Enter only numeric characters!", vbInformation, "Information"
    End If
End Sub
