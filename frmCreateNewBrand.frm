VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCreateNewBrand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Brand"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7710
   HelpContextID   =   20
   Icon            =   "frmCreateNewBrand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtFounded 
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      ToolTipText     =   "Date when the brand was created"
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
      Format          =   112132099
      CurrentDate     =   44707
   End
   Begin MSWLess.WLCombo cmbParent 
      Height          =   390
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Manufacturer"
      Top             =   4080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   27764
      Text            =   "cmbParent"
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
      List            =   "frmCreateNewBrand.frx":3AFA
   End
   Begin MSWLess.WLText txtNumberEmployees 
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "Number of employees working there"
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
   Begin MSWLess.WLText txtName 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Name of the brand"
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
   Begin VB.Label Label9 
      Caption         =   "Parent Company"
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
      Width           =   3135
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
   Begin MSWLess.WLText txtWebsite 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Website"
      Top             =   3120
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
   Begin MSWLess.WLText txtAreaServed 
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Countries where this brand is selling products"
      Top             =   2160
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
   Begin MSWLess.WLText txtHeadquarters 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Country where the headquarter is located"
      Top             =   2160
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
   Begin MSWLess.WLText txtOwner 
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      ToolTipText     =   "Name of the brand owner"
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
      Caption         =   "Brand"
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
   Begin VB.Label lblOwner 
      Caption         =   "Owner"
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
      TabIndex        =   11
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblHeadquarter 
      Caption         =   "Headquarter"
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
      TabIndex        =   12
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label lblAreaServed 
      Caption         =   "Area Served"
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
      TabIndex        =   13
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Founded"
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
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label lblWebsite 
      Caption         =   "Website"
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
      Width           =   3135
   End
   Begin VB.Label lblNumberEmployees 
      Caption         =   "Number of Employees"
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
      Top             =   3840
      Width           =   3135
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
      TabIndex        =   18
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "frmCreateNewBrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PreviousName As String
Dim AlreadyMarked As Boolean

Private Sub btnCreate_Click()
    On Error GoTo ErrHandler
        ExecuteSQL "Select * from Brand where Brand_Name = '" & txtName & "'"
        If IsInformationValid Then
            If btnCreate.Caption = "&Create" Then
                If Not rs.EOF Then
                    MsgBox "Brand already exists!", vbCritical, "Error"
                    Exit Sub
                End If
                rs.AddNew
                
                MsgBox "New brand added successfully!", vbOKOnly, "Information"
            ElseIf btnCreate.Caption = "&Update" Then
                If BrandExists(txtName) Then Exit Sub
                If cmbParent.Text = txtName Then
                    MsgBox "Parent Company can't be the same company", vbExclamation, "Error"
                    Exit Sub
                End If
                
                MsgBox "Brand updated successfully!", vbOKOnly, "Information"
            End If
        Else
            MsgBox "Fill all the required spaces in the form", vbInformation, "Information"
            CheckEmptyText
            AlreadyMarked = True
            Exit Sub
        End If
        
        rs.Fields!Brand_Name = txtName
        rs.Fields!Owner = txtOwner
        rs.Fields!Headquarter = txtHeadquarters
        rs.Fields!Area_Served = txtAreaServed
        rs.Fields!Website = txtWebsite
        rs.Fields!Founded = CDate(dtFounded)
        If cmbParent.ListIndex <> -1 Then
            rs.Fields!Parent_Company = GetManufacturerID
        End If
        rs.Fields!Number_Employees = Replace(txtNumberEmployees.Text, ",", "")
        rs.Fields!Available = True
        rs.Update
        
        AlreadyMarked = False
        RemoveMark Me
        ClearForm
        LoadParents
        Exit Sub
ErrHandler:
    MsgBox "There was an error during the operation", vbCritical, "Error"
    Exit Sub
End Sub

Function BrandExists(BrandName As String) As Boolean
    Dim result As Boolean
    ExecuteSQL2 "Select * from Brand where Brand_Name = '" & BrandName & "'"
    If BrandName = PreviousName Then
        result = False
    ElseIf BrandName <> PreviousName And Not rs2.EOF Then
        MsgBox "Brand already exists!", vbCritical, "Error"
        result = True
    End If
    BrandExists = result
End Function

Sub CheckEmptyText()
    AddRequiredMark lblName, vbRed, txtName
    AddRequiredMark lblOwner, vbRed, txtOwner
    AddRequiredMark lblHeadquarter, vbRed, txtHeadquarters
    AddRequiredMark lblAreaServed, vbRed, txtAreaServed
    AddRequiredMark lblWebsite, vbRed, txtWebsite
    AddRequiredMark lblNumberEmployees, vbRed, txtNumberEmployees
End Sub

Function GetManufacturerID() As Integer
    ExecuteSQL2 "Select * from Brand where Brand_Name = '" & cmbParent.Text & "'"
    GetManufacturerID = rs2.Fields!ID
End Function

Function IsInformationValid() As Boolean
    If txtName.Text <> "" And txtOwner.Text <> "" And txtHeadquarters.Text <> "" And _
        txtAreaServed.Text <> "" And txtWebsite.Text <> "" And _
        txtNumberEmployees.Text <> "" Then
        IsInformationValid = True
    Else
        IsInformationValid = False
    End If
End Function

Private Sub btnReset_Click()
    ClearForm
End Sub

Sub ClearForm()
    RemoveMark Me
    AlreadyMarked = False
    txtName.Text = ""
    txtOwner.Text = ""
    txtHeadquarters.Text = ""
    txtAreaServed.Text = ""
    txtWebsite.Text = ""
    dtFounded.value = DateTime.Date
    cmbParent.ListIndex = -1
    txtNumberEmployees.Text = ""
End Sub

Private Sub Form_Load()
    dtFounded.value = DateTime.Date
    LoadParents
    VerifyCurrentRole
End Sub

Sub VerifyCurrentRole()
    If frmMain.CurrentUserRoleID = 1 Then
        MsgBox "The current user does not have permission to add or modify brands information", vbInformation, "Information"
        btnCreate.Enabled = False
    End If
End Sub

Sub LoadParents()
    ExecuteSQL "Select * from Brand order by Brand_Name asc"
    cmbParent.Clear
    While Not rs.EOF
        cmbParent.AddItem (rs.Fields!Brand_Name)
        rs.MoveNext
    Wend
End Sub

Private Sub txtNumberEmployees_GotFocus()
    If txtNumberEmployees.Text <> "" Then txtNumberEmployees.Text = Format(txtNumberEmployees.Text, "")
End Sub

Private Sub txtNumberEmployees_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
        MsgBox "Enter only numeric characters!", vbInformation, "Information"
    End If
End Sub

Public Sub txtNumberEmployees_LostFocus()
    txtNumberEmployees.Text = Format(txtNumberEmployees.Text, "###,###")
End Sub
