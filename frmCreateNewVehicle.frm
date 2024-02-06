VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCreateNewVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Vehicle"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7755
   HelpContextID   =   20
   Icon            =   "frmCreateNewVehicle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLength 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   11
      Top             =   5880
      Width           =   3135
   End
   Begin VB.TextBox txtWidth 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4200
      TabIndex        =   12
      Top             =   5880
      Width           =   3135
   End
   Begin VB.TextBox txtQuantity 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4200
      TabIndex        =   5
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker dtProductionStarted 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Year when the production started"
      Top             =   2160
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
      CustomFormat    =   "MM/yyyy"
      Format          =   97255427
      CurrentDate     =   44707
   End
   Begin MSComCtl2.DTPicker dtProductionEnded 
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Year when the production ended"
      Top             =   2160
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
      CustomFormat    =   "MM/yyyy"
      Format          =   97255427
      CurrentDate     =   44707
   End
   Begin MSWLess.WLText txtPrice 
      Height          =   495
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Price per unit"
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
   Begin VB.Label lblQuantity 
      Caption         =   "Quantity"
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
      TabIndex        =   27
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label lblYearProduction 
      Caption         =   "Year of Production"
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
      TabIndex        =   26
      Top             =   2880
      Width           =   3135
   End
   Begin MSWLess.WLCommand btnReset 
      Height          =   975
      Left            =   4080
      TabIndex        =   14
      Top             =   6720
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
   Begin MSWLess.WLCommand btnCreate 
      Height          =   975
      Left            =   960
      TabIndex        =   13
      Top             =   6720
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
   Begin VB.Label lblWidth 
      Caption         =   "Width"
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
      TabIndex        =   25
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label lblLength 
      Caption         =   "Length"
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
      TabIndex        =   24
      Top             =   5640
      Width           =   3135
   End
   Begin MSWLess.WLOption optTransmission 
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   8
      ToolTipText     =   "Manual transmission"
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393216
      Caption         =   "Manual"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
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
   Begin MSWLess.WLOption optTransmission 
      Height          =   495
      Index           =   0
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "Automatic or CVT tansmission"
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      Caption         =   "Automatic"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
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
   Begin VB.Label lblPrice 
      Caption         =   "Price"
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
      TabIndex        =   23
      Top             =   3840
      Width           =   3135
   End
   Begin MSWLess.WLCombo cmbManufacturer 
      Height          =   390
      Left            =   4200
      TabIndex        =   1
      ToolTipText     =   "Manufacturer"
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   -14340
      Text            =   "cmbManufacturer"
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
      List            =   "frmCreateNewVehicle.frx":3AFA
   End
   Begin MSWLess.WLCombo cmbBodyStyle 
      Height          =   390
      Left            =   4200
      TabIndex        =   10
      ToolTipText     =   "Body Style"
      Top             =   5040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   -14300
      Text            =   "cmbBodyStyle"
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
      List            =   "frmCreateNewVehicle.frx":3B16
   End
   Begin MSWLess.WLCombo cmbClass 
      Height          =   390
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "Class"
      Top             =   5040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   -13860
      Text            =   "cmbClass"
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
      List            =   "frmCreateNewVehicle.frx":3D22
   End
   Begin VB.Label Label3 
      Caption         =   "Production Ended"
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
      Top             =   1920
      Width           =   3135
   End
   Begin MSWLess.WLText txtName 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Name of the vehicle model"
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
      Caption         =   "Vehicle"
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
      TabIndex        =   15
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
      TabIndex        =   16
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Production Started"
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
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label lblClass 
      Caption         =   "Class"
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
      TabIndex        =   20
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label lblBodyStyle 
      Caption         =   "Body Style"
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
      TabIndex        =   19
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label lblManufacturer 
      Caption         =   "Manufacturer"
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
      TabIndex        =   21
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblTransmission 
      Caption         =   "Transmission"
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
      TabIndex        =   22
      Top             =   3840
      Width           =   3135
   End
End
Attribute VB_Name = "frmCreateNewVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PreviousName As String
Dim AlreadyMarked As Boolean

Private Sub btnCreate_Click()
    On Error GoTo ErrHandler
        ExecuteSQL "Select * from Vehicle where Vehicle_Name = '" & txtName & "'"
        If IsInformationValid Then
            If btnCreate.Caption = "&Create" Then
                If Not rs.EOF Then
                    MsgBox "Vehicle model already exists!", vbCritical, "Error"
                    Exit Sub
                End If
                rs.AddNew
                
                MsgBox "New vehicle model added successfully!", vbOKOnly, "Information"
            ElseIf btnCreate.Caption = "&Update" Then
                If ModelExists(txtName) Then Exit Sub
                
                MsgBox "Vehicle model updated successfully!", vbOKOnly, "Information"
            End If
        Else
            MsgBox "Fill all the required spaces in the form", vbInformation, "Information"
            CheckEmptyText
            AlreadyMarked = True
            Exit Sub
        End If
        
        rs!Vehicle_Name = txtName
        rs!Manufacturer_ID = GetManufacturerID
        rs!Production_Started = Format(CDate(dtProductionStarted), "MM/dd/YYYY")
        rs!Production_Ended = Format(CDate(dtProductionEnded), "MM/dd/YYYY")
        rs!Produced = txtYear
        rs!Quantity = txtQuantity
        rs!Price = CDbl(Replace(txtPrice, "$", ""))
        If optTransmission(0).value = True Then
            rs!Transmission = "Automatic"
        Else
            rs!Transmission = "Manual"
        End If
        rs!Class_ID = GetClassID
        rs!Body_Style = cmbBodyStyle.Text
        rs!Length = CDbl(Replace(txtLength.Text, " mm", ""))
        rs!Width = CDbl(Replace(txtWidth.Text, " mm", ""))
        rs!Available = True
        rs.Update
        
        AlreadyMarked = False
        RemoveMark Me
        ClearForm
        Exit Sub
ErrHandler:
    MsgBox "There was an error during the operation", vbCritical, "Error"
    Exit Sub
End Sub

Function ModelExists(ModelName As String) As Boolean
    Dim result As Boolean
    ExecuteSQL2 "Select * from Vehicle where Vehicle_Name = '" & ModelName & "'"
    If ModelName = PreviousName Then
        result = False
    ElseIf Not rs2.EOF And ModelName <> PreviousName Then
        MsgBox "Vehicle model already exists!", vbCritical, "Error"
        result = True
    End If
    ModelExists = result
End Function

Sub CheckEmptyText()
    If Not AlreadyMarked Then
        AddRequiredMark lblName, vbRed, txtName
        AddRequiredMark lblManufacturer, vbRed, , cmbManufacturer
        AddRequiredMark lblYearProduction, vbRed, txtYear
        AddRequiredMark lblQuantity, vbRed, txtQuantity
        AddRequiredMark lblPrice, vbRed, txtPrice
        AddRequiredMark lblTransmission, vbRed, , , optTransmission
        AddRequiredMark lblClass, vbRed, , cmbClass
        AddRequiredMark lblBodyStyle, vbRed, , cmbBodyStyle
        AddRequiredMark lblLength, vbRed, txtLength
        AddRequiredMark lblWidth, vbRed, txtWidth
    End If
End Sub

Function GetManufacturerID() As Integer
    ExecuteSQL2 "Select * from Brand where Brand_Name = '" & cmbManufacturer.Text & "'"
    GetManufacturerID = rs2.Fields!ID
End Function

Function GetClassID() As Integer
    ExecuteSQL2 "Select * from Class where Class_Name = '" & cmbClass.Text & "'"
    GetClassID = rs2.Fields!ID
End Function

Function IsInformationValid() As Boolean
    If txtName.Text <> "" And cmbManufacturer.ListIndex <> -1 And _
        txtQuantity.Text <> "" And txtPrice.Text <> "" And txtYear.Text <> "" And _
        (optTransmission(0).value <> False Or optTransmission(1).value <> False) And _
        cmbClass.ListIndex <> -1 And cmbBodyStyle.ListIndex <> -1 And _
        txtLength.Text <> "" And txtWidth.Text <> "" Then
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
    cmbManufacturer.ListIndex = -1
    optTransmission(0).value = False
    optTransmission(1).value = False
    cmbClass.ListIndex = -1
    cmbBodyStyle.ListIndex = -1
    dtProductionEnded.value = DateTime.Date
    dtProductionStarted.value = DateTime.Date
    txtYear.Text = ""
    txtQuantity.Text = ""
    txtPrice.Text = ""
    txtLength.Text = ""
    txtWidth.Text = ""
End Sub

Private Sub Form_Load()
    LoadManufacturers
    LoadClasses
    VerifyCurrentRole
End Sub

Sub VerifyCurrentRole()
    If frmMain.CurrentUserRoleID = 1 Then
        MsgBox "The current user does not have permission to add or modify vehicles information", vbInformation, "Information"
        btnCreate.Enabled = False
    End If
End Sub

Sub LoadManufacturers()
    ExecuteSQL "Select * from Brand order by Brand_Name asc"
    cmbManufacturer.Clear
    While Not rs.EOF
        cmbManufacturer.AddItem (rs.Fields!Brand_Name)
        rs.MoveNext
    Wend
End Sub

Sub LoadClasses()
    ExecuteSQL "Select * from Class"
    cmbClass.Clear
    While Not rs.EOF
        cmbClass.AddItem (rs.Fields!Class_Name)
        rs.MoveNext
    Wend
End Sub

Private Sub txtLength_GotFocus()
    If txtLength.Text <> "" Then
        txtLength.Text = Replace(txtLength.Text, " mm", "")
        txtLength.Text = FormatNumber(txtLength.Text, 0, vbFalse, vbFalse, vbFalse)
        txtLength.MaxLength = 6
    End If
End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)
    VerifyChar KeyAscii
End Sub

Public Sub txtLength_LostFocus()
    txtLength.MaxLength = 10
    txtLength.Text = Format(txtLength.Text, "##,### mm")
End Sub

Private Sub txtPrice_GotFocus()
    If txtPrice.Text <> "" Then txtPrice.Text = Replace(FormatNumber(txtPrice.Text, 0), ",", "")
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    VerifyChar KeyAscii
End Sub

Public Sub txtPrice_LostFocus()
    txtPrice.Text = Format(txtPrice.Text, "$#,###")
End Sub

Private Sub txtQuantity_Click()
    txtQuantity.MaxLength = 6
End Sub

Private Sub txtQuantity_GotFocus()
    If txtQuantity.Text <> "" Then txtQuantity.Text = Replace(FormatNumber(txtQuantity.Text, 0), ",", "")
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    VerifyChar KeyAscii
End Sub

Public Sub txtQuantity_LostFocus()
    If Len(txtQuantity) Mod 3 = 0 Then txtQuantity.MaxLength = 7
    txtQuantity.Text = Format(txtQuantity.Text, "#,##0")
End Sub

Private Sub txtWidth_GotFocus()
    If txtWidth.Text <> "" Then
        txtWidth.Text = Replace(txtWidth.Text, " mm", "")
        txtWidth.Text = Format(txtWidth.Text, "")
        txtWidth.MaxLength = 6
    End If
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    VerifyChar KeyAscii
End Sub

Sub VerifyChar(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
        MsgBox "Enter only numeric characters!", vbInformation, "Information"
    End If
End Sub

Public Sub txtWidth_LostFocus()
    txtWidth.MaxLength = 10
    txtWidth.Text = Format(txtWidth.Text, "##,### mm")
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
        MsgBox "Enter only numeric characters!", vbInformation, "Information"
    End If
    If Len(txtYear) = 4 And KeyAscii <> 8 Then
        KeyAscii = 0
        MsgBox "Enter a valid year of 4 characters"
    End If
End Sub
