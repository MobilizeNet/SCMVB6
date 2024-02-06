VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeleteVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Vehicle Model"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7755
   HelpContextID   =   20
   Icon            =   "frmDeleteVehicle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameFilters 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   6975
      Begin VB.Frame pnlResults 
         Caption         =   "Results"
         Height          =   2535
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   6735
         Begin MSComctlLib.ListView lstResults 
            Height          =   2175
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3836
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "colModel"
               Text            =   "Model"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "colmanufacturer"
               Text            =   "Manufacturer"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Key             =   "colClass"
               Text            =   "Class"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "colBodyStyle"
               Text            =   "Body Style"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Key             =   "colTransmission"
               Text            =   "Transmission"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Key             =   "colPrice"
               Text            =   "Price"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Key             =   "colYearProduction"
               Text            =   "Year of Production"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Key             =   "colQuantity"
               Text            =   "Quantity"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame pnlFilters 
         Caption         =   "Filters"
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   6735
         Begin VB.ComboBox cmbYear 
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   2640
            Width           =   2535
         End
         Begin VB.ComboBox cmbManufacturer 
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "frmDeleteVehicle.frx":3AFA
            Left            =   240
            List            =   "frmDeleteVehicle.frx":3AFC
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   2655
         End
         Begin VB.CommandButton cmdResetFilters 
            Caption         =   "Reset Filters"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label11 
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
            Left            =   3960
            TabIndex        =   21
            Top             =   2400
            Width           =   2415
         End
         Begin MSWLess.WLCombo cmbPrice 
            Height          =   390
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "Manufacturer"
            Top             =   2640
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   688
            _Version        =   393216
            Style           =   2
            ListCount       =   -19956
            Text            =   "cmbPrice"
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
            List            =   "frmDeleteVehicle.frx":3AFE
         End
         Begin VB.Label Label5 
            Caption         =   "Price Range"
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
            Left            =   240
            TabIndex        =   17
            Top             =   2400
            Width           =   2655
         End
         Begin MSWLess.WLCombo cmbTransmission 
            Height          =   390
            Left            =   3960
            TabIndex        =   16
            ToolTipText     =   "Manufacturer"
            Top             =   1800
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   688
            _Version        =   393216
            Style           =   2
            ListCount       =   -19476
            Text            =   "cmbTransmission"
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
            List            =   "frmDeleteVehicle.frx":3B1A
         End
         Begin VB.Label Label4 
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
            Left            =   3960
            TabIndex        =   15
            Top             =   1560
            Width           =   2655
         End
         Begin MSWLess.WLCombo cmbBodyStyle 
            Height          =   390
            Left            =   240
            TabIndex        =   14
            ToolTipText     =   "Manufacturer"
            Top             =   1800
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   688
            _Version        =   393216
            Style           =   2
            ListCount       =   -19636
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
            List            =   "frmDeleteVehicle.frx":3B64
         End
         Begin MSWLess.WLCombo cmbClass 
            Height          =   390
            Left            =   3960
            TabIndex        =   13
            ToolTipText     =   "Manufacturer"
            Top             =   960
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   688
            _Version        =   393216
            Style           =   2
            ListCount       =   -19356
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
            List            =   "frmDeleteVehicle.frx":3B80
         End
         Begin VB.Label Label3 
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
            Left            =   240
            TabIndex        =   12
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label1 
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
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label2 
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
            Left            =   3960
            TabIndex        =   10
            Top             =   720
            Width           =   2535
         End
         Begin MSWLess.WLCheck chkUseAllFilters 
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            Caption         =   "Use all filters?"
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
      End
   End
   Begin VB.CheckBox chkFilters 
      Caption         =   "Show filters"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin MSWLess.WLCombo cmbModel 
      Height          =   390
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Manufacturer"
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   -19636
      Text            =   "cmbModel"
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
      List            =   "frmDeleteVehicle.frx":3B9C
   End
   Begin MSWLess.WLCommand btnDelete 
      Height          =   975
      Left            =   2520
      TabIndex        =   4
      Top             =   8160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "&Delete"
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
   Begin VB.Label Label6 
      Caption         =   "Model Name"
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
      TabIndex        =   3
      Top             =   960
      Width           =   3135
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
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmDeleteVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDelete_Click()
    If cmbModel.ListIndex <> -1 Then
        If MsgBox("Are you sure you want to delete the vehicle model '" _
            & cmbModel.Text & "'?", vbExclamation + vbYesNo) = vbYes Then
            If AreThereVehiclesInStock(cmbModel) Then
                MsgBox "The selected model " & cmbManufacturer.Text & " has vehicles pending to sell!" _
                & vbCrLf & "Remove or sell them to delete this model", vbExclamation, "Error deleting vehicle model"
            Else
                ExecuteSQL "Select * from Vehicle where Vehicle_Name = '" & cmbModel.Text & "'"
                rs.Fields!Available = False
                rs.Fields!Quantity = 0
                rs.Update
                MsgBox "Model deleted successfully!", vbInformation, "Information"
                LoadModels
            End If
        End If
    ElseIf lstResults.ListItems.Count > 0 And Not lstResults.SelectedItem Is Nothing Then
        If MsgBox("Are you sure you want to delete the selected vehicle model '" _
            & lstResults.SelectedItem.Text & "'?", vbExclamation + vbYesNo) = vbYes Then
            If AreThereVehiclesInStock(cmbModel) Then
                MsgBox "The selected model " & lstResults.SelectedItem.Text & " has vehicles pending to sell!" _
                & vbCrLf & "Remove or sell them to delete this model", vbExclamation, "Error deleting vehicle model"
            Else
                ExecuteSQL "Select * from Vehicle where Vehicle_Name = '" & lstResults.SelectedItem & "'"
                rs.Fields!Available = False
                rs.Fields!Quantity = 0
                rs.Update
                MsgBox "Model deleted successfully!", vbInformation, "Information"
                LoadBodyStyles
                LoadClasses
                LoadManufacturers
                LoadYears
                LoadTransmissions
                LoadPrices
                cmdResetFilters_Click
            End If
        End If
    Else
        MsgBox "Select a vehicle model to delete it", vbInformation, "Information"
    End If
End Sub

Function AreThereVehiclesInStock(Model As String) As Boolean
    Dim VehicleQuantity As Integer
    ExecuteSQL2 "Select * from Vehicle where Vehicle_Name = '" & Model & "'"
    While Not rs2.EOF
        VehicleQuantity = rs2.Fields!Quantity
        If VehicleQuantity > 0 Then
            result = True
        End If
        rs2.MoveNext
    Wend
    AreThereVehiclesInStock = result
End Function

Private Sub chkFilters_Click()
    If chkFilters.value <> 0 Then
        frameFilters.Enabled = True
        frameFilters.Visible = True
        Me.Height = 9930
        btnDelete.Top = 8160
        cmbModel.Enabled = False
        cmbModel.ListIndex = -1
    Else
        frameFilters.Enabled = False
        frameFilters.Visible = False
        Me.Height = 4000
        btnDelete.Top = 2000
        cmbModel.Enabled = True
    End If
End Sub

Private Sub chkUseAllFilters_Click()
    LoadClasses
    LoadTransmissions
    LoadBodyStyles
End Sub

Private Sub cmbBodyStyle_Click()
    If chkUseAllFilters.value = 0 Then
        LoadManufacturers
        LoadClasses
        LoadPrices
        LoadTransmissions
        LoadYears
    Else
        If cmbManufacturer.Text = "" Then LoadManufacturers cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
        If cmbClass.Text = "" Then LoadClasses cmbManufacturer.Text, cmbBodyStyle.Text, cmbTransmission.Text, cmbPrice.Text, cmbYear.Text
        If cmbPrice.Text = "" Then LoadPrices cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text
        If cmbTransmission.Text = "" Then LoadTransmissions cmbManufacturer.Text, cmbClass, cmbBodyStyle.Text, cmbPrice, cmbYear.Text
        If cmbYear.Text = "" Then LoadYears cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text, cmbPrice.Text
    End If
    ShowResults cmbManufacturer, cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
End Sub

Private Sub cmbClass_Click()
    If chkUseAllFilters.value = 0 Then
        LoadManufacturers
        LoadBodyStyles
        LoadPrices
        LoadTransmissions
        LoadYears
    Else
        If cmbManufacturer.Text = "" Then LoadManufacturers cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
        If cmbBodyStyle.Text = "" Then LoadBodyStyles cmbManufacturer, cmbClass, cmbTransmission, cmbPrice, cmbYear
        If cmbPrice.Text = "" Then LoadPrices cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text
        If cmbTransmission.Text = "" Then LoadTransmissions cmbManufacturer.Text, cmbClass, cmbBodyStyle.Text, cmbPrice, cmbYear.Text
        If cmbYear.Text = "" Then LoadYears cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text, cmbPrice.Text
    End If
    ShowResults cmbManufacturer, cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
End Sub

Private Sub cmbManufacturer_Click()
    If chkUseAllFilters.value = 0 Then
        LoadClasses
        LoadBodyStyles
        LoadPrices
        LoadTransmissions
        LoadYears
    Else
        If cmbClass.Text = "" Then LoadClasses cmbManufacturer.Text, cmbBodyStyle.Text, cmbTransmission.Text, cmbPrice.Text, cmbYear.Text
        If cmbBodyStyle.Text = "" Then LoadBodyStyles cmbManufacturer, cmbClass, cmbTransmission, cmbPrice, cmbYear
        If cmbPrice.Text = "" Then LoadPrices cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text
        If cmbTransmission.Text = "" Then LoadTransmissions cmbManufacturer.Text, cmbClass, cmbBodyStyle.Text, cmbPrice, cmbYear.Text
        If cmbYear.Text = "" Then LoadYears cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text, cmbPrice.Text
    End If
    ShowResults cmbManufacturer, cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
End Sub

Private Sub cmbPrice_Click()
    If chkUseAllFilters.value = 0 Then
        LoadClasses
        LoadBodyStyles
        LoadManufacturers
        LoadTransmissions
        LoadYears
    Else
        If cmbClass.Text = "" Then LoadClasses cmbManufacturer.Text, cmbBodyStyle.Text, cmbTransmission.Text, cmbPrice.Text, cmbYear.Text
        If cmbBodyStyle.Text = "" Then LoadBodyStyles cmbManufacturer, cmbClass, cmbTransmission, cmbPrice, cmbYear
        If cmbManufacturer.Text = "" Then LoadManufacturers cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
        If cmbTransmission.Text = "" Then LoadTransmissions cmbManufacturer.Text, cmbClass, cmbBodyStyle.Text, cmbPrice, cmbYear.Text
        If cmbYear.Text = "" Then LoadYears cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text, cmbPrice.Text
    End If
    ShowResults cmbManufacturer, cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
End Sub

Private Sub cmbTransmission_Click()
    If chkUseAllFilters.value = 0 Then
        LoadClasses
        LoadBodyStyles
        LoadManufacturers
        LoadPrices
        LoadYears
    Else
        If cmbClass.Text = "" Then LoadClasses cmbManufacturer.Text, cmbBodyStyle.Text, cmbTransmission.Text, cmbPrice.Text, cmbYear.Text
        If cmbBodyStyle.Text = "" Then LoadBodyStyles cmbManufacturer, cmbClass, cmbTransmission, cmbPrice, cmbYear
        If cmbManufacturer.Text = "" Then LoadManufacturers cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
        If cmbPrice.Text = "" Then LoadPrices cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text
        If cmbYear.Text = "" Then LoadYears cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text, cmbPrice.Text
    End If
    ShowResults cmbManufacturer, cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
End Sub

Private Sub cmbYear_Click()
    If chkUseAllFilters.value = 0 Then
        LoadClasses
        LoadBodyStyles
        LoadManufacturers
        LoadPrices
        LoadTransmissions
    Else
        If cmbClass.Text = "" Then LoadClasses cmbManufacturer.Text, cmbBodyStyle.Text, cmbTransmission.Text, cmbPrice.Text, cmbYear.Text
        If cmbBodyStyle.Text = "" Then LoadBodyStyles cmbManufacturer, cmbClass, cmbTransmission, cmbPrice, cmbYear
        If cmbManufacturer.Text = "" Then LoadManufacturers cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
        If cmbPrice.Text = "" Then LoadPrices cmbManufacturer.Text, cmbClass.Text, cmbTransmission.Text, cmbBodyStyle.Text
        If cmbTransmission.Text = "" Then LoadTransmissions cmbManufacturer.Text, cmbClass, cmbBodyStyle.Text, cmbPrice, cmbYear.Text
    End If
    ShowResults cmbManufacturer, cmbClass, cmbBodyStyle, cmbTransmission, cmbPrice, cmbYear
End Sub

Private Sub cmdResetFilters_Click()
    cmbManufacturer.ListIndex = -1
    cmbClass.ListIndex = -1
    cmbBodyStyle.ListIndex = -1
    cmbTransmission.ListIndex = -1
    cmbPrice.ListIndex = -1
    cmbYear.ListIndex = -1
    lstResults.ListItems.Clear
End Sub

Private Sub Form_Load()
    LoadModels
    LoadManufacturers
    LoadClasses
    LoadBodyStyles
    LoadTransmissions
    LoadPrices
    LoadYears
    frameFilters.Enabled = False
    frameFilters.Visible = False
    Me.Height = 4000
    btnDelete.Top = 2000
    VerifyCurrentRole
End Sub

Sub VerifyCurrentRole()
    If frmMain.CurrentUserRoleID = 1 Then
        MsgBox "The current user does not have permission to delete vehicles", vbInformation, "Information"
        btnDelete.Enabled = False
    End If
End Sub

Sub LoadModels()
    ExecuteSQL "Select * from Vehicle where Available = True Order By Vehicle_Name ASC"
    cmbModel.Clear
    While Not rs.EOF
        cmbModel.AddItem (rs.Fields!Vehicle_Name)
        rs.MoveNext
    Wend
End Sub

Sub LoadManufacturers(Optional class As String, Optional bodyStyle As String, _
    Optional Transmission As String, Optional Price As String, Optional pYear As String)
    Dim query As String
    Dim ManufacturerName As String
    
    CreateQuery query, , class, bodyStyle, Transmission, Price, pYear

    ExecuteSQL query
    
    cmbManufacturer.Clear
    While Not rs.EOF
        ManufacturerName = GetManufacturerName(rs.Fields!Manufacturer_ID)
        If Not ExistsInCombo(ManufacturerName, cmbManufacturer) Then
            cmbManufacturer.AddItem (ManufacturerName)
        End If
        rs.MoveNext
    Wend
End Sub

Sub LoadClasses(Optional manufacturer As String, Optional bodyStyle As String, _
    Optional Transmission As String, Optional Price As String, Optional pYear As String)
    Dim query As String
    Dim ClassName As String
    
    If chkUseAllFilters.value = wlChecked Then
        CreateQuery query, manufacturer, , bodyStyle, Transmission, Price, pYear
    
        ExecuteSQL query
        
        cmbClass.Clear
        While Not rs.EOF
            ClassName = GetClassName(rs.Fields!Class_ID)
            If Not ExistsInCombo(ClassName, , cmbClass) Then cmbClass.AddItem (ClassName)
            rs.MoveNext
        Wend
    Else
        ExecuteSQL "Select * from Class"
        cmbClass.Clear
        While Not rs.EOF
            ClassName = rs.Fields!Class_Name
            If Not ExistsInCombo(ClassName, , cmbClass) Then cmbClass.AddItem (ClassName)
            rs.MoveNext
        Wend
    End If
End Sub


Sub LoadBodyStyles(Optional manufacturer As String, Optional class As String, _
    Optional Transmission As String, Optional Price As String, Optional pYear As String)
    Dim query As String
    Dim value As String
    
    If chkUseAllFilters.value = wlChecked Then
        CreateQuery query, manufacturer, class, , Transmission, Price, pYear
        
        ExecuteSQL query
        
        cmbBodyStyle.Clear
        While Not rs.EOF
            value = rs.Fields!Body_Style
            If Not ExistsInCombo(value, , cmbBodyStyle) Then
                cmbBodyStyle.AddItem (value)
            End If
            rs.MoveNext
        Wend
        'rs.Close
    Else
        cmbBodyStyle.Clear
        cmbBodyStyle.AddItem "Buggy"
        cmbBodyStyle.AddItem "Convertible"
        cmbBodyStyle.AddItem "Coupé"
        cmbBodyStyle.AddItem "Flower Car / Hearse"
        cmbBodyStyle.AddItem "Hatchback"
        cmbBodyStyle.AddItem "Limousine"
        cmbBodyStyle.AddItem "Microvan"
        cmbBodyStyle.AddItem "Minivan"
        cmbBodyStyle.AddItem "Panel van"
        cmbBodyStyle.AddItem "Panel truck"
        cmbBodyStyle.AddItem "Pickup truck"
        cmbBodyStyle.AddItem "Roadster"
        cmbBodyStyle.AddItem "Sedan"
        cmbBodyStyle.AddItem "Shooting -brake"
        cmbBodyStyle.AddItem "SUV"
        cmbBodyStyle.AddItem "Station wagon"
        cmbBodyStyle.AddItem "Targa Top"
        cmbBodyStyle.AddItem "Ute / Coupe utility"
    End If
End Sub

Sub LoadTransmissions(Optional manufacturer As String, Optional class As String, _
    Optional bodyStyle As String, Optional Price As String, Optional pYear As String)
    Dim query As String
    Dim value As String
    
    If chkUseAllFilters.value = wlChecked Then
        
        CreateQuery query, manufacturer, class, bodyStyle, , Price, pYear
        
        ExecuteSQL query
        
        cmbTransmission.Clear
        While Not rs.EOF
            value = rs.Fields!Transmission
            If Not ExistsInCombo(value, , cmbTransmission) Then cmbTransmission.AddItem (value)
            rs.MoveNext
        Wend
    Else
        cmbTransmission.Clear
        cmbTransmission.AddItem "Automatic"
        cmbTransmission.AddItem "Manual"
    End If
End Sub

Sub LoadPrices(Optional manufacturer As String, Optional class As String, _
    Optional Transmission As String, Optional bodyStyle As String, Optional pYear As String)
    Dim query As String
    Dim value As String
    
    If chkUseAllFilters.value = wlChecked Then
        
        CreateQuery query, manufacturer, class, bodyStyle, Transmission, , pYear
        
        ExecuteSQL query
        
        cmbPrice.Clear
        While Not rs.EOF
            value = rs.Fields!Price
            If Not ExistsInCombo(value, , cmbPrice) Then cmbPrice.AddItem (value)
            rs.MoveNext
        Wend
    Else
        cmbPrice.Clear
        cmbPrice.AddItem "0 - 1000"
        cmbPrice.AddItem "1000 - 2000"
        cmbPrice.AddItem "2000 - 3000"
        cmbPrice.AddItem "3000 - 5000"
        cmbPrice.AddItem "5000 - 10000"
        cmbPrice.AddItem "10000 - 20000"
        cmbPrice.AddItem "20000 - 50000"
        cmbPrice.AddItem "50000+"
    End If
End Sub

Sub LoadYears(Optional manufacturer As String, Optional class As String, _
    Optional Transmission As String, Optional bodyStyle As String, Optional Price As String)
    Dim query As String, value As String
    
    CreateQuery query, manufacturer, class, bodyStyle, Transmission, Price
    query = query & " order by Produced asc"
    
    ExecuteSQL query
    
    cmbYear.Clear
    While Not rs.EOF
        value = rs.Fields!Produced
        If Not ExistsInCombo(value, cmbYear) Then
            cmbYear.AddItem (value)
        End If
        rs.MoveNext
    Wend

End Sub

Sub ShowResults(Optional manufacturer As String, Optional class As String, Optional bodyStyle As String, _
    Optional pTransmission As String, Optional pPrice As String, Optional pYear As String)
    Dim li As ListItem
    Dim ManufacturerName As String, ClassName As String
    Dim query As String
    On Error GoTo DoNothing
    
    CreateQuery query, manufacturer, class, bodyStyle, pTransmission, pPrice, pYear
    
    ExecuteSQL query
    lstResults.ListItems.Clear
    While Not rs.EOF
        Set li = lstResults.ListItems.Add(, , rs.Fields!Vehicle_Name)
        ManufacturerName = GetManufacturerName(rs.Fields!Manufacturer_ID)
        li.SubItems(1) = ManufacturerName
        ClassName = GetClassName(rs.Fields!Class_ID)
        li.SubItems(2) = ClassName
        li.SubItems(3) = rs.Fields!Body_Style
        li.SubItems(4) = rs.Fields!Transmission
        li.SubItems(5) = rs.Fields!Price
        li.SubItems(6) = rs.Fields!Produced
        li.SubItems(7) = rs.Fields!Quantity
        rs.MoveNext
    Wend
    Exit Sub
DoNothing:
    Exit Sub
End Sub

Sub CreateQuery(ByRef queryOut As String, Optional manufacturer As String, Optional class As String, _
    Optional bodyStyle As String, Optional Transmission As String, Optional Price As String, _
    Optional pYear As String)
    Dim IsFirstFilter As Boolean
    IsFirstFilter = True
    queryOut = "Select * from Vehicle where Available = True"
    
    If Not manufacturer = "" Then queryOut = queryOut & " and Manufacturer_ID = " & GetManufacturerID(manufacturer)
    If Not class = "" Then queryOut = queryOut & " and Class_ID = " & GetClassID(class)
    If Not bodyStyle = "" Then queryOut = queryOut & " and Body_Style = '" & bodyStyle & "'"
    If Not Transmission = "" Then queryOut = queryOut & " and Transmission = '" & Transmission & "'"
    If Not Price = "" Then
        If Price <> "50000+" Then
            GetPriceQuery queryOut, Price, False
        Else
            GetPriceQuery queryOut, Price, True
        End If
    End If
    If Not pYear = "" Then queryOut = queryOut & " and Produced = " & pYear
    
End Sub

Sub GetPriceQuery(ByRef queryOut As String, Price As String, IsHighPrice As Boolean)
    Price = Replace(Price, " ", "")
    Prices = Split(Price, "-")
    If chkUseAllFilters.value = wlChecked Then
        queryOut = queryOut & " and Price = " & CDbl(Prices(0))
    Else
        If Not IsHighPrice Then
            queryOut = queryOut & " and Price >= " & CDbl(Prices(0)) & " and Price <= " & CDbl(Prices(1))
        Else
            queryOut = queryOut & " and Price >= " & CDbl(Replace(Price, "+", ""))
        End If
    End If
End Sub

Function GetManufacturerID(ManufacturerName As String) As Integer
    ExecuteSQL2 "Select * from Brand where Brand_Name = '" & ManufacturerName & "'"
    GetManufacturerID = rs2.Fields!ID
End Function

Function GetManufacturerName(ManufacturerID As Integer) As String
    ExecuteSQL2 "Select * from Brand where ID = " & ManufacturerID
    GetManufacturerName = rs2.Fields!Brand_Name
End Function

Function GetClassID(ClassName As String) As Integer
    ExecuteSQL2 "Select * from Class where Class_Name = '" & ClassName & "'"
    GetClassID = rs2.Fields!ID
End Function

Function GetClassName(ClassID As String) As String
    ExecuteSQL2 "Select * from Class where ID = " & ClassID
    GetClassName = rs2.Fields!Class_Name
End Function
