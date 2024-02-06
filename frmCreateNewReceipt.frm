VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Begin VB.Form frmCreateNewReceipt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Receipt"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7725
   HelpContextID   =   20
   Icon            =   "frmCreateNewReceipt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin MSWLess.WLText txtQuantity 
      Height          =   495
      Left            =   360
      TabIndex        =   23
      ToolTipText     =   "Name of the client"
      Top             =   3960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   6
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
   Begin MSWLess.WLText txtID 
      Height          =   495
      Left            =   360
      TabIndex        =   21
      ToolTipText     =   "Receipt ID"
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      Alignment       =   1
      Text            =   "ID"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
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
   Begin MSWLess.WLCheck chk3PersonInsurance 
      Height          =   375
      Left            =   360
      TabIndex        =   20
      ToolTipText     =   "Insurance in case of accidents with people"
      Top             =   5520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      Caption         =   "Third Person Insurance"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin MSWLess.WLCommand btnCreate 
      Height          =   975
      Left            =   1080
      TabIndex        =   19
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
   Begin MSWLess.WLCommand btnReset 
      Height          =   975
      Left            =   4200
      TabIndex        =   18
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
   Begin MSWLess.WLText txtTotal 
      Height          =   495
      Left            =   4200
      TabIndex        =   16
      ToolTipText     =   "Total (Includes 13% taxes)"
      Top             =   5880
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
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
   Begin MSWLess.WLText txtSubTotal 
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Subtotal to pay"
      Top             =   4920
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      ForeColor       =   255
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
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
   Begin MSWLess.WLCombo cmbSeller 
      Height          =   390
      Left            =   4200
      TabIndex        =   12
      ToolTipText     =   "Seller"
      Top             =   3960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   -9740
      Text            =   "cmbSeller"
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
      List            =   "frmCreateNewReceipt.frx":3AFA
   End
   Begin MSWLess.WLCheck chkVehicleInsurance 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Insurance in case of crashes"
      Top             =   4920
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      Caption         =   "Vehicle Insurance"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Extras"
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
      TabIndex        =   9
      Top             =   4680
      Width           =   3255
   End
   Begin MSWLess.WLCombo cmbModel 
      Height          =   390
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "Model"
      Top             =   3120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   -9860
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
      List            =   "frmCreateNewReceipt.frx":3B16
   End
   Begin MSWLess.WLCombo cmbManufacturer 
      Height          =   390
      Left            =   360
      TabIndex        =   5
      ToolTipText     =   "Manufacturer"
      Top             =   3120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   -10660
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
      List            =   "frmCreateNewReceipt.frx":3B64
   End
   Begin MSWLess.WLText txtLastName 
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Last name of the client"
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
   Begin MSWLess.WLText txtName 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Name of the client"
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
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Receipt"
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
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
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
      TabIndex        =   4
      Top             =   1920
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
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label lblModel 
      Caption         =   "Model"
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
      TabIndex        =   8
      Top             =   2880
      Width           =   3135
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
      Left            =   360
      TabIndex        =   11
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label lblSeller 
      Caption         =   "Seller"
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
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label8 
      Caption         =   "SubTotal"
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
      TabIndex        =   15
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "Total"
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
      TabIndex        =   17
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label10 
      Caption         =   "Receipt ID"
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
      TabIndex        =   22
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "frmCreateNewReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AlreadyMarked As Boolean
Public SubTotalValue As Double
Public TotalValue As Double
Public PreviousStatus As String
Const VehicleInsuranceValue = 4000
Const ThirdPersonInsuranceValue = 7500

Private Sub btnCreate_Click()
    On Error GoTo ErrHandler
        Dim Subtotal As Long
        Dim Total As Long
        If IsInformationValid And IsQuantityAvailable Then
            If btnCreate.Caption = "&Create" Then
                ExecuteSQL "Select * from Receipt"
                rs.AddNew
                rs!Receipt_ID = txtID
                PreviousStatus = "Approved"
                MsgBox "New receipt created successfully!", vbOKOnly, "Information"
            ElseIf btnCreate.Caption = "&Update" Then
                ExecuteSQL "Select * from Receipt where Receipt_ID = '" & txtID & "'"
                MsgBox "Receipt updated successfully!", vbOKOnly, "Information"
            End If
        Else
            MsgBox "Fill all the required spaces in the form", vbInformation, "Information"
            CheckEmptyText
            AlreadyMarked = True
            Exit Sub
        End If
        rs!Client_Name = txtName
        rs!Client_LastName = txtLastName
        rs!Manufacturer_ID = GetManufacturerID
        rs!Model_ID = GetModelID
        rs!Quantity = FormatNumber(txtQuantity.Text, 0)
        rs!Seller_ID = GetSellerID
        rs!Vehicle_Insurance = chkVehicleInsurance.value
        rs!Third_Person_Insurance = chk3PersonInsurance.value
        Subtotal = FormatNumber(Replace(txtSubTotal.Text, "$", ""), 0)
        rs!Subtotal = Subtotal
        Total = FormatNumber(Replace(txtTotal.Text, "$", ""), 0)
        rs!Total = Total
        rs!Status = PreviousStatus
        
        rs.Update
        
        UpdateVehicleQuantity
        
        AlreadyMarked = False
        RemoveMark Me
        ClearForm
        If btnCreate.Caption = "&Update" Then
            Unload Me
        End If
        Exit Sub
ErrHandler:
    MsgBox "There was an error during the operation", vbCritical, "Error"
    Exit Sub
End Sub

Function IsQuantityAvailable() As Boolean
    Dim NextAvailable As Variant
    Dim value As Boolean
    value = True
    ExecuteSQL "Select * from Vehicle where ID = " & GetModelID
    NextAvailable = rs.Fields!Quantity - CDbl(txtQuantity)
    If NextAvailable < 0 Then
        MsgBox "Quantity is not available or vehicle is currently Out of Stock"
        value = False
    End If
    IsQuantityAvailable = value
End Function

Sub UpdateVehicleQuantity()
    ExecuteSQL "Select * from Vehicle where ID = " & GetModelID
    rs.Fields!Quantity = rs.Fields!Quantity - CDbl(txtQuantity)
    rs.Update
End Sub

Sub CheckEmptyText()
    If Not AlreadyMarked Then
        AddRequiredMark lblName, vbRed, txtName
        AddRequiredMark lblLastName, vbRed, txtLastName
        AddRequiredMark lblManufacturer, vbRed, , cmbManufacturer
        AddRequiredMark lblModel, vbRed, , cmbModel
        AddRequiredMark lblQuantity, vbRed, txtQuantity
        AddRequiredMark lblSeller, vbRed, , cmbSeller
    End If
End Sub

Private Sub chk3PersonInsurance_Click()
    On Error GoTo ErrHandler
    If chk3PersonInsurance.value = wlChecked Then
        SubTotalValue = Replace(txtSubTotal, "$", "")
        SubTotalValue = SubTotalValue + ThirdPersonInsuranceValue
        txtSubTotal.Text = ""
        FormatSubTotalValue
    Else
        SubTotalValue = Replace(txtSubTotal, "$", "")
        SubTotalValue = SubTotalValue - ThirdPersonInsuranceValue
        txtSubTotal.Text = ""
        FormatSubTotalValue
    End If
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub txtQuantity_LostFocus()
    Dim currentQuantity As Variant
    If cmbManufacturer.ListIndex = -1 Or cmbModel.ListIndex = -1 Or txtQuantity.Text = "" Then
        Exit Sub
    End If
    If IsQuantityValid(currentQuantity) Then
        txtQuantity.Text = Format(txtQuantity.Text, "###,#")
        SetSubTotal
    Else
        MsgBox "There are not enough vehicles, maximum available is " & currentQuantity, vbOKOnly, "Information"
        txtQuantity = "0"
        txtQuantity.SetFocus
    End If
End Sub

Function IsQuantityValid(ByRef currentQuantity As Variant) As Boolean
    ExecuteSQL2 "Select * from Vehicle where Vehicle_Name = '" & cmbModel.Text & "'"
    currentQuantity = rs2.Fields!Quantity
    If currentQuantity >= CDbl(txtQuantity.Text) Then
        IsQuantityValid = True
    Else
        IsQuantityValid = False
    End If
End Function

Function GetManufacturerID() As Integer
    ExecuteSQL3 "Select * from Brand where Brand_Name = '" & cmbManufacturer.Text & "'"
    GetManufacturerID = rs3.Fields!ID
End Function

Function GetModelID() As Integer
    ExecuteSQL3 "Select * from Vehicle where Vehicle_Name = '" & cmbModel.Text & "'"
    GetModelID = rs3.Fields!ID
End Function

Function GetSellerID() As Integer
    Dim SellerDNI As String
    SellerDNI = Right(cmbSeller, 9)
    ExecuteSQL3 "Select * from Staff where DNI = '" & SellerDNI & "'"
    GetSellerID = rs3.Fields!ID
End Function

Private Sub btnReset_Click()
    ClearForm
End Sub

Sub ClearForm()
    RemoveMark Me
    AlreadyMarked = False
    txtID.Text = ""
    txtName.Text = ""
    txtLastName.Text = ""
    txtQuantity.Text = ""
    cmbManufacturer.ListIndex = -1
    cmbModel.ListIndex = -1
    cmbSeller.ListIndex = -1
    txtSubTotal.Text = ""
    txtTotal.Text = ""
    chkVehicleInsurance.value = wlUnchecked
    chk3PersonInsurance.value = wlUnchecked
    LoadReceiptID
End Sub

Function IsInformationValid() As Boolean
    If txtName.Text <> "" And txtLastName.Text <> "" And _
        cmbManufacturer.ListIndex <> -1 And cmbModel.ListIndex <> -1 And _
        txtQuantity.Text <> "" And cmbSeller.ListIndex <> -1 Then
        IsInformationValid = True
    Else
        IsInformationValid = False
    End If
End Function

Private Sub chkVehicleInsurance_Click()
    On Error GoTo ErrHandler
    If chkVehicleInsurance.value = wlChecked Then
        SubTotalValue = Replace(txtSubTotal, "$", "")
        SubTotalValue = SubTotalValue + VehicleInsuranceValue
        txtSubTotal.Text = ""
        FormatSubTotalValue
    Else
        SubTotalValue = Replace(txtSubTotal, "$", "")
        SubTotalValue = SubTotalValue - VehicleInsuranceValue
        txtSubTotal.Text = ""
        FormatSubTotalValue
    End If
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub cmbManufacturer_Click()
    If cmbManufacturer.Text <> "" Then
        LoadModels
        txtQuantity.Text = ""
        txtSubTotal.Text = ""
        txtTotal.Text = ""
        chk3PersonInsurance.value = wlUnchecked
        chkVehicleInsurance.value = wlUnchecked
    End If
End Sub

Private Sub cmbModel_Click()
    If cmbModel = "" Then
        Exit Sub
    End If
    txtQuantity.Text = ""
    txtSubTotal.Text = ""
    txtTotal.Text = ""
    chk3PersonInsurance.value = wlUnchecked
    chkVehicleInsurance.value = wlUnchecked
End Sub

Sub SetSubTotal()
    If txtQuantity.Text <> "" Then
        ExecuteSQL2 "Select * from Vehicle where Vehicle_Name = '" & cmbModel.Text & "'"
        vehiclePrice = rs2.Fields!Price
        SubTotalValue = CDbl(vehiclePrice * txtQuantity)
        If chkVehicleInsurance.value = wlChecked Then
            SubTotalValue = SubTotalValue + VehicleInsuranceValue
        End If
        If chk3PersonInsurance.value = wlChecked Then
            SubTotalValue = SubTotalValue + ThirdPersonInsuranceValue
        End If
        FormatSubTotalValue
    Else
        SubTotalValue = 0
        TotalValue = 0
        FormatSubTotalValue
    End If
End Sub

Sub FormatSubTotalValue()
    txtSubTotal = Format(SubTotalValue, "$#,###")
    TotalValue = SubTotalValue + SubTotalValue * 0.13
    txtTotal = Format(TotalValue, "$#,###")
End Sub

Private Sub Form_Load()
    LoadReceiptID
    LoadManufacturers
    LoadSellers
    cmbModel.ListIndex = 0
    VerifyCurrentRole
End Sub

Sub VerifyCurrentRole()
    If frmMain.CurrentUserRoleID = 2 Then
        MsgBox "The current user does not have permission to sell cars", vbInformation, "Information"
        btnCreate.Enabled = False
    End If
End Sub

Sub LoadReceiptID()
    ExecuteSQL "Select * from Receipt"
    txtID = "V3UC44P" & CStr(rs.RecordCount) + 1
    If Len(CStr(rs.RecordCount)) = 2 Then
        txtID = "V3UC44" & CStr(rs.RecordCount) + 1
    ElseIf Len(CStr(rs.RecordCount)) = 3 Then
        txtID = "V3UC4" & CStr(rs.RecordCount) + 1
    ElseIf Len(CStr(rs.RecordCount)) = 4 Then
        txtID = "V3UC" & CStr(rs.RecordCount) + 1
    End If
End Sub

Sub LoadManufacturers()
    ExecuteSQL "Select * from Brand"
    cmbManufacturer.Clear
    While Not rs.EOF
        cmbManufacturer.AddItem (rs.Fields!Brand_Name)
        rs.MoveNext
    Wend
End Sub

Sub LoadModels()
    ExecuteSQL "Select * from Vehicle where Manufacturer_ID = " & GetManufacturerID
    cmbModel.Clear
    While Not rs.EOF
        cmbModel.AddItem (rs.Fields!Vehicle_Name)
        rs.MoveNext
    Wend
End Sub

Sub LoadSellers()
    Dim FullName As String
    ExecuteSQL "Select * from Staff where Role_ID = 1 or Role_ID = 3"
    cmbSeller.Clear
    While Not rs.EOF
        FullName = rs.Fields!Staff_Name & " " & rs.Fields!Staff_LastName & " - " & rs.Fields!DNI
        cmbSeller.AddItem (FullName)
        rs.MoveNext
    Wend
End Sub

