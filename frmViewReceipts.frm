VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewReceipts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Receipts Details"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17430
   HelpContextID   =   30
   Icon            =   "frmViewReceipts.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   17430
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridResults 
      Height          =   5415
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   9551
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWLess.WLCommand btnShowAll 
      Height          =   975
      Left            =   2760
      TabIndex        =   5
      Top             =   6600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "&Show Approved"
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
   Begin MSWLess.WLCommand btnExit 
      Height          =   975
      Left            =   12480
      TabIndex        =   4
      Top             =   6600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "E&xit"
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
   Begin MSWLess.WLCommand btnEdit 
      Height          =   975
      Left            =   6000
      TabIndex        =   3
      Top             =   6600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "&Edit"
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
   Begin MSWLess.WLCommand btnChangeStatus 
      Height          =   975
      Left            =   9240
      TabIndex        =   2
      Top             =   6600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "&Change Status"
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
      Width           =   16935
   End
End
Attribute VB_Name = "frmViewReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelectedRow As Integer
Dim query As String

Private Sub btnChangeStatus_Click()
    Dim currentStatus As String
    SelectedReceipt = gridResults.TextMatrix(gridResults.Row, 0)
    ExecuteSQL2 "Select * from Receipt where Receipt_ID = '" & SelectedReceipt & "'"
    currentStatus = rs2!Status
    If currentStatus = "Approved" Then
        rs2.Fields!Status = "Rejected"
        UpdateQuantity rs2.Fields!Model_ID, False, rs2.Fields!Quantity
    Else
        rs2!Status = "Approved"
        UpdateQuantity rs2.Fields!Model_ID, True, rs2.Fields!Quantity
    End If
    rs2.Update
    MsgBox "Status changed", vbInformation, "Information"
    FillGrid
End Sub

Sub UpdateQuantity(ModelID As Integer, IsRemovingVehicles As Boolean, CarsSold As Integer)
    Dim currentQuantity As Double
    ExecuteSQL "Select * from Vehicle where ID = " & ModelID
    currentQuantity = CDbl(rs.Fields!Quantity)
    If IsRemovingVehicles Then
        currentQuantity = currentQuantity - CarsSold
    Else
        currentQuantity = currentQuantity + CarsSold
    End If
    rs.Fields!Quantity = currentQuantity
    rs.Update
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnShowAll_Click()
    If btnShowAll.Caption = "&Show Approved" Then
        query = "Select * from Receipt where Status = 'Approved' Order by Receipt_ID Asc"
        FillGrid
        btnShowAll.Caption = "&Show Rejected"
    ElseIf btnShowAll.Caption = "&Show Rejected" Then
        query = "Select * from Receipt where Status = 'Rejected' Order by Receipt_ID Asc"
        FillGrid
        btnShowAll.Caption = "&Show All"
    ElseIf btnShowAll.Caption = "&Show All" Then
        query = "Select * from Receipt Order by Receipt_ID Asc"
        FillGrid
        btnShowAll.Caption = "&Show Approved"
    End If
End Sub

Private Sub Form_Load()
    query = "Select * from Receipt Order by Receipt_ID Asc"
    FillGrid
End Sub

Sub FillGrid()
    Dim ManufacturerName As String, ModelName As String, SellerName As String
    Dim value As String
    Dim BooleanValue As Variant
    gridResults.Clear
    ExecuteSQL query
    With gridResults
        .Cols = 12
        .FixedCols = 0
        .Rows = 0
        .AddItem "Receipt ID" & vbTab & "Name" & vbTab & "Last Name" _
            & vbTab & "Manufacturer" & vbTab & "Model" & vbTab & "Quantity" _
            & vbTab & "Seller" & vbTab & "Vehicle Insurance" & vbTab & "Third Person Insurance" _
            & vbTab & "SubTotal" & vbTab & "Total" & vbTab & "Status"
        AutoFitRows
        Dim i As Integer, j As Integer, k As Integer
        i = 1
        k = 0
        If rs.RecordCount = 0 Then
            .Rows = 2
            .FixedRows = 1
            .Row = 0
            Exit Sub
        End If
        While Not rs.EOF
            .Rows = .Rows + 1
            For j = 1 To .Cols
                If j = 4 Then
                    ManufacturerName = GetManufacturerName(rs.Fields(j))
                    value = ManufacturerName
                ElseIf j = 5 Then
                    ModelName = GetModelName(rs.Fields(j))
                    value = ModelName
                    .ColAlignment(k) = flexAlignLeftCenter
                ElseIf j = 6 Then
                    value = rs.Fields(j)
                    value = Format(value, "#,###")
                ElseIf j = 7 Then
                    SellerName = GetSellerName(rs.Fields(j))
                    value = SellerName
                ElseIf j = 8 Or j = 9 Then
                    BooleanValue = rs.Fields(j)
                    If BooleanValue = True Then
                        value = "Yes"
                    Else
                        value = "No"
                    End If
                ElseIf j = 10 Or j = 11 Then
                    value = rs.Fields(j)
                    value = Format(value, "$#,###")
                Else
                    value = rs.Fields(j)
                End If
                .TextMatrix(i, k) = value
                k = k + 1
            Next
            rs.MoveNext
            i = i + 1
            k = 0
        Wend
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
    End With
End Sub

Sub AutoFitRows()
    Dim i As Integer
    Dim Twips As Integer
    
    With gridResults
        For i = 0 To .Cols - 1
        Twips = Me.TextWidth(.TextMatrix(0, i))
        .ColWidth(i) = Twips * Me.gridResults.Font.Size / Me.Font.Size + 530 '* Screen.TwipsPerPixelX
        Next i
    End With
End Sub

Function GetSellerName(SellerID As Integer) As String
    ExecuteSQL3 "Select * from Staff where ID = " & SellerID
    GetSellerName = rs3.Fields!Staff_Name & " " & rs3.Fields!Staff_LastName
End Function

Function GetModelName(ModelID As Integer) As String
    ExecuteSQL3 "Select * from Vehicle where ID = " & ModelID
    GetModelName = rs3.Fields!Vehicle_Name
End Function

Function GetManufacturerName(ManufacturerID As Integer) As String
    ExecuteSQL3 "Select * from Brand where ID = " & ManufacturerID
    GetManufacturerName = rs3.Fields!Brand_Name
End Function

Private Sub gridResults_DblClick()
    Dim SelectedReceipt As String
    SelectedReceipt = gridResults.TextMatrix(gridResults.Row, 0)
    ExecuteSQL2 "Select * from Receipt where Receipt_ID = '" & SelectedReceipt & "'"
    If rs2.EOF Or rs2.RecordCount = 0 Then
        MsgBox "Please select a valid item", vbCritical, "Error"
        Exit Sub
    End If
    Dim ManufacturerName As String, ModelName As String, SellerName As String, SellerDNI As String
    Dim ManufacturerIndex As Integer, ModelIndex As Integer, SellerIndex As Integer
    Dim f As frmCreateNewReceipt
    Set f = New frmCreateNewReceipt
    f.txtID = rs2.Fields!Receipt_ID
    f.txtName = rs2.Fields!Client_Name
    f.txtLastName = rs2.Fields!Client_LastName
    ManufacturerName = GetManufacturerName(rs2.Fields!Manufacturer_ID)
    ManufacturerIndex = GetManufacturerIndex(ManufacturerName, f)
    f.cmbManufacturer.Text = f.cmbManufacturer.List(ManufacturerIndex)
    ModelName = GetModelName(rs2.Fields!Model_ID)
    ModelIndex = GetModelIndex(ModelName, f)
    f.cmbModel.Text = f.cmbModel.List(ModelIndex)
    f.txtQuantity = rs2.Fields!Quantity
    SellerDNI = GetSellerDNI(rs2.Fields!Seller_ID)
    SellerIndex = GetSellerIndex(SellerDNI, f)
    f.cmbSeller.Text = f.cmbSeller.List(SellerIndex)
    f.chk3PersonInsurance = rs2.Fields!Third_Person_Insurance
    f.chkVehicleInsurance = rs2.Fields!Vehicle_Insurance
    f.txtSubTotal = rs2.Fields!Subtotal
    f.txtTotal = rs2.Fields!Total
    f.btnCreate.Caption = "&Update"
    f.btnReset.Enabled = False
    f.PreviousStatus = rs2.Fields!Status
    f.SubTotalValue = f.txtSubTotal
    f.TotalValue = f.txtTotal
    f.FormatSubTotalValue
    f.Show vbModal, Me
    SelectedRow = gridResults.Row
    FillGrid
    gridResults.Row = SelectedRow
End Sub

Function GetSellerDNI(SellerID As Integer) As String
    ExecuteSQL3 "Select * from Staff where ID = " & SellerID
    GetSellerDNI = rs3.Fields!DNI
End Function

Function GetSellerIndex(SellerDNI As String, CreateReceiptForm As frmCreateNewReceipt) As Integer
    Dim i As Integer, value As Integer
    For i = 0 To CreateReceiptForm.cmbSeller.ListCount - 1
        If InStr(1, CreateReceiptForm.cmbModel.List(i), SellerDNI, vbTextCompare) > 0 Then
            value = i
            Exit For
        End If
    Next i
    GetSellerIndex = value
End Function

Function GetModelIndex(ModelName As String, CreateReceiptForm As frmCreateNewReceipt) As Integer
    Dim i As Integer, value As Integer
    For i = 0 To CreateReceiptForm.cmbModel.ListCount - 1
        If CreateReceiptForm.cmbModel.List(i) = ModelName Then
            value = i
            Exit For
        End If
    Next i
    GetModelIndex = value
End Function

Function GetManufacturerIndex(ManufacturerName As String, CreateReceiptForm As frmCreateNewReceipt) As Integer
    Dim i As Integer, value As Integer
    For i = 0 To CreateReceiptForm.cmbManufacturer.ListCount - 1
        If CreateReceiptForm.cmbManufacturer.List(i) = ManufacturerName Then
            value = i
            Exit For
        End If
    Next i
    GetManufacturerIndex = value
End Function

Private Sub btnEdit_Click()
    gridResults_DblClick
End Sub
