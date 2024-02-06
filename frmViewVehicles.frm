VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewVehicles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Vehicle Models Details"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20355
   HelpContextID   =   30
   Icon            =   "frmViewVehicles.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   20355
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridResults 
      Height          =   5415
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   19575
      _ExtentX        =   34528
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
   Begin MSWLess.WLCommand btnShowHiddenElements 
      Height          =   975
      Left            =   2520
      TabIndex        =   6
      Top             =   6600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "&Show Deleted"
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
      Left            =   14760
      TabIndex        =   5
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
   Begin MSWLess.WLCommand btnChangeStock 
      Height          =   975
      Left            =   11640
      TabIndex        =   4
      Top             =   6600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "Change &Stock"
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
      Left            =   5160
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
   Begin MSWLess.WLCommand btnDelete 
      Height          =   975
      Left            =   8400
      TabIndex        =   2
      Top             =   6600
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
      Width           =   19815
   End
End
Attribute VB_Name = "frmViewVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelectedRow As Integer
Dim query As String

Private Sub btnChangeStock_Click()
    Dim SelectedModel As String
    SelectedModel = gridResults.TextMatrix(gridResults.Row, 0)
    SelectedRow = gridResults.Row
    Dim f As frmChangeStock
    Set f = New frmChangeStock
    f.LoadQuery "Select * from Vehicle where Vehicle_Name = '" & SelectedModel & "'"
    f.Show vbModal, Me
    FillGrid
    gridResults.Row = SelectedRow
End Sub

Private Sub btnShowHiddenElements_Click()
    If btnShowHiddenElements.Caption = "&Show Deleted" Then
        query = "Select * from Vehicle where Available = False Order By Vehicle_Name Asc"
        FillGrid
        btnShowHiddenElements.Caption = "Show &Both"
    ElseIf btnShowHiddenElements.Caption = "Show &Both" Then
        query = "Select * from Vehicle order by Vehicle_Name asc"
        FillGrid
        btnShowHiddenElements.Caption = "&Hide Deleted"
    ElseIf btnShowHiddenElements.Caption = "&Hide Deleted" Then
        query = "Select * from Vehicle where Available = True Order By Vehicle_Name Asc"
        FillGrid
        btnShowHiddenElements.Caption = "&Show Deleted"
    End If
    gridResults_SelChange
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    query = "Select * from Vehicle where Available = True Order By Vehicle_Name Asc"
    FillGrid
    gridResults_SelChange
End Sub

Sub FillGrid()
    Dim ManufacturerName As String, ClassName As String, value As String
    gridResults.Clear
    ExecuteSQL query
    With gridResults
        .Cols = 12
        .FixedCols = 0
        .Rows = 0
        .AddItem "Model Name" & vbTab & "Manufacturer" & vbTab & "Production Started" _
            & vbTab & "Production Ended" & vbTab & "Year of Production" _
            & vbTab & "Quantity Available" & vbTab & "Price per Unit" _
            & vbTab & "Transmission" & vbTab & "Class" & vbTab & "Body Style" _
            & vbTab & "Length" & vbTab & "Width"
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
            For j = 1 To 12
                value = rs.Fields(j)
                If j = 2 Then
                    ManufacturerName = GetManufacturerName(rs.Fields(j))
                    value = ManufacturerName
                    .ColWidth(j - 1) = 1500
                ElseIf j = 3 Or j = 4 Then
                    value = Format(value, "MM/YYYY")
                    If j = 3 Then .ColWidth(j - 1) = 2000
                    If j = 4 Then .ColWidth(j - 1) = 1800
                ElseIf j = 5 Then .ColWidth(j - 1) = 2000
                ElseIf j = 6 Then
                    If value = 0 Then value = "Out of Stock"
                    value = Format(value, "#,###")
                    .ColWidth(j - 1) = 2000
                ElseIf j = 7 Then
                    value = Format(value, "$#,###")
                    .ColWidth(j - 1) = 1600
                ElseIf j = 8 Then .ColWidth(j - 1) = 1400
                ElseIf j = 9 Then
                    ClassName = GetClassName(rs.Fields(j))
                    value = ClassName
                    .ColWidth(j - 1) = 1800
                ElseIf j = 11 Or j = 12 Then
                    value = Format(value, "#,### mm")
                End If
                .TextMatrix(i, k) = value
                If k >= 2 And k <= 6 Or k = 10 Or k = 11 Then
                    .ColAlignment(k) = flexAlignRightCenter
                Else
                    .ColAlignment(k) = flexAlignLeftCenter
                End If
                k = k + 1
            Next
            rs.MoveNext
            i = i + 1
            k = 0
        Wend
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        '.Row = 0
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

Private Sub btnEdit_Click()
    gridResults_DblClick
End Sub

Function GetManufacturerName(ManufacturerID As Integer) As String
    ExecuteSQL3 "Select * from Brand where ID = " & ManufacturerID
    GetManufacturerName = rs3.Fields!Brand_Name
End Function

Function GetManufacturerID(ManufacturerName As Integer) As Integer
    ExecuteSQL3 "Select * from Brand where Brand_Name = " & ManufacturerName
    GetManufacturerID = rs3.Fields!ID
End Function

Function GetClassName(ClassID As Integer) As String
    ExecuteSQL3 "Select * from Class where ID = " & ClassID
    GetClassName = rs3.Fields!Class_Name
End Function

Function GetClassID(ClassName As Integer) As String
    ExecuteSQL3 "Select * from Class where Class_Name = " & ClassName
    GetClassID = rs3.Fields!ID
End Function

Function GetManufacturerIndex(ManufacturerName As String, CreateVehicleForm As frmCreateNewVehicle) As Integer
    Dim i As Integer, value As Integer
    For i = 0 To CreateVehicleForm.cmbManufacturer.ListCount - 1
        If CreateVehicleForm.cmbManufacturer.List(i) = ManufacturerName Then
            value = i
            Exit For
        End If
    Next i
    GetManufacturerIndex = value
End Function

Function GetClassIndex(ClassName As String, CreateVehicleForm As frmCreateNewVehicle) As Integer
    Dim i As Integer, value As Integer
    For i = 0 To CreateVehicleForm.cmbClass.ListCount - 1
        If CreateVehicleForm.cmbClass.List(i) = ClassName Then
            value = i
            Exit For
        End If
    Next i
    GetClassIndex = value
End Function

Function GetBodyStyleIndex(BodyStyleName As String, CreateVehicleForm As frmCreateNewVehicle) As Integer
    Dim i As Integer, value As Integer
    For i = 0 To CreateVehicleForm.cmbBodyStyle.ListCount - 1
        If CreateVehicleForm.cmbBodyStyle.List(i) = BodyStyleName Then
            value = i
            Exit For
        End If
    Next i
    GetBodyStyleIndex = value
End Function

Function GetModelIndex(ModelName As String, DeleteVehicleForm As frmDeleteVehicle) As Integer
    Dim i As Integer, value As Integer
    For i = 0 To DeleteVehicleForm.cmbModel.ListCount - 1
        If DeleteVehicleForm.cmbModel.List(i) = ModelName Then
            value = i
            Exit For
        End If
    Next i
    GetModelIndex = value
End Function

Private Sub btnDelete_Click()
    Dim SelectedModel As String, ModelIndex As Integer
    SelectedModel = gridResults.TextMatrix(gridResults.Row, 0)
    Dim f As frmDeleteVehicle
    Set f = New frmDeleteVehicle
    ModelIndex = GetModelIndex(SelectedModel, f)
    f.cmbModel = f.cmbModel.List(ModelIndex)
    SelectedRow = gridResults.Row
    f.Show vbModal, Me
    FillGrid
    SelectLastRow
End Sub

Sub SelectLastRow()
    If gridResults.Rows > SelectedRow Then
        gridResults.Row = SelectedRow
    Else
        gridResults.Row = gridResults.Rows - 1
    End If
End Sub

Private Sub gridResults_DblClick()
    Dim SelectedModel As String
    SelectedModel = gridResults.TextMatrix(gridResults.Row, 0)
    ExecuteSQL2 "Select * from Vehicle where Vehicle_Name = '" & SelectedModel & "'"
    If rs2.EOF Or rs2.RecordCount = 0 Then
        MsgBox "Please select a valid item", vbCritical, "Error"
        Exit Sub
    End If
    Dim ManufacturerName As String, ClassName As String
    Dim ManufacturerIndex As Integer, ClassIndex As Integer, BodyStyleIndex As Integer
    Dim Transmission As String
    If btnEdit.Caption = "&Edit" Then
        Dim f As frmCreateNewVehicle
        Set f = New frmCreateNewVehicle
        f.txtName = rs2.Fields!Vehicle_Name
        ManufacturerName = GetManufacturerName(rs2.Fields!Manufacturer_ID)
        ManufacturerIndex = GetManufacturerIndex(ManufacturerName, f)
        f.cmbManufacturer.Text = f.cmbManufacturer.List(ManufacturerIndex)
        f.dtProductionStarted = rs2.Fields!Production_Started
        f.dtProductionEnded = rs2.Fields!Production_Ended
        f.txtYear = rs2.Fields!Produced
        f.txtQuantity = rs2.Fields!Quantity
        f.txtQuantity.Enabled = False
        f.txtPrice = rs2.Fields!Price
        Transmission = rs2.Fields!Transmission
        If Transmission = "Manual" Then
            f.optTransmission(1).value = True
        Else
            f.optTransmission(0).value = True
        End If
        ClassName = GetClassName(rs2.Fields!Class_ID)
        ClassIndex = GetClassIndex(ClassName, f)
        f.cmbClass.Text = f.cmbClass.List(ClassIndex)
        BodyStyleIndex = GetBodyStyleIndex(rs2.Fields!Body_Style, f)
        f.cmbBodyStyle = f.cmbBodyStyle.List(BodyStyleIndex)
        f.txtLength = rs2.Fields!Length
        f.txtWidth = rs2.Fields!Width
        f.txtLength_LostFocus
        f.txtPrice_LostFocus
        f.txtQuantity_LostFocus
        f.txtWidth_LostFocus
        f.PreviousName = rs2.Fields!Vehicle_Name
        f.btnCreate.Caption = "&Update"
        f.btnReset.Enabled = False
        f.Show vbModal, Me
    ElseIf btnEdit.Caption = "&Restore model" Then
        rs2.Fields!Available = True
        rs2.Update
        
        MsgBox "Model restored successfully!", vbInformation, "Information"
        'btnEdit.Caption = "&Edit"
        'btnDelete.Enabled = True
    End If
    SelectedRow = gridResults.Row
    FillGrid
    SelectLastRow
End Sub

Private Sub gridResults_Click()
    gridResults_SelChange
End Sub

Private Sub gridResults_SelChange()
    Dim CurrentVehicle As String
    CurrentVehicle = gridResults.TextMatrix(gridResults.Row, 0)
    If gridResults.Row > 0 And Not CurrentVehicle = "" And Not CurrentVehicle = "Model Name" Then
        Dim SelectedModel As String, currentBool As Boolean
        SelectedModel = CurrentVehicle
        ExecuteSQL3 "Select * from Vehicle where Vehicle_Name = '" & SelectedModel & "'"
        currentBool = rs3.Fields!Available
        If currentBool = False Then
            btnEdit.Caption = "&Restore model"
            btnDelete.Enabled = False
            btnChangeStock.Enabled = False
            btnEdit.Enabled = True
        ElseIf currentBool = True Then
            btnEdit.Caption = "&Edit"
            btnDelete.Enabled = True
            btnChangeStock.Enabled = True
            btnEdit.Enabled = True
        End If
    Else
        btnDelete.Enabled = False
        btnEdit.Enabled = False
        btnChangeStock.Enabled = False
    End If
End Sub

