VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewBrands 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Brands Details"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15285
   HelpContextID   =   30
   Icon            =   "frmViewBrands.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   15285
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridResults 
      Height          =   5415
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   14655
      _ExtentX        =   25850
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
      Left            =   1320
      TabIndex        =   5
      Top             =   6600
      Width           =   2535
      _ExtentX        =   4471
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
      Left            =   11040
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
   Begin MSWLess.WLCommand btnDelete 
      Height          =   975
      Left            =   7800
      TabIndex        =   3
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
   Begin MSWLess.WLCommand btnEdit 
      Height          =   975
      Left            =   4560
      TabIndex        =   2
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
      TabIndex        =   0
      Top             =   0
      Width           =   14775
   End
End
Attribute VB_Name = "frmViewBrands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelectedRow As Integer
Dim query As String

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnShowHiddenElements_Click()
    If btnShowHiddenElements.Caption = "&Show Deleted" Then
        query = "Select * from Brand where Available = False Order By Brand_Name Asc"
        FillGrid
        btnShowHiddenElements.Caption = "Show &Both"
        'btnDelete.Enabled = False
    ElseIf btnShowHiddenElements.Caption = "Show &Both" Then
        query = "Select * from Brand order by Brand_Name asc"
        FillGrid
        btnShowHiddenElements.Caption = "&Hide Deleted"
        'btnDelete.Enabled = True
    ElseIf btnShowHiddenElements.Caption = "&Hide Deleted" Then
        query = "Select * from Brand where Available = True Order By Brand_Name Asc"
        FillGrid
        btnShowHiddenElements.Caption = "&Show Deleted"
        'btnDelete.Enabled = True
    End If
    gridResults_SelChange
End Sub

Private Sub Form_Load()
    query = "Select * from Brand where Available = True Order By Brand_Name Asc"
    FillGrid
    gridResults_SelChange
End Sub

Function GetParentName(ParentID As Integer) As String
    ExecuteSQL3 "Select * from Brand where ID = " & ParentID
    GetParentName = rs3.Fields!Brand_Name
End Function

Sub FillGrid()
    Dim value As String
    gridResults.Clear
    ExecuteSQL query
    With gridResults
        .Cols = 8
        .FixedCols = 0
        .Rows = 0
        .AddItem "Brand Name" & vbTab & "Owner" & vbTab & "Headquarter Location" _
            & vbTab & "Area Served" & vbTab & "Website" & vbTab & "Year of Foundation" _
            & vbTab & "Parent Company" & vbTab & "Number of Employees"
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
            For j = 1 To 8
                If j = 2 Then
                    value = rs.Fields(j)
                    .ColWidth(j - 1) = 1600
                ElseIf j = 3 Then
                    value = rs.Fields(j)
                    .ColWidth(j - 1) = 2200
                ElseIf j = 4 Then
                    value = rs.Fields(j)
                    .ColWidth(j - 1) = 1300
                ElseIf j = 5 Then
                    value = rs.Fields(j)
                    value = Mid(value, 5)
                   .ColWidth(j - 1) = 1800
                ElseIf j = 6 Then
                    value = rs.Fields(j)
                    value = Format(value, "MM/dd/YYYY")
                   .ColWidth(j - 1) = 2000
                ElseIf j = 7 And Not IsNull(rs.Fields(j)) Then
                    If rs.Fields(j) = 0 Then
                        value = ""
                    Else
                        value = GetParentName(rs.Fields(j))
                    End If
                    .ColWidth(j - 1) = 1600
                ElseIf j = 7 And IsNull(rs.Fields(j)) Then value = ""
                ElseIf j = 8 Then
                    value = rs.Fields(j)
                    value = Format(value, "#,###")
                    .ColWidth(j - 1) = 2100
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

Private Sub gridResults_Click()
    gridResults_SelChange
End Sub

Private Sub gridResults_DblClick()
    Dim SelectedBrand As String
    SelectedBrand = gridResults.TextMatrix(gridResults.Row, 0)
    ExecuteSQL2 "Select * from Brand where Brand_Name = '" & SelectedBrand & "'"
    If rs2.EOF Or rs2.RecordCount = 0 Then
        MsgBox "Please select a valid item", vbCritical, "Error"
        Exit Sub
    End If
    If btnEdit.Caption = "&Edit" Then
        Dim ParentName As String, ParentIndex As Integer
        Dim f As frmCreateNewBrand
        Set f = New frmCreateNewBrand
        f.txtName = rs2.Fields!Brand_Name
        f.txtOwner = rs2.Fields!Owner
        f.txtHeadquarters = rs2.Fields!Headquarter
        f.txtAreaServed = rs2.Fields!Area_Served
        f.txtWebsite = rs2.Fields!Website
        f.dtFounded = rs2.Fields!Founded
        If Not IsNull(rs2.Fields!Parent_Company) And Not rs2.Fields!Parent_Company = 0 Then
            ParentName = GetParentName(rs2.Fields!Parent_Company)
            ParentIndex = GetBrandIndex(ParentName, f)
            f.cmbParent.Text = f.cmbParent.List(ParentIndex)
        Else
            f.cmbParent.ListIndex = -1
        End If
        f.txtNumberEmployees = rs2.Fields!Number_Employees
        f.btnCreate.Caption = "&Update"
        f.btnReset.Enabled = False
        f.txtNumberEmployees_LostFocus
        f.PreviousName = rs2.Fields!Brand_Name
        f.Show vbModal, Me
    ElseIf btnEdit.Caption = "&Restore brand" Then
        rs2.Fields!Available = True
        rs2.Update
        
        MsgBox "Brand restored successfully!", vbInformation, "Information"
        'btnEdit.Caption = "&Edit"
        'btnDelete.Enabled = False
    End If
    SelectedRow = gridResults.Row
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

Private Sub btnDelete_Click()
    Dim SelectedBrand As String, ParentIndex As Integer
    SelectedBrand = gridResults.TextMatrix(gridResults.Row, 0)
    Dim f As frmDeleteBrand
    Set f = New frmDeleteBrand
    ParentIndex = GetBrandIndex(SelectedBrand, , f)
    f.cmbManufacturer.Text = f.cmbManufacturer.List(ParentIndex)
    f.Show vbModal, Me
    SelectedRow = gridResults.Row
    FillGrid
    SelectLastRow
End Sub

Private Sub btnEdit_Click()
    gridResults_DblClick
End Sub

Function GetBrandIndex(BrandName As String, Optional CreateBrandForm As frmCreateNewBrand, Optional DeleteBrandForm As frmDeleteBrand) As Integer
    Dim i As Integer, value As Integer
    If Not CreateBrandForm Is Nothing Then
        For i = 0 To CreateBrandForm.cmbParent.ListCount - 1
            If CreateBrandForm.cmbParent.List(i) = BrandName Then
                value = i
                Exit For
            End If
        Next i
    ElseIf Not DeleteBrandForm Is Nothing Then
        For i = 0 To DeleteBrandForm.cmbManufacturer.ListCount - 1
            If DeleteBrandForm.cmbManufacturer.List(i) = BrandName Then
                value = i
                Exit For
            End If
        Next i
    End If
    GetBrandIndex = value
End Function

Private Sub gridResults_SelChange()
    Dim CurrentBrand As String
    CurrentBrand = gridResults.TextMatrix(gridResults.Row, 0)
    If gridResults.Row > 0 And Not CurrentBrand = "" And Not CurrentBrand = "Brand Name" Then
        Dim SelectedBrand As String, currentBool As Boolean
        SelectedBrand = CurrentBrand
        If SelectedBrand = "" Then Exit Sub
        ExecuteSQL3 "Select * from Brand where Brand_Name = '" & SelectedBrand & "'"
        currentBool = rs3.Fields!Available
        If currentBool = False Then
            btnEdit.Caption = "&Restore brand"
            btnDelete.Enabled = False
            btnEdit.Enabled = True
        ElseIf currentBool = True Then
            btnEdit.Caption = "&Edit"
            btnDelete.Enabled = True
            btnEdit.Enabled = True
        End If
    Else
        btnDelete.Enabled = False
        btnEdit.Enabled = False
    End If
End Sub
