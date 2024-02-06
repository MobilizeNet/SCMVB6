VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeleteBrand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Brand"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7710
   HelpContextID   =   20
   Icon            =   "frmDeleteBrand.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameFilters 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   6975
      Begin VB.Frame pnlFilters 
         Caption         =   "Filters"
         Height          =   2535
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   6735
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
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox cmbHeadquarter 
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
            ItemData        =   "frmDeleteBrand.frx":3AFA
            Left            =   240
            List            =   "frmDeleteBrand.frx":3AFC
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   960
            Width           =   2655
         End
         Begin MSWLess.WLCheck chkUseAllFilters 
            Height          =   375
            Left            =   240
            TabIndex        =   16
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
         Begin MSWLess.WLCombo cmbParent 
            Height          =   390
            Left            =   3960
            TabIndex        =   10
            ToolTipText     =   "Manufacturer"
            Top             =   960
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   688
            _Version        =   393216
            Style           =   2
            ListCount       =   -19796
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
            List            =   "frmDeleteBrand.frx":3AFE
         End
         Begin MSWLess.WLCombo cmbAreaServed 
            Height          =   390
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "Manufacturer"
            Top             =   1800
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   688
            _Version        =   393216
            Style           =   2
            ListCount       =   -19956
            Text            =   "cmbAreaServed"
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
            List            =   "frmDeleteBrand.frx":3B1A
         End
         Begin VB.Label Label2 
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
            Left            =   3960
            TabIndex        =   14
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label1 
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
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label3 
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
            Left            =   240
            TabIndex        =   12
            Top             =   1560
            Width           =   2655
         End
      End
      Begin VB.Frame pnlResults 
         Caption         =   "Results"
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   6735
         Begin MSComctlLib.ListView lstResults 
            Height          =   2655
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   4683
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "colName"
               Text            =   "Name"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "colHeadquarter"
               Text            =   "Headquarter"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Key             =   "colParent"
               Text            =   "Parent Company"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "colAreaServed"
               Text            =   "Area Served"
               Object.Width           =   2646
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
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin MSWLess.WLCombo cmbManufacturer 
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
      ListCount       =   -19396
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
      List            =   "frmDeleteBrand.frx":3B36
   End
   Begin MSWLess.WLCommand btnDelete 
      Height          =   975
      Left            =   2400
      TabIndex        =   3
      Top             =   7920
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
      TabIndex        =   1
      Top             =   960
      Width           =   3135
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
      Width           =   7215
   End
End
Attribute VB_Name = "frmDeleteBrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDelete_Click()
    If cmbManufacturer.ListIndex <> -1 Then
        If MsgBox("Are you sure you want to delete the brand '" _
            & cmbManufacturer.Text & "'?", vbExclamation + vbYesNo) = vbYes Then
            If DoesBrandHasChildren(cmbManufacturer.Text) Then
                MsgBox "The selected brand has other brands associated, it can't be deleted!" _
                    & vbCrLf & "Remove the associated brands to remove this one", vbExclamation, "Error deleting brand"
            Else
                If AreThereVehiclesInStock(cmbManufacturer) Then
                    MsgBox "The brand " & cmbManufacturer.Text & " has vehicles pending to sell!" _
                    & vbCrLf & "Remove or sell them to remove this brand", vbExclamation, "Error deleting brand"
                Else
                    ExecuteSQL "Select * from Brand where Brand_Name = '" & cmbManufacturer.Text & "'"
                    rs!Available = False
                    rs.Update
                    MsgBox "Brand deleted successfully!", vbInformation, "Information"
                    LoadManufacturers
                End If
            End If
        End If
    ElseIf lstResults.ListItems.Count > 0 And Not lstResults.SelectedItem Is Nothing Then
        If MsgBox("Are you sure you want to delete the brand '" _
            & lstResults.SelectedItem & "'?", vbExclamation + vbYesNo) = vbYes Then
            If DoesBrandHasChildren(lstResults.SelectedItem.Text) Then
                MsgBox "The selected brand has other brands associated, it can't be deleted!" _
                    & vbCrLf & "Remove the associated brands to remove this one", vbExclamation, "Error deleting brand"
            Else
                If AreThereVehiclesInStock(lstResults.SelectedItem.Text) Then
                    MsgBox "The brand " & lstResults.SelectedItem & " has vehicles pending to sell!" _
                    & vbCrLf & "Remove or sell them to remove this brand", vbExclamation, "Error deleting brand"
                Else
                    ExecuteSQL "Select * from Brand where Brand_Name = '" & lstResults.SelectedItem & "'"
                    rs!Available = False
                    rs.Update
                    MsgBox "Brand deleted successfully!", vbInformation, "Information"
                    LoadManufacturers
                    LoadHeadquarters
                    LoadParents
                    LoadAreasServed
                    cmdResetFilters_Click
                End If
            End If
        End If
    Else
        MsgBox "Select a brand to delete it", vbInformation, "Information"
    End If
End Sub

Function DoesBrandHasChildren(value As String) As Boolean
    Dim ParentID As Integer
    Dim result As Boolean
    ParentID = GetManufacturerID(value)
    ExecuteSQL2 "Select * from Brand where Parent_Company = " & ParentID
    If rs2.EOF Then
        result = False
    Else
        result = True
    End If
    DoesBrandHasChildren = result
End Function

Function AreThereVehiclesInStock(value As String) As Boolean
    Dim BrandID As Integer
    Dim result As Boolean
    Dim VehicleQuantity As Integer
    result = False
    BrandID = GetManufacturerID(value)
    ExecuteSQL2 "Select * from Vehicle where Manufacturer_ID = " & BrandID
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
        Me.Height = 9690
        btnDelete.Top = 7920
        cmbManufacturer.Enabled = False
        cmbManufacturer.ListIndex = -1
    Else
        frameFilters.Enabled = False
        frameFilters.Visible = False
        Me.Height = 4000
        btnDelete.Top = 2000
        cmbManufacturer.Enabled = True
    End If
End Sub

Private Sub cmbAreaServed_Click()
    If cmbAreaServed.Text <> "" Then
        If chkUseAllFilters.value = 0 Then
            LoadParents
            LoadHeadquarters
        Else
            If cmbParent.Text = "" Then LoadParents cmbHeadquarter.Text, cmbAreaServed.Text
            If cmbHeadquarter.Text = "" Then LoadHeadquarters cmbParent.Text, cmbAreaServed.Text
        End If
        ShowResults cmbHeadquarter.Text, cmbParent.Text, cmbAreaServed.Text
    End If
End Sub

Private Sub cmbHeadquarter_Click()
    If cmbHeadquarter = "" Then
        Exit Sub
    End If
    If chkUseAllFilters.value = 0 Then
        LoadParents
        LoadAreasServed
    Else
        If cmbParent = "" Then LoadParents cmbHeadquarter.Text, cmbAreaServed.Text
        If cmbAreaServed = "" Then LoadAreasServed cmbHeadquarter.Text, cmbParent.Text
    End If
    ShowResults cmbHeadquarter.Text, cmbParent.Text, cmbAreaServed.Text
End Sub

Private Sub cmbParent_Click()
    If cmbParent <> "" Then
        If chkUseAllFilters.value = 0 Then
            LoadAreasServed
            LoadHeadquarters
        Else
            If cmbHeadquarter = "" Then LoadHeadquarters cmbParent.Text, cmbAreaServed.Text
            If cmbAreaServed.Text = "" Then LoadAreasServed cmbHeadquarter.Text, cmbParent.Text
        End If
        ShowResults cmbHeadquarter.Text, cmbParent.Text, cmbAreaServed.Text
    End If
End Sub

Function GetManufacturerID(ManufacturerName As String) As Integer
    ExecuteSQL3 "Select * from Brand where Brand_Name = '" & ManufacturerName & "'"
    GetManufacturerID = rs3.Fields!ID
End Function

Sub ShowResults(Optional HQValue As String, Optional Parent As String, Optional Area As String)
    Dim li As ListItem
    Dim ParentName As String
    Dim query As String
    On Error GoTo DoNothing
    
    CreateQuery query, HQValue, Parent, Area
    
    ExecuteSQL query
    lstResults.ListItems.Clear
    While Not rs.EOF
        Set li = lstResults.ListItems.Add(, , rs.Fields!Brand_Name)
        li.SubItems(1) = rs.Fields!Headquarter
        If Not IsNull(rs.Fields!Parent_Company) And rs.Fields!Parent_Company <> 0 Then
            ParentName = GetParentName(rs.Fields!Parent_Company)
            li.SubItems(2) = ParentName
        End If
        li.SubItems(3) = rs.Fields!Area_Served
        rs.MoveNext
    Wend
    Exit Sub
DoNothing:
    Exit Sub
End Sub

Private Sub cmdResetFilters_Click()
    cmbAreaServed.ListIndex = -1
    cmbHeadquarter.ListIndex = -1
    cmbParent.ListIndex = -1
    lstResults.ListItems.Clear
End Sub

Private Sub Form_Load()
    LoadManufacturers
    LoadHeadquarters
    LoadParents
    LoadAreasServed
    frameFilters.Enabled = False
    frameFilters.Visible = False
    Me.Height = 4000
    btnDelete.Top = 2000
    VerifyCurrentRole
End Sub

Sub VerifyCurrentRole()
    If frmMain.CurrentUserRoleID = 1 Then
        MsgBox "The current user does not have permission to delete brands", vbInformation, "Information"
        btnDelete.Enabled = False
    End If
End Sub

Sub LoadManufacturers()
    ExecuteSQL2 "Select * from Brand where Available = True order by Brand_Name asc "
    cmbManufacturer.Clear
    While Not rs2.EOF
        cmbManufacturer.AddItem (rs2.Fields!Brand_Name)
        rs2.MoveNext
    Wend
End Sub

Function GetParentName(ParentID As Integer) As String
    ExecuteSQL3 "Select * from Brand where ID = " & ParentID
    GetParentName = rs3.Fields!Brand_Name
End Function

Sub LoadParents(Optional HeadquarterLocation As String, Optional AreaServed As String)
    Dim ParentName As String
    Dim query As String
    CreateQuery query, HeadquarterLocation, , AreaServed
    
    ExecuteSQL2 query
    
    cmbParent.Clear
    While Not rs2.EOF
        If Not IsNull(rs2.Fields!Parent_Company) And rs2.Fields!Parent_Company <> 0 Then
            ParentName = GetParentName(rs2.Fields!Parent_Company)
            If Not ExistsInCombo(ParentName, , cmbParent) Then
                cmbParent.AddItem (ParentName)
            End If
        End If
        rs2.MoveNext
    Wend
End Sub

Sub LoadHeadquarters(Optional ParentCompany As String, Optional AreaServed As String)
    Dim query As String
    Dim value As String
    CreateQuery query, , ParentCompany, AreaServed
    
    ExecuteSQL query
    
    cmbHeadquarter.Clear
    While Not rs.EOF
        value = rs.Fields!Headquarter
        If Not ExistsInCombo(value, cmbHeadquarter) Then
            cmbHeadquarter.AddItem (value)
        End If
        rs.MoveNext
    Wend
End Sub

Sub LoadAreasServed(Optional HeadquarterLocation As String, Optional ParentCompany As String)
    Dim value As String
    Dim query As String
    CreateQuery query, HeadquarterLocation, ParentCompany
    
    ExecuteSQL query
    
    cmbAreaServed.Clear
    While Not rs.EOF
        value = rs.Fields!Area_Served
        If Not ExistsInCombo(value, , cmbAreaServed) Then
            cmbAreaServed.AddItem (value)
        End If
        rs.MoveNext
    Wend
End Sub

Sub CreateQuery(ByRef queryOut As String, Optional HQValue As String, Optional Parent As String, Optional Area As String)
    queryOut = "Select * from Brand where Available = True"
    
    If Not HQValue = "" Then queryOut = queryOut & " and Headquarter = '" & HQValue & "'"
    If Not Parent = "" Then queryOut = queryOut & " and Parent_Company = " & GetManufacturerID(Parent)
    If Not Area = "" Then queryOut = queryOut & " and Area_Served = '" & Area & "'"
End Sub

