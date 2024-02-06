VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeleteStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Employee"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7695
   HelpContextID   =   20
   Icon            =   "frmDeleteStaff.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   11
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Frame frameFilters 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   6975
      Begin VB.Frame pnlResults 
         Caption         =   "Results"
         Height          =   3015
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   6735
         Begin MSComctlLib.ListView lstResults 
            Height          =   2655
            Left            =   240
            TabIndex        =   10
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
               Text            =   "DNI"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "colLastName"
               Text            =   "Name"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Key             =   "colDNI"
               Text            =   "Last Name"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "colRole"
               Text            =   "Role"
               Object.Width           =   2646
            EndProperty
         End
      End
      Begin VB.Frame pnlFilters 
         Caption         =   "Filters"
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   6735
         Begin VB.ComboBox cmbLastName 
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
            ItemData        =   "frmDeleteStaff.frx":3AFA
            Left            =   240
            List            =   "frmDeleteStaff.frx":3AFC
            Style           =   2  'Dropdown List
            TabIndex        =   4
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
            TabIndex        =   3
            Top             =   240
            Width           =   2055
         End
         Begin MSWLess.WLCombo cmbRole 
            Height          =   390
            Left            =   3960
            TabIndex        =   6
            ToolTipText     =   "Manufacturer"
            Top             =   960
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   688
            _Version        =   393216
            Style           =   2
            ListCount       =   -19436
            Text            =   "cmbRole"
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
            List            =   "frmDeleteStaff.frx":3AFE
         End
         Begin VB.Label Label1 
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
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "Role"
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
            TabIndex        =   7
            Top             =   720
            Width           =   2535
         End
         Begin MSWLess.WLCheck chkUseAllFilters 
            Height          =   375
            Left            =   240
            TabIndex        =   5
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
   Begin MSWLess.WLCommand btnDelete 
      Height          =   975
      Left            =   2520
      TabIndex        =   14
      Top             =   6960
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
   Begin MSWLess.WLCombo cmbFullName 
      Height          =   390
      Left            =   360
      TabIndex        =   12
      ToolTipText     =   "Manufacturer"
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListCount       =   -19396
      Text            =   "cmbFullName"
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
      List            =   "frmDeleteStaff.frx":3B1A
   End
   Begin VB.Label Label6 
      Caption         =   "Employee Full Name"
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
      TabIndex        =   13
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Staff"
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
Attribute VB_Name = "frmDeleteStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDelete_Click()
    Dim userDNI As String
    Dim IsFullNameDuplicated As Boolean
    If cmbFullName.ListIndex <> -1 Then
        If MsgBox("Are you sure you want to delete the user '" _
            & cmbFullName.Text & "'?", vbExclamation + vbYesNo) = vbYes Then
            userDNI = GetUserDNI(IsFullNameDuplicated)
            If IsFullNameDuplicated Then
                MsgBox "There are more than 1 user with this name." & vbCrLf & _
                    "Please use the filters to select the right user", vbInformation, "Information"
                Exit Sub
            Else
                If IsUserLastAdministrator Then
                    MsgBox "User can't be deleted because there are no more Administrators", vbCritical, "Erorr"
                Else
                    ExecuteSQL "Select * from Staff where DNI = '" & userDNI & "'"
                    rs.Fields!Available = False
                    rs.Fields!Role_ID = 4
                    rs.Update
                    MsgBox "User deleted successfully!", vbInformation, "Information"
                    LoadFullNames
                End If
            End If
        End If
    ElseIf lstResults.ListItems.Count > 0 And Not lstResults.SelectedItem Is Nothing Then
        If MsgBox("Are you sure you want to delete the user '" _
            & GetUserFullName(lstResults.SelectedItem.Text) & "'?", vbExclamation + vbYesNo) = vbYes Then
            If IsUserLastAdministrator Then
                MsgBox "User can't be deleted because there are no more Administrators", vbCritical, "Erorr"
            Else
                ExecuteSQL "Select * from Staff where DNI = '" & lstResults.SelectedItem & "'"
                rs.Fields!Available = False
                rs.Fields!Role_ID = 4
                rs.Update
                MsgBox "User deleted successfully!", vbInformation, "Information"
                LoadLastNames
                LoadRoles
                cmdResetFilters_Click
            End If
        End If
    Else
        MsgBox "Select an user to delete it", vbInformation, "Information"
    End If
End Sub

Function IsUserLastAdministrator() As Boolean
    Dim result As Boolean
    result = False
    ExecuteSQL2 "Select * from Staff where Role_ID = 3"
    If rs2.RecordCount = 1 Then
        result = True
    End If
    IsUserLastAdministrator = result
End Function

Function GetUserDNI(ByRef IsFullNameDuplicated As Boolean) As String
    Dim dividedName As Variant
    dividedName = Split(cmbFullName.Text, " ")
    ExecuteSQL2 "Select * from Staff where Staff_Name = '" & dividedName(0) & "' and Staff_LastName = '" & dividedName(1) & "'"
    If rs2.RecordCount > 1 Then
        IsFullNameDuplicated = True
        Exit Function
    End If
    GetUserDNI = rs2.Fields!DNI
End Function

Function GetUserFullName(UserDNI As String) As String
    Dim fullName As String
    ExecuteSQL2 "Select * from Staff where DNI = '" & UserDNI & "'"
    fullName = rs2.Fields!Staff_Name & " " & rs2.Fields!Staff_LastName
    GetUserFullName = fullName
End Function

Private Sub chkFilters_Click()
    If chkFilters.value <> 0 Then
        frameFilters.Enabled = True
        frameFilters.Visible = True
        Me.Height = 8715
        btnDelete.Top = 6960
        cmbFullName.Enabled = False
        cmbFullName.ListIndex = -1
    Else
        frameFilters.Enabled = False
        frameFilters.Visible = False
        Me.Height = 4000
        btnDelete.Top = 2000
        cmbFullName.Enabled = True
    End If
End Sub

Private Sub chkUseAllFilters_Click()
    LoadLastNames
    LoadRoles
End Sub

Private Sub cmbLastName_Click()
    If chkUseAllFilters.value = wlUnchecked Then
        LoadRoles
    Else
        If cmbRole.Text = "" Then LoadRoles cmbLastName.Text
    End If
    ShowResults cmbLastName, cmbRole
End Sub

Private Sub cmbRole_Click()
    If chkUseAllFilters.value = wlUnchecked Then
        LoadLastNames
    Else
        If cmbLastName.Text = "" Then LoadLastNames cmbRole.Text
    End If
    ShowResults cmbLastName, cmbRole
End Sub

Private Sub cmdResetFilters_Click()
    LoadLastNames
    LoadRoles
    lstResults.ListItems.Clear
End Sub

Private Sub Form_Load()
    LoadFullNames
    LoadLastNames
    LoadRoles
    frameFilters.Enabled = False
    frameFilters.Visible = False
    Me.Height = 4000
    btnDelete.Top = 2000
    VerifyCurrentRole
End Sub

Sub VerifyCurrentRole()
    If frmMain.CurrentUserRoleID = 1 Then
        MsgBox "The current user does not have permission to delete employees", vbInformation, "Information"
        btnDelete.Enabled = False
    End If
End Sub

Sub LoadFullNames()
    Dim FullName As String
    ExecuteSQL "Select * from Staff where Available = True order by Staff_LastName asc"
    cmbFullName.Clear
    While Not rs.EOF
        FullName = rs.Fields!Staff_Name & " " & rs.Fields!Staff_LastName
        cmbFullName.AddItem FullName
        rs.MoveNext
    Wend
End Sub

Sub LoadLastNames(Optional Role As String)
    Dim query As String
    Dim value As String
    
    If chkUseAllFilters.value <> 0 Then
        CreateQuery query, , Role
        
        ExecuteSQL query
        
        cmbLastName.Clear
        While Not rs.EOF
            value = rs.Fields!Staff_LastName
            If Not ExistsInCombo(value, cmbLastName) Then
                cmbLastName.AddItem (value)
            End If
            rs.MoveNext
        Wend
    Else
        cmbLastName.Clear
        cmbLastName.AddItem "A-E"
        cmbLastName.AddItem "F-J"
        cmbLastName.AddItem "K-O"
        cmbLastName.AddItem "P-T"
        cmbLastName.AddItem "U-Z"
    End If
End Sub

Sub LoadRoles(Optional lastName As String)
    Dim query As String
    Dim value As String
    
    If chkUseAllFilters.value <> 0 Then
        CreateQuery query, lastName
        
        ExecuteSQL query
        
        cmbRole.Clear
        While Not rs.EOF
            value = GetRoleName(rs.Fields!Role_ID)
            If Not ExistsInCombo(value, , cmbRole) Then
                cmbRole.AddItem value
            End If
            rs.MoveNext
        Wend
    Else
        cmbRole.Clear
        cmbRole.AddItem "Seller"
        cmbRole.AddItem "Manager"
        cmbRole.AddItem "Administrator"
    End If
End Sub

Function GetRoleName(RoleID As Integer) As String
    ExecuteSQL2 "Select * from Role where ID = " & RoleID
    GetRoleName = rs2.Fields!Role
End Function

Function GetRoleID(RoleName As String) As Integer
    ExecuteSQL2 "Select * from Role where Role = '" & RoleName & "'"
    GetRoleID = rs2.Fields!ID
End Function

Sub CreateQuery(ByRef queryOut As String, Optional lastName As String, Optional Role As String)
    queryOut = "Select * from Staff where Available = True"
    
    If Not lastName = "" Then GetLastNameQuery queryOut, lastName
    If Not Role = "" Then queryOut = queryOut & " and Role_ID = " & GetRoleID(Role)
End Sub

Sub ShowResults(Optional lastName As String, Optional Role As String)
    Dim li As ListItem
    Dim RoleName As String
    Dim query As String
    On Error GoTo DoNothing
    
    CreateQuery query, lastName, Role
    
    ExecuteSQL query
    lstResults.ListItems.Clear
    While Not rs.EOF
        Set li = lstResults.ListItems.Add(, , rs.Fields!DNI)
        li.SubItems(1) = rs.Fields!Staff_Name
        li.SubItems(2) = rs.Fields!Staff_LastName
        RoleName = GetRoleName(rs.Fields!Role_ID)
        li.SubItems(3) = RoleName
        rs.MoveNext
    Wend
    Exit Sub
DoNothing:
    Exit Sub
End Sub

Sub GetLastNameQuery(ByRef queryOut, lastName As String)
    Dim Initials As Variant
    Dim SearchedCharacter As String
    If Len(lastName) = 3 And InStr(1, lastName, "-", vbTextCompare) = 2 Then
        Initials = Split(lastName, "-")
        queryOut = queryOut & " and Staff_LastName >= '" & Initials(0) & "' and Staff_LastName <= '" & Initials(1) & "'"
    Else
        queryOut = queryOut & " and Staff_LastName = '" & lastName & "'"
    End If
End Sub
