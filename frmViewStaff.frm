VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Employees Details"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11940
   HelpContextID   =   30
   Icon            =   "frmViewStaff.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid gridResults 
      Height          =   5415
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   11175
      _ExtentX        =   19711
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
      Left            =   960
      TabIndex        =   5
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
      Left            =   8880
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   3600
      TabIndex        =   3
      Top             =   6600
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   6240
      TabIndex        =   2
      Top             =   6600
      Width           =   1935
      _ExtentX        =   3413
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
      Width           =   11415
   End
End
Attribute VB_Name = "frmViewStaff"
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
        query = "Select * from Staff where Available = False Order By Staff_LastName Asc"
        FillGrid
        btnShowHiddenElements.Caption = "Show &Both"
    ElseIf btnShowHiddenElements.Caption = "Show &Both" Then
        query = "Select * from Staff order by Staff_LastName asc"
        FillGrid
        btnShowHiddenElements.Caption = "&Hide Deleted"
    ElseIf btnShowHiddenElements.Caption = "&Hide Deleted" Then
        query = "Select * from Staff where Available = True Order By Staff_LastName Asc"
        FillGrid
        btnShowHiddenElements.Caption = "&Show Deleted"
    End If
    gridResults_SelChange
End Sub

Private Sub Form_Load()
    query = "Select * from Staff where Available = True Order By Staff_LastName Asc"
    FillGrid
    gridResults_SelChange
End Sub

Sub FillGrid()
    Dim RoleName As String
    Dim value As String
    gridResults.Clear
    ExecuteSQL query
    With gridResults
        .Cols = 7
        .FixedCols = 0
        .Rows = 0
        .AddItem "Name" & vbTab & "Last Name" & vbTab & "Identification Number" _
            & vbTab & "Phone Number" & vbTab & "Date of Birth" & vbTab & "Role" _
            & vbTab & "Username"
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
            For j = 1 To 7
                value = rs.Fields(j)
                If j = 1 Or j = 2 Then
                    .ColWidth(j - 1) = 1500
                ElseIf j = 3 Then
                    value = Format(value, "#-####-####")
                    .ColWidth(j - 1) = 1600
                ElseIf j = 4 Then
                    value = Format(value, "####-####")
                    .ColWidth(j - 1) = 1400
                ElseIf j = 5 Then
                    value = Format(value, "MM/dd/YYYY")
                    .ColWidth(j - 1) = 1500
                ElseIf j = 6 Then
                    RoleName = GetRoleName(rs.Fields(j))
                    value = RoleName
                    .ColWidth(j - 1) = 1500
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

Function GetRoleName(RoleID As Integer) As String
    ExecuteSQL3 "Select * from Role where ID = " & RoleID
    GetRoleName = rs3.Fields!Role
End Function

Function GetRoleIndex(RoleName As String, CreateUserForm As frmCreateNewStaff) As Integer
    Dim i As Integer, value As Integer
    For i = 0 To CreateUserForm.cmbRole.ListCount - 1
        If CreateUserForm.cmbRole.List(i) = RoleName Then
            value = i
            Exit For
        End If
    Next i
    GetRoleIndex = value
End Function

Function GetUserIndex(UserName As String, DeleteUserForm As frmDeleteStaff) As Integer
    Dim i As Integer, value As Integer
    For i = 0 To DeleteUserForm.cmbFullName.ListCount - 1
        If DeleteUserForm.cmbFullName.List(i) = UserName Then
            value = i
            Exit For
        End If
    Next i
    GetUserIndex = value
End Function

Private Sub gridResults_DblClick()
    Dim SelectedUser As String
    SelectedUser = Replace(gridResults.TextMatrix(gridResults.Row, 2), "-", "")
    ExecuteSQL2 "Select * from Staff where DNI = '" & SelectedUser & "'"
    If rs2.EOF Or rs2.RecordCount = 0 Then
        MsgBox "Please select a valid item", vbCritical, "Error"
        Exit Sub
    End If
    Dim RoleName As String
    Dim RoleIndex As Integer
    If btnEdit.Caption = "&Edit" Then
        Dim f As frmCreateNewStaff
        Set f = New frmCreateNewStaff
        f.txtName = rs2.Fields!Staff_Name
        f.txtLastName = rs2.Fields!Staff_LastName
        f.txtDNI = rs2.Fields!DNI
        f.txtPhoneNumber = rs2.Fields!Phone_Number
        f.dtDateBirth = rs2.Fields!DateBirth
        RoleName = GetRoleName(rs2.Fields!Role_ID)
        RoleIndex = GetRoleIndex(RoleName, f)
        f.cmbRole = f.cmbRole.List(RoleIndex)
        f.txtUsername = rs2.Fields!UserName
        f.PreviousDNI = rs2.Fields!DNI
        f.PreviousUsername = rs2.Fields!UserName
        f.btnCreate.Caption = "&Update"
        f.btnReset.Enabled = False
        f.txtDNI_LostFocus
        f.txtPhoneNumber_LostFocus
        f.Show vbModal, Me
    ElseIf btnEdit.Caption = "&Restore user" Then
        rs2.Fields!Available = True
        rs2.Fields!Role_ID = rs2.Fields!PreviousRole_ID
        rs2.Update
        
        MsgBox "User restored successfully!", vbInformation, "Information"
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

Sub SelectLastRow()
    If gridResults.Rows > SelectedRow Then
        gridResults.Row = SelectedRow
    Else
        gridResults.Row = gridResults.Rows - 1
    End If
End Sub

Private Sub btnDelete_Click()
    Dim SelectedUser As String, UserIndex As Integer
    SelectedUser = gridResults.TextMatrix(gridResults.Row, 0) & " " & gridResults.TextMatrix(gridResults.Row, 1)
    Dim f As frmDeleteStaff
    Set f = New frmDeleteStaff
    UserIndex = GetUserIndex(SelectedUser, f)
    f.cmbFullName = f.cmbFullName.List(UserIndex)
    SelectedRow = gridResults.Row
    f.Show vbModal, Me
    FillGrid
    SelectLastRow
End Sub

Private Sub btnEdit_Click()
    gridResults_DblClick
End Sub

Private Sub gridResults_SelChange()
    Dim CurrentEmployee As String
    CurrentEmployee = gridResults.TextMatrix(gridResults.Row, 2)
    If gridResults.Row > 0 And Not CurrentEmployee = "" And Not CurrentEmployee = "Identification Number" Then
        Dim SelectedUser As String, currentBool As Boolean
        SelectedUser = Replace(CurrentEmployee, "-", "")
        If SelectedUser = "" Then Exit Sub
        ExecuteSQL3 "Select * from Staff where DNI = '" & SelectedUser & "'"
        currentBool = rs3.Fields!Available
        If currentBool = False Then
            btnEdit.Caption = "&Restore user"
            btnDelete.Enabled = False
            btnEdit.Enabled = True
        ElseIf currentBool = True Then
            btnEdit.Caption = "&Edit"
            btnEdit.Enabled = True
            btnDelete.Enabled = True
        End If
    Else
        btnDelete.Enabled = False
        btnEdit.Enabled = False
    End If
End Sub
