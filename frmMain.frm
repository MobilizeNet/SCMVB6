VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   5985
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   10140
   HelpContextID   =   10
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblRole 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1080
      TabIndex        =   1
      Top             =   3000
      Width           =   7815
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   7815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuChangeUser 
         Caption         =   "Change User"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuManage 
      Caption         =   "Manage"
      Begin VB.Menu mnuBrands 
         Caption         =   "Brands"
         Begin VB.Menu mnuCreateBrand 
            Caption         =   "Create new"
         End
         Begin VB.Menu mnuDeleteBrand 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu mnuVehicles 
         Caption         =   "Vehicles"
         Begin VB.Menu mnuCreateVehicle 
            Caption         =   "Create new"
         End
         Begin VB.Menu mnuDeleteVehicle 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu mnuStaff 
         Caption         =   "Staff"
         Begin VB.Menu mnuCreateStaff 
            Caption         =   "Create new"
         End
         Begin VB.Menu mnuDeleteStaff 
            Caption         =   "Delete"
         End
      End
      Begin VB.Menu mnuReceipts 
         Caption         =   "Receipts"
         Begin VB.Menu menuCreateReceipts 
            Caption         =   "Create new"
         End
      End
   End
   Begin VB.Menu mnuConsults 
      Caption         =   "Consults"
      Begin VB.Menu mnuShowBrands 
         Caption         =   "Brands"
      End
      Begin VB.Menu mnuShowVehicles 
         Caption         =   "Vehicles"
      End
      Begin VB.Menu mnuShowStaff 
         Caption         =   "Staff"
      End
      Begin VB.Menu mnuShowReceipts 
         Caption         =   "Receipts"
      End
      Begin VB.Menu mnuDetailedInformation 
         Caption         =   "Detailed Information"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CurrentUserRoleID As Integer

Private Sub Form_Load()
    Dim f As frmLogin
    Set f = New frmLogin
    f.Show vbModal, Me
    If lblUser.Caption = "" Then
        If Not f.LoginSucceeded Then
            Unload Me
        End If
    End If
End Sub

Private Sub menuCreateReceipts_Click()
    Dim f As frmCreateNewReceipt
    Set f = New frmCreateNewReceipt
    f.Show vbModal, Me
End Sub

Private Sub mnuAbout_Click()
    Dim f As frmAbout
    Set f = New frmAbout
    f.Show vbModal, Me
End Sub

Private Sub mnuChangeUser_Click()
    Form_Load
End Sub

Private Sub mnuCreateBrand_Click()
    Dim f As frmCreateNewBrand
    Set f = New frmCreateNewBrand
    f.Show vbModal, Me
End Sub

Private Sub mnuCreateStaff_Click()
    Dim f As frmCreateNewStaff
    Set f = New frmCreateNewStaff
    f.Show vbModal, Me
End Sub

Private Sub mnuCreateVehicle_Click()
    Dim f As frmCreateNewVehicle
    Set f = New frmCreateNewVehicle
    f.Show vbModal, Me
End Sub

Private Sub mnuDeleteBrand_Click()
    Dim f As frmDeleteBrand
    Set f = New frmDeleteBrand
    f.Show vbModal, Me
End Sub

Private Sub mnuDeleteStaff_Click()
    Dim f As frmDeleteStaff
    Set f = New frmDeleteStaff
    f.Show vbModal, Me
End Sub

Private Sub mnuDeleteVehicle_Click()
    Dim f As frmDeleteVehicle
    Set f = New frmDeleteVehicle
    f.Show vbModal, Me
End Sub

Private Sub mnuDetailedInformation_Click()
    Dim f As frmDetailedInformation
    Set f = New frmDetailedInformation
    f.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuShowBrands_Click()
    Dim f As frmViewBrands
    Set f = New frmViewBrands
    f.Show vbModal, Me
End Sub

Private Sub mnuShowReceipts_Click()
    Dim f As frmViewReceipts
    Set f = New frmViewReceipts
    f.Show vbModal, Me
End Sub

Private Sub mnuShowStaff_Click()
    Dim f As frmViewStaff
    Set f = New frmViewStaff
    f.Show vbModal, Me
End Sub

Private Sub mnuShowVehicles_Click()
    Dim f As frmViewVehicles
    Set f = New frmViewVehicles
    f.Show vbModal, Me
End Sub
