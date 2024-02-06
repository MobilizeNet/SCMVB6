VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetailedInformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detailed Information"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15645
   HelpContextID   =   30
   Icon            =   "frmDetailedInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   15645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12720
      TabIndex        =   6
      Top             =   6360
      Width           =   2535
   End
   Begin VB.CommandButton cmdSoldByModel 
      Caption         =   "Most Sold by &Model"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12720
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdSoldByManufacturer 
      Caption         =   "Most Sold by M&anufacturer"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12720
      TabIndex        =   4
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdCompaniesByCountry 
      Caption         =   "Most &Companies by Country"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12720
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton cmdSoldBySeller 
      Caption         =   "Sold by &Seller"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12720
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   360
      ScaleHeight     =   6315
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   960
      Width           =   12015
      Begin MSComctlLib.ListView lstResults 
         Height          =   6015
         Left            =   2400
         TabIndex        =   7
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   10610
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
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Detailed Queries"
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
      TabIndex        =   1
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "frmDetailedInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim query As String, SellerName As String
Dim li As ListItem
    
Private Sub cmdCompaniesByCountry_Click()
    ClearListView
    query = "SELECT Count(Brand.Brand_Name) AS CountOfBrand_Name, Brand.Headquarter " & _
            "From Brand " & _
            "GROUP BY Brand.Headquarter " & _
            "ORDER BY Count(Brand.Brand_Name) DESC"
    
    ExecuteSQL query
    
    lstResults.ColumnHeaders.Add , , "Country", 4500
    lstResults.ColumnHeaders.Add , , "Quantity", 1500
    
    While Not rs.EOF
        Set li = lstResults.ListItems.Add(, , rs.Fields!Headquarter)
        li.SubItems(1) = rs.Fields!CountOfBrand_Name
        rs.MoveNext
    Wend
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSoldByManufacturer_Click()
    ClearListView
    query = "SELECT DISTINCTROW Brand.Brand_Name, Sum(Receipt.Quantity) AS TotalSold " & _
            "FROM Brand INNER JOIN Receipt ON Brand.[ID] = Receipt.[Manufacturer_ID] " & _
            "WHERE Receipt.Status = 'Approved' " & _
            "GROUP BY Brand.Brand_Name " & _
            "ORDER BY Sum(Receipt.Quantity) DESC"
    
    ExecuteSQL query
    
    lstResults.ColumnHeaders.Add , , "Manufacturer", 4500
    lstResults.ColumnHeaders.Add , , "Quantity", 1500
    
    While Not rs.EOF
        Set li = lstResults.ListItems.Add(, , rs.Fields!Brand_Name)
        li.SubItems(1) = Format(rs.Fields!TotalSold, "#,###")
        rs.MoveNext
    Wend
End Sub

Private Sub cmdSoldByModel_Click()
    ClearListView
    query = "SELECT DISTINCTROW Vehicle.Vehicle_Name, Sum(Receipt.Quantity) AS TotalSold " & _
            "FROM Vehicle INNER JOIN Receipt ON Vehicle.[ID] = Receipt.[Model_ID] " & _
            "WHERE (((Receipt.Status)='Approved')) " & _
            "GROUP BY Receipt.Model_ID, Vehicle.ID, Vehicle.Vehicle_Name " & _
            "ORDER BY Sum(Receipt.Quantity) DESC "
    
    ExecuteSQL query
    
    lstResults.ColumnHeaders.Add , , "Model", 4500
    lstResults.ColumnHeaders.Add , , "Quantity", 1500
    
    While Not rs.EOF
        Set li = lstResults.ListItems.Add(, , rs.Fields!Vehicle_Name)
        li.SubItems(1) = Format(rs.Fields!TotalSold, "#,###")
        rs.MoveNext
    Wend

End Sub

Sub ClearListView()
    lstResults.ListItems.Clear
    lstResults.ColumnHeaders.Clear
End Sub

Private Sub cmdSoldBySeller_Click()
    ClearListView
    query = "SELECT DISTINCTROW Staff.DNI, Sum(Receipt.Quantity) AS TotalSold " & _
            "FROM Staff INNER JOIN Receipt ON Staff.[ID] = Receipt.[Seller_ID] " & _
            "WHERE Receipt.Status = 'Approved' " & _
            "GROUP BY Staff.DNI, Receipt.Seller_ID, Staff.ID " & _
            "ORDER BY Sum(Receipt.Quantity) DESC "
    
    ExecuteSQL query
    
    lstResults.ColumnHeaders.Add , , "Seller DNI", 2000
    lstResults.ColumnHeaders.Add , , "Seller Full Name", 3500
    lstResults.ColumnHeaders.Add , , "Quantity", 1500
    
    While Not rs.EOF
        Set li = lstResults.ListItems.Add(, , Format(rs.Fields!DNI, "#-####-####"))
        SellerName = GetSellerName(rs.Fields!DNI)
        li.SubItems(1) = SellerName
        li.SubItems(2) = Format(rs.Fields!TotalSold, "#,###")
        rs.MoveNext
    Wend
End Sub

Function GetSellerName(SellerDNI As String) As String
    ExecuteSQL3 "Select * from Staff where DNI = '" & SellerDNI & "'"
    GetSellerName = rs3.Fields!Staff_Name & " " & rs3.Fields!Staff_LastName
End Function
