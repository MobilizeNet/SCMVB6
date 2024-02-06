VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Begin VB.Form frmGraphics 
   Caption         =   "Detailed Graphics"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   16335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
   Begin MSWLess.WLCommand btnDelete 
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   5760
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      _Version        =   393216
      Caption         =   "Delete"
      ForeColor       =   -1
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDelete_Click()
    DrawLineGraph
End Sub

Private Sub DrawLineGraph()
  Dim I As Long
  With Draw1
    'Graph is a line graph
    .GraphType = dgtLine

    'Set the origin and axes
    .OriginY = 150
    .YAxisNegative = 100
    .YTop = 50
    .YGrad = 10
    .XTop = 11
    .XGrad = 1

    'Set some properties based on the check boxes
    If Check11 Then
      .LineWidth = 2
    Else
      .LineWidth = 1
    End If
    If Check12 Then
      .PointSize = 4
      .PointStyle = dgpsDot
    Else
      .PointSize = 0
    End If
      .ShowLine = Check13.value
      .ShowGrid = Check14.value
      .ShowLegend = Check15.value

    'This option replaces the x-axis values with the values
    'of the Months array
    If Check16 Then
      .UseXAxisLabels = True
      For I = 0 To 11
        .AddXValue I, Months(I)
      Next I
    Else
      .UseXAxisLabels = False
    End If

    'Add the data points
    For I = 0 To 11
      .AddPoint I, Line1Values(I), vbRed, "Line 1"
      .AddPoint I, Line2Values(I), vbBlue, "Line 2"
    Next I

    'Draw the graph
    .DrawGraph
  End With
End Sub
