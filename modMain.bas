Attribute VB_Name = "modMain"
Option Explicit

Public connString As String
Public conn As ADODB.Connection
Public rs As ADODB.Recordset, rs2 As ADODB.Recordset, rs3 As ADODB.Recordset

Public Sub OpenConnection()
    Set conn = New ADODB.Connection
    conn.Open connString
End Sub

Sub main()
    App.HelpFile = App.Path + "\SCM_Help.chm"
    connString = ReadFile("db.txt") + App.Path + "\database.mdb;Persist Security Info=False"
    
    OpenConnection
    On Error GoTo FinishExecution
    frmMain.Show
FinishExecution:
    Exit Sub
End Sub

Public Sub ExecuteSQL(query As String)
    Set rs = New ADODB.Recordset
    rs.Open query, conn, adOpenKeyset, adLockPessimistic
End Sub

Public Sub ExecuteSQL2(query As String)
    Set rs2 = New ADODB.Recordset
    rs2.Open query, conn, adOpenKeyset, adLockPessimistic
End Sub

Public Sub ExecuteSQL3(query As String)
    Set rs3 = New ADODB.Recordset
    rs3.Open query, conn, adOpenKeyset, adLockPessimistic
End Sub
