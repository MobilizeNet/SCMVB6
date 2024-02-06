Attribute VB_Name = "CommonFunctions"
Option Explicit

Public Function ExistsInCombo(v As String, Optional combo As ComboBox, Optional WLcombo As WLcombo) As Boolean
    Dim element As String
    Dim result As Boolean
    Dim i As Integer
    result = False
    If Not combo Is Nothing Then
        For i = 0 To combo.ListCount
            element = combo.List(i)
            If v = element Then
                result = True
                Exit For
            End If
        Next i
    ElseIf Not WLcombo Is Nothing Then
        For i = 0 To WLcombo.ListCount
        element = WLcombo.List(i)
        If v = element Then
            result = True
            Exit For
        End If
        Next i
    End If
    ExistsInCombo = result
End Function

Public Sub AddRequiredMark(label As Variant, colorToUse As Variant, Optional textbox As Variant, Optional combo As Variant, Optional optionArray As Variant)
    If Not IsMissing(textbox) Then
        If textbox.Text = "" Then
            label.Caption = label.Caption & " *"
            label.ForeColor = colorToUse
        End If
    ElseIf Not IsMissing(optionArray) Then
        If optionArray(0).value = 0 And optionArray(1).value = 0 Then
            label.Caption = label.Caption & " *"
            label.ForeColor = colorToUse
        End If
    ElseIf Not IsMissing(combo) Then
        If combo.ListIndex = -1 Then
            label.Caption = label.Caption & " *"
            label.ForeColor = colorToUse
        End If
    End If
End Sub

Public Sub RemoveMark(frm As Form)
    Dim ctrl As Control
    For Each ctrl In frm
        If TypeOf ctrl Is label Then
            ctrl.ForeColor = vbBlack
            ctrl.Caption = Replace(ctrl.Caption, " *", "")
        End If
    Next
End Sub

Public Function ReadFile(strFilename As String) As String
    Dim FileHandle As Integer, TextLine As String
    FileHandle = 1
    Open strFilename$ For Input As #FileHandle
    
    Do While Not EOF(FileHandle)        ' Loop until end of file
        Line Input #FileHandle, TextLine  ' Read line into variable
    Loop
        ReadFile = TextLine$
    Close #FileHandle
End Function

