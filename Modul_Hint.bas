Attribute VB_Name = "Modul_hint"
Dim val_hint As Integer

Sub Hint_Stage1()
Dim i1 As Integer
val_hint = 20

If val_hint >= 0 Then
    If frm_Stage1.got Then
        frm_Stage1.txt_1.Text = "I"
        val_hint = val_hint - 1
        frm_Stage1.lbl_val_hint.Caption = val_hint
    Else
    End If
Else
    MsgBox "Hint sudah habis"
End If

For i1 = 2 To 41

    

Next i


End Sub
