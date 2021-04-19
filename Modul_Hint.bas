Attribute VB_Name = "Modul_hint"
Dim val_hint As Integer

Sub Hint_Stage1()
val_hint = frm_Stage1.lbl_val_hint.Caption

If val_hint >= 0 Then
    val_hint = val_hint - 1
    frm_Stage1.lbl_val_hint.Caption = val_hint

If frm_Stage1.txt_1.Text = "" Then
    frm_Stage1.txt_1.Text = "I"
    frm_Stage1.txt_1.SetFocus
    
ElseIf Not frm_Stage1.txt_1.Text = "I" Then
    frm_Stage1.txt_1.Text = "I"
    frm_Stage1.txt_1.SetFocus

ElseIf frm_Stage1.txt_2.Text = "" Then
    frm_Stage1.txt_2.Text = "R"
    frm_Stage1.txt_2.SetFocus

ElseIf Not frm_Stage1.txt_2.Text = "R" Then
    frm_Stage1.txt_2.Text = "R"
    frm_Stage1.txt_2.SetFocus

ElseIf frm_Stage1.txt_3.Text = "" Then
    frm_Stage1.txt_3.Text = "I"
    frm_Stage1.txt_3.SetFocus

ElseIf Not frm_Stage1.txt_3.Text = "I" Then
    frm_Stage1.txt_3.Text = "I"
    frm_Stage1.txt_3.SetFocus

ElseIf frm_Stage1.txt_4.Text = "" Then
    frm_Stage1.txt_4.Text = "T"
    frm_Stage1.txt_4.SetFocus

ElseIf Not frm_Stage1.txt_4.Text = "T" Then
    frm_Stage1.txt_4.Text = "T"
    frm_Stage1.txt_4.SetFocus
    
ElseIf frm_Stage1.txt_5.Text = "" Then
    frm_Stage1.txt_5.Text = "A"
    frm_Stage1.txt_5.SetFocus

ElseIf Not frm_Stage1.txt_5.Text = "A" Then
    frm_Stage1.txt_5.Text = "A"
    frm_Stage1.txt_5.SetFocus

ElseIf frm_Stage1.txt_6.Text = "" Then
    frm_Stage1.txt_6.Text = "N"
    frm_Stage1.txt_6.SetFocus

ElseIf Not frm_Stage1.txt_6.Text = "N" Then
    frm_Stage1.txt_6.Text = "N"
    frm_Stage1.txt_6.SetFocus

ElseIf frm_Stage1.txt_7.Text = "" Then
    frm_Stage1.txt_7.Text = "J"
    frm_Stage1.txt_7.SetFocus

ElseIf Not frm_Stage1.txt_7.Text = "J" Then
    frm_Stage1.txt_7.Text = "J"
    frm_Stage1.txt_7.SetFocus

ElseIf frm_Stage1.txt_8.Text = "" Then
    frm_Stage1.txt_8.Text = "I"
    frm_Stage1.txt_8.SetFocus

ElseIf Not frm_Stage1.txt_8.Text = "I" Then
    frm_Stage1.txt_8.Text = "I"
    frm_Stage1.txt_8.SetFocus
    
ElseIf frm_Stage1.txt_9.Text = "" Then
    frm_Stage1.txt_9.Text = "N"
    frm_Stage1.txt_9.SetFocus

ElseIf Not frm_Stage1.txt_9.Text = "N" Then
    frm_Stage1.txt_9.Text = "N"
    frm_Stage1.txt_9.SetFocus

ElseIf frm_Stage1.txt_10.Text = "" Then
    frm_Stage1.txt_10.Text = "G"
    frm_Stage1.txt_10.SetFocus

ElseIf Not frm_Stage1.txt_10.Text = "G" Then
    frm_Stage1.txt_10.Text = "G"
    frm_Stage1.txt_10.SetFocus

ElseIf frm_Stage1.txt_11.Text = "" Then
    frm_Stage1.txt_11.Text = "K"
    frm_Stage1.txt_11.SetFocus

ElseIf Not frm_Stage1.txt_11.Text = "K" Then
    frm_Stage1.txt_11.Text = "K"
    frm_Stage1.txt_11.SetFocus

ElseIf frm_Stage1.txt_12.Text = "" Then
    frm_Stage1.txt_12.Text = "A"
    frm_Stage1.txt_12.SetFocus

ElseIf Not frm_Stage1.txt_12.Text = "A" Then
    frm_Stage1.txt_12.Text = "A"
    frm_Stage1.txt_12.SetFocus

ElseIf frm_Stage1.txt_13.Text = "" Then
    frm_Stage1.txt_13.Text = "I"
    frm_Stage1.txt_13.SetFocus

ElseIf Not frm_Stage1.txt_13.Text = "I" Then
    frm_Stage1.txt_13.Text = "I"
    frm_Stage1.txt_13.SetFocus

ElseIf frm_Stage1.txt_14.Text = "" Then
    frm_Stage1.txt_14.Text = "I"
    frm_Stage1.txt_14.SetFocus

ElseIf Not frm_Stage1.txt_14.Text = "I" Then
    frm_Stage1.txt_14.Text = "I"
    frm_Stage1.txt_14.SetFocus

ElseIf frm_Stage1.txt_15.Text = "" Then
    frm_Stage1.txt_15.Text = "A"
    frm_Stage1.txt_15.SetFocus

ElseIf Not frm_Stage1.txt_15.Text = "A" Then
    frm_Stage1.txt_5.Text = "I"
    frm_Stage1.txt_5.SetFocus

ElseIf frm_Stage1.txt_16.Text = "" Then
    frm_Stage1.txt_16.Text = "A"
    frm_Stage1.txt_16.SetFocus

ElseIf Not frm_Stage1.txt_16.Text = "A" Then
    frm_Stage1.txt_16.Text = "A"
    frm_Stage1.txt_16.SetFocus

ElseIf frm_Stage1.txt_17.Text = "" Then
    frm_Stage1.txt_17.Text = "M"
    frm_Stage1.txt_17.SetFocus

ElseIf Not frm_Stage1.txt_17.Text = "M" Then
    frm_Stage1.txt_17.Text = "M"
    frm_Stage1.txt_17.SetFocus

ElseIf frm_Stage1.txt_18.Text = "" Then
    frm_Stage1.txt_18.Text = "A"
    frm_Stage1.txt_18.SetFocus

ElseIf Not frm_Stage1.txt_18.Text = "A" Then
    frm_Stage1.txt_18.Text = "A"
    frm_Stage1.txt_18.SetFocus

ElseIf frm_Stage1.txt_19.Text = "" Then
    frm_Stage1.txt_19.Text = "R"
    frm_Stage1.txt_19.SetFocus

ElseIf Not frm_Stage1.txt_19.Text = "R" Then
    frm_Stage1.txt_19.Text = "R"
    frm_Stage1.txt_19.SetFocus

ElseIf frm_Stage1.txt_20.Text = "" Then
    frm_Stage1.txt_20.Text = "M"
    frm_Stage1.txt_20.SetFocus

ElseIf Not frm_Stage1.txt_20.Text = "M" Then
    frm_Stage1.txt_20.Text = "M"
    frm_Stage1.txt_20.SetFocus

ElseIf frm_Stage1.txt_21.Text = "" Then
    frm_Stage1.txt_21.Text = "N"
    frm_Stage1.txt_21.SetFocus

ElseIf Not frm_Stage1.txt_21.Text = "N" Then
    frm_Stage1.txt_21.Text = "N"
    frm_Stage1.txt_21.SetFocus

ElseIf frm_Stage1.txt_22.Text = "" Then
    frm_Stage1.txt_22.Text = "N"
    frm_Stage1.txt_22.SetFocus

ElseIf Not frm_Stage1.txt_22.Text = "N" Then
    frm_Stage1.txt_22.Text = "N"
    frm_Stage1.txt_22.SetFocus

ElseIf frm_Stage1.txt_23.Text = "" Then
    frm_Stage1.txt_23.Text = "P"
    frm_Stage1.txt_23.SetFocus

ElseIf Not frm_Stage1.txt_23.Text = "P" Then
    frm_Stage1.txt_23.Text = "P"
    frm_Stage1.txt_23.SetFocus

ElseIf frm_Stage1.txt_24.Text = "" Then
    frm_Stage1.txt_24.Text = "A"
    frm_Stage1.txt_24.SetFocus

ElseIf Not frm_Stage1.txt_24.Text = "A" Then
    frm_Stage1.txt_24.Text = "A"
    frm_Stage1.txt_24.SetFocus

ElseIf frm_Stage1.txt_25.Text = "" Then
    frm_Stage1.txt_25.Text = "A"
    frm_Stage1.txt_25.SetFocus

ElseIf Not frm_Stage1.txt_25.Text = "A" Then
    frm_Stage1.txt_25.Text = "A"
    frm_Stage1.txt_25.SetFocus

ElseIf frm_Stage1.txt_26.Text = "" Then
    frm_Stage1.txt_26.Text = "N"
    frm_Stage1.txt_26.SetFocus

ElseIf Not frm_Stage1.txt_26.Text = "N" Then
    frm_Stage1.txt_26.Text = "N"
    frm_Stage1.txt_26.SetFocus

ElseIf frm_Stage1.txt_27.Text = "" Then
    frm_Stage1.txt_27.Text = "G"
    frm_Stage1.txt_27.SetFocus

ElseIf Not frm_Stage1.txt_27.Text = "G" Then
    frm_Stage1.txt_27.Text = "G"
    frm_Stage1.txt_27.SetFocus

ElseIf frm_Stage1.txt_28.Text = "" Then
    frm_Stage1.txt_28.Text = "I"
    frm_Stage1.txt_28.SetFocus

ElseIf Not frm_Stage1.txt_28.Text = "I" Then
    frm_Stage1.txt_28.Text = "I"
    frm_Stage1.txt_28.SetFocus

ElseIf frm_Stage1.txt_29.Text = "" Then
    frm_Stage1.txt_29.Text = "N"
    frm_Stage1.txt_29.SetFocus

ElseIf Not frm_Stage1.txt_29.Text = "N" Then
    frm_Stage1.txt_29.Text = "N"
    frm_Stage1.txt_29.SetFocus

ElseIf frm_Stage1.txt_30.Text = "" Then
    frm_Stage1.txt_30.Text = "A"
    frm_Stage1.txt_30.SetFocus

ElseIf Not frm_Stage1.txt_30.Text = "A" Then
    frm_Stage1.txt_30.Text = "A"
    frm_Stage1.txt_30.SetFocus

ElseIf frm_Stage1.txt_31.Text = "" Then
    frm_Stage1.txt_31.Text = "A"
    frm_Stage1.txt_31.SetFocus

ElseIf Not frm_Stage1.txt_31.Text = "A" Then
    frm_Stage1.txt_31.Text = "A"
    frm_Stage1.txt_31.SetFocus

ElseIf frm_Stage1.txt_32.Text = "" Then
    frm_Stage1.txt_32.Text = "T"
    frm_Stage1.txt_32.SetFocus

ElseIf Not frm_Stage1.txt_32.Text = "T" Then
    frm_Stage1.txt_32.Text = "T"
    frm_Stage1.txt_32.SetFocus

ElseIf frm_Stage1.txt_33.Text = "" Then
    frm_Stage1.txt_33.Text = "A"
    frm_Stage1.txt_33.SetFocus

ElseIf Not frm_Stage1.txt_33.Text = "A" Then
    frm_Stage1.txt_33.Text = "A"
    frm_Stage1.txt_33.SetFocus

ElseIf frm_Stage1.txt_34.Text = "" Then
    frm_Stage1.txt_34.Text = "S"
    frm_Stage1.txt_34.SetFocus

ElseIf Not frm_Stage1.txt_34.Text = "S" Then
    frm_Stage1.txt_34.Text = "S"
    frm_Stage1.txt_34.SetFocus

ElseIf frm_Stage1.txt_35.Text = "" Then
    frm_Stage1.txt_35.Text = "S"
    frm_Stage1.txt_35.SetFocus

ElseIf Not frm_Stage1.txt_35.Text = "S" Then
    frm_Stage1.txt_35.Text = "S"
    frm_Stage1.txt_35.SetFocus

ElseIf frm_Stage1.txt_36.Text = "" Then
    frm_Stage1.txt_36.Text = "A"
    frm_Stage1.txt_36.SetFocus
    
ElseIf Not frm_Stage1.txt_36.Text = "A" Then
    frm_Stage1.txt_36.Text = "A"
    frm_Stage1.txt_36.SetFocus

ElseIf frm_Stage1.txt_37.Text = "" Then
    frm_Stage1.txt_37.Text = "R"
    frm_Stage1.txt_37.SetFocus

ElseIf Not frm_Stage1.txt_37.Text = "R" Then
    frm_Stage1.txt_37.Text = "R"
    frm_Stage1.txt_37.SetFocus

ElseIf frm_Stage1.txt_38.Text = "" Then
    frm_Stage1.txt_38.Text = "A"
    frm_Stage1.txt_38.SetFocus

ElseIf Not frm_Stage1.txt_38.Text = "A" Then
    frm_Stage1.txt_38.Text = "A"
    frm_Stage1.txt_38.SetFocus

ElseIf frm_Stage1.txt_39.Text = "" Then
    frm_Stage1.txt_39.Text = "W"
    frm_Stage1.txt_39.SetFocus
    
ElseIf Not frm_Stage1.txt_39.Text = "W" Then
    frm_Stage1.txt_39.Text = "W"
    frm_Stage1.txt_39.SetFocus

ElseIf frm_Stage1.txt_40.Text = "" Then
    frm_Stage1.txt_40.Text = "A"
    frm_Stage1.txt_40.SetFocus

ElseIf Not frm_Stage1.txt_40.Text = "A" Then
    frm_Stage1.txt_40.Text = "A"
    frm_Stage1.txt_40.SetFocus

ElseIf frm_Stage1.txt_41.Text = "" Then
    frm_Stage1.txt_41.Text = "N"
    frm_Stage1.txt_41.SetFocus

ElseIf Not frm_Stage1.txt_41.Text = "N" Then
    frm_Stage1.txt_41.Text = "N"
    frm_Stage1.txt_41.SetFocus
    
Else
    MsgBox "Selamat anda berhasil menyelesaikan puzle ini"
    MsgBox "Stage selanjutnya sedang dalam tahapn pembuatan"
End If

Else
MsgBox "Hint anda sudah habis"
End If

End Sub
