Attribute VB_Name = "Module_Cek_Jawaban"
Sub SelectAllText(tb As TextBox)

tb.SelStart = 0
tb.SelLength = Len(tb.Text)

End Sub

Sub Cek_Jawaban_Stage1()

If frm_Stage1.txt_1.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_1.SetFocus
    
ElseIf Not frm_Stage1.txt_1.Text = "I" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_1.SetFocus

ElseIf frm_Stage1.txt_2.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_2.SetFocus

ElseIf Not frm_Stage1.txt_2.Text = "R" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_2.SetFocus

ElseIf frm_Stage1.txt_3.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_3.SetFocus

ElseIf Not frm_Stage1.txt_3.Text = "I" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_3.SetFocus

ElseIf frm_Stage1.txt_4.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_4.SetFocus

ElseIf Not frm_Stage1.txt_4.Text = "T" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_4.SetFocus
    
ElseIf frm_Stage1.txt_5.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_5.SetFocus

ElseIf Not frm_Stage1.txt_5.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_5.SetFocus

ElseIf frm_Stage1.txt_6.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_6.SetFocus

ElseIf Not frm_Stage1.txt_6.Text = "N" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_6.SetFocus

ElseIf frm_Stage1.txt_7.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_7.SetFocus

ElseIf Not frm_Stage1.txt_7.Text = "J" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_7.SetFocus

ElseIf frm_Stage1.txt_8.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_8.SetFocus

ElseIf Not frm_Stage1.txt_8.Text = "I" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_8.SetFocus
    
ElseIf frm_Stage1.txt_9.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_9.SetFocus

ElseIf Not frm_Stage1.txt_9.Text = "N" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_9.SetFocus

ElseIf frm_Stage1.txt_10.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_10.SetFocus

ElseIf Not frm_Stage1.txt_10.Text = "G" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_10.SetFocus

ElseIf frm_Stage1.txt_11.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_11.SetFocus

ElseIf Not frm_Stage1.txt_11.Text = "K" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_11.SetFocus

ElseIf frm_Stage1.txt_12.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_12.SetFocus

ElseIf Not frm_Stage1.txt_12.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_12.SetFocus

ElseIf frm_Stage1.txt_13.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_13.SetFocus

ElseIf Not frm_Stage1.txt_13.Text = "I" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_13.SetFocus

ElseIf frm_Stage1.txt_14.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_14.SetFocus

ElseIf Not frm_Stage1.txt_14.Text = "I" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_14.SetFocus

ElseIf frm_Stage1.txt_15.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_15.SetFocus

ElseIf Not frm_Stage1.txt_15.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_15.SetFocus

ElseIf frm_Stage1.txt_16.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_16.SetFocus

ElseIf Not frm_Stage1.txt_16.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_16.SetFocus

ElseIf frm_Stage1.txt_17.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_17.SetFocus

ElseIf Not frm_Stage1.txt_17.Text = "M" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_17.SetFocus

ElseIf frm_Stage1.txt_18.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_18.SetFocus

ElseIf Not frm_Stage1.txt_18.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_18.SetFocus

ElseIf frm_Stage1.txt_19.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_19.SetFocus

ElseIf Not frm_Stage1.txt_19.Text = "R" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_19.SetFocus

ElseIf frm_Stage1.txt_20.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_20.SetFocus

ElseIf Not frm_Stage1.txt_20.Text = "M" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_20.SetFocus

ElseIf frm_Stage1.txt_21.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_21.SetFocus

ElseIf Not frm_Stage1.txt_21.Text = "N" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_21.SetFocus

ElseIf frm_Stage1.txt_22.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_22.SetFocus

ElseIf Not frm_Stage1.txt_22.Text = "N" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_22.SetFocus

ElseIf frm_Stage1.txt_23.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_23.SetFocus

ElseIf Not frm_Stage1.txt_23.Text = "P" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_23.SetFocus

ElseIf frm_Stage1.txt_24.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_24.SetFocus

ElseIf Not frm_Stage1.txt_24.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_24.SetFocus

ElseIf frm_Stage1.txt_25.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_25.SetFocus

ElseIf Not frm_Stage1.txt_25.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_25.SetFocus

ElseIf frm_Stage1.txt_26.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_26.SetFocus

ElseIf Not frm_Stage1.txt_26.Text = "N" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_26.SetFocus

ElseIf frm_Stage1.txt_27.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_27.SetFocus

ElseIf Not frm_Stage1.txt_27.Text = "G" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_27.SetFocus

ElseIf frm_Stage1.txt_28.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_28.SetFocus

ElseIf Not frm_Stage1.txt_28.Text = "I" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_28.SetFocus

ElseIf frm_Stage1.txt_29.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_2.SetFocus

ElseIf Not frm_Stage1.txt_29.Text = "N" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_29.SetFocus

ElseIf frm_Stage1.txt_30.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_30.SetFocus

ElseIf Not frm_Stage1.txt_30.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_30.SetFocus

ElseIf frm_Stage1.txt_31.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_31.SetFocus

ElseIf Not frm_Stage1.txt_31.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_31.SetFocus

ElseIf frm_Stage1.txt_32.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_32.SetFocus

ElseIf Not frm_Stage1.txt_32.Text = "T" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_32.SetFocus

ElseIf frm_Stage1.txt_33.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_33.SetFocus

ElseIf Not frm_Stage1.txt_33.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_33.SetFocus

ElseIf frm_Stage1.txt_34.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_34.SetFocus

ElseIf Not frm_Stage1.txt_34.Text = "S" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_34.SetFocus

ElseIf frm_Stage1.txt_35.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_35.SetFocus

ElseIf Not frm_Stage1.txt_35.Text = "S" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_35.SetFocus

ElseIf frm_Stage1.txt_36.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_36.SetFocus
    
ElseIf Not frm_Stage1.txt_36.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_36.SetFocus

ElseIf frm_Stage1.txt_37.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_37.SetFocus

ElseIf Not frm_Stage1.txt_37.Text = "R" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_37.SetFocus

ElseIf frm_Stage1.txt_38.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_38.SetFocus

ElseIf Not frm_Stage1.txt_38.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_38.SetFocus

ElseIf frm_Stage1.txt_39.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_39.SetFocus
    
ElseIf Not frm_Stage1.txt_39.Text = "W" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_39.SetFocus

ElseIf frm_Stage1.txt_40.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_40.SetFocus

ElseIf Not frm_Stage1.txt_40.Text = "A" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_40.SetFocus

ElseIf frm_Stage1.txt_41.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_41.SetFocus

ElseIf Not frm_Stage1.txt_41.Text = "N" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_41.SetFocus
    
Else
    MsgBox "Selamat anda berhasil menyelesaikan puzle ini"
    frm_Stage2.Show
    frm_Stage1.Hide
End If

End Sub

Sub test_ceck()
If frm_Stage1.txt_1.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_1.SetFocus
    
ElseIf Not frm_Stage1.txt_1.Text = "I" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_1.SetFocus

ElseIf frm_Stage1.txt_2.Text = "" Then
    MsgBox "Jawaban anda kosong"
    frm_Stage1.txt_2.SetFocus
    
ElseIf Not frm_Stage1.txt_2.Text = "R" Then
    MsgBox "Ada jawaban yang salah"
    frm_Stage1.txt_2.SetFocus
    
Else
    MsgBox "Selamat anda berhasil menyelesaikan puzle ini"
    MsgBox "Stage selanjutnya sedang dalam tahapn pembuatan"
End If
End Sub


