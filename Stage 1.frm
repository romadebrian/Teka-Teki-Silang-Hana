VERSION 5.00
Begin VB.Form frm_Stage1 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TTS 1"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14520
   Icon            =   "Stage 1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_check 
      BackColor       =   &H8000000D&
      Caption         =   "CHECK"
      BeginProperty Font 
         Name            =   "Corporate Logo Rounded"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton btn_hint 
      BackColor       =   &H8000000D&
      Caption         =   "HINT"
      BeginProperty Font 
         Name            =   "Corporate Logo Rounded"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txt_41 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   41
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txt_40 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   40
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txt_39 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   39
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txt_33 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   38
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txt_31 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   37
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txt_32 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   36
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txt_24 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   35
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txt_20 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   34
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txt_14 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   33
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txt_29 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   32
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txt_28 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   31
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txt_27 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   30
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txt_26 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   29
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txt_34 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   28
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txt_25 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   27
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txt_21 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   26
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txt_15 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   25
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txt_10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   24
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt_9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   23
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt_8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   22
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt_7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   21
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt_6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   20
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt_5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   19
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt_13 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   18
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txt_35 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txt_38 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   16
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txt_36 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   15
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txt_19 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txt_18 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   13
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txt_37 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   12
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txt_30 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txt_23 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txt_17 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txt_12 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txt_22 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txt_16 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txt_11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txt_4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt_3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt_2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt_focus 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox txt_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lbl_val_hint 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   12120
      TabIndex        =   46
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1695
      Left            =   10800
      Picture         =   "Stage 1.frx":C84A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lbl_hint2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   45
      Top             =   6840
      Width           =   6975
   End
   Begin VB.Label lbl_hint1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3240
      TabIndex        =   44
      Top             =   6240
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   2880
      Top             =   6080
      Width           =   8895
   End
   Begin VB.Image Image1 
      Height          =   7935
      Left            =   0
      Picture         =   "Stage 1.frx":16050
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14535
   End
   Begin VB.Image Image2 
      Height          =   4440
      Left            =   360
      Picture         =   "Stage 1.frx":25A1D
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   3570
   End
   Begin VB.Image Image4 
      Height          =   3015
      Left            =   9600
      Picture         =   "Stage 1.frx":CDEF1
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4575
   End
End
Attribute VB_Name = "frm_Stage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_check_Click()
Cek_Jawaban_Stage1
'test_ceck
End Sub

Private Sub btn_hint_Click()
'MsgBox "Oh tuhan, berikanlah hamba petunjuk"
Hint_Stage1
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub txt_1_GotFocus()
lbl_hint1.Caption = "Berhemat"
lbl_hint2.Caption = "Hewan yang hidup di air"

'if hint_mode = on then
End Sub

Private Sub txt_1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_1_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_2_GotFocus()
lbl_hint1.Caption = "Berhemat"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_2_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_3_GotFocus()
lbl_hint1.Caption = "Berhemat"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_3_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_4_GotFocus()
lbl_hint1.Caption = "Berhemat"
lbl_hint2.Caption = "Memukul pipi orang lain"
End Sub

Private Sub txt_4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_4_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_5_GotFocus()
lbl_hint1.Caption = "Hewan yang meng guk guk"
lbl_hint2.Caption = "Yang biasa turun ketika hujan"
End Sub

Private Sub txt_5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_5_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_6_GotFocus()
lbl_hint1.Caption = "Hewan yang meng guk guk"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_6_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_6_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_7_GotFocus()
lbl_hint1.Caption = "Hewan yang meng guk guk"
lbl_hint2.Caption = "Yang biasa dimiliki dukun"
End Sub

Private Sub txt_7_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_7_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_8_GotFocus()
lbl_hint1.Caption = "Hewan yang meng guk guk"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_8_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_8_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_9_GotFocus()
lbl_hint1.Caption = "Hewan yang meng guk guk"
lbl_hint2.Caption = "pineapple (indonesia)"
End Sub

Private Sub txt_9_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_9_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_10_GotFocus()
lbl_hint1.Caption = "Hewan yang meng guk guk"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_10_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_10_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_11_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Hewan yang hidup di air"
End Sub

Private Sub txt_11_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_11_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_12_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Memukul pipi orang lain"
End Sub

Private Sub txt_12_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_12_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_13_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Yang biasa turun ketika hujan"
End Sub

Private Sub txt_13_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_13_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_14_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Yang biasa dimiliki dukun"
End Sub

Private Sub txt_14_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_14_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_15_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "pineapple (indonesia)"
End Sub

Private Sub txt_15_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_15_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_16_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Hewan yang hidup di air"
End Sub

Private Sub txt_16_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_16_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_17_GotFocus()
lbl_hint1.Caption = "Bulan setelah bulan februari"
lbl_hint2.Caption = "Memukul pipi orang lain"
End Sub

Private Sub txt_17_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_17_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_18_GotFocus()
lbl_hint1.Caption = "Bulan setelah bulan februari"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_18_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_18_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_19_GotFocus()
lbl_hint1.Caption = "Bulan setelah bulan februari"
lbl_hint2.Caption = "Yang biasa turun ketika hujan"
End Sub

Private Sub txt_19_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_19_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_20_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Yang biasa dimiliki dukun"
End Sub

Private Sub txt_20_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_20_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_21_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "pineapple (indonesia)"
End Sub

Private Sub txt_21_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_21_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_22_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Hewan yang hidup di air"
End Sub

Private Sub txt_22_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_22_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_23_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Memukul pipi orang lain"
End Sub

Private Sub txt_23_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_23_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_24_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Yang biasa dimiliki dukun"
End Sub

Private Sub txt_24_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_24_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_25_GotFocus()
lbl_hint1.Caption = "Yang dihembuskan kipas"
lbl_hint2.Caption = "pineapple (indonesia)"
End Sub

Private Sub txt_25_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_25_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_26_GotFocus()
lbl_hint1.Caption = "Yang dihembuskan kipas"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_26_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_26_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_27_GotFocus()
lbl_hint1.Caption = "Yang dihembuskan kipas"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_27_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_27_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_28_GotFocus()
lbl_hint1.Caption = "Yang dihembuskan kipas"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_28_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_28_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_29_GotFocus()
lbl_hint1.Caption = "Yang dihembuskan kipas"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_29_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_29_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_30_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Memukul pipi orang lain"
End Sub

Private Sub txt_30_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_30_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_31_GotFocus()
lbl_hint1.Caption = "Antonim dari bawah"
lbl_hint2.Caption = "Hasil dari air laut yang menguap"
End Sub

Private Sub txt_31_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_31_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_32_GotFocus()
lbl_hint1.Caption = "Antonim dari bawah"
lbl_hint2.Caption = "Yang biasa dimiliki dukun"
End Sub

Private Sub txt_32_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_32_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_33_GotFocus()
lbl_hint1.Caption = "Antonim dari bawah"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_33_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_33_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_34_GotFocus()
lbl_hint1.Caption = "Antonim dari bawah"
lbl_hint2.Caption = "pineapple (indonesia)"
End Sub

Private Sub txt_34_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_34_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_35_GotFocus()
lbl_hint1.Caption = "Suku, Ras, Agama, dan Antar Golongan"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_35_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_35_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_36_GotFocus()
lbl_hint1.Caption = "Suku, Ras, Agama, dan Antar Golongan"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_36_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_36_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_37_GotFocus()
lbl_hint1.Caption = "Suku, Ras, Agama, dan Antar Golongan"
lbl_hint2.Caption = "Memukul pipi orang lain"
End Sub

Private Sub txt_37_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_37_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_38_GotFocus()
lbl_hint1.Caption = "Suku, Ras, Agama, dan Antar Golongan"
lbl_hint2.Caption = ""
End Sub

Private Sub txt_38_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_38_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_39_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Hasil dari air laut yang menguap"
End Sub

Private Sub txt_39_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_39_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_40_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Hasil dari air laut yang menguap"
End Sub

Private Sub txt_40_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_40_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub

Private Sub txt_41_GotFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = "Hasil dari air laut yang menguap"
End Sub

Private Sub txt_41_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
txt_focus.SetFocus
End Sub

Private Sub txt_41_LostFocus()
lbl_hint1.Caption = ""
lbl_hint2.Caption = ""
End Sub


