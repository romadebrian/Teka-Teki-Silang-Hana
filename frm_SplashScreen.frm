VERSION 5.00
Begin VB.Form frm_SplashScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14370
   Icon            =   "frm_SplashScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   14370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3000
      Top             =   6720
   End
   Begin VB.Image Image1 
      Height          =   4260
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   11055
   End
End
Attribute VB_Name = "frm_SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Dim i As Integer

Private Sub Timer1_Timer()
i = i + 1

Select Case i
    Case 1
        Image1.Picture = LoadPicture(App.Path & "\Resource\Splash Screen\1.jpg")
    Case 2
        Image1.Picture = LoadPicture(App.Path & "\Resource\Splash Screen\2.jpg")
    Case 3
        Image1.Picture = LoadPicture(App.Path & "\Resource\Splash Screen\3.jpg")
    Case 4
        Image1.Picture = LoadPicture(App.Path & "\Resource\Splash Screen\4.jpg")
    Case 5
        Sleep 500
    Case 6
        Image1.Picture = LoadPicture(App.Path & "\Resource\Splash Screen\3.jpg")
    Case 7
        Image1.Picture = LoadPicture(App.Path & "\Resource\Splash Screen\2.jpg")
    Case 8
        Image1.Picture = LoadPicture(App.Path & "\Resource\Splash Screen\1.jpg")
    Case 9
        Image1.Visible = False
    Case Else
    frm_Stage1.Show
    Unload frm_SplashScreen
End Select
End Sub
