VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Apakah Mouse Terinstall"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPeriksa 
      Caption         =   "Periksa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function CheckMouse() As Boolean
   If GetSystemMetrics(SM_CMOUSEBUTTONS) > 0 Then
      CheckMouse = True  'Mouse terinstall
   Else
      CheckMouse = False 'Mouse tidak terinstall
   End If
End Function

Private Sub cmdPeriksa_Click()
    If CheckMouse Then
        MsgBox "Mouse terinstall di PC Anda!", vbInformation, "Terinstall"
    Else
        MsgBox "Mouse belum terinstall di PC Anda!", vbCritical, "Belum Terinstall"
    End If
End Sub

