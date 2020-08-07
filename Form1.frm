VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Apakah Form Diload ke Memory"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Load Form2"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Periksa"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  'Ganti 'Form2' dengan nama form yang akan Anda
  'periksa.
  If FormLoadedByName("Form2") = True Then
     MsgBox "Form2 sedang diload!", vbInformation, "Sedang Diload"
  Else
     MsgBox "Form2 sedang tidak di-load!", vbCritical, "Tidak Diload"
  End If
End Sub

Private Sub Command2_Click()
   Load Form2
End Sub


