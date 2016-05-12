VERSION 5.00
Begin VB.Form frnNeed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "На диску залишилось мало місця!"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmSpace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ок"
      Default         =   -1  'True
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   300
      Picture         =   "frmSpace.frx":014A
      Top             =   420
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Увага! На диску 1: залишилось мало місця. Вам необхідно його почистити."
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   420
      Width           =   3375
   End
End
Attribute VB_Name = "frnNeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Blnk As Integer

Private Sub cmdOk_Click()
Unload Me
End Sub



Private Sub tmrBlink_Timer()

End Sub

