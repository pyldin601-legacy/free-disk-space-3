VERSION 5.00
Begin VB.Form cfg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Настройка программы"
   ClientHeight    =   3840
   ClientLeft      =   3510
   ClientTop       =   5250
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "cfg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Стиль текста"
      Height          =   3015
      Left            =   3480
      TabIndex        =   4
      Top             =   180
      Width           =   3075
      Begin VB.Frame Frame3 
         Caption         =   "Пример"
         Height          =   1095
         Left            =   180
         TabIndex        =   10
         Top             =   1680
         Width           =   2715
         Begin VB.Label Label2 
            Caption         =   "123,456 GB AaBbCcDdEe"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   60
            TabIndex        =   11
            Top             =   240
            Width           =   2565
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1260
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   840
         ScaleHeight     =   915
         ScaleWidth      =   1455
         TabIndex        =   5
         Top             =   300
         Width           =   1455
         Begin VB.OptionButton Option1 
            Caption         =   "Тень"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Контур"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Размер шрифта:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Отображать диски"
      Height          =   3015
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   3135
      Begin VB.ListBox lstDscs 
         Columns         =   3
         Height          =   2535
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   300
         Width           =   2775
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4020
      TabIndex        =   0
      Top             =   3300
      Width           =   1215
   End
End
Attribute VB_Name = "cfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub



Private Sub Combo1_Click()
Label2.FontSize = Val(Combo1.Text)
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Z, BZ As Integer
Dim ConfigFile As String, ZM, ZN
Me.Show

For Z = 0 To 25
  lstDscs.List(Z) = Chr(65 + Z) + ":"
Next

For Z = 8 To 25
  Combo1.AddItem Format(Z, "0")
Next Z

ConfigFile = LowPath(App.Path) + "FDSI.cfg"

Dim DrvCount As Integer
Dim Intervl As Integer
Dim Inform As Integer
Dim DiskName As String
Dim TopMost

DrvCount = Val(GetIniRecord("DRIVES=", ConfigFile, "1"))
Combo1.Text = GetIniRecord("SIZE=", ConfigFile, "10")

Dim tmpA As String

tmpA = GetIniRecord("FONT=", ConfigFile, "0")

If Val(tmpA) = 0 Then Option1(0).Value = True Else Option1(1).Value = True

For Z = 0 To DrvCount - 1
 DiskName = GetIniRecord("DRV_" + Format(Z, "0") + "=", ConfigFile, "C")
 
 For BZ = 0 To lstDscs.ListCount - 1
  If UCase(Left(lstDscs.List(BZ), 1)) = UCase(DiskName) Then lstDscs.Selected(BZ) = True
 Next
Next

End Sub


Private Sub OKButton_Click()
' On Error Resume Next
Dim ConfigFile As String, ZM, ZN

ConfigFile = LowPath(App.Path) + "FDSI.cfg"

Open ConfigFile For Output As #1
 Print #1, "DRIVES=" + Str(lstDscs.SelCount)
 Print #1, "SIZE=" + Combo1.Text
 If Option1(0).Value = True Then Print #1, "FONT=0" Else Print #1, "FONT=1"
 
 Print #1, ""
 
 ZN = -1
 
 For ZM = 0 To lstDscs.ListCount - 1
  If lstDscs.Selected(ZM) = True Then
   ZN = ZN + 1
   Print #1, "DRV_" + Format(ZN, "0") + "=" + UCase(Left(lstDscs.List(ZM), 1))
  End If
 Next
 
Close #1


FDSI_2.ReConf

Unload Me

End Sub



