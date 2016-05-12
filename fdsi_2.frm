VERSION 5.00
Begin VB.Form FDSI_2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   885
   ClientLeft      =   5250
   ClientTop       =   12135
   ClientWidth     =   6285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "fdsi_2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   900
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3420
      Picture         =   "fdsi_2.frx":0CCA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmr 
      Interval        =   250
      Left            =   1380
      Top             =   360
   End
   Begin VB.Label dsc 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "999 999 999 999 999 bytes free on drive C:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.Menu filo 
      Caption         =   "fl"
      Visible         =   0   'False
      Begin VB.Menu mnuTITLE 
         Caption         =   "Woobind Disk Space Info"
         Enabled         =   0   'False
      End
      Begin VB.Menu poiuyt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSELDISKS 
         Caption         =   "Диски"
         Begin VB.Menu mnuDRIVES 
            Caption         =   "C:"
            Index           =   0
         End
      End
      Begin VB.Menu mnuVIEW 
         Caption         =   "Отображение"
         Begin VB.Menu mnuLABEL 
            Caption         =   "Имена томов вместо букв дисков"
         End
         Begin VB.Menu mnuAPPROX 
            Caption         =   "Не округлять значения"
         End
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTEXTILE 
         Caption         =   "Стиль текста"
         Begin VB.Menu mnuTSTYLE 
            Caption         =   "Контур"
            Index           =   0
         End
         Begin VB.Menu mnuTSTYLE 
            Caption         =   "Тень"
            Index           =   1
         End
      End
      Begin VB.Menu mnuCOLOR 
         Caption         =   "Инвертировать цвета"
      End
      Begin VB.Menu mnuCLICKS 
         Caption         =   "Прозрачный режим"
      End
      Begin VB.Menu mnuHIDE 
         Caption         =   "Показывать только при наведении курсора"
      End
      Begin VB.Menu sdfsfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе..."
      End
      Begin VB.Menu mnuAUTO 
         Caption         =   "Запускать вместе с Windows"
      End
      Begin VB.Menu sepp 
         Caption         =   "-"
      End
      Begin VB.Menu fvig 
         Caption         =   "Закрыть"
      End
   End
End
Attribute VB_Name = "FDSI_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DrvCount As Integer
Dim bytes As String
Dim free As String
Dim DRVS(22) As String
Dim ChangedDisk As Currency
Dim OldVal(27) As Currency
Dim OldValX(27) As Currency
Dim StoreValue(27) As STO

Dim Kaza(25) As Currency
Dim Pitz(25) As String
Dim Lapo(25) As Integer
Dim oldx, oldy
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const SW_SHOW = 5
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Dim FSize As Integer
Dim FColor As Integer
Dim FType As Integer
Dim MoveMode As Boolean
Dim MIP As Boolean

Dim FormVisible As Boolean
Dim FormLevel As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long


Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Function GetDriveVolume(inDrive As String) As String

Dim N As Integer
Dim tmp As String

Drive1.Refresh

For N = 0 To Drive1.ListCount
  If UCase(Mid(Drive1.List(N), 1, 1)) = UCase(inDrive) Then _
     tmp = Mid(Drive1.List(N), 4): _
     tmp = Replace(tmp, "[", ""): _
     GetDriveVolume = Replace(tmp, "]", ""): _
     Exit Function
Next N

End Function


Sub ShowForm(inLvl As Long)

SetFormTColorXP Me, RGB(255, 0, 0), inLvl

End Sub

Sub ReConf()
On Error Resume Next

Dim ConfigFile As String, MX As Integer, LDX
Dim MeWidth As Single, MeHeight As Single

 Me.Cls


If DrvCount > 0 Then Me.Height = DrvCount * (dsc(0).Height + 30) + 40
If DrvCount = 0 Then Me.Height = (dsc(0).Height + 30) + 40

 Me.Top = Val(GetSetting("WDSI", "Settings", "TOP", Screen.Height - 1000)) - Me.Height
 Me.Left = Val(GetSetting("WDSI", "Settings", "LEFT", Screen.Width - 500)) - Me.Width

dsc(0).Caption = "Updating..."
dsc(0).Top = 0

For MX = 1 To DrvCount - 1
  Load dsc(MX)
  dsc(MX).Caption = "Updating..."
  dsc(MX).Visible = False
  dsc(MX).Top = (dsc(MX).Height + 30) * MX
Next



End Sub

Function GetMaxWidth() As Long
  For YoY = dsc.LBound To dsc.UBound
    If dsc(YoY).Width > GetMaxWidth Then GetMaxWidth = dsc(YoY).Width
  Next YoY
End Function


Private Sub dsc_Change(Index As Integer)

dsc(Index).Left = (Me.Width - dsc(Index).Width) - 50
Me.Line (0, dsc(Index).Top)-(Me.Width, dsc(Index).Height + dsc(Index).Top), Me.BackColor, BF

N = GetMaxWidth

If FType = 0 Then
    Me.ForeColor = IIf(FColor = 0, vbBlack, vbWhite)
    Me.CurrentX = dsc(Index).Left
    Me.CurrentY = dsc(Index).Top + 15
    Me.Print dsc(Index).Caption

    Me.CurrentX = dsc(Index).Left
    Me.CurrentY = dsc(Index).Top - 15
    Me.Print dsc(Index).Caption

    Me.CurrentX = dsc(Index).Left + 15
    Me.CurrentY = dsc(Index).Top
    Me.Print dsc(Index).Caption

    Me.CurrentX = dsc(Index).Left - 15
    Me.CurrentY = dsc(Index).Top
    Me.Print dsc(Index).Caption
    
    Me.ForeColor = IIf(FColor = 0, vbWhite, vbBlack)
    Me.CurrentX = dsc(Index).Left
    Me.CurrentY = dsc(Index).Top
    Me.Print dsc(Index).Caption
Else
    Me.ForeColor = IIf(FColor = 0, vbBlack, vbWhite)
    Me.CurrentX = dsc(Index).Left + 15
    Me.CurrentY = dsc(Index).Top + 15
    Me.Print dsc(Index).Caption
    
    Me.ForeColor = IIf(FColor = 0, vbWhite, bBlack)
    Me.CurrentX = dsc(Index).Left
    Me.CurrentY = dsc(Index).Top
    Me.Print dsc(Index).Caption
End If

End Sub

Private Sub dsc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Form_MouseDown Button, Shift, X, Y


End Sub


Private Sub dsc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Form_MouseMove Button, Shift, X, Y


End Sub

Private Sub dsc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Form_MouseUp Button, Shift, X, Y

End Sub

Private Sub fcfg_Click()
On Error Resume Next
If Me.Visible = False Then cfg.Show vbModal, Me Else cfg.SetFocus

End Sub

Function DataDevi(InLong As Currency) As String
If (InLong) >= 1000000000000# Then DataDevi = FormatEx(InLong / 1000000000000#, "0.00") + " TB": Exit Function
If (InLong) >= 1000000000 Then DataDevi = FormatEx(InLong / 1000000000, "0.00") + " GB": Exit Function
If (InLong) >= 1000000 Then DataDevi = FormatEx(InLong / 1000000, "0.00") + " MB":  Exit Function
If (InLong) >= 1000 Then DataDevi = FormatEx(InLong / 1000, "0.0") + " KB":  Exit Function
If (InLong) >= 0 Then DataDevi = FormatEx(InLong, "0") + " Bytes": Exit Function
End Function


Private Sub Form_Load()

Me.BackColor = vbRed

LoadDrives
TrayAdd Picture1, "Woobind Disk Space Info"

App.TaskVisible = False
LoadSettings
 
ShowForm 0
 
End Sub

Sub LoadDrives()
On Error Resume Next
For N = 0 To 23
 
 If Not mnuDRIVES(N) Then _
   Load mnuDRIVES(N)
 mnuDRIVES(N).Caption = Chr(67 + N) + ": " + vbTab + GetDriveVolume(Chr(67 + N))
 mnuDRIVES(N).Checked = False

Next N

End Sub

Function CountDrives()

CountDrives = 0
For N = 0 To 23
    If mnuDRIVES(N).Checked Then CountDrives = CountDrives + 1
Next N


End Function

Sub LoadSettings()
' On Error Resume Next
Dim Z, BZ As Integer
Dim ConfigFile As String, ZM, ZN

ConfigFile = LowPath(App.Path) + "wdsi.cfg"


Dim Intervl As Integer
Dim Inform As Integer
Dim DiskName As String

DrvCount = Val(GetIniRecord("DRIVES=", ConfigFile, "1"))
FType = Val(GetIniRecord("FONT=", ConfigFile, "0"))
FColor = Val(GetIniRecord("COLOR=", ConfigFile, "0"))
mnuHIDE.Checked = CBol(GetIniRecord("HIDE=", ConfigFile, "0"))
mnuCLICKS.Checked = CBol(GetIniRecord("TRANS=", ConfigFile, "0"))
mnuAPPROX.Checked = CBol(GetIniRecord("APPROX=", ConfigFile, "0"))
mnuLABEL.Checked = CBol(GetIniRecord("LABEL=", ConfigFile, "0"))

If DrvCount = 0 Then DrvCount = 1

mnuCOLOR.Checked = CBol(FColor)

If FType = 0 Then
 mnuTSTYLE(0).Checked = True
 mnuTSTYLE(1).Checked = False
Else
 mnuTSTYLE(1).Checked = True
 mnuTSTYLE(0).Checked = False
End If

For Z = 0 To DrvCount - 1
 DiskName = GetIniRecord("DRV_" + Format(Z, "0") + "=", ConfigFile, "C")
 DRVS(Z) = DiskName
 For BZ = 0 To mnuDRIVES.Count - 1
  If UCase(Left(mnuDRIVES(BZ).Caption, 1)) = UCase(DiskName) Then mnuDRIVES(BZ).Checked = True
 Next
Next

Call ReConf

End Sub

Sub SaveSettings()
' On Error Resume Next
Dim ConfigFile As String, ZM, ZN

ConfigFile = LowPath(App.Path) + "WDSI.cfg"

Open ConfigFile For Output As #1
 Print #1, "LABEL=" + Str(-CInt(mnuLABEL.Checked))
 Print #1, "APPROX=" + Str(-CInt(mnuAPPROX.Checked))
 Print #1, "HIDE=" + Str(-CInt(mnuHIDE.Checked))
 Print #1, "DRIVES=" + Str(CountDrives)
 Print #1, "FONT=" + Str(FType)
 Print #1, "COLOR=" + Str(FColor)
 Print #1, "TRANS=" + Str(-CInt(mnuCLICKS.Checked))
 
 Print #1, ""
 ZN = -1
 
 For ZM = 0 To mnuDRIVES.Count - 1
  If mnuDRIVES(ZM).Checked Then
   ZN = ZN + 1
   Print #1, "DRV_" + Format(ZN, "0") + "=" + UCase(Left(mnuDRIVES(ZM).Caption, 1))
  End If
 Next
 
Close #1


FDSI_2.ReConf

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = 2 And Not mnuCLICKS.Checked Then Picture1_MouseMove 0, 0, WindowsMessage.WM_RBUTTONUP, 0

If Button = 1 Then
    oldx = X
    oldy = Y
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 And Not mnuCLICKS.Checked Then
 Me.Move Me.Left + (X - oldx), Me.Top + (Y - oldy)
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then

SaveSetting "WDSI", "Settings", "TOP", Str(Me.Top + Me.Height)
SaveSetting "WDSI", "Settings", "LEFT", Str(Me.Left + Me.Width)

End If

End Sub

Private Sub Form_Resize()


 Dim ret As Long

 ret = CreateRectRgn(2, 2, (Me.Width / Screen.TwipsPerPixelX) - 1, (Me.Height / Screen.TwipsPerPixelY) - 1)
 SetWindowRgn Me.hwnd, ret, True


End Sub

Private Sub Form_Unload(Cancel As Integer)
TrayRemove
End
End Sub

Private Sub fvig_Click()
On Error Resume Next
Unload cfg
Unload FDSI_2
End
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAPPROX_Click()
mnuAPPROX.Checked = Not mnuAPPROX.Checked
SaveSettings
End Sub

Private Sub mnuAUTO_Click()
mnuAUTO.Checked = Not mnuAUTO.Checked
If mnuAUTO.Checked Then SetAutorun Else KillAutorun
End Sub

Private Sub mnuCLICKS_Click()
mnuCLICKS.Checked = Not mnuCLICKS.Checked
SaveSettings

End Sub

Private Sub mnuCOLOR_Click()
mnuCOLOR.Checked = Not mnuCOLOR.Checked
FColor = -CInt(mnuCOLOR.Checked)
End Sub

Private Sub mnuDRIVES_Click(Index As Integer)
mnuDRIVES(Index).Checked = Not mnuDRIVES(Index).Checked
SaveSettings
LoadSettings
End Sub

Private Sub mnuUpdate_Click()
 ReConf
End Sub


Private Sub mnuHIDE_Click()
mnuHIDE.Checked = Not mnuHIDE.Checked
SaveSettings
End Sub

Private Sub mnuLABEL_Click()
mnuLABEL.Checked = Not mnuLABEL.Checked
SaveSettings
End Sub

Private Sub mnuTSTYLE_Click(Index As Integer)
FType = Index
SaveSettings
LoadSettings
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
If X = WindowsMessage.WM_RBUTTONUP Then _
  mnuAUTO.Checked = IfAutorun: _
  Me.SetFocus: _
  PopupMenu Me.filo, , , , fvig

End Sub


Private Sub Timer1_Timer()
' // VISIBLITY CONTROL
If FormVisible And FormLevel < 180 Then FormLevel = FormLevel + 5: ShowForm FormLevel
If Not FormVisible And FormLevel >= 5 Then FormLevel = FormLevel - 5: ShowForm FormLevel

If mnuHIDE.Checked Then
 FormVisible = False
 For R = 0 To dsc.Count - 1
  FormVisible = FormVisible Or IsMouseInControl(Me, dsc(R))
 Next R
Else
  FormVisible = True
End If

For ZZ = 0 To DrvCount - 1
    dsc_Change (ZZ)
Next ZZ
    
    Dim mpos As POINTAPI
    GetCursorPos mpos
    
   If (mpos.X) * 15 >= Me.Left + Me.Width - GetMaxWidth Then
    If mnuCLICKS.Checked = False Then _
       Me.PSet ((mpos.X - 2 - Fix(Me.Left / 15)) * 15, (mpos.Y - 2 - Fix(Me.Top / 15)) * 15), RGB(0, 0, 0) _
    Else _
       Me.PSet ((mpos.X - 2 - Fix(Me.Left / 15)) * 15, (mpos.Y - 2 - Fix(Me.Top / 15)) * 15), RGB(255, 0, 0)
   End If

End Sub

Private Sub Tmr_Timer()
On Error Resume Next

For id = 0 To 25
 If Kaza(id) > 0 Then Kaza(id) = Kaza(id) + 1
 If Kaza(id) >= 5 Then Kaza(id) = 0
Next id

Dim ZZ As Integer, DriveLetter As String
Dim TotalSPC, FreeSPC, CurrSpc

For ZZ = 0 To DrvCount - 1

DriveLetter = DRVS(ZZ) + ":\"

TotalSPC = GotTotalDiskSpace(DriveLetter)
FreeSPC = GotFreeDiskSpace(DriveLetter)

Pitz(ZZ) = "     "

If OldVal(ZZ) < FreeSPC And Kaza(ZZ) = 0 Then
 Lapo(ZZ) = 1: Kaza(ZZ) = 1
End If

If OldVal(ZZ) > FreeSPC And Kaza(ZZ) = 0 Then
 Lapo(ZZ) = 2: Kaza(ZZ) = 1
End If

If Lapo(ZZ) = 1 Then
 Mid(Pitz(ZZ), 5 - Fix(Kaza(ZZ)), 1) = "<"
End If

If Lapo(ZZ) = 2 Then
 Mid(Pitz(ZZ), 1 + Fix(Kaza(ZZ)), 1) = ">"
End If


If Kaza(ZZ) = 0 Then Pitz(ZZ) = "     "

If OldVal(ZZ) > FreeSPC Then dsc(ZZ).ForeColor = RGB(255, 100, 100)
If OldVal(ZZ) < FreeSPC Then dsc(ZZ).ForeColor = RGB(100, 255, 100)
If OldVal(ZZ) = FreeSPC Then dsc(ZZ).ForeColor = RGB(255, 255, 255)

If TotalSPC > 0 Then
 If mnuLABEL.Checked Then
   dsc(ZZ).Caption = Pitz(ZZ) + IIf(mnuAPPROX.Checked, FormatEx(FreeSPC, "### ### ### ### ##0") + " bytes", DataDevi(CCur(FreeSPC))) + " free on " + _
   IIf(GetDriveVolume(UCase(DRVS(ZZ))) > "", GetDriveVolume(UCase(DRVS(ZZ))), "drive " + _
   UCase(DRVS(ZZ)) + ":")
 Else
   dsc(ZZ).Caption = Pitz(ZZ) + IIf(mnuAPPROX.Checked, FormatEx(FreeSPC, "### ### ### ### ##0") + " bytes", DataDevi(CCur(FreeSPC))) + " free on drive " + _
   UCase(DRVS(ZZ)) + ":"
 End If
End If

If TotalSPC = 0 Then dsc(ZZ).Caption = UCase(DRVS(ZZ)) + ": is not present!"
OldVal(ZZ) = FreeSPC
Next ZZ

End Sub



