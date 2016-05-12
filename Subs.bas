Attribute VB_Name = "Subs"
Type STO
 stoSIZE As Currency
End Type

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0
Public Const HWND_TOP = 0

Function FormatTime(Seconds As Long) As String
Dim HR, MN, SC
HR = Fix(Seconds / 3600)
MN = Fix(Seconds / 60) Mod 60
SC = Seconds Mod 60
FormatTime = Format(HR, "0") + ":" + Format(MN, "00") + ":" + Format(SC, "00")

End Function


Function GetTimeFromMinutes(vMinutes As Integer)
 GetTimeFromMinutes = Format$(Fix(vMinutes / 60), "00") & ":" & Format$(Fix(vMinutes Mod 60), "00")
End Function



Function Language(Rec As String)

Language = GetIniRecord(Rec, LowPath(App.Path) + GetIniRecord("LANG_FILE:", LowPath(App.Path) + "LANGUAGE.SEL"))

End Function

Public Function GetVersion() As String
GetVersion = Format$(App.Major, "0") + "." + Format$(App.Minor, "00")
End Function
Function PathHead$(Filename As String)
Dim Names As Integer
For Names = Len(Filename) To 1 Step -1
 If Mid$(Filename, Names, 1) = "\" Then
  PathHead$ = Mid$(Filename, 1, (Names) - 1)
  If PathHead$ = "$APPDIR$" Then PathHead$ = App.Path
  Exit For
 End If
Next
End Function
Function FileExists(Path$) As Boolean
    Dim x As Integer

    x = FreeFile

    On Error Resume Next
    Open Path$ For Input As x
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close x

End Function
Public Function FileHead$(Filename As String)
Dim Names As Integer
For Names = Len(Filename) To 1 Step -1
If Mid$(Filename, Names, 1) = "\" Then FileHead$ = Right$(Filename, Len(Filename) - (Names)): Exit Function
Next
End Function

Public Function LowPath(InPath As String) As String
If Right$(InPath, 1) = "\" Then LowPath = InPath
If Right$(InPath, 1) <> "\" Then LowPath = InPath + "\"
End Function

Public Function GetIniRecord(Record As String, INIFile As String)
Dim CfgLine As String, G As Integer
On Error Resume Next
G = FreeFile
Open INIFile For Input As #G
Do
Line Input #G, CfgLine
If UCase$(Mid$(CfgLine, 1, Len(Record))) = UCase(Record) Then
   GetIniRecord = Mid$(CfgLine, Len(Record) + 1)
End If
Loop While Not EOF(G)
Close G
End Function

Function ReadCommand(ByRef GetCommand As String, ByRef GetValue As Boolean)
 If GetValue = True Then ReadCommand = Right$(GetCommand, Len(GetCommand) - 12)
 If GetValue = False Then ReadCommand = Mid$(GetCommand, 1, 11)
End Function

Function FilterName(Text As String) As String

Dim Ls, Bs, Variants, Bizer
On Error Resume Next

For Ls = 1 To Len(Text)
Bs = Mid$(Text, Ls, 1)

 For Variants = 0 To 47
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 91 To 96
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 58 To 63
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 123 To 191
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 
Mid$(Text, Ls, 1) = Bs

Next

If Text = "" Then Text = "Unnamed"
FilterName = Text

End Function



