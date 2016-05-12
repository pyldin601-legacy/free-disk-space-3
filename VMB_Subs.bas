Attribute VB_Name = "VMB_Subs"

Function GetTimeFromMinutes(vMinutes As Integer)
 GetTimeFromMinutes = Format$(Fix(vMinutes / 60), "00") & ":" & Format$(Fix(vMinutes Mod 60), "00")
End Function

Function Language(Rec As String)

Language = GetIniRecord(Rec, LowPath(App.Path) + GetIniRecord("LANG_FILE:", LowPath(App.Path) + "LANGUAGE.SEL"))

End Function

Function WINTODOS(Tex$) As String
Dim Char, Cit, Zen As String
Zen = Tex$

For X = 1 To Len(Tex$)

Char = Asc(Mid$(Zen, X, 1))
If Char >= 192 And Char <= 239 Then
  Cit = Char - 64
  GoTo OK
Else
  Cit = Char
End If

If Char >= 240 And Char <= 255 Then
  Cit = Char - 16
  GoTo OK
Else
  Cit = Char
End If

OK:
Mid$(Zen, X, 1) = Chr$(Cit)

Next

WINTODOS = Zen

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
    Dim X As Integer

    X = FreeFile

    On Error Resume Next
    Open Path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X

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

Public Function GetMp3Song(ByRef Song As String)

On Error Resume Next

Dim RU As String
Dim TT As String, AU As String, AL As String
Dim YR As String, CM As String, Tip As String

Tip = Right(UCase(Song), 3)

If Tip = "MP3" Or Tip = "MP2" Or Tip = "MP1" Or Tip = "WMA" Then
 GetMP3idTAG Song, TT, AU, AL, YR, CM
End If

If Tip = "MID" Or Tip = "RMI" Then
 TT = MIDI_NAME(Song)
 YR = Midi_Size(Song)
End If

If TT = "" Then TT = Mid(FileHead(Song), 1, Len(FileHead(Song)) - 4)
If AU > "" Then AU = AU + " - "
If YR > "" Then YR = " (" + YR + ")"


If Tip = "WAV" Then RU = "[WAVE] "
If Tip = "MID" Then RU = "[MIDI] "
If Tip = "RMI" Then RU = "[MIDI] "
If Tip = "MP3" Then RU = "[MPEG] "
If Tip = "MP2" Then RU = "[MPEG] "
If Tip = "MP1" Then RU = "[MPEG] "
If Tip = "MPG" Then RU = "[VIDEO] "
If Tip = "MPG" Then RU = "[VDEO] "
If Tip = "WMA" Then RU = "[WMED] "

If RU = "" Then RU = "[DEFL] "

GetMp3Song = RU + AU + TT + YR

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

Sub ExchangeFiles(SI As Integer, di As Integer, Sources As ListBox)
On Error Resume Next
Dim a, B, ASel As Boolean, BSel As Boolean
a = Sources.List(di)
B = Sources.List(SI)
ASel = Sources.Selected(di)
BSel = Sources.Selected(SI)
Sources.List(di) = B
Sources.List(SI) = a
Sources.Selected(di) = BSel
Sources.Selected(SI) = ASel

End Sub


