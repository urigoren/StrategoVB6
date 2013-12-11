Attribute VB_Name = "Subs"
Option Explicit
Public Map(0 To 99) As Byte, AttackMap(0 To 99) As Byte, SelPlace(1 To 2) As Byte
'pawns 1-10, unknown 11, bomb 12, flag 13, blank 0 or 255

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Sub DisableX(hWnd As Long)
ModifyMenu GetSystemMenu(hWnd, 0), &HF060, &H0& Or &H1&, -10, "Close"
End Sub

Public Sub Trans_Form(hWnd As Long, Visible As Boolean)
Dim InitStyle As Long
InitStyle = GetWindowLong(hWnd, (-20))
If Visible = False Then
    SetWindowLong hWnd, (-20), InitStyle Or &H20
    SetWindowPos hWnd, 0, 0, 0, 0, 0, &H20 Or &H20 Or &H4 Or &H1
Else
    SetWindowLong hWnd, (-20), InitStyle
    SetWindowPos hWnd, 0, 0, 0, 0, 0, &H20 Or &H20 Or &H4 Or &H1
End If
End Sub

Public Sub MoveForm(hWnd As Long, Button As Integer)
 If Button = 1 Then
    ReleaseCapture
    SendMessage hWnd, &HA1, 2, 0&
 End If
End Sub


Public Sub AlwaysOnTop(hWnd As Long, Currect As Boolean)
If Currect = True Then
SetWindowPos hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
Else
SetWindowPos hWnd, -2, 0, 0, 0, 0, &H1 Or &H2
End If
End Sub

Public Sub NegativePic(PlaceHDC As Long)
Dim r As RECT
    With r
        .Top = 0
        .Left = 0
        .Bottom = fMain.Place(0).Height
        .Right = fMain.Place(0).Width
    End With
    InvertRect PlaceHDC, r
End Sub

'NO API
Private Sub Bar(P As PictureBox, ByVal n As Byte, Optional from0 As Boolean, Optional C As Long)
Dim i As Integer
Dim A As Integer, B As Integer
If n < 0 Or n > 100 Then Exit Sub
P.Cls
P.AutoRedraw = True
If from0 = True Then
A = 0: B = n * P.ScaleWidth / 100
Else
A = P.ScaleWidth: B = n * P.ScaleWidth / 100
End If
For i = A To B Step IIf(from0, 1, -1)
P.Line (i, 0)-(i, P.ScaleHeight), C
Next i
End Sub


Public Function CutStr(ByVal Str As String, ByVal PartNum As Integer, Optional Delimiter As String = " ") As String
Dim fStr As String, pos(1 To 2) As Integer, i As Integer
fStr = Str: pos(1) = 1
If Left(Str, Len(Delimiter)) <> Delimiter Then fStr = Delimiter & fStr
If Right(Str, Len(Delimiter)) <> Delimiter Then fStr = fStr & Delimiter
For i = 1 To PartNum + 1
If i <= PartNum Then
pos(1) = InStr(pos(1), fStr, Delimiter) + Len(Delimiter)
Else
pos(2) = InStr(pos(1), fStr, Delimiter)
If pos(2) = 0 Then pos(2) = pos(1): pos(1) = 1 + Len(Delimiter)
End If
Next i
CutStr = Mid(fStr, pos(1), pos(2) - pos(1))
End Function

Public Function MultiChar(ByVal Number As Integer, ByVal Char As String) As String
    Dim i As Integer
    If Len(Char) > 1 Then Char = Mid(Char, 1, 1)
    For i = 1 To Number Step 1
        MultiChar = MultiChar & Char
    Next i
End Function

Public Function CancelChar(ByVal Text As String, Optional Char As String = " ") As String
Dim i As Integer
If Len(Char) > 1 Then Char = Mid(Char, 1, 1)
For i = 1 To Len(Text) Step 1
If Not Mid(Text, i, 1) = Char Then CancelChar = CancelChar & Mid(Text, i, 1)
Next i
End Function

Public Function Flip(ByVal Text As String) As String
Dim i As Integer
For i = Len(Text) To 1 Step -1
Flip = Flip & Mid(Text, i, 1)
Next i
End Function

Public Function GetMaxCharNum(ByVal Text As String, ByVal Char As String) As Integer
Dim i As Integer
If Len(Char) > 1 Then Char = Mid(Char, 1, 1)
For i = 1 To Len(Text)
If Mid(Text, i, 1) = Char Then GetMaxCharNum = GetMaxCharNum + 1
Next i
End Function

Public Function GetMaxStrNum(ByVal Text As String, ByVal Str As String) As Integer
Dim pos As Integer
pos = 1
Do
If InStr(pos, Text, Str) <> 0 Then GetMaxStrNum = GetMaxStrNum + 1
pos = InStr(pos, Text, Str) + Len(Str)
Loop Until InStr(pos, Text, Str) = 0
End Function
