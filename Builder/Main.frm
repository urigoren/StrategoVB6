VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Stratego - Strategy Builder"
   ClientHeight    =   3105
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6285
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton New_S 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Open_S 
      Caption         =   "Open"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Save_S 
      Caption         =   "Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox Place 
      Height          =   630
      Index           =   0
      Left            =   0
      Picture         =   "Main.frx":030A
      ScaleHeight     =   570
      ScaleWidth      =   570
      TabIndex        =   3
      Top             =   0
      Width           =   630
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Click the Board"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   6240
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Menu pawns 
      Caption         =   "Pawns"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "1(1)"
      End
      Begin VB.Menu mnu2 
         Caption         =   "2(1)"
      End
      Begin VB.Menu mnu3 
         Caption         =   "3(2)"
      End
      Begin VB.Menu mnu4 
         Caption         =   "4(3)"
      End
      Begin VB.Menu mnu5 
         Caption         =   "5(4)"
      End
      Begin VB.Menu mnu6 
         Caption         =   "6(4)"
      End
      Begin VB.Menu mnu7 
         Caption         =   "7(4)"
      End
      Begin VB.Menu mnu8 
         Caption         =   "8(5)"
      End
      Begin VB.Menu mnu9 
         Caption         =   "9(8)"
      End
      Begin VB.Menu mnu10 
         Caption         =   "10(1)"
      End
      Begin VB.Menu mnuF 
         Caption         =   "Flag (1)"
      End
      Begin VB.Menu mnuB 
         Caption         =   "Bomb(6)"
      End
      Begin VB.Menu mnudel 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cmndlg As New CMNDialog, CurrIndex As Integer
Dim pawn(1 To 12) As Byte
Private Map(0 To 99) As Byte, Param As String
'1: 1, 2: 1
'3: 2, 4: 3
'5: 4, 6: 4
'7: 4, 8: 5
'9: 8, 10: 1
'F: 1, B: 6


Private Sub Form_Load()
Dim X As Integer, Y As Integer, i As Integer
For Y = 0 To 3
For X = 0 To 9
i = X + 10 * Y
If i > 0 Then Load Place(i)
With Place(i)
.Left = X * Place(0).Width
.Top = Y * Place(0).Height
.Visible = True
End With
Next X
Next Y
cmndlg.Filter = "Stratego Strategy Seting|*.sss"
cmndlg.InitDir = App.Path
pawn(1) = 1
pawn(2) = 1
pawn(3) = 2
pawn(4) = 3
pawn(5) = 4
pawn(6) = 4
pawn(7) = 4
pawn(8) = 5
pawn(9) = 8
pawn(10) = 1
pawn(11) = 6 'B
pawn(12) = 1 'F
If Command() <> "" And Dir(Command()) <> "" Then Param = Command(): Call Open_S_Click
End Sub

Private Sub mnu1_Click()
If pawn(1) = 0 Then Exit Sub
Map(CurrIndex) = 1
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\1.jpg")
updatemnu
End Sub

Private Sub mnu2_Click()
If pawn(2) = 0 Then Exit Sub
Map(CurrIndex) = 2
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\2.jpg")
updatemnu
End Sub

Private Sub mnu3_Click()
If pawn(3) = 0 Then Exit Sub
Map(CurrIndex) = 3
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\3.jpg")
updatemnu
End Sub

Private Sub mnu4_Click()
If pawn(4) = 0 Then Exit Sub
Map(CurrIndex) = 4
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\4.jpg")
updatemnu
End Sub

Private Sub mnu5_Click()
If pawn(5) = 0 Then Exit Sub
Map(CurrIndex) = 5
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\5.jpg")
updatemnu
End Sub

Private Sub mnu6_Click()
If pawn(6) = 0 Then Exit Sub
Map(CurrIndex) = 6
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\6.jpg")
updatemnu
End Sub

Private Sub mnu7_Click()
If pawn(7) = 0 Then Exit Sub
Map(CurrIndex) = 7
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\7.jpg")
updatemnu
End Sub

Private Sub mnu8_Click()
If pawn(8) = 0 Then Exit Sub
Map(CurrIndex) = 8
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\8.jpg")
updatemnu
End Sub

Private Sub mnu9_Click()
If pawn(9) = 0 Then Exit Sub
Map(CurrIndex) = 9
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\9.jpg")
updatemnu
End Sub

Private Sub mnu10_Click()
If pawn(10) = 0 Then Exit Sub
Map(CurrIndex) = 10
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\10.jpg")
updatemnu
End Sub

Private Sub mnuB_Click()
If pawn(11) = 0 Then Exit Sub
Map(CurrIndex) = 11
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\12.jpg")
updatemnu
End Sub

Private Sub mnuF_Click()
If pawn(12) = 0 Then Exit Sub
Map(CurrIndex) = 12
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\13.jpg")
updatemnu
End Sub

Private Sub mnudel_Click()
Map(CurrIndex) = 0
Place(CurrIndex).Picture = LoadPicture(App.Path & "\images\0.jpg")
updatemnu
End Sub


Private Sub New_S_Click()
For i = 0 To 39
Place(i).Picture = LoadPicture(App.Path & "\images\0.jpg")
Map(i) = 0
Next i
updatemnu
End Sub

Private Sub Open_S_Click()
Dim S As String, FileName As String
If Param = "" Then
cmndlg.DialogTitle = "Open a Strategy"
cmndlg.Flags = OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_PATHMUSTEXIST
cmndlg.ShowOpen
FileName = Replace(cmndlg.FileName, Chr(0), vbNullString)
If FileName = vbNullString Then Exit Sub
Else
FileName = Param
Param = ""
End If
Open FileName For Input As #2
Input #2, S
Close #2
For i = 1 To 40
Map(i - 1) = Val(CutStr(S, i, ";"))
Place(i - 1).Picture = LoadPicture(App.Path & "\images\" & Map(i - 1) & ".jpg")
If Map(i - 1) = 12 Then Map(i - 1) = 11
If Map(i - 1) = 13 Then Map(i - 1) = 12
Next i
Call updatemnu
End Sub

Private Sub Place_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
CurrIndex = Index
PopupMenu pawns
End Sub

Private Sub Save_S_Click()
Dim S As String, i As Integer, FileName As String, fixed As Byte
S = 0
If Map(0) + Map(1) + Map(4) + Map(5) + Map(8) + Map(9) >= 71 Then Call MsgBox("Illegal setting", vbCritical, "Error"): Exit Sub
For i = 1 To 12
S = S + pawn(i)
Next i
If S > 0 Then Call MsgBox("Illegal setting", vbCritical, "Error"): Exit Sub
cmndlg.DialogTitle = "Save New strategy"
cmndlg.Flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_OVERWRITEPROMPT
cmndlg.ShowSave
FileName = Replace(cmndlg.FileName, Chr(0), vbNullString)
If FileName = vbNullString Then Exit Sub
If LCase(Right(FileName, 4)) <> ".sss" Then FileName = FileName & ".sss"
Open FileName For Output As #1
For i = 0 To 39
If Map(i) = 11 Then fixed = 12
If Map(i) = 12 Then fixed = 13
If Map(i) < 11 Then fixed = Map(i)
Print #1, fixed & ";";
Next i
Close #1
Call MsgBox("Done")
End Sub

Private Sub updatemnu()
pawn(1) = 1
pawn(2) = 1
pawn(3) = 2
pawn(4) = 3
pawn(5) = 4
pawn(6) = 4
pawn(7) = 4
pawn(8) = 5
pawn(9) = 8
pawn(10) = 1
pawn(11) = 6 'B
pawn(12) = 1 'F

For i = 0 To 39
If Map(i) = 1 Then pawn(1) = pawn(1) - 1
If Map(i) = 2 Then pawn(2) = pawn(2) - 1
If Map(i) = 3 Then pawn(3) = pawn(3) - 1
If Map(i) = 4 Then pawn(4) = pawn(4) - 1
If Map(i) = 5 Then pawn(5) = pawn(5) - 1
If Map(i) = 6 Then pawn(6) = pawn(6) - 1
If Map(i) = 7 Then pawn(7) = pawn(7) - 1
If Map(i) = 8 Then pawn(8) = pawn(8) - 1
If Map(i) = 9 Then pawn(9) = pawn(9) - 1
If Map(i) = 10 Then pawn(10) = pawn(10) - 1
If Map(i) = 11 Then pawn(11) = pawn(11) - 1
If Map(i) = 12 Then pawn(12) = pawn(12) - 1
Next i
mnu1.Caption = "1(" & pawn(1) & ")"
mnu2.Caption = "2(" & pawn(2) & ")"
mnu3.Caption = "3(" & pawn(3) & ")"
mnu4.Caption = "4(" & pawn(4) & ")"
mnu5.Caption = "5(" & pawn(5) & ")"
mnu6.Caption = "6(" & pawn(6) & ")"
mnu7.Caption = "7(" & pawn(7) & ")"
mnu8.Caption = "8(" & pawn(8) & ")"
mnu9.Caption = "9(" & pawn(9) & ")"
mnu10.Caption = "10(" & pawn(10) & ")"
mnuB.Caption = "Bomb(" & pawn(11) & ")"
mnuF.Caption = "Flag(" & pawn(12) & ")"
End Sub
