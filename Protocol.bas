Attribute VB_Name = "Protocol"
' Web Stratego Protocol
'"A:" = Answer
'"B:" = Update Battle
'"C:" = Chat Messag
'"D:" = Draw
'"G:" get whole map
'"K:" = Kill
'"L:" = you Lose
'"M:" = Move
'"P:" = Place a certin pawn
'"Q:" = Question - Done
'"S:" = Special effect
'"T:" = your Turn
'"U:" send whole map
'"V:" = I win
'MEMO - Opponent's map is 255-my map
'Every command must end with ;

Public Sub P(Data As String)
Dim pos As Byte, pic As Byte
pos = 99 - Val(CutStr(Data, 1, "-"))
pic = 255 - Val(CutStr(Data, 2, "-"))
Map(pos) = pic
AttackMap(pos) = pic
DrawMap
End Sub

Public Sub B(Data As String)
Dim pic(1 To 2) As Integer
pic(1) = 255 - Val(CutStr(Data, 1, "-"))
pic(2) = 255 - Val(CutStr(Data, 2, "-"))
fMain.battle(1).Picture = LoadPicture(App.Path & "\images\" & pic(1) & ".jpg")
fMain.battle(2).Picture = LoadPicture(App.Path & "\images\" & pic(2) & ".jpg")
End Sub

Public Sub C(Data As String) 'V
ChatFrm.Add_Line (Data)
ChatFrm.Show
End Sub

Public Sub A(Data As String) 'V
Dim pos As Integer
pos = Val(CutStr(Data, 1, "-"))
AttackMap(pos) = Val(CutStr(Data, 2, "-"))
If AttackMap(pos) = 255 Then AttackMap(pos) = 0
DrawMap
End Sub

Public Sub Q(Data As String) 'V
fMain.Socket.SendData "A:" & Data & "-" & (255 - Map(99 - Val(Data))) & ";"
'the 99-pos is to convert player's map to opponent's
'the 255-map is to convert enemy's pawns to correct
End Sub

Public Sub T(Data As String)
If UCase(Data) = "CM" Then
If CanMove Then
Call MsgBox("Your Opponent Can't move, you get an extra turn", vbInformation, "Extra turn")
Else 'Draw
fMain.Socket.SendData "D:;"
Protocol.D
End If
Else
fMain.Socket.SendData "Q:" & (99 - Val(CutStr(Data, 1, "-"))) & ";Q:" & (99 - Val(CutStr(Data, 2, "-"))) & ";"
End If
End Sub

Public Sub K(Data As String) 'V
Map(99 - Val(Data)) = 0
AttackMap(99 - Val(Data)) = 0
DrawMap
End Sub

Public Sub M(Data As String) 'V
Dim pos(1 To 2) As Integer
pos(1) = 99 - Val(CutStr(Data, 1, "-"))
pos(2) = 99 - Val(CutStr(Data, 2, "-"))
Map(pos(2)) = Map(pos(1))
AttackMap(pos(2)) = AttackMap(pos(1))
Map(pos(1)) = 0
AttackMap(pos(1)) = 0
DrawMap
End Sub

Public Sub L(Optional Data As String) 'Defeat
Dim i As Byte
FrmFX.Play_File ("Defeat")
If Data <> "" Then Protocol.M (Data)
For i = 0 To 99
Map(i) = 0
AttackMap(i) = 0
Next i
DrawMap
fMain.Socket.SendData "V:;"
fMain.Socket.Close
'Do something incase of a re-match
End Sub

Public Sub D() 'Draw
Dim i As Byte
FrmFX.Play_File ("Draw")
For i = 0 To 99
Map(i) = 0
AttackMap(i) = 0
Next i
DrawMap
fMain.Socket.Close
'Do something incase of a re-match
End Sub

Public Sub V(Optional Data As String) 'Victory
Dim i As Byte
FrmFX.Play_File ("Victory")
If Data <> "" Then Protocol.M (Data)
For i = 0 To 99
Map(i) = 0
AttackMap(i) = 0
Next i
DrawMap
fMain.Socket.SendData "L:;"
fMain.Socket.Close
'Do something incase of a re-match
End Sub

Public Sub GetMap()
Dim i As Integer, s As String
For i = 0 To 99
Select Case Map(i)
Case 1 To 9
s = s & Map(i)
Case 10
s = s & "0"
Case 0, 127 To 128, 242 To 254
s = s & "K"
Case 12
s = s & "B"
Case 13
s = s & "F"
End Select
Next i
fMain.Socket.SendData "U:" & Flip(s) & ";"
End Sub

Public Sub DrawMap()
Dim i As Integer
For i = 0 To 99
If Map(i) = 255 Then Map(i) = 0
If Map(i) <> 244 Then AttackMap(i) = Map(i)

fMain.Place(i).Picture = LoadPicture(App.Path & "\images\" & Map(i) & ".jpg")
If i = SelPlace(1) Or i = SelPlace(2) Then NegativePic (fMain.Place(i).hdc)
Next i
End Sub

Public Function CheckVictory() As Boolean
Dim i As Byte, flag As Boolean, s As Integer
For i = 0 To 99
If AttackMap(i) = 242 Then flag = True
If AttackMap(i) >= 242 Then s = s + 1 'Every enemy piece
If AttackMap(i) = 242 Or AttackMap(i) = 243 Then s = s - 1  'bombs and flag
Next i
If s = 0 Or flag = False Then CheckVictory = True
End Function

Public Function CanMove(Optional Pl As Byte = 100) As Boolean
   Dim i As Byte
   If Pl >= 100 Then
      CanMove = False
      For i = 0 To 99
         If CanMove(i) = True Then CanMove = True
      Next i
   Else
      If Map(Pl) > 12 Or Map(Pl) = 0 Then CanMove = False: Exit Function
      Select Case Pl
      Case 0 '2 way
         CanMove = MovablePlace(1) Or MovablePlace(10)
      Case 90 '2 way
         CanMove = MovablePlace(91) Or MovablePlace(80)
      Case 9 '2 way
         CanMove = MovablePlace(8) Or MovablePlace(19)
      Case 99 '2 way
         CanMove = MovablePlace(98) Or MovablePlace(89)
      Case 10, 20, 30, 40, 50, 60, 70, 80 '3 way
         CanMove = MovablePlace(Pl + 1) Or MovablePlace(Pl + 10) Or MovablePlace(Pl - 10)
      Case 19, 29, 39, 49, 59, 69, 79, 89 '3 way
         CanMove = MovablePlace(Pl - 1) Or MovablePlace(Pl + 10) Or MovablePlace(Pl - 10)
      Case 1 To 8 '3 way
         CanMove = MovablePlace(Pl + 1) Or MovablePlace(Pl - 1) Or MovablePlace(Pl + 10)
      Case 91 To 98 '3 way
         CanMove = MovablePlace(Pl + 1) Or MovablePlace(Pl - 1) Or MovablePlace(Pl - 10)
      Case Else '4 way
         CanMove = MovablePlace(Pl + 1) Or MovablePlace(Pl - 1) Or MovablePlace(Pl + 10) Or MovablePlace(Pl - 10)
      End Select
   End If
End Function

Private Function MovablePlace(ByVal Pl As Byte) As Boolean
MovablePlace = (Map(Pl) = 0) Or (Map(Pl) > 13 And Map(Pl) <> 127 And Map(Pl) <> 128)
End Function
