VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form fMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Stratego"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SkinS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Skins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6600
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1635
      Width           =   735
   End
   Begin VB.CommandButton cmdStat 
      BackColor       =   &H00FFC0C0&
      Caption         =   "%"
      Height          =   495
      Left            =   7250
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "View statistics"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Place 
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   0
      Left            =   0
      Picture         =   "Main.frx":08CA
      ScaleHeight     =   630
      ScaleWidth      =   630
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   630
   End
   Begin VB.CommandButton New_S 
      Caption         =   "&New"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      ToolTipText     =   "Create New strategy file"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton ChatB 
      Caption         =   "C&hat"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "Web Stratego Chat"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Load_S 
      Caption         =   "&Load"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      ToolTipText     =   "Load existing strategy file"
      Top             =   2400
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   120
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Help 
      Caption         =   "&Help"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      ToolTipText     =   "Web Stratego Help"
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton ENDit 
      Caption         =   "&Disconnect"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      ToolTipText     =   "Disconnect existing connection"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Passive 
      Caption         =   "&Host a Game"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "Wait for someone to connect you"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Send 
      BackColor       =   &H00FFFFFF&
      Caption         =   "O&K"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Action 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      ToolTipText     =   "Connect to someone"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Domain 
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Last Battle:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblTurn 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Turn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   6480
      TabIndex        =   17
      ToolTipText     =   "Who's Turn ?"
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image battle 
      Height          =   615
      Index           =   1
      Left            =   6600
      Stretch         =   -1  'True
      ToolTipText     =   "Attacker"
      Top             =   5280
      Width           =   615
   End
   Begin VB.Image battle 
      Height          =   615
      Index           =   2
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Goren4U Product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   6480
      MouseIcon       =   "Main.frx":1191
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   6360
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Game:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Strategy:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Software:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   6360
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "User's IP:"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Status 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Connection:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHlpRules 
         Caption         =   "Rules"
      End
      Begin VB.Menu MnuHlpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'_______________________________________'
' Web Stratego was written by Uri Goren '
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
'
' When Selplace is over 99 it marks nothing
Private cmndlg As New CMNDialog

Sub D_Buttons()
Load_S.Enabled = True
Passive.Enabled = True
Action.Enabled = True
ENDit.Enabled = False
ChatB.Enabled = False
SkinS.Visible = True
Status.Caption = "Disconnected"
End Sub

Sub C_Buttons()
Load_S.Enabled = False
Passive.Enabled = False
Action.Enabled = False
ENDit.Enabled = True
ChatB.Enabled = True
SkinS.Visible = False
Status.Caption = "Connected"
End Sub


Private Sub Action_Click()
On Error GoTo ErrorHandle:
If Map(0) = 0 Then Call MsgBox("Please load your strategy first", vbExclamation, "Can not Connect")
If Domain.Text = "" Then
MsgBox ("No address entered")
Else
Socket.Close
Socket.RemotePort = 1066
Socket.RemoteHost = Domain.Text
Socket.Connect
End If
Send.Enabled = True
Exit Sub
ErrorHandle:
Call Socket_Error(443, "No Description", 442, "No source", "No help", 441, True)
End Sub

Private Sub ChatB_Click()
ChatFrm.Show
End Sub

Private Sub cmdStat_Click()
fStat.Show
End Sub

Private Sub ENDit_Click()
Socket.Close
End Sub

Private Sub Form_Activate()
Call Form_MouseMove(0, 0, 0, 0)
End Sub

Private Sub Form_Load()
Dim X, Y, i As Integer
If Socket.LocalIP = "127.0.0.1" Then
Me.Caption = "Web Stratego"
Else
Me.Caption = "Web Stratego - " & Socket.LocalIP
End If
For Y = 0 To 9
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
SelPlace(1) = 100 'none
SelPlace(2) = 100 'none
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MoveForm(Me.hWnd, Button)
Select Case Socket.State
Case 7
Status.Caption = "Connected"
Call C_Buttons
Case 0, 8
Status.Caption = "Disconnected"
Call D_Buttons
Case 4, 6
Status.Caption = "Connecting"
Case 2
Status.Caption = "Listenning"
ENDit.Enabled = True
Case Else
Status.Caption = Socket.State
End Select
End Sub

Private Sub Form_Paint()
Call DrawMap
End Sub

Private Sub Form_Unload(Cancel As Integer)
Socket.Close
End
End Sub

Private Sub Help_Click()
PopupMenu mnuHelp, , Help.Left, Help.Top + Help.Height
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Dim i As Byte, Str As String
Str = InputBox("Please enter the direct message", "Protocol")
If Str = vbNullString Then Exit Sub
If Mid(Str, 2, 1) <> ":" And LCase(Str) <> "uri goren" Then Exit Sub
If LCase(Str) = "uri goren" Then
For i = 0 To 99
Map(i) = AttackMap(i)
Next i
DrawMap
Else
If Right(Str, 1) <> ";" Then Str = Str & ";"
Socket.SendData Str
End If
Else
Call ShellExecute(Me.hWnd, "", "http://www.goren4u.com/", "", App.Path, vbNormalFocus)
End If
End Sub

Private Sub Load_S_Click()
Dim FileName As String, Data As String
cmndlg.Filter = "Strategy File (*.sss)|*.sss"
cmndlg.InitDir = App.Path
cmndlg.Flags = OFN_PATHMUSTEXIST + OFN_HIDEREADONLY + OFN_FILEMUSTEXIST
cmndlg.ShowOpen
FileName = Replace(cmndlg.FileName, Chr(0), "")
If FileName = "" Then Exit Sub
Open FileName For Input As #1
Input #1, Data
Close #1
For i = 60 To 99
Map(i) = Val(CutStr(Data, i - 59, ";"))
AttackMap(i) = Map(i)
Next i
For i = 40 To 59
Map(i) = 0
AttackMap(i) = 0
Next i
AttackMap(46) = 128
AttackMap(47) = 127
AttackMap(42) = 128
AttackMap(43) = 127
AttackMap(56) = 127
AttackMap(57) = 128
AttackMap(52) = 127
AttackMap(53) = 128
Map(46) = 128
Map(47) = 127
Map(42) = 128
Map(43) = 127
Map(56) = 127
Map(57) = 128
Map(52) = 127
Map(53) = 128
For i = 0 To 39
Map(i) = 244
AttackMap(i) = 244
Next i
Call DrawMap
End Sub

Private Sub MnuHlpAbout_Click()
Call MsgBox(Space(20) & "A Goren4U Product" & Space(20) & vbCrLf & Space(20) & "---------------------" & vbCrLf & Space(14) & "Programming by Uri Goren" & vbCrLf & Space(16) & "Graphics by Or Unger" & vbCrLf & vbCrLf & "Quality authorization:" & vbCrLf & Space(15) & "Noam Arkind" & vbCrLf & Space(15) & "Eli Sigal", vbInformation, "About")
End Sub

Private Sub mnuHlpRules_Click()
Call ShellExecute(Me.hWnd, "", "http://www.goren4u.com/games/strarules.htm", "", App.Path, vbNormalFocus)
End Sub

Private Sub New_S_Click()
Shell App.Path & "\Sbuild.exe", vbNormalFocus
End Sub

Private Sub Passive_Click()
On Error GoTo ErrorHandle
If Map(0) = 0 Then Call MsgBox("Please load your strategy first", vbExclamation, "Can not Connect")
Socket.Close
Socket.LocalPort = 1066
Socket.Listen
ENDit.Enabled = True
Status.Caption = "Listenning"
Domain.Text = ""
Exit Sub
ErrorHandle:
Call Socket_Error(443, "No Description", 442, "No source", "No help", 441, True)
End Sub

Private Sub Place_Click(Index As Integer)
 Select Case Index
  Case Is = SelPlace(1)
   SelPlace(1) = SelPlace(2)
   SelPlace(2) = 100
  Case Is = SelPlace(2)
   SelPlace(2) = 100
  Case Else
   If SelPlace(1) > 99 And SelPlace(2) > 99 Then SelPlace(1) = Index
   If SelPlace(1) < 100 And SelPlace(2) > 99 Then SelPlace(2) = Index
   If SelPlace(1) < 100 And SelPlace(2) < 100 Then SelPlace(1) = SelPlace(2): SelPlace(2) = Index
 End Select
 Call DrawMap
 Debug.Print SelPlace(1), SelPlace(2)
End Sub

Private Sub Send_Click()
    Dim i As Integer
   If Socket.State <> sckConnected Then Call MsgBox("Not Connected", vbApplicationModal, "Error"): Exit Sub
'-----------Check movement---------------
    If SelPlace(1) = 100 Or SelPlace(2) = 100 Then Call MsgBox("Mark your move", vbExclamation, "Error"): Exit Sub
   If Map(SelPlace(1)) > 13 And Map(SelPlace(2)) <= 13 Then
      i = SelPlace(1)
      SelPlace(1) = SelPlace(2)
      SelPlace(2) = i
   End If 'Make sure that Selplace(2) contains the forign parts
   If Map(SelPlace(1)) > 13 And Map(SelPlace(2)) > 13 Or Map(SelPlace(2)) = Map(SelPlace(1)) Then Call MsgBox("Illegal move - You must move yourown peaces", vbExclamation, "Error"): Exit Sub
   If Map(SelPlace(1)) > 10 And Map(SelPlace(1)) <= 13 Then Call MsgBox("You can't move flags and bombs", vbExclamation, "Error"): Exit Sub

   If Map(SelPlace(1)) = 9 Then
      If SelPlace(1) > SelPlace(2) Then
         i = SelPlace(1) - SelPlace(2)
      Else
         i = SelPlace(2) - SelPlace(1)
      End If
      If i > 10 And Int(i / 10) <> i / 10 Then Call MsgBox("Illegal move", vbExclamation, "Error"): Exit Sub
   Else ' Pawn isn't 9
      If SelPlace(1) <> SelPlace(2) + 10 And SelPlace(1) <> SelPlace(2) - 10 And SelPlace(1) <> SelPlace(2) + 1 And SelPlace(1) <> SelPlace(2) - 1 Then Call MsgBox("Illegal move", vbExclamation, "Error"): Exit Sub
   End If
'Barriers
  If SelPlace(1) = 42 Or SelPlace(1) = 43 Or SelPlace(1) = 46 Or SelPlace(1) = 47 Or SelPlace(1) = 52 Or SelPlace(1) = 53 Or SelPlace(1) = 56 Or SelPlace(1) = 57 Then Call MsgBox("You can not move on the barriers", vbExclamation, "Error"): Exit Sub
  If SelPlace(2) = 42 Or SelPlace(2) = 43 Or SelPlace(2) = 46 Or SelPlace(2) = 47 Or SelPlace(2) = 52 Or SelPlace(2) = 53 Or SelPlace(2) = 56 Or SelPlace(2) = 57 Then Call MsgBox("You can not move on the barriers", vbExclamation, "Error"): Exit Sub
'Can move ?
  If CanMove(SelPlace(1)) = False Then Call MsgBox("The move you chose is Illegal."): Exit Sub

   Select Case Map(SelPlace(2))
   Case Is = 0
      Socket.SendData "M:" & SelPlace(1) & "-" & SelPlace(2) & ";"
      Map(SelPlace(2)) = Map(SelPlace(1))
      AttackMap(SelPlace(2)) = AttackMap(SelPlace(1))
      Map(SelPlace(1)) = 0
      AttackMap(SelPlace(1)) = 0
   Case Is > 13 'Enemy's pawn
      Map(SelPlace(2)) = AttackMap(SelPlace(2))
      battle(1) = LoadPicture(App.Path & "\images\" & AttackMap(SelPlace(1)) & ".jpg")
      battle(2) = LoadPicture(App.Path & "\images\" & AttackMap(SelPlace(2)) & ".jpg")
      Socket.SendData "B:" & AttackMap(SelPlace(1)) & "-" & AttackMap(SelPlace(2)) & ";"

      If AttackMap(SelPlace(2)) = 254 And Map(SelPlace(1)) = 10 Then '10 attacks 11
         Socket.SendData "P:" & SelPlace(2) & "-" & Map(SelPlace(1)) & ";K:" & SelPlace(1) & ";"
         Map(SelPlace(2)) = Map(SelPlace(1))
         AttackMap(SelPlace(2)) = AttackMap(SelPlace(1))
         Map(SelPlace(1)) = 0
         AttackMap(SelPlace(1)) = 0
         GoTo postattack
      End If
      If AttackMap(SelPlace(2)) = 242 Then 'Flag
         Socket.SendData "M:" & SelPlace(1) & "-" & SelPlace(2) & ";"
         Socket.SendData "L:;T:;"
         Map(SelPlace(2)) = Map(SelPlace(1))
         AttackMap(SelPlace(2)) = AttackMap(SelPlace(1))
         Map(SelPlace(1)) = 0
         AttackMap(SelPlace(1)) = 0
         Protocol.V
         Exit Sub
      End If
      If AttackMap(SelPlace(2)) = 243 Then 'Bomb
         If Map(SelPlace(1)) <> 8 Then
            Socket.SendData "K:" & SelPlace(1) & ";"
         Else
            Socket.SendData "K:" & SelPlace(1) & ";K:" & SelPlace(2) & ";"
            Socket.SendData "P:" & SelPlace(2) & "-" & Map(SelPlace(1)) & ";"
            Map(SelPlace(2)) = Map(SelPlace(1))
            AttackMap(SelPlace(2)) = AttackMap(SelPlace(1))
         End If
         Map(SelPlace(1)) = 0
         AttackMap(SelPlace(1)) = 0
         GoTo postattack
      End If
      Select Case 255 - Map(SelPlace(2))
      Case Is > Map(SelPlace(1)) 'Kill
         Map(SelPlace(2)) = Map(SelPlace(1))
         AttackMap(SelPlace(2)) = AttackMap(SelPlace(1))
         Map(SelPlace(1)) = 0
         Socket.SendData "M:" & SelPlace(1) & "-" & SelPlace(2) & ";P:" & SelPlace(2) & "-" & Map(SelPlace(2)) & ";"
         AttackMap(SelPlace(1)) = 0
      Case Is < Map(SelPlace(1)) 'Be killed
         Map(SelPlace(1)) = 0
         AttackMap(SelPlace(1)) = 0
         Map(SelPlace(2)) = AttackMap(SelPlace(2))
         Socket.SendData "K:" & SelPlace(1) & ";"
      Case Is = Map(SelPlace(1))
         Map(SelPlace(1)) = 0
         AttackMap(SelPlace(1)) = 0
         Map(SelPlace(2)) = 0
         AttackMap(SelPlace(2)) = 0
         Socket.SendData "K:" & SelPlace(1) & ";K:" & SelPlace(2) & ";"
      End Select
   Case Else 'Error
      Call MsgBox("Illegal move", vbExclamation, "Error"): Exit Sub
      Debug.Print "CASE ERROR"
   End Select
postattack:
   DrawMap
   fMain.Send.Enabled = False
   If CheckVictory = False Then
   Socket.SendData "T:" & SelPlace(1) & "-" & SelPlace(2) & ";"
   Else
   Socket.SendData "M:" & SelPlace(1) & "-" & SelPlace(2) & ";"
   Socket.SendData "S:defeat;L:;"
   Protocol.V
   End If
   lblTurn.ForeColor = RGB(255, 0, 0)
   lblTurn.Caption = "Please wait"
   lblTurn.Font.Size = 12
   SelPlace(1) = 100: SelPlace(2) = 100
End Sub

Private Sub SkinS_Click()
Shell App.Path & "\SkinS.exe", vbNormalFocus
End Sub

Private Sub Socket_Close()
Call D_Buttons
lblTurn.Visible = False
SkinS.Visible = True
Call Form_MouseMove(0, 0, 0, 0)
End Sub

Private Sub Socket_Connect()
Call C_Buttons
MsgBox ("Connection Success")
Protocol.GetMap
lblTurn.Visible = True
SkinS.Visible = False
Call Form_MouseMove(0, 0, 0, 0)
End Sub

Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
If Socket.State <> sckclose Then Socket.Close
Socket.Accept (requestID)
Status.Caption = "Connected"
MsgBox ("Connection Success")
Protocol.GetMap
Domain.Text = Socket.RemoteHostIP
Call C_Buttons
lblTurn.Visible = True
 lblTurn.ForeColor = RGB(255, 0, 0)
 lblTurn.Caption = "Please wait"
 lblTurn.Font.Size = 12
Send.Enabled = False
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
   Dim Cmd() As String, inp As String, i As Integer, j As Integer
   Socket.GetData inp, vbString
   Debug.Print inp
   Cmd() = Split(inp, ";")
   For j = 0 To UBound(Cmd()) - 1
      Data = Right(Cmd(j), Len(Cmd(j)) - 2)
      Select Case UCase(Mid(Cmd(j), 1, 2))
      Case "A:"
         Protocol.A (Data)
      Case "B:"
         Protocol.B (Data)
      Case "C:"
         Protocol.C (Data)
      Case "D:"
         Protocol.D
      Case "G:"
         Protocol.GetMap
      Case "K:"
         Protocol.K (Data)
      Case "L:"
         Protocol.L (Data)
      Case "M:"
         Protocol.M (Data)
      Case "Q:"
         Protocol.Q (Data)
      Case "S:"
         FrmFX.Play_File (Data)
      Case "T:"
         If CanMove = True Then
         Send.Enabled = True
         lblTurn.ForeColor = RGB(0, 255, 0)
         lblTurn.Caption = "Your Turn"
         lblTurn.Font.Size = 14
         Else
         fMain.Socket.SendData "T:CM;"
         End If
         If Data <> "" Then Protocol.T (Data)
      Case "P:"
         Protocol.P (Data)
      Case "U:"
         For i = 1 To 100
            tmp = UCase(Mid(Data, i, 1))
            If tmp = "K" Then
               AttackMap(i - 1) = Map(i - 1)
            Else
               If Asc(tmp) > 48 And Asc(tmp) < 58 Then AttackMap(i - 1) = 255 - Val(tmp)
               If tmp = "0" Then AttackMap(i - 1) = 245
               If tmp = "F" Then AttackMap(i - 1) = 242
               If tmp = "B" Then AttackMap(i - 1) = 243
            End If
         Next i
      Case "V:"
         Protocol.V (Data)
      Case Else
         Socket.SendData ("illegal input.")
      End Select
   Next j
End Sub


Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox ("Connection Error")
Socket.Close
lblTurn.Visible = False
Call D_Buttons
End Sub

Private Sub Status_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(0, 0, 0, 0)
End Sub
