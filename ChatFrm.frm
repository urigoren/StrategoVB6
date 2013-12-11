VERSION 5.00
Begin VB.Form ChatFrm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Web Stratego Chat"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   ClipControls    =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Clr 
      Caption         =   "Clr"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.ListBox Tprotocol 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   2400
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Action 
      Caption         =   "&Send"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox TInput 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   3735
   End
End
Attribute VB_Name = "ChatFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Action_Click()
If fMain.Socket.State <> sckConnected Then MsgBox ("There isn't any connection established"): Exit Sub
Tprotocol.AddItem "Player: " & TInput.Text
Tprotocol.ListIndex = Tprotocol.ListCount - 1
fMain.Socket.SendData ("C:" & TInput.Text & ";")
TInput.Text = vbNullString
End Sub

Public Sub Add_Line(incomming As String)
Tprotocol.AddItem "Opponent: " & incomming
Tprotocol.ListIndex = Tprotocol.ListCount - 1
End Sub

Private Sub Clr_Click()
Tprotocol.Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MoveForm(Me.hWnd, Button)
End Sub

Private Sub TInput_KeyPress(KeyAscii As Integer)
Dim tmp As Integer
If KeyAscii = 10 Or KeyAscii = 13 Then Call Action_Click
If InStr(1, TInput.Text, ";") > 0 Then
tmp = TInput.SelStart
TInput.Text = Replace(TInput.Text, ";", ":")
TInput.SelStart = tmp
End If
End Sub
