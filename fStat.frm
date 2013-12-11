VERSION 5.00
Begin VB.Form fStat 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Statics"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.PictureBox Place 
      BorderStyle     =   0  'None
      Height          =   630
      Index           =   0
      Left            =   0
      Picture         =   "fStat.frx":0000
      ScaleHeight     =   630
      ScaleWidth      =   630
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   630
   End
   Begin VB.Label tot 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label tot 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Qty 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "fStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAction_Click()
Dim i As Byte
For i = 0 To 23
Qty(i).Caption = "0"
Next i
For i = 0 To 99
Select Case AttackMap(i)
Case 1 To 10
Qty(AttackMap(i) - 1) = Val(Qty(AttackMap(i) - 1)) + 1
Case 12
Qty(11) = Val(Qty(11)) + 1
Case 13
Qty(10) = Val(Qty(10)) + 1
Case 245 To 255
Qty(267 - AttackMap(i)) = Val(Qty(267 - AttackMap(i))) + 1
Case 242
Qty(22) = Val(Qty(22)) + 1
Case 243
Qty(23) = Val(Qty(23)) + 1
End Select
Next i
tot(0).Caption = "0": tot(1).Caption = "0"
For i = 0 To 23
tot(Int(i / 12)) = Val(tot(Int(i / 12))) + Val(Qty(i).Caption)
Next i
tot(0).Caption = "Total peaces on board: " & tot(0).Caption
tot(1).Caption = "Total peaces on board: " & tot(1).Caption
End Sub

Private Sub cmdClose_Click()
Unload Me
fMain.Show
End Sub

Private Sub Form_Load()
Dim i As Byte
For i = 1 To 23
Load Place(i)
Load Qty(i)
Place(i).Visible = True
Qty(i).Visible = True
Next i
For i = 0 To 11
Place(i).Top = 0
Place(i).Left = i * Place(0).Width
Place(i) = LoadPicture(App.Path & "\images\" & (i + 1) & ".jpg")
Place(i).Tag = i + 1
Qty(i).ForeColor = RGB(0, 0, 255)
Qty(i).Left = Place(i).Left
Qty(i).Top = Place(i).Top + Place(0).Height
Next i
For i = 12 To 23
Place(i).Top = 900
Place(i).Left = (i - 12) * Place(0).Width
Place(i) = LoadPicture(App.Path & "\images\" & (266 - i) & ".jpg")
Place(i).Tag = 266 - i
Qty(i).ForeColor = RGB(255, 0, 0)
Qty(i).Left = Place(i).Left
Qty(i).Top = Place(i).Top + Place(0).Height
Next i
Place(10) = LoadPicture(App.Path & "\images\13.jpg")
Place(10).Tag = 13
Place(22) = LoadPicture(App.Path & "\images\242.jpg")
Place(22).Tag = 242
Call cmdAction_Click
End Sub

