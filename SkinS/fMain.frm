VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skin Selector"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1845
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   1845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkNeg 
      Caption         =   "Flip Skin"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdDep 
      Caption         =   "O&K"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   615
   End
   Begin VB.FileListBox SkinList 
      Height          =   1845
      Left            =   120
      Pattern         =   "*.ssz"
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Please select the skin you wish to deploy"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
End
End Sub

Private Sub cmdDep_Click()
Dim i As Integer
If SkinList.ListIndex < 0 Then
Call MsgBox("No skin selected", vbExclamation, "Error")
Else
Call VBUnzip(App.Path & "\" & SkinList.List(SkinList.ListIndex), App.Path & "\images", 0, 1, 0, 0, 0, 0)
End If
If ChkNeg = 1 Then
MkDir App.Path & "\images\tmp"
For i = 1 To 255
If Dir(App.Path & "\images\" & i & ".jpg") <> vbNullString Then _
Name App.Path & "\images\" & i & ".jpg" As App.Path & "\images\tmp\" & i & ".jpg"
Next i
For i = 1 To 255
If Dir(App.Path & "\images\tmp\" & i & ".jpg") <> vbNullString Then _
Name App.Path & "\images\tmp\" & i & ".jpg" As App.Path & "\images\" & (255 - i) & ".jpg"
Next i
RmDir App.Path & "\images\tmp"
End If
End Sub

Private Sub Form_Load()
Dim dest As String
SkinList.Path = App.Path
If Command <> "" And Dir(Command) <> "" Then
If UCase(Right(Command(), 4)) <> ".SSZ" Then Exit Sub
If InStr(1, LCase(Command), LCase(App.Path)) > 0 Then Exit Sub
dest = App.Path & "\images\" & GetFileTitle(Command)
Call FileCopy(Command, dest)
Call VBUnzip(dest, App.Path & "\images", 0, 1, 0, 0, 0, 0)
End
End If
End Sub

Private Function GetFileTitle(ByVal Filename As String) As String
Dim i As Integer, p As Integer
For i = 1 To Len(Filename)
If Mid(Filename, i, 1) = "\" Then p = i
Next i
GetFileTitle = Right(Filename, Len(Filename) - p)
End Function
