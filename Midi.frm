VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Midi player"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   Icon            =   "Midi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select midi file"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
T = mciSendString("play " & Label1.Caption, 0&, 0, 0)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Temp = mciSendString("close " & Label1.Caption, 0&, 0, 0)
End Sub

Private Sub Command4_Click()
MsgBox "Created by Johannes.B 2000"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Temp = mciSendString("close " & Label1.Caption, 0&, 0, 0)
End
End Sub
