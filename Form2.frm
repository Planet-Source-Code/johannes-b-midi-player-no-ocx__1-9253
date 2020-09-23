VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open midi file"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   2880
      Pattern         =   "*.mid"
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim JB As String

Private Sub Command1_Click()
Form1.Label1.Caption = JB
Form2.Hide
End Sub

Private Sub Command2_Click()
Form2.Hide
End Sub

Private Sub Dir1_Change()
     File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveHandler
Dir1.Path = Drive1.Drive
Exit Sub

DriveHandler:
    Error.Show
    Error.Label1.Caption = "Drive not ready!"
    Exit Sub
End Sub

Private Sub File1_Click()
If Right(File1.Path, 1) <> "\" Then
        JB = File1.Path & "\" & File1.FileName
      Else
        JB = File1.Path & File1.FileName
        End If
        
        End Sub


Private Sub File1_DblClick()
Command1.Value = True
End Sub


