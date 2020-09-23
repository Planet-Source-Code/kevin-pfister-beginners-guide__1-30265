VERSION 5.00
Begin VB.Form FrmCopying 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Copying Examples"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdHard 
      Caption         =   "Harder Way (*.Txt Only)"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton CmdEasy 
      Caption         =   "Easy Way"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Frame FrameDest 
      Caption         =   "Destination File"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin VB.TextBox TxtDest 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "C:\..."
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame FrameSource 
      Caption         =   "Source File"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox TxtSource 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "C:\..."
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "FrmCopying"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdEasy_Click()
        Call FileCopy(TxtSource, TxtDest)
End Sub

Private Sub CmdHard_Click()
        Open TxtSource For Input As #1
        Open TxtDest For Output As #2
        Do While EOF(1)
                Input #1, Fcopy$
                Print #2, Fcopy$
        Loop
        Close #1
        Close #2
End Sub
