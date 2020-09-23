VERSION 5.00
Begin VB.Form FrmApi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Api Examples"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Swap 
      Caption         =   "Swap Buttons"
      Height          =   1215
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton CmdNorm 
         Caption         =   "Normal"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdSwap 
         Caption         =   "Swap"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Hide 
      Caption         =   "Cursor"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton CmdShow 
         Caption         =   "Show"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdHide 
         Caption         =   "Hide"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdHide_Click()
ShowCursor& (0)
End Sub

Private Sub CmdNorm_Click()
SwapMouseButton& (0)
End Sub

Private Sub CmdShow_Click()
ShowCursor& (1)
End Sub

Private Sub CmdSwap_Click()
SwapMouseButton& (1)
End Sub
