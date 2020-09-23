VERSION 5.00
Begin VB.Form FrmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example of Text File Opening"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGo 
      Caption         =   "Load"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TxtOutput 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FrmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGo_Click()
   
        F = FreeFile    'Assigns a Free File Number
        FileName$ = "C:\..."    'Place the Filename Here

        Open FileName$ For Input As #F  'Opens the File to read from
                TxtOutput.Text = Input$(LOF(F), F) 'Places the contents of the file into the Textbox(Txtoutput)
        Close #F        'Close the File
End Sub
