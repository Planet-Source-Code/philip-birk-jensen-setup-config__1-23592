VERSION 5.00
Begin VB.Form About 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Setup Config"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   240
      Left            =   1710
      TabIndex        =   1
      Top             =   2925
      Width           =   960
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"about.frx":0000
      Height          =   1590
      Left            =   45
      TabIndex        =   0
      Top             =   1530
      Width           =   2625
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2745
      Y1              =   1485
      Y2              =   1485
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   0
      Picture         =   "about.frx":0104
      Top             =   0
      Width           =   2745
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()

   Unload Me

End Sub
