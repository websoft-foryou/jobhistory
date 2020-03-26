VERSION 5.00
Begin VB.Form frmImage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin VB.Image imgView 
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub
