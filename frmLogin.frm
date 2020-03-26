VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log In"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7170
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   3360
      Width           =   2400
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   3360
      Width           =   2400
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2520
      Width           =   5175
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   3120
      Picture         =   "frmLogin.frx":28190
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Please Login To Continue"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2415
      TabIndex        =   6
      Top             =   1200
      Width           =   2520
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS As Recordset
Dim sSQL As String

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdLogin_Click()
    If txtUsername = "" Then
        MsgBox "Please enter the user name", vbExclamation, App.Title
        txtUsername.SetFocus
        Exit Sub
    End If
    
    If txtPassword = "" Then
        MsgBox "Please enter the password", vbExclamation, App.Title
        txtPassword.SetFocus
        Exit Sub
    End If
    
    Call db_connection
    
    sSQL = "SELECT * FROM tblUser WHERE userid='" & txtUsername & "' AND passwd='" & txtPassword & "'"
    Set RS = gDB.OpenRecordset(sSQL)
    
    If RS.RecordCount = 0 Then
        MsgBox "Incorrect the user name or password", vbCritical, App.Title
        RS.Close
        Exit Sub
    End If
    
    gUserId = txtUsername
    gPermission = RS.Fields("permission").Value
    
    RS.Close
    
    Call db_close
    
    Unload Me
    
    If gPermission = "normal" Then
        mainForm.tbMain.Buttons("tbAddExpenses").Visible = False
        mainForm.tbMain.Buttons("tbAllExpenses").Visible = False
        mainForm.tbMain.Buttons("tbManageUsers").Visible = False
    Else
        mainForm.tbMain.Buttons("tbAddExpenses").Visible = True
        mainForm.tbMain.Buttons("tbAllExpenses").Visible = True
        mainForm.tbMain.Buttons("tbManageUsers").Visible = True
    End If
    mainForm.Show
End Sub


Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
End Sub
