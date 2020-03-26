VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmManageUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Users"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10305
   Icon            =   "frmManageUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flexList 
      Height          =   4095
      Left            =   360
      TabIndex        =   9
      Top             =   3120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   4
      ForeColor       =   -2147483647
      BackColorBkg    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   7440
      TabIndex        =   8
      Top             =   1800
      Width           =   2400
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create User"
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
      Left            =   7440
      TabIndex        =   7
      Top             =   1200
      Width           =   2400
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Password"
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
      Left            =   7440
      TabIndex        =   6
      Top             =   600
      Width           =   2400
   End
   Begin VB.OptionButton optNormal 
      Caption         =   "Normal User"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.OptionButton optAdmin 
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Value           =   -1  'True
      Width           =   1695
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
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   6495
   End
   Begin VB.TextBox txtUserid 
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
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "UserID"
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
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmManageUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS As Recordset
Dim sSQL As String


Private Sub cmdCreate_Click()

    If Trim(txtUserid) = "" Then
        MsgBox "Please enter the user id.", vbExclamation, App.Title
        txtUserid.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPassword) = "" Then
        MsgBox "Please enter the password.", vbExclamation, App.Title
        txtPassword.SetFocus
        Exit Sub
    End If
    
    Dim sPermission As String
    
    sSQL = "SELECT * FROM tblUser WHERE userid='" & txtUserid & "'"
    Set RS = gDB.OpenRecordset(sSQL)
    If RS.RecordCount > 0 Then
        MsgBox "The user already exist. Please enter the another user id.", vbExclamation, App.Title
        RS.Close
        Exit Sub
    End If
    RS.Close
    
    If optAdmin.Value = True Then
        sPermission = "admin"
    End If
    If optNormal.Value = True Then
        sPermission = "normal"
    End If
    
    sSQL = "INSERT INTO tblUser(userid, passwd, permission) VALUES('" & txtUserid & "', '" & txtPassword & "','" & sPermission & "')"
    gDB.Execute sSQL
    
    Call LoadData
    
End Sub



Private Sub cmdDelete_Click()

    If flexList.Row < 1 Then
        MsgBox "Please select data for remove.", vbExclamation, App.Title
        Exit Sub
    End If
    
    If MsgBox("Are you really remove the user [" & flexList.TextMatrix(flexList.Row, 1) & "]?", vbYesNo + vbInformation, App.Title) = vbNo Then
        Exit Sub
    End If
    
    sSQL = "SELECT * FROM tblProject WHERE userid='" & flexList.TextMatrix(flexList.Row, 1) & "'"
    Set RS = gDB.OpenRecordset(sSQL)
    Do While Not RS.EOF
        sSQL = "DELETE FROM tblExpense WHERE projectid=" & RS.Fields("id").Value
        gDB.Execute sSQL
        RS.MoveNext
    Loop
    RS.Close
    
    sSQL = "DELETE FROM tblProject WHERE userid ='" & flexList.TextMatrix(flexList.Row, 1) & "'"
    gDB.Execute sSQL
    
    
    sSQL = "DELETE FROM tblUser WHERE userid='" & flexList.TextMatrix(flexList.Row, 1) & "'"
    gDB.Execute sSQL
    
    Call LoadData
End Sub



Private Sub cmdReset_Click()

    If Trim(txtPassword) = "" Then
        MsgBox "Please enter the password", vbExclamation, App.Title
        Exit Sub
    End If
    
    If MsgBox("Are you really reset password of user [" & flexList.TextMatrix(flexList.Row, 1) & "]?", vbYesNo + vbInformation, App.Title) = vbNo Then
        Exit Sub
    End If
    
    sSQL = "UPDATE tblUser SET passwd='" & txtPassword & "' WHERE userid='" & flexList.TextMatrix(flexList.Row, 1) & "'"
    gDB.Execute sSQL
    
    flexList.TextMatrix(flexList.Row, 2) = txtPassword
End Sub



Private Sub flexList_Click()
    txtUserid = flexList.TextMatrix(flexList.Row, 1)
    txtPassword = flexList.TextMatrix(flexList.Row, 2)
    
    If flexList.TextMatrix(flexList.Row, 3) = "admin" Then
        optAdmin.Value = True
    End If
    
    If flexList.TextMatrix(flexList.Row, 3) = "normal" Then
        optNormal.Value = True
    End If
End Sub




Private Sub Form_Load()
    Dim aColumnWidth As Variant
    Dim aColumnText As Variant
    Dim i As Integer
    
    aColumnWidth = Array(1000, 3000, 1800, 3000)
    aColumnText = Array("No", "User id", "Password", "Access")
    
    With flexList
        For i = 0 To flexList.Cols - 1
            .ColWidth(i) = aColumnWidth(i)
            .TextMatrix(0, i) = aColumnText(i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        .RowHeight(0) = 350
        .Rows = 1
        .ColAlignment(0) = flexAlignCenterCenter
    End With
    
    Call db_connection
    Call LoadData
    
End Sub


Private Sub LoadData()
    
    flexList.Rows = 1
    
    If optAdmin.Value = True Then
        sSQL = "SELECT * FROM tblUser WHERE permission='admin'"
    End If
    If optNormal.Value = True Then
        sSQL = "SELECT * FROM tblUser WHERE permission='normal'"
    End If
    
    Set RS = gDB.OpenRecordset(sSQL)
    
    Do While Not RS.EOF
        flexList.AddItem flexList.Rows & vbTab & RS.Fields("userid").Value & vbTab & RS.Fields("passwd") & vbTab & RS.Fields("permission").Value
        flexList.RowHeight(flexList.Rows - 1) = 350
        RS.MoveNext
    Loop
    
    RS.Close
    
    txtUserid = ""
    txtPassword = ""
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call db_close
End Sub



Private Sub optAdmin_Click()
    Call LoadData
End Sub

Private Sub optNormal_Click()
    Call LoadData
End Sub
