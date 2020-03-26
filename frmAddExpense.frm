VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddExpense 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Expense"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCleared 
      Cancel          =   -1  'True
      Caption         =   "Cleared"
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
      Left            =   5280
      TabIndex        =   11
      Top             =   5880
      Width           =   2160
   End
   Begin VB.CommandButton cmdSubmitted 
      Caption         =   "Signed not Submitted"
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
      Left            =   2640
      TabIndex        =   10
      Top             =   5880
      Width           =   2520
   End
   Begin VB.CommandButton cmdSigned 
      Caption         =   "To be Signed"
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
      Left            =   360
      TabIndex        =   9
      Top             =   5880
      Width           =   2160
   End
   Begin VB.TextBox txtAmount 
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
      TabIndex        =   7
      Top             =   5040
      Width           =   7095
   End
   Begin VB.TextBox txtDescription 
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
      TabIndex        =   5
      Top             =   3960
      Width           =   7095
   End
   Begin VB.TextBox txtCompany 
      Enabled         =   0   'False
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
      TabIndex        =   2
      Top             =   1800
      Width           =   7095
   End
   Begin VB.TextBox txtJobNo 
      Enabled         =   0   'False
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
      Top             =   720
      Width           =   7095
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   2895
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   118030337
      CurrentDate     =   43902
   End
   Begin VB.Label lblId 
      Caption         =   "lblId"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblProjectId 
      Caption         =   "lblProjectId"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Total Amount"
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
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Expense Description"
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
      TabIndex        =   6
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
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
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Company Name"
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
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Job Number"
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
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmAddExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim RS As Recordset
Dim save_flag As Integer



Private Sub cmdSigned_Click()
    save_flag = 0
    Call SaveData
End Sub



Private Sub cmdSubmitted_Click()
    save_flag = 1
    Call SaveData
End Sub




Private Sub cmdCleared_Click()
    save_flag = 2
    Call SaveData
End Sub



Private Sub SaveData()

    If Trim(txtDescription) = "" Then
        MsgBox "Please enter the expense description.", vbExclamation, App.Title
        txtDescription.SetFocus
        Exit Sub
    End If
    
    If Trim(txtAmount) = "" Then
        MsgBox "Please enter the amount.", vbExclamation, App.Title
        txtAmount.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtAmount) = False Then
        MsgBox "Incorrect data type of total amount. It have to numeric.", vbExclamation, App.Title
        txtAmount.SetFocus
        Exit Sub
    End If
    
    Call db_connection
    
    If Me.Caption = "Add New Expense" Then
        sSQL = "INSERT INTO tblExpense(projectid, description, amount, edate, amount_type) VALUES(" & _
            lblProjectId & ", '" & txtDescription & "'," & txtAmount & ",#" & dtpDate.Value & "#," & save_flag & ")"
    Else
        sSQL = "UPDATE tblExpense SET description='" & txtDescription & "', amount=" & txtAmount & ", edate=#" & dtpDate.Value & "#, amount_type=" & save_flag & " WHERE id=" & lblId
    End If
    
    gDB.Execute sSQL
    
    Call db_close
    
    
    
    If frmExpense.Visible = True Then
        Call frmExpense.LoadData
    Else
        MsgBox "Saved Successfully.", vbInformation, App.Title
    End If
    
    Unload Me
    
End Sub


Private Sub Form_Load()

    dtpDate.Value = Date
End Sub
