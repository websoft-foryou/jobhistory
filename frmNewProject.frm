VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNewProject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Project"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7605
   Icon            =   "frmNewProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
      _ExtentX        =   5741
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
      Format          =   169279489
      CurrentDate     =   43902
   End
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
      Left            =   3960
      TabIndex        =   21
      Top             =   9000
      Width           =   2400
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   1200
      TabIndex        =   20
      Top             =   9000
      Width           =   2400
   End
   Begin VB.TextBox txtRemark 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   6720
      Width           =   6855
   End
   Begin VB.TextBox txtWorkStatus 
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
      TabIndex        =   16
      Top             =   5520
      Width           =   6855
   End
   Begin VB.TextBox txtReceiveAmt 
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
      Left            =   3960
      TabIndex        =   14
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox txtTotalAmt 
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
      TabIndex        =   12
      Top             =   4320
      Width           =   3135
   End
   Begin VB.TextBox txtInvoiceNo 
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
      Left            =   3960
      TabIndex        =   10
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtPurchaseNo 
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
      TabIndex        =   8
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox txtWorkType 
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
      TabIndex        =   4
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtCompany 
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
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtJobNo 
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
      Width           =   3135
   End
   Begin VB.Label lblId 
      Caption         =   "lblId"
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Remark"
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
      TabIndex        =   19
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Work Status"
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
      TabIndex        =   17
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Payment Received"
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
      Left            =   3960
      TabIndex        =   15
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label7 
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
      TabIndex        =   13
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Invoice No"
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
      Left            =   3960
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Purchase Order No"
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
      TabIndex        =   9
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label4 
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
      Left            =   3960
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Type of work"
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
      Top             =   1560
      Width           =   1815
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
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Job No"
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
      Width           =   855
   End
End
Attribute VB_Name = "frmNewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim RS As Recordset


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdSave_Click()
    Dim update_where As String
    Dim insert_where1 As String, insert_where2 As String
    
    If Trim(txtJobNo) = "" Then
        MsgBox "Please enter the job no.", vbExclamation, App.Title
        txtJobNo.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtJobNo) = False Then
        MsgBox "Incorrect data type of job no. It have to numeric.", vbExclamation, App.Title
        txtJobNo.SetFocus
        Exit Sub
    End If
    
    If Trim(txtCompany) = "" Then
        MsgBox "Please enter the company name.", vbExclamation, App.Title
        txtCompany.SetFocus
        Exit Sub
    End If
    
    If Trim(txtWorkType) = "" Then
        MsgBox "Please enter the type of work.", vbExclamation, App.Title
        txtWorkType.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPurchaseNo) = "" Then
        MsgBox "Please enter the purchase no.", vbExclamation, App.Title
        txtPurchaseNo.SetFocus
        Exit Sub
    End If

    If Trim(txtTotalAmt) = "" Then
        MsgBox "Please enter the total amount.", vbExclamation, App.Title
        txtTotalAmt.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtTotalAmt) = False Then
        MsgBox "Incorrect data type of total amount. It have to numeric.", vbExclamation, App.Title
        txtTotalAmt.SetFocus
        Exit Sub
    End If
    
    
    If Trim(txtReceiveAmt) <> "" And IsNumeric(txtReceiveAmt) = False Then
        MsgBox "Incorrect data type of payment received. It have to numeric.", vbExclamation, App.Title
        txtReceiveAmt.SetFocus
        Exit Sub
    End If
 
    
    If Me.Caption = "Edit Project" Then
        sSQL = "SELECT * FROM tblProject WHERE jobno=" & txtJobNo & " AND id <>" & lblId
    Else
        sSQL = "SELECT * FROM tblProject WHERE jobno=" & txtJobNo
    End If
    
    Set RS = gDB.OpenRecordset(sSQL)
    If RS.RecordCount > 0 Then
        MsgBox "The job no already exist. Please enter the another job no.", vbExclamation, App.Title
        txtJobNo.SetFocus
        RS.Close
        Exit Sub
    End If
    RS.Close
    
    
    If Trim(txtInvoiceNo) <> "" Then
        update_where = update_where & ",invoice_no='" & txtInvoiceNo & "'"
        insert_where1 = insert_where1 & ", invoice_no"
        insert_where2 = insert_where2 & ",'" & txtInvoiceNo & "'"
    End If
    If Trim(txtReceiveAmt) <> "" Then
        update_where = update_where & ", receive_amt=" & txtReceiveAmt
        insert_where1 = insert_where1 & ", receive_amt"
        insert_where2 = insert_where2 & "," & txtReceiveAmt
    End If

    
    If Me.Caption = "Edit Project" Then
        sSQL = "UPDATE tblProject SET jobno=" & txtJobNo & ", company_name='" & txtCompany & "', work_type='" & txtWorkType & "', work_date=#" & dtpDate.Value & "#, purchase_no='" & txtPurchaseNo & _
            "', total_amt=" & txtTotalAmt & ", work_status='" & txtWorkStatus & "', remark='" & txtRemark & "'" & update_where & " WHERE id=" & lblId
    Else
        sSQL = "INSERT INTO tblProject(userid, jobno, company_name, work_type, work_date, purchase_no, total_amt, work_status, remark, project_status" & insert_where1 & ") VALUES('" & _
            gUserId & "', " & txtJobNo & ", '" & txtCompany & "','" & txtWorkType & "',#" & dtpDate.Value & "#,'" & txtPurchaseNo & "'," & txtTotalAmt & _
            ",'" & txtWorkStatus & "','" & txtRemark & "', 0" & insert_where2 & ")"
    End If
    gDB.Execute sSQL
    
    Unload Me
    
    Call mainForm.LoadData
End Sub

Private Sub Form_Load()

    dtpDate.Value = Date
    Call db_connection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call db_close
End Sub
