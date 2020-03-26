VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAllExpenses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "All Expenses"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdViewAll 
      Caption         =   "All Data"
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
      Left            =   9720
      TabIndex        =   4
      Top             =   240
      Width           =   1920
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
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
      TabIndex        =   3
      Top             =   240
      Width           =   1920
   End
   Begin MSFlexGridLib.MSFlexGrid flexList 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   16325
      _Version        =   393216
      Cols            =   5
      BackColorBkg    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.Label Label1 
      Caption         =   "To"
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
      TabIndex        =   6
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "From"
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
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmAllExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim RS As Recordset

Private Sub cmdSubmit_Click()

    Call LoadData("search")
End Sub

Private Sub cmdViewAll_Click()
    Call LoadData("all")
End Sub

Private Sub Form_Load()

    Dim aColumnWidth As Variant
    Dim aColumnText As Variant
    Dim i As Integer
    
    aColumnWidth = Array(1000, 2800, 2800, 2800, 2800)
    aColumnText = Array("No", "Company", "Job No", "Work", "Total")
    
    With flexList
        For i = 0 To .Cols - 1
            .ColWidth(i) = aColumnWidth(i)
            .TextMatrix(0, i) = aColumnText(i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        .RowHeight(0) = 350
        .ColAlignment(0) = flexAlignCenterCenter
    End With
    
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
    
End Sub



Private Sub LoadData(flag As String)
    Call db_connection
    
    flexList.Rows = 1
    If flag = "search" Then
        sSQL = "SELECT  * FROM tblproject P INNER JOIN (" & _
                    "SELECT projectid, SUM(amount) AS expense_amount FROM tblExpense WHERE edate BETWEEN #" & dtpFromDate.Value & "# AND #" & dtpToDate.Value & "# GROUP BY projectid" & _
                ") E ON P.id=E.projectid ORDER BY jobno"
    Else
        sSQL = "SELECT  * FROM tblproject P INNER JOIN (" & _
                    "SELECT projectid, SUM(amount) AS expense_amount FROM tblExpense GROUP BY projectid" & _
                ") E ON P.id=E.projectid ORDER BY jobno"
    End If
    Set RS = gDB.OpenRecordset(sSQL)
    
    Do While Not RS.EOF
        flexList.AddItem flexList.Rows & vbTab & RS.Fields("company_name").Value & vbTab & RS.Fields("jobno") & vbTab & RS.Fields("work_type").Value & vbTab & RS.Fields("expense_amount").Value
        flexList.RowHeight(flexList.Rows - 1) = 350
        RS.MoveNext
    Loop
    
    RS.Close
    Call db_close
End Sub
