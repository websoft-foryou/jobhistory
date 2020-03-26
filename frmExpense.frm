VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExpense 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expenses"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   8625
      Left            =   9570
      TabIndex        =   2
      Top             =   570
      Width           =   3495
      Begin VB.Label lblProfit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Job Number"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   21
         Top             =   6915
         Width           =   3330
      End
      Begin VB.Label lblTotalExpense 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Job Number"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   20
         Top             =   5715
         Width           =   3330
      End
      Begin VB.Label lblTotalPaid 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Job Number"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   19
         Top             =   4515
         Width           =   3330
      End
      Begin VB.Label lblTotalAmt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Job Number"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   18
         Top             =   3300
         Width           =   3330
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Job Number"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   17
         Top             =   2115
         Width           =   3330
      End
      Begin VB.Label lblJobNo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Job Number"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   16
         Top             =   840
         Width           =   3330
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Expenses"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   795
         TabIndex        =   15
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Paid"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1065
         TabIndex        =   14
         Top             =   3840
         Width           =   1365
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   930
         TabIndex        =   13
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Profit"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1335
         TabIndex        =   11
         Top             =   6240
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1290
         Left            =   0
         TabIndex        =   10
         Top             =   7320
         Width           =   3480
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   0
         TabIndex        =   9
         Top             =   2520
         Width           =   3480
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   0
         TabIndex        =   8
         Top             =   3720
         Width           =   3480
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   0
         TabIndex        =   7
         Top             =   4920
         Width           =   3480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   0
         TabIndex        =   6
         Top             =   6120
         Width           =   3480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   0
         TabIndex        =   5
         Top             =   1320
         Width           =   3480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Number"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   4
         Top             =   255
         Width           =   3390
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   0
         TabIndex        =   3
         Top             =   90
         Width           =   3480
      End
   End
   Begin MSComctlLib.Toolbar tbExpense 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   1376
      ButtonWidth     =   1482
      ButtonHeight    =   1217
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New"
            Key             =   "tbExpenseNew"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Key             =   "tbExpenseEdit"
            ImageKey        =   "edit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "tbExpenseDelete"
            ImageKey        =   "delete"
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8640
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExpense.frx":0000
               Key             =   "new"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExpense.frx":02D5
               Key             =   "edit"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExpense.frx":07D8
               Key             =   "delete"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblProjectId 
         Height          =   255
         Left            =   8280
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flexList 
      Height          =   7830
      Left            =   0
      TabIndex        =   1
      Top             =   750
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   13811
      _Version        =   393216
      Cols            =   7
      BackColorBkg    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Label lblCleared 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   8055
      TabIndex        =   33
      Top             =   8760
      Width           =   1470
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSubmitted 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   4875
      TabIndex        =   32
      Top             =   8760
      Width           =   1470
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSigned 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1665
      TabIndex        =   31
      Top             =   8760
      Width           =   1470
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cleared"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6390
      TabIndex        =   30
      Top             =   8745
      Width           =   1635
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Signed Not Submitted "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   3315
      TabIndex        =   29
      Top             =   8595
      Width           =   1530
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To be Signed"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      Left            =   105
      TabIndex        =   28
      Top             =   8580
      Width           =   1470
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   7980
      TabIndex        =   27
      Top             =   8580
      Width           =   1605
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   6390
      TabIndex        =   26
      Top             =   8580
      Width           =   1605
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   4800
      TabIndex        =   25
      Top             =   8580
      Width           =   1605
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   3210
      TabIndex        =   24
      Top             =   8580
      Width           =   1605
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1620
      TabIndex        =   23
      Top             =   8580
      Width           =   1605
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   30
      TabIndex        =   22
      Top             =   8580
      Width           =   1600
   End
End
Attribute VB_Name = "frmExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim RS As Recordset

Private Sub Form_Load()

    Dim aColumnWidth As Variant
    Dim aColumnText As Variant
    Dim i As Integer
    
    aColumnWidth = Array(700, 1000, 2200, 1200, 1500, 2500, 0)
    aColumnText = Array("No", "Job No", "Expense Description", "Amount", "Date", "Category", "ID")
    
    With flexList
        For i = 0 To .Cols - 1
            .ColWidth(i) = aColumnWidth(i)
            .TextMatrix(0, i) = aColumnText(i)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .RowHeight(0) = 350
        .Rows = 1
    End With
        
End Sub



Public Sub LoadData()
    Dim category As String
    Dim total_signed As Double, total_submitted As Double, total_cleared As Double, total_expense As Double
    
    If lblProjectId = "" Then Exit Sub
    
    Call db_connection
    
    flexList.Rows = 1
    
    sSQL = "SELECT * FROM tblExpense WHERE projectid=" & lblProjectId
    Set RS = gDB.OpenRecordset(sSQL)
    
    Do While Not RS.EOF
        If RS.Fields("amount_type").Value = 0 Then
            category = "To be Signed"
            total_signed = total_signed + RS.Fields("amount").Value
        ElseIf RS.Fields("amount_type").Value = 1 Then
            category = "Signed not Submitted"
            total_submitted = total_submitted + RS.Fields("amount").Value
        Else
            category = "Cleared"
            total_cleared = total_cleared + RS.Fields("amount").Value
        End If
        
        flexList.AddItem flexList.Rows & vbTab & lblJobNo & vbTab & RS.Fields("description").Value & vbTab & RS.Fields("amount").Value & vbTab & RS.Fields("edate").Value & vbTab & category & vbTab & RS.Fields("id").Value
        flexList.RowHeight(flexList.Rows - 1) = 350
        RS.MoveNext
    Loop
    
    RS.Close
    
    total_expense = total_signed + total_submitted + total_cleared
    lblSigned = CStr(total_signed)
    lblSubmitted = CStr(total_submitted)
    lblCleared = CStr(total_cleared)
    lblTotalExpense = CStr(total_expense)
    lblProfit = CStr(Val(lblTotalAmt) - total_expense)
    
    Call db_close
    
End Sub



Private Sub tbExpense_ButtonClick(ByVal Button As MSComctlLib.Button)

    With frmAddExpense
        .lblProjectId = lblProjectId
                
        If Button.Key = "tbExpenseNew" Then
            .Caption = "Add New Expense"
            .txtJobNo = lblJobNo
            .txtCompany = lblCompany
            .dtpDate.Value = Date
            .txtDescription = ""
            .txtAmount = ""
            .Show vbModal
        ElseIf Button.Key = "tbExpenseEdit" Then
            If flexList.Row < 1 Then
                MsgBox "Please select data to edit.", vbExclamation, App.Title
                Exit Sub
            End If
            .Caption = "Edit Expense"
            .lblId = flexList.TextMatrix(flexList.Row, 6)
            .txtJobNo = lblJobNo
            .txtCompany = lblCompany
            .dtpDate.Value = Date
            .txtDescription = flexList.TextMatrix(flexList.Row, 2)
            .txtAmount = flexList.TextMatrix(flexList.Row, 3)
            .Show vbModal
        ElseIf Button.Key = "tbExpenseDelete" Then
            If flexList.Row < 1 Then
                MsgBox "Please select data to remove.", vbExclamation, App.Title
                Exit Sub
            End If
            
            If MsgBox("Are you really delete it?", vbYesNo + vbInformation, App.Title) = vbNo Then Exit Sub
            Call db_connection
            sSQL = "DELETE FROM tblExpense WHERE id=" & flexList.TextMatrix(flexList.Row, 6)
            gDB.Execute sSQL
            Call db_close
            
            Call LoadData
        End If
    End With
    
End Sub
