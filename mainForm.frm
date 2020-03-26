VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mainForm 
   BackColor       =   &H8000000C&
   Caption         =   "Account Software"
   ClientHeight    =   10140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21480
   Icon            =   "mainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   2880
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   21420
      TabIndex        =   4
      Top             =   7935
      Width           =   21480
      Begin VB.Label lblReceiveAmt 
         BackStyle       =   0  'Transparent
         Caption         =   "2000"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   19560
         TabIndex        =   8
         Top             =   75
         Width           =   3855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payment Received:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   14880
         TabIndex        =   7
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label lblTotalAmt 
         BackStyle       =   0  'Transparent
         Caption         =   "2000"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4185
         TabIndex        =   6
         Top             =   90
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Project Value:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   9240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainForm.frx":14D73
            Key             =   "new_project"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainForm.frx":15048
            Key             =   "logout"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainForm.frx":15501
            Key             =   "edit_project"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainForm.frx":15A04
            Key             =   "delete_project"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainForm.frx":15EA5
            Key             =   "open_project"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainForm.frx":1638F
            Key             =   "finish_project"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainForm.frx":16816
            Key             =   "manage_user"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7155
      ScaleWidth      =   21420
      TabIndex        =   1
      Top             =   720
      Width           =   21480
      Begin VB.ListBox lstID 
         Height          =   450
         Left            =   2400
         TabIndex        =   3
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid flexMain 
         Height          =   5415
         Left            =   -120
         TabIndex        =   2
         Top             =   0
         Width           =   21015
         _ExtentX        =   37068
         _ExtentY        =   9551
         _Version        =   393216
         Cols            =   13
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
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   21480
      _ExtentX        =   37888
      _ExtentY        =   1270
      ButtonWidth     =   2540
      ButtonHeight    =   1217
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New"
            Key             =   "tbAddNew"
            ImageKey        =   "new_project"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Key             =   "tbEdit"
            ImageKey        =   "edit_project"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "tbDelete"
            ImageKey        =   "delete_project"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open Project"
            Key             =   "tbOpenProject"
            ImageKey        =   "open_project"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Finish Project"
            Key             =   "tbFinishProject"
            ImageKey        =   "finish_project"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Expenses"
            Key             =   "tbAddExpenses"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pending Projects"
            Key             =   "tbPendingProject"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "All Expenses"
            Key             =   "tbAllExpenses"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "All Projects"
            Key             =   "tbAllProjects"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Manage Users"
            Key             =   "tbManageUsers"
            ImageKey        =   "manage_user"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log Off"
            Key             =   "tbLogoff"
            ImageKey        =   "logout"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopupmenu 
      Caption         =   "popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuInvoice 
         Caption         =   "Invoice"
         Begin VB.Menu mnuViewInvoice 
            Caption         =   "View Invoice"
         End
         Begin VB.Menu mnuAddInvoice 
            Caption         =   "Add Invoice"
         End
         Begin VB.Menu mnuRemoveInvoice 
            Caption         =   "Remove Invoice"
         End
      End
      Begin VB.Menu mnuPurchase 
         Caption         =   "Purchase"
         Begin VB.Menu mnuViewPurchase 
            Caption         =   "View Purchase"
         End
         Begin VB.Menu mnuAddPurchase 
            Caption         =   "Add Purchase"
         End
         Begin VB.Menu mnuRemovePurchase 
            Caption         =   "Remove Purchase"
         End
      End
      Begin VB.Menu mnuBusinessCard 
         Caption         =   "Business Card"
         Begin VB.Menu mnuViewBusinessCard 
            Caption         =   "View Business card"
         End
         Begin VB.Menu mnuAddBusinessCard 
            Caption         =   "Add Business card"
         End
         Begin VB.Menu mnuRemoveBusinessCard 
            Caption         =   "Remove Business card"
         End
      End
      Begin VB.Menu mnuSeparate3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddExpenses 
         Caption         =   "Add Expenses"
      End
      Begin VB.Menu mnuViewExpenses 
         Caption         =   "View Expenses"
      End
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim RS As Recordset



Private Sub flexMain_Click()
    If flexMain.Row < 1 Then Exit Sub
    If flexMain.TextMatrix(flexMain.Row, 12) = 1 Then
        tbMain.Buttons("tbEdit").Enabled = False
        tbMain.Buttons("tbDelete").Enabled = False
        tbMain.Buttons("tbOpenProject").Enabled = True
        tbMain.Buttons("tbFinishProject").Enabled = False
    Else
        tbMain.Buttons("tbEdit").Enabled = True
        tbMain.Buttons("tbDelete").Enabled = True
        tbMain.Buttons("tbOpenProject").Enabled = False
        tbMain.Buttons("tbFinishProject").Enabled = True
    End If
    
End Sub




Private Sub flexMain_DblClick()
    If flexMain.Row < 1 Then
        MsgBox "Please select a project to edit.", vbExclamation, App.Title
        Exit Sub
    End If
    If flexMain.TextMatrix(flexMain.Row, 12) = 1 Then Exit Sub          ' if project_staus is finish, can't edit
    
    frmNewProject.lblId = lstID.List(flexMain.Row - 1)
    frmNewProject.txtJobNo = flexMain.TextMatrix(flexMain.Row, 1)
    frmNewProject.txtCompany = flexMain.TextMatrix(flexMain.Row, 2)
    frmNewProject.txtWorkType = flexMain.TextMatrix(flexMain.Row, 3)
    frmNewProject.dtpDate.Value = flexMain.TextMatrix(flexMain.Row, 4)
    frmNewProject.txtPurchaseNo = Replace(Replace(flexMain.TextMatrix(flexMain.Row, 5), "(with image)", ""), "(with pdf)", "")
    frmNewProject.txtTotalAmt = flexMain.TextMatrix(flexMain.Row, 6)
    frmNewProject.txtInvoiceNo = Replace(Replace(flexMain.TextMatrix(flexMain.Row, 8), "(with image)", ""), "(with pdf)", "")
    frmNewProject.txtReceiveAmt = flexMain.TextMatrix(flexMain.Row, 7)
    frmNewProject.txtWorkStatus = flexMain.TextMatrix(flexMain.Row, 9)
    frmNewProject.txtRemark = flexMain.TextMatrix(flexMain.Row, 10)
    
    frmNewProject.Caption = "Edit Project"
    frmNewProject.Show vbModal
End Sub




Private Sub flexMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> vbRightButton Then Exit Sub
    If flexMain.Row < 1 Then Exit Sub
    
    flexMain.Row = flexMain.MouseRow
    flexMain.Col = flexMain.MouseCol
    
    
    If flexMain.TextMatrix(flexMain.Row, 12) = 1 Then
        mnuAddInvoice.Enabled = False
        mnuAddPurchase.Enabled = False
        mnuAddBusinessCard.Enabled = False
        mnuRemoveInvoice.Enabled = False
        mnuRemovePurchase.Enabled = False
        mnuRemoveBusinessCard.Enabled = False
        mnuAddExpenses.Enabled = False
    Else
        mnuAddInvoice.Enabled = True
        mnuAddPurchase.Enabled = True
        mnuAddBusinessCard.Enabled = True
        mnuRemoveInvoice.Enabled = True
        mnuRemovePurchase.Enabled = True
        mnuRemoveBusinessCard.Enabled = True
        mnuAddExpenses.Enabled = True
    End If
    
    If tbMain.Buttons("tbAddExpenses").Visible = False Then
        mnuAddExpenses.Visible = False
        mnuViewExpenses.Visible = False
        mnuSeparate3.Visible = False
    Else
        mnuAddExpenses.Visible = True
        mnuViewExpenses.Visible = True
        mnuSeparate3.Visible = True
    End If
    PopupMenu mnuPopupmenu, vbPopupMenuRightButton
    
End Sub



Private Sub MDIForm_Load()

    Me.Move 0, 0, Screen.Width, Screen.Height
    
    Picture1.Height = Screen.Height - tbMain.Height - Picture2.Height - 1000
    flexMain.Left = 0
    flexMain.Width = Me.Width
    flexMain.Height = Picture1.Height
    
    Dim aColumnWidth As Variant
    Dim aColumnText As Variant
    Dim i As Integer
    
    aColumnWidth = Array(1000, 1500, 1800, 1800, 1500, 2500, 1200, 1200, 2000, 1200, 2500, 1500, 0)
    aColumnText = Array("No", "Job No", "Company", "Type of Work", "Date", "PurchaseNo", "Total", "Payment Received", "Inovice No", "Status", "Remark", "User Id", "Project Status")
    
    With flexMain
        For i = 0 To flexMain.Cols - 1
            .ColWidth(i) = aColumnWidth(i)
            .TextMatrix(0, i) = aColumnText(i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        .ColAlignment(0) = flexAlignCenterCenter
        .RowHeight(0) = 350
    End With
    
    Call LoadData
End Sub





Private Sub MDIForm_Terminate()
    End
End Sub

Private Sub mnuAddInvoice_Click()
    Call AddImage("invoice_image")
    
End Sub

Private Sub mnuRemoveInvoice_Click()
    Call RemoveImage("invoice_image")
    
End Sub

Private Sub mnuViewInvoice_Click()
    Call ViewImage("invoice_image")
    
End Sub





Private Sub mnuAddPurchase_Click()
    Call AddImage("purchase_image")
    
End Sub

Private Sub mnuRemovePurchase_Click()
    Call RemoveImage("purchase_image")
    
End Sub

Private Sub mnuViewPurchase_Click()
    Call ViewImage("purchase_image")
    
End Sub




Private Sub mnuAddBusinessCard_Click()
    Call AddImage("business_card_image")
    
End Sub

Private Sub mnuRemoveBusinessCard_Click()
    Call RemoveImage("business_card_image")
    
End Sub

Private Sub mnuViewBusinessCard_Click()
    Call ViewImage("business_card_image")
    
End Sub



Private Sub mnuAddExpenses_click()
    If flexMain.Row < 1 Then
        MsgBox "Please select a project to add expense.", vbExclamation, App.Title
        Exit Sub
    End If
    frmAddExpense.lblProjectId = lstID.List(flexMain.Row - 1)
    frmAddExpense.txtJobNo = flexMain.TextMatrix(flexMain.Row, 1)
    frmAddExpense.txtCompany = flexMain.TextMatrix(flexMain.Row, 2)
    frmAddExpense.Show vbModal
End Sub



Private Sub mnuViewExpenses_click()
    If flexMain.Row < 1 Then
        MsgBox "Please select a project to view expense.", vbExclamation, App.Title
        Exit Sub
    End If
    
    With frmExpense
        .lblProjectId = lstID.List(flexMain.Row - 1)
        .lblJobNo = flexMain.TextMatrix(flexMain.Row, 1)
        .lblCompany = flexMain.TextMatrix(flexMain.Row, 2)
        .lblTotalAmt = flexMain.TextMatrix(flexMain.Row, 6)
        .lblTotalPaid = flexMain.TextMatrix(flexMain.Row, 8)
        .lblTotalExpense = "-"
        .lblProfit = "-"
    
        If flexMain.TextMatrix(flexMain.Row, 12) = 1 Then
            .tbExpense.Buttons("tbExpenseNew").Enabled = False
            .tbExpense.Buttons("tbExpenseEdit").Enabled = False
            .tbExpense.Buttons("tbExpenseDelete").Enabled = False
        Else
            .tbExpense.Buttons("tbExpenseNew").Enabled = True
            .tbExpense.Buttons("tbExpenseEdit").Enabled = True
            .tbExpense.Buttons("tbExpenseDelete").Enabled = True
        End If
    End With
    
    Call frmExpense.LoadData
    frmExpense.Show vbModal
    
End Sub




Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Integer
    
    If Button.Key = "tbAddNew" Then
        frmNewProject.Caption = "New Project"
        frmNewProject.Show vbModal
        
    ElseIf Button.Key = "tbEdit" Then
        Call flexMain_DblClick
    
    ElseIf Button.Key = "tbDelete" Then
        If flexMain.Row < 1 Then
            MsgBox "Please select a project to delete.", vbExclamation, App.Title
            Exit Sub
        End If
        
        If MsgBox("Are you really the project? (Job No: " & flexMain.TextMatrix(flexMain.Row, 1) & ")", vbYesNo + vbInformation, App.Title) = vbNo Then
            Exit Sub
        End If
        
        Call db_connection

        sSQL = "DELETE FROM tblExpense WHERE projectid=" & lstID.List(flexMain.Row - 1)
        gDB.Execute sSQL
        
        sSQL = "DELETE FROM tblProject WHERE id=" & lstID.List(flexMain.Row - 1)
        gDB.Execute sSQL

        Call db_close
        Call LoadData
    
    ElseIf Button.Key = "tbOpenProject" Then
        If flexMain.Row < 1 Then
            MsgBox "Please select a project to open.", vbExclamation, App.Title
            Exit Sub
        End If
        
        Call db_connection

        sSQL = "UPDATE tblProject SET project_status=0 WHERE id=" & lstID.List(flexMain.Row - 1)
        gDB.Execute sSQL

        Call db_close
        Call LoadData
        
    ElseIf Button.Key = "tbFinishProject" Then
        If flexMain.Row < 1 Then
            MsgBox "Please select a project to finish.", vbExclamation, App.Title
            Exit Sub
        End If
        
        Call db_connection

        sSQL = "UPDATE tblProject SET project_status=1 WHERE id=" & lstID.List(flexMain.Row - 1)
        gDB.Execute sSQL

        Call db_close
        Call LoadData
        
    ElseIf Button.Key = "tbManageUsers" Then
        frmManageUser.Show vbModal
        
    ElseIf Button.Key = "tbAddExpenses" Then
        Call mnuAddExpenses_click
        
    ElseIf Button.Key = "tbAllExpenses" Then
        frmAllExpenses.Show vbModal
    ElseIf Button.Key = "tbLogoff" Then
        Unload Me
        frmLogin.Show vbModal
    End If
    
End Sub



Private Sub AddImage(field_name As String)
    Dim image_path As String, image_extend As String
    Dim pic_stream As ADODB.Stream
    On Error GoTo ErrProc
    
    With cd
        .FileName = ""
        .Filter = "Image (*.jpg; *.bmp; *.gif, *.pdf) | *.jpg; *.bmp; *.gif; *.pdf"
        .ShowOpen
        
        If Len(.FileName) = 0 Then Exit Sub
        image_path = .FileName
        image_extend = Mid(.FileName, InStrRev(.FileName, ".") + 1, Len(.FileName))
    End With
    
    Set pic_stream = New ADODB.Stream
    pic_stream.Type = adTypeBinary
    
    pic_stream.Open
    pic_stream.LoadFromFile image_path
    
    Call db_connection
    
    sSQL = "SELECT * FROM tblProject WHERE id=" & lstID.List(flexMain.Row - 1)
    Set RS = gDB.OpenRecordset(sSQL)
    
    RS.Edit
    RS.Fields(field_name) = pic_stream.Read
    RS.Fields(field_name & "_ext") = image_extend
    RS.Update
    RS.Close
    
    Call db_close
    
    pic_stream.Close
    Set pic_stream = Nothing
    
    Call LoadData
    MsgBox "Added image successfully.", vbInformation, App.Title
    Exit Sub
ErrProc:
    If Err.Number = 3002 Then
        MsgBox "The file name isn't supported. Please check file name.", vbExclamation, App.Title
    End If
End Sub



Private Sub RemoveImage(field_name As String)
    Call db_connection
    
    sSQL = "SELECT * FROM tblProject WHERE id=" & lstID.List(flexMain.Row - 1)
    Set RS = gDB.OpenRecordset(sSQL)
    
    RS.Edit
    RS.Fields(field_name) = ""
    RS.Update
    RS.Close
    
    Call db_close
    
    Call LoadData
    MsgBox "Removed image successfully.", vbInformation, App.Title
    
End Sub


Private Sub ViewImage(field_name As String)
    
    Dim pic_stream As ADODB.Stream, temp_picture As StdPicture
    Dim image_extend As String
    On Error GoTo ErrProc
    
    Call db_connection
    
    sSQL = "SELECT " & field_name & " AS tmpimage, " & field_name & "_ext AS tmpext FROM tblProject WHERE id=" & lstID.List(flexMain.Row - 1)
    Set RS = gDB.OpenRecordset(sSQL)
    
    If Not RS.EOF Then
        Set pic_stream = New ADODB.Stream
        pic_stream.Type = adTypeBinary
        pic_stream.Open
        
        image_extend = RS.Fields("tmpext").Value
        If LCase(image_extend) = "pdf" Then
            If IsNull(RS.Fields("tmpimage").Value) = False Then
                pic_stream.Write RS.Fields("tmpimage").Value
                pic_stream.SaveToFile App.Path & "\resource.pdf", adSaveCreateOverWrite
                
                frmPdf.WebBrowser1.Navigate App.Path & "\resource.pdf"
                frmPdf.Show
                
            End If
        Else
            If IsNull(RS.Fields("tmpimage").Value) = False Then
                pic_stream.Write RS.Fields("tmpimage").Value
                pic_stream.SaveToFile App.Path & "\resource.jpg", adSaveCreateOverWrite
                
                Set temp_picture = LoadPicture(App.Path & "\resource.jpg")
                
                frmImage.Width = frmImage.ScaleX(temp_picture.Width, vbHimetric, vbTwips)
                frmImage.Height = frmImage.ScaleY(temp_picture.Height, vbHimetric, vbTwips)
                frmImage.Move Screen.Width / 2 - frmImage.Width / 2, Screen.Height / 2 - frmImage.Height / 2
                frmImage.imgView.Picture = LoadPicture(App.Path & "\resource.jpg")
                frmImage.Show
            End If
        End If
        
        pic_stream.Close
        Set pic_stream = Nothing
        
        
    End If
    
    RS.Close
    Call db_close
ErrProc:
    If Err.Number = 3004 Then
        MsgBox "The file already opened. After close the file, please try again.", vbExclamation, App.Title
    End If
End Sub


Public Sub LoadData()

    Dim i As Integer, open_project_nums As Integer
    Dim sum_total_amt As Double, sum_receive_amt As Double
    Dim purchase_type As String, invoice_type As String
    
    Call db_connection
    
    flexMain.Rows = 1
    lstID.Clear
    
    If gPermission = "admin" Then
        sSQL = "SELECT * FROM tblProject"
    Else
        sSQL = "SELECT * FROM tblProject WHERE userid='" & gUserId & "'"
    End If
    
    Set RS = gDB.OpenRecordset(sSQL)
    
    Do While Not RS.EOF
        If IsNull(RS.Fields("purchase_image").Value) = True Then
            purchase_type = ""
        Else
            purchase_type = IIf(RS.Fields("purchase_image_ext").Value = "pdf", "(with pdf)", "(with image)")
        End If
        If IsNull(RS.Fields("invoice_image").Value) = True Then
            invoice_type = ""
        Else
            invoice_type = IIf(RS.Fields("invoice_image_ext").Value = "pdf", "(with pdf)", "(with image)")
        End If
        
        flexMain.AddItem flexMain.Rows & vbTab & RS.Fields("jobno").Value & vbTab & RS.Fields("company_name") & vbTab & RS.Fields("work_type").Value & vbTab & RS.Fields("work_date").Value & vbTab & _
            RS.Fields("purchase_no").Value & purchase_type & vbTab & RS.Fields("total_amt").Value & vbTab & RS.Fields("receive_amt").Value & vbTab & _
            RS.Fields("invoice_no").Value & invoice_type & vbTab & _
            RS.Fields("work_status").Value & vbTab & RS.Fields("remark").Value & vbTab & RS.Fields("userid").Value & vbTab & RS.Fields("project_status").Value
        flexMain.RowHeight(flexMain.Rows - 1) = 350
        
        flexMain.Row = flexMain.Rows - 1
        
        If RS.Fields("project_status").Value = 0 Then
            For i = 1 To flexMain.Cols - 1
                flexMain.Col = i
                flexMain.CellForeColor = vbRed
            Next
        
            open_project_nums = open_project_nums + 1
        End If
        
        flexMain.Row = flexMain.Rows - 1
        
        sum_total_amt = sum_total_amt + RS.Fields("total_amt").Value
        sum_receive_amt = sum_receive_amt + RS.Fields("receive_amt").Value
        
        lstID.AddItem RS.Fields("id").Value
        
        RS.MoveNext
    Loop
    
    RS.Close
    
    tbMain.Buttons("tbPendingProject").Caption = " Pendign Projects: " & CStr(open_project_nums) & " "
    tbMain.Buttons("tbAllProjects").Caption = " All Projects: " & CStr(flexMain.Rows - 1)
    
    lblTotalAmt = CStr(sum_total_amt)
    lblReceiveAmt = CStr(sum_receive_amt)
    
    Call db_close
End Sub

