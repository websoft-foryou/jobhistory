Attribute VB_Name = "mainModule"
Option Explicit
Dim ws As Workspace
Public gDB As DAO.Database
Public gPermission As String
Public gUserId As String

Public Sub Main()
    
    'mainForm.Show
    frmLogin.Show vbModal
End Sub


Public Sub db_connection()
    Set ws = DBEngine.Workspaces(0)
    Set gDB = ws.OpenDatabase(App.Path & "\account.mdb", False, False, ";pwd=Asdf1234!!")
End Sub

Public Sub db_close()
    gDB.Close
    ws.Close
End Sub

