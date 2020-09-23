Attribute VB_Name = "Module1"
Option Explicit

Global Const LISTVIEW_MODE0 = "View Large Icons"
Global Const LISTVIEW_MODE1 = "View Small Icons"
Global Const LISTVIEW_MODE2 = "View List"
Global Const LISTVIEW_MODE3 = "View Details"
Public fMainForm As frmMain
Public o1 As Outlook.Application
Public n1 As Outlook.NameSpace

Sub Main()
    sbMsg "Trying to Open Outlook..."
    Set o1 = New Outlook.Application
    Set n1 = o1.GetNamespace("MAPI")
    
    frmMain.Show
    FillTree
    'DoLogin
    
End Sub

Public Sub EndAll()
    On Error Resume Next
    'o1.Quit
    Set o1 = Nothing
    End
End Sub
Public Sub DoLogin(Optional Forced As Boolean = False)
    Dim fLogin As New frmLogin
    
    If n1.CurrentUser = vbNullString Or Forced Then
        fLogin.Show vbModal
        sbMsg "Logging on"
        n1.Logon fLogin.txtUserName, fLogin.txtPassword, False, False
        sbMsg "Done"
        Unload fLogin
    Else
        sbMsg "Using current profile"
        
    End If
    
End Sub

Public Sub sbMsg(msg As String)
    frmMain.sbStatusBar.Panels(1).Text = msg
End Sub

Public Sub FillTree()
    sbMsg "Retrieving Outlook folders..."
    frmMain.tvTreeView.Nodes.Clear
    'NodeCount = 0
    
    frmMain.tvTreeView.Nodes.Add , tvwFirst, "Root", "Outlook", 9
    frmMain.tvTreeView.Nodes("Root").Expanded = True
    GetKids n1, "Root"
    sbMsg "Done"
End Sub

Private Sub GetKids(fs As Object, ByVal ThisNode As String)
    Dim f1 As MAPIFolder
    Dim n As Node
    Dim i1 As Integer
    Dim i2 As Integer
    For Each f1 In fs.Folders
        Select Case f1.DefaultItemType
            Case olMailItem:        i1 = 10: i2 = 11
            Case olAppointmentItem: i1 = 12: i2 = 12
            Case olContactItem:     i1 = 17: i2 = 17
            Case olJournalItem:     i1 = 13: i2 = 14
            Case olNoteItem:        i1 = 16: i2 = 16
            Case olPostItem:        i1 = 16: i2 = 16
            Case olTaskItem:        i1 = 15: i2 = 15
        End Select
        frmMain.tvTreeView.Nodes.Add ThisNode, tvwChild, f1.EntryId, f1.Name, i1, i2
        GetKids f1, f1.EntryId
    Next
End Sub

Public Sub GetItems(EntryId As String)
    Dim f1 As MAPIFolder
    Dim m1 As Variant
    Dim d1 As MailItem
    Dim l1 As ListItem
    
    Set f1 = n1.GetFolderFromID(EntryId)
    frmMain.lvListView.ListItems.Clear
    
    For Each m1 In f1.Items
        Set l1 = frmMain.lvListView.ListItems.Add(, m1.EntryId, m1.Subject, 7)
        If m1.MessageClass = "IPM.Note" Then
            l1.SubItems(1) = m1.SenderName
            l1.SubItems(2) = m1.ReceivedTime
        End If
    Next
End Sub
