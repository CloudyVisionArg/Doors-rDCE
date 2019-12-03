VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExplorer 
   Caption         =   "Remote DCE"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8340
   Icon            =   "frmExplorer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   8340
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   325
      Left            =   0
      TabIndex        =   9
      Top             =   5577
      Width           =   8333
      _ExtentX        =   14711
      _ExtentY        =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin rDCE.SplitPanel SplitPanel1 
      Height          =   5172
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   7812
      _ExtentX        =   13785
      _ExtentY        =   9128
      Begin rDCE.SplitPanel SplitPanel2 
         Height          =   4452
         Left            =   2640
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   4812
         _ExtentX        =   8493
         _ExtentY        =   7858
         Begin VB.CheckBox chkAcl 
            Caption         =   "Acl"
            Height          =   192
            Left            =   2640
            TabIndex        =   10
            Top             =   2160
            Visible         =   0   'False
            Width           =   732
         End
         Begin MSComctlLib.ListView lstViews 
            Height          =   852
            Left            =   2520
            TabIndex        =   5
            Top             =   2640
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   1508
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lstDocuments 
            Height          =   852
            Left            =   600
            TabIndex        =   4
            Top             =   2640
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   1508
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.TabStrip TabStrip2 
            Height          =   1812
            Left            =   240
            TabIndex        =   6
            Top             =   2400
            Width           =   4092
            _ExtentX        =   7223
            _ExtentY        =   3175
            MultiRow        =   -1  'True
            Placement       =   1
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Documents"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Views"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lstAsyncEvents 
            Height          =   1212
            Left            =   1680
            TabIndex        =   2
            Top             =   360
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   2143
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lstSyncEvents 
            Height          =   1212
            Left            =   360
            TabIndex        =   1
            Top             =   360
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   2143
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.TabStrip TabStrip1 
            Height          =   2172
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   4212
            _ExtentX        =   7435
            _ExtentY        =   3836
            MultiRow        =   -1  'True
            Placement       =   1
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Sync Events"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Async Events"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2055
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnKeyboard As Boolean
Public blnLoaded As Boolean
Public LastFocus As Object

Private Sub chkAcl_Click()
    Dim n As Object
    Dim fldId As Long
    Dim sXPath As String
    
    If chkAcl.Tag Then
        fldId = Mid(TreeView1.SelectedItem.Key, 5)
        If fldId > 1000 Then
            sXPath = GSelected.FolderXPath(fldId)
            Set n = GSelected.dom.selectSingleNode(sXPath)
            If chkAcl.Value = 1 Then
                If n Is Nothing Then
                    Set n = GSelected.AddFolder(fldId, False)
                End If
                n.setAttribute "acl", "1"
            ElseIf chkAcl.Value = 0 Then
                If Not n Is Nothing Then
                    n.setAttribute "acl", "0"
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
'    If Not GSession.IsLogged Then
'        MsgBox "Sesion no iniciada", vbExclamation
'        Unload Me
'        'frmLogon.Show
'        Exit Sub
'    End If

    On Error GoTo Error
    
    If Not blnLoaded Then
        Screen.MousePointer = vbHourglass
    
        Caption = "EXPLORE"
    
        LoadTree
        TabStrip1_Click
        TabStrip2_Click
        
        'Chequear la tabla DCE_HISTORY
        Dim strSQL As String
        Dim oRcs As Object
        Dim lngErr As Long
        
        strSQL = "select * from DCE_HISTORY"
        On Error Resume Next
        Set oRcs = GSession.Db.OpenRecordset(strSQL, Array(2, Empty, Empty, Empty, 1)) ' CommandTimeout = 2, MaxRecords = 1
        lngErr = Err.Number
        oRcs.Close
        If lngErr <> 0 Then
            If GSession.Db.DbType = 6 Then ' SqlServer
                strSQL = "create table dbo.DCE_HISTORY (TIMESTAMP datetime, ACC_ID int, ACC_NAME varchar(50), " & _
                    "CODETYPE int, FRM_ID int, FLD_ID int, SEV_ID int, DOC_ID int, CODE text)"
                GSession.Db.Execute strSQL
            
            ElseIf GSession.Db.DbType = 5 Then ' Oracle
                strSQL = "create table DCE_HISTORY (TIMESTAMP date, ACC_ID number(10), ACC_NAME varchar2(50), " & _
                    "CODETYPE number(10), FRM_ID number(10), FLD_ID number(10), SEV_ID number(10), DOC_ID number(10), CODE clob)"
                GSession.Db.Execute strSQL
            End If
        End If
        
        Screen.MousePointer = vbNormal
        blnLoaded = True
    End If
    
    If Not LastFocus Is Nothing Then
        If LastFocus.Enabled Then LastFocus.SetFocus
    Else
        TreeView1.SetFocus
    End If

    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim frmSearch1 As frmSearch
    
    If KeyCode = vbKeyF3 Or KeyCode = vbKeyF And Shift = vbCtrlMask Then
        Set frmSearch1 = New frmSearch
        frmSearch1.Show
    End If
End Sub

Private Sub Form_Load()
    blnLoaded = False
    Set LastFocus = Nothing
    StatusBar1.Height = 300
    
    With SplitPanel1
        Set .Control1 = TreeView1
        Set .Control2 = SplitPanel2
        .Position = 3000
        .SplitterWidth = 50
    End With
    
    With SplitPanel2
        Set .Control1 = TabStrip1
        Set .Control2 = TabStrip2
        .Horizontal = True
        .SplitterWidth = 50
    End With
    
    TreeView1.Indentation = 350
    
    lstSyncEvents.View = lvwReport
    lstAsyncEvents.View = lvwReport
    lstDocuments.View = lvwReport
    lstViews.View = lvwReport
    
    lstSyncEvents.MultiSelect = True
    lstAsyncEvents.MultiSelect = True
    lstDocuments.MultiSelect = True
    lstViews.MultiSelect = True
    
    MDIForm1.mnuSetupCheckboxes.Checked = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        With SplitPanel1
            .Top = 1
            .Left = 1
            .Width = ScaleWidth
            .Height = ScaleHeight - StatusBar1.Height
        End With
        SplitPanel2.Position = ScaleHeight / 2
    End If
End Sub

Sub LoadTree()
    Dim oNode As Object
    Dim strAux As String
    Dim oTreeNode As MSComctlLib.node
    Dim lngId As Long
    Dim oForm As Object
    Dim blnSystem As Boolean
    Dim oTree As Object
    Dim oChildNode As Object
    Dim oNodes As Object
    Dim oDom As Object
    
    On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    TreeView1.Nodes.Clear
    Set oTreeNode = TreeView1.Nodes.Add(, , "FLD-1", "System Folders")
    oTreeNode.Expanded = True
    Set oDom = GSession.FoldersTree()
    Set oNodes = oDom.selectNodes("//d:folder")
    For Each oNode In oNodes
        lngId = oNode.getAttribute("id")
        blnSystem = Val(oNode.getAttribute("system") & "")
        'If oNode.getAttribute("parent_folder") & "" <> "1" Then
            If Not blnSystem Or lngId = 1001 Then
                If oNode.getAttribute("description") & "" <> "" Then
                    strAux = oNode.getAttribute("description") & " (" & oNode.getAttribute("name") & ")"
                Else
                    strAux = oNode.getAttribute("name")
                End If
                
                If oNode.getAttribute("parent_folder") & "" = "" Then
                    Set oTreeNode = TreeView1.Nodes.Add(, , "FLD-" & lngId, strAux)
                    oTreeNode.Expanded = True
                Else
                    Dim prtFolderId As String
                    prtFolderId = oNode.getAttribute("parent_folder")
                    If NodeExists("FLD-" & prtFolderId) Then
                        If Not NodeExists("FLD-" & lngId) Then
                            Set oTreeNode = TreeView1.Nodes.Add("FLD-" & prtFolderId, tvwChild, "FLD-" & lngId, strAux)
                            'oTreeNode.Expanded = True
                        End If
                    Else
                        
                    End If
                End If
                oTreeNode.Checked = GSelected.Checked(GSelected.FolderXPath(oNode))
            End If
        'End If
    Next
    
   
    oTreeNode.Expanded = True
    Set oTreeNode = TreeView1.Nodes.Add("FLD-1", tvwChild, "FLD-5", "Forms")
    Set oTreeNode = TreeView1.Nodes.Add("FLD-1", tvwChild, "FLD-11", "CodeLib")
    Set oTreeNode = TreeView1.Nodes.Add("FLD-1", tvwChild, "FLD-3", "Directory")
    
    For Each oNode In GSession.FormsList.documentElement.childNodes
        lngId = oNode.getAttribute("id")
        strAux = oNode.getAttribute("name") & " (" & lngId & ")"
        Set oTreeNode = TreeView1.Nodes.Add("FLD-5", tvwChild, _
            "FRM-" & lngId, strAux)
        oTreeNode.Checked = GSelected.Checked(GSelected.FormXPath(oNode))
    Next
    
    TreeView1.Nodes("FLD-1001").Selected = True
    RefreshList
    
    Screen.MousePointer = vbNormal
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Private Function NodeExists(ByVal strKey As String) As Boolean
    Dim node As MSComctlLib.node
    On Error Resume Next
    Set node = TreeView1.Nodes(strKey)
    Select Case Err.Number
        Case 0
            NodeExists = True
        Case Else
            NodeExists = False
    End Select
End Function

Private Sub lstAsyncEvents_GotFocus()
    Set LastFocus = lstAsyncEvents
End Sub

Public Sub lstAsyncEvents_ItemCheck(ByVal item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass

    If item.Checked Then
        GSelected.AddFolderItem "AsyncEvent", Mid(lstAsyncEvents.Tag, 5), Mid(item.Key, 4)
    Else
        GSelected.Remove GSelected.AsyncEventXPath(Mid(lstAsyncEvents.Tag, 5), Mid(item.Key, 4))
    End If

    Screen.MousePointer = vbNormal
End Sub

Private Sub lstDocuments_GotFocus()
    Set LastFocus = lstDocuments
End Sub

Public Sub lstDocuments_ItemCheck(ByVal item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    
    If lstDocuments.Tag = "FLD-3" Then ' Directory
        If item.Checked Then
            GSelected.AddAccount item.ListSubItems(3)
        Else
            GSelected.Remove GSelected.AccountXPath(item.ListSubItems(3))
        End If
    Else
        If item.Checked Then
            GSelected.AddFolderItem "Document", Mid(lstDocuments.Tag, 5), item.Tag
        Else
            GSelected.Remove GSelected.DocumentXPath(Mid(lstDocuments.Tag, 5), item.Tag)
        End If
    End If

    Screen.MousePointer = vbNormal
End Sub

Private Sub lstSyncEvents_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick lstSyncEvents, ColumnHeader
End Sub

Private Sub lstAsyncEvents_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick lstAsyncEvents, ColumnHeader
End Sub

Private Sub lstDocuments_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick lstDocuments, ColumnHeader
End Sub

Private Sub lstSyncEvents_GotFocus()
    Set LastFocus = lstSyncEvents
End Sub

Public Sub lstSyncEvents_ItemCheck(ByVal item As MSComctlLib.ListItem)
    Dim sPath As String
    
    Screen.MousePointer = vbHourglass
    
    If item.Checked Then
        If Left(lstSyncEvents.Tag, 4) = "FLD-" Then
            GSelected.AddFolderItem "SyncEvent", Mid(lstSyncEvents.Tag, 5), Mid(item.Key, 4)
        ElseIf Left(lstSyncEvents.Tag, 4) = "FRM-" Then
            GSelected.AddFormItem "SyncEvent", Mid(lstSyncEvents.Tag, 5), Mid(item.Key, 4)
        End If
    Else
        If Left(lstSyncEvents.Tag, 4) = "FLD-" Then
            sPath = GSelected.FolderEventXPath(Mid(lstSyncEvents.Tag, 5), Mid(item.Key, 4))
        ElseIf Left(lstSyncEvents.Tag, 4) = "FRM-" Then
            sPath = GSelected.FormEventXPath(Mid(lstSyncEvents.Tag, 5), Mid(item.Key, 4))
        End If
        GSelected.Remove sPath
    End If

    Screen.MousePointer = vbNormal
End Sub

Private Sub lstViews_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick lstViews, ColumnHeader
End Sub

Private Sub lstSyncEvents_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstSyncEvents_DblClick
End Sub

Private Sub lstAsyncEvents_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstAsyncEvents_DblClick
End Sub

Private Sub lstDocuments_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstDocuments_DblClick
End Sub

Private Sub lstViews_GotFocus()
    Set LastFocus = lstViews
End Sub

Public Sub lstViews_ItemCheck(ByVal item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    
    If item.Checked Then
        GSelected.AddFolderItem "View", Mid(lstViews.Tag, 5), item.ListSubItems(1)
    Else
        GSelected.Remove GSelected.ViewXPath(Mid(lstViews.Tag, 5), item.ListSubItems(1))
    End If

    Screen.MousePointer = vbNormal
End Sub

Private Sub lstViews_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstViews_DblClick
End Sub

Private Sub lstSyncEvents_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If Not lstSyncEvents.SelectedItem Is Nothing Then
            Me.PopupMenu MDIForm1.mnuPopup, , , , MDIForm1.mnuPopupEdit
        End If
    End If
End Sub

Private Sub lstAsyncEvents_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If Not lstAsyncEvents.SelectedItem Is Nothing Then
            Me.PopupMenu MDIForm1.mnuPopup, , , , MDIForm1.mnuPopupEdit
        End If
    End If
End Sub

Private Sub lstDocuments_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If Not lstDocuments.SelectedItem Is Nothing Then
            Me.PopupMenu MDIForm1.mnuPopup, , , , MDIForm1.mnuPopupEdit
        End If
    End If
End Sub

Private Sub lstViews_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If Not lstViews.SelectedItem Is Nothing Then
            Me.PopupMenu MDIForm1.mnuPopup, , , , MDIForm1.mnuPopupEdit
        End If
    End If
End Sub

Private Sub lstSyncEvents_DblClick()
    Dim TreeKey As String
    Dim fld As Object
    Dim frm As Object
    Dim frmCode As frmEditor
    Dim li As MSComctlLib.ListItem
    Dim dom As Object
    
    On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    
    TreeKey = lstSyncEvents.Tag & ""
    
    Set li = lstSyncEvents.SelectedItem
    If Not li Is Nothing Then
        If Left(TreeKey, 4) = "FLD-" Then ' Folder
            Set fld = FolderCache(Mid(TreeKey, 5))
            
            Set frmCode = New frmEditor
            With frmCode
                .Caption = "EDIT /" & fld.Path & "/" & li.Text
                .CodeMax1.Text = fld.Events(li.Key).Code
                .CodeType = 1
                Set .Folder = fld
                .EventKey = li.Key
                .Show
            End With
            
        ElseIf Left(TreeKey, 4) = "FRM-" Then ' Form
            Set frm = FormCache(Mid(TreeKey, 5))
            Set frmCode = New frmEditor
            With frmCode
                .Caption = "EDIT //Forms/" & frm.Name & "/" & li.Text
                .CodeMax1.Text = frm.Events(li.Key).Code
                .CodeType = 2
                Set .dForm = frm
                .EventKey = li.Key
                .Show
            End With
        End If
    End If

    Screen.MousePointer = vbNormal
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Private Sub lstAsyncEvents_DblClick()
    Dim TreeKey As String
    Dim fld As Object
    Dim frmCode As frmEditor
    Dim li As MSComctlLib.ListItem
    Dim evn As Object
    
    On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    
    TreeKey = lstAsyncEvents.Tag & ""
    
    Set li = lstAsyncEvents.SelectedItem
    If Not li Is Nothing Then
        Set fld = FolderCache(Mid(TreeKey, 5))
        If fld.id <> li.Tag Then
            Screen.MousePointer = vbNormal
            MsgBox "Este evento es heredado", vbExclamation
            Exit Sub
        End If
        Set evn = fld.AsyncEvents(li.Key)
        If evn.IsCom = True Then
            Screen.MousePointer = vbNormal
            MsgBox "Este es un evento COM", vbExclamation
            Exit Sub
        End If
        Set frmCode = New frmEditor
        With frmCode
            .Caption = "EDIT /" & fld.Path & "/AsyncEvent-" & li.Text
            .CodeMax1.Text = evn.Code
            .CodeType = 4
            Set .Folder = fld
            .EventKey = li.Key
            .Show
        End With
    End If

    Screen.MousePointer = vbNormal
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Private Sub lstDocuments_DblClick()
    Dim TreeKey As String
    Dim fld As Object
    Dim frm As Object
    Dim frmCode As frmEditor
    Dim li As MSComctlLib.ListItem
    Dim dom As Object
    Dim id As Long
    Dim sCodeCol As String
    Dim doc As Object
    
    On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    
    TreeKey = lstDocuments.Tag & ""
    If TreeKey = "FLD-3" Then
        Screen.MousePointer = vbNormal
        MsgBox "No tiene codigo", vbExclamation
        Exit Sub
    End If
    
    Set li = lstDocuments.SelectedItem
    If Not li Is Nothing Then
        id = Mid(li.Key, 4)
        Set fld = FolderCache(Mid(TreeKey, 5))
        Set frmCode = New frmEditor
        
        ' CodeLib
        If LCase(fld.Form.Guid) = LCase("F89ECD42FAFF48FDA229E4D5C5F433ED") Then
            Set doc = fld.Documents(id)
            
            With frmCode
                .Caption = "EDIT /" & fld.Path & "/" & li.Text
                .CodeMax1.Text = doc("code").Value & ""
                .CodeType = 3
                Set .Folder = fld
                .DocId = id
                .Field = "code"
                .Show
            End With
        
        ' Controls
        ElseIf LCase(fld.Form.Guid) = LCase("EAC99A4211204E1D8EEFEB8273174AC4") Then
            Set doc = fld.Documents(id)
            With frmCode
                .Caption = "EDIT /" & fld.Path & "/" & li.Text
                .CodeMax1.Text = doc("scriptbeforerender").Value & ""
                .CodeType = 3
                Set .Folder = fld
                .DocId = id
                .Field = "scriptbeforerender"
                .Show
            End With
            
        ' DCE_CodeColumn
        ElseIf fld.Form.Properties.Exists("DCE_CodeColumn") Then
            sCodeCol = fld.Form.Properties("DCE_CodeColumn").Value
            Set doc = fld.Documents(id)
            With frmCode
                .Caption = "EDIT /" & fld.Path & "/" & li.Text
                .CodeMax1.Text = doc(sCodeCol).Value & ""
                .CodeType = 3
                Set .Folder = fld
                .DocId = id
                .Field = sCodeCol
                .Show
            End With
        Else
            MsgBox "No hay campo de codigo definido", vbExclamation
        End If
    End If

    Screen.MousePointer = vbNormal
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Private Sub lstViews_DblClick()
    MsgBox "TODO"
End Sub

Private Sub SplitPanel2_Change(Position As Long)
    With lstSyncEvents
        .Top = 50
        .Left = 50
        .Width = TabStrip1.Width - 100
        .Height = TabStrip1.Height - 425
    End With
    
    With lstAsyncEvents
        .Top = 50
        .Left = 50
        .Width = TabStrip1.Width - 100
        .Height = TabStrip1.Height - 425
    End With

    With chkAcl
        .Top = TabStrip1.Top + TabStrip1.Height - 210
        .Left = TabStrip1.Left + 2700
    End With
    
    With lstDocuments
        .Top = TabStrip2.Top + 50
        .Left = TabStrip2.Left + 50
        .Width = TabStrip2.Width - 100
        .Height = TabStrip2.Height - 425
    End With
    
    With lstViews
        .Top = TabStrip2.Top + 50
        .Left = TabStrip2.Left + 50
        .Width = TabStrip2.Width - 100
        .Height = TabStrip2.Height - 425
    End With
End Sub

Private Sub TabStrip1_Click()
    lstSyncEvents.Visible = (TabStrip1.SelectedItem.Index = 1)
    lstAsyncEvents.Visible = (TabStrip1.SelectedItem.Index = 2)
End Sub

Private Sub TabStrip2_Click()
    lstDocuments.Visible = (TabStrip2.SelectedItem.Index = 1)
    lstViews.Visible = (TabStrip2.SelectedItem.Index = 2)
End Sub

Private Sub TreeView1_GotFocus()
    Set LastFocus = TreeView1
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
    blnKeyboard = (KeyCode <> 13)
    If KeyCode = vbKeyF5 Then LoadTree
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TreeView1_NodeClick TreeView1.SelectedItem
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnKeyboard = False
    
    If Button = vbRightButton Then
        If Not TreeView1.SelectedItem Is Nothing Then
            Me.PopupMenu MDIForm1.mnuPopupTree
        End If
    End If
End Sub

Private Sub TreeView1_NodeCheck(ByVal node As MSComctlLib.node)
    Screen.MousePointer = vbHourglass
    
    If node.Checked Then
        If Left(node.Key, 4) = "FLD-" Then
            GSelected.AddFolder Mid(node.Key, 5)
        ElseIf Left(node.Key, 4) = "FRM-" Then
            GSelected.AddForm Mid(node.Key, 5)
        End If
    Else
        If Left(node.Key, 4) = "FLD-" Then
            GSelected.Remove GSelected.FolderXPath(Mid(node.Key, 5))
        ElseIf Left(node.Key, 4) = "FRM-" Then
            GSelected.Remove GSelected.FormXPath(Mid(node.Key, 5))
        End If
    End If

    Screen.MousePointer = vbNormal
End Sub

Private Sub TreeView1_NodeClick(ByVal node As MSComctlLib.node)
    Dim fld As Object
    Dim dom As Object
    Dim n As Object
    Dim li As ListItem
    Dim ch As ColumnHeader
    Dim frm As Object
    Dim lsi As ListSubItem
    Dim sAux As String
    Dim arr As Variant
    Dim sWidth As String
    Dim i As Long
    Dim oFie As Object
    Dim sCols As String
    Dim sCodeCol As String
    Dim bHasCode As Boolean
    Dim oFList As clsFieldList
    Dim arrWidths
    Dim tOut As Long
    
    If blnKeyboard Then Exit Sub ' No se cambia al navegar con el teclado
    
    On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    tOut = GSession.HttpRequestTimeout
    
    With lstSyncEvents
        .ListItems.Clear
        .ColumnHeaders.Clear
        .Tag = node.Key
        .Sorted = False
    End With
    
    With lstAsyncEvents
        .ListItems.Clear
        .ColumnHeaders.Clear
        .Tag = node.Key
        .Sorted = False
    End With
    
    With lstDocuments
        .ListItems.Clear
        .ColumnHeaders.Clear
        .Tag = node.Key
        .Sorted = False
    End With
    
    With lstViews
        .ListItems.Clear
        .ColumnHeaders.Clear
        .Tag = node.Key
        .Sorted = False
    End With
    
    If Left(node.Key, 4) = "FLD-" Then ' Folder
        Set fld = FolderCache(Mid(node.Key, 5))
        Caption = "EXPLORE /" & fld.Path
        
        If fld.System Then
            chkAcl.Enabled = False
            chkAcl.Tag = False ' Flag para evitar el evento Click
            chkAcl.Value = 0
            chkAcl.Tag = True
        Else
            chkAcl.Enabled = True
            chkAcl.Tag = False ' Flag para evitar el evento Click
            Set n = GSelected.dom.selectSingleNode(GSelected.FolderXPath(fld))
            If Not n Is Nothing Then
                chkAcl.Value = IIf(n.getAttribute("acl") & "" = "1", 1, 0)
            Else
                chkAcl.Value = 0
            End If
            chkAcl.Tag = True
        End If
        
        With lstAsyncEvents
            Set ch = .ColumnHeaders.Add(, , "Id")
            ch.Width = 500
            Set ch = .ColumnHeaders.Add(, , "Type")
            ch.Width = 700
            Set ch = .ColumnHeaders.Add(, , "Login")
            ch.Width = 1300
            Set ch = .ColumnHeaders.Add(, , "Is COM")
            ch.Width = 750
            Set ch = .ColumnHeaders.Add(, , "Class")
            ch.Width = 2000
            Set ch = .ColumnHeaders.Add(, , "Method")
            ch.Width = 1500
            Set ch = .ColumnHeaders.Add(, , "Timeout")
            ch.Width = 800
            Set ch = .ColumnHeaders.Add(, , "Created")
            ch.Width = 1800
            Set ch = .ColumnHeaders.Add(, , "Modified")
            ch.Width = 1800
    
            For Each n In fld.AsyncEventsList.documentElement.childNodes
                Set li = .ListItems.Add(, "ID=" & n.getAttribute("id"), n.getAttribute("id"))
                li.Checked = GSelected.Checked(GSelected.AsyncEventXPath(fld, n.getAttribute("id")))
                
                li.Tag = CDbl(n.getAttribute("fld_id"))
                
                sAux = n.getAttribute("type")
                If sAux = "0" Then
                    sAux = "TMR"
                ElseIf sAux = "1" Then
                    sAux = "TRG"
                End If
                li.ListSubItems.Add , , sAux
                
                li.ListSubItems.Add , , n.getAttribute("login")
                
                sAux = n.getAttribute("is_com")
                If sAux = "0" Then
                    sAux = "N"
                ElseIf sAux = "1" Then
                    sAux = "Y"
                End If
                li.ListSubItems.Add , , sAux
                
                li.ListSubItems.Add , , n.getAttribute("class")
                li.ListSubItems.Add , , n.getAttribute("method")
                li.ListSubItems.Add , , n.getAttribute("code_timeout")
                li.ListSubItems.Add , , n.getAttribute("created")
                li.ListSubItems.Add , , n.getAttribute("modified")
                If fld.id <> li.Tag Then
                    li.ForeColor = vbGrayText
                    For Each lsi In li.ListSubItems
                        lsi.ForeColor = vbGrayText
                    Next
                End If
                If n.getAttribute("code") = "1" Then
                    li.Bold = True
                    li.ListSubItems(1).Bold = True
                End If
            Next
        End With
        
        If fld.FolderType = 1 Then
            With lstSyncEvents
                Set ch = .ColumnHeaders.Add(, , "Event")
                ch.Width = 2500
                Set ch = .ColumnHeaders.Add(, , "Overrides")
                ch.Width = 1000
                Set ch = .ColumnHeaders.Add(, , "Created")
                ch.Width = 1800
                Set ch = .ColumnHeaders.Add(, , "Modified")
                ch.Width = 1800
        
                For Each n In fld.EventsList.documentElement.childNodes
                    Set li = .ListItems.Add(, "ID=" & n.getAttribute("id"), n.getAttribute("name"))
                    li.Checked = GSelected.Checked(GSelected.FolderEventXPath(fld, n.getAttribute("id")))

                    If n.getAttribute("code") = "1" Then li.Bold = True
                    li.ListSubItems.Add , , n.getAttribute("overrides")
                    li.ListSubItems.Add , , n.getAttribute("created")
                    li.ListSubItems.Add , , n.getAttribute("modified")
                Next
            End With
        
            ' CodeLib
            If LCase(fld.Form.Guid) = LCase("F89ECD42FAFF48FDA229E4D5C5F433ED") Then
                Set dom = fld.Search("doc_id,name,code,created,modified", , "name", GlngMaxDocs)
                CheckMaxDocs dom
                
                With lstDocuments
                    Set ch = .ColumnHeaders.Add(, , "Code")
                    ch.Width = 2500
                    Set ch = .ColumnHeaders.Add(, , "Created")
                    ch.Width = 1800
                    Set ch = .ColumnHeaders.Add(, , "Modified")
                    ch.Width = 1800
                    
                    For Each n In dom.documentElement.childNodes
                        Set li = .ListItems.Add(, "ID=" & n.getAttribute("doc_id"), n.getAttribute("name"))
                        Set li.Tag = n
                        li.Checked = GSelected.Checked(GSelected.DocumentXPath(fld, n))
                        
                        If n.getAttribute("code") <> "" Then li.Bold = True
                        li.ListSubItems.Add , , n.getAttribute("created")
                        li.ListSubItems.Add , , n.getAttribute("modified")
                    Next
                End With
            
            ' Controls
            ElseIf LCase(fld.Form.Guid) = LCase("EAC99A4211204E1D8EEFEB8273174AC4") Then
                Set dom = fld.Search("doc_id,name,control,scriptbeforerender,created,modified", , "name", GlngMaxDocs)
                CheckMaxDocs dom
                
                With lstDocuments
                    Set ch = .ColumnHeaders.Add(, , "Name")
                    ch.Width = 2500
                    Set ch = .ColumnHeaders.Add(, , "Control")
                    ch.Width = 2500
                    Set ch = .ColumnHeaders.Add(, , "Created")
                    ch.Width = 1800
                    Set ch = .ColumnHeaders.Add(, , "Modified")
                    ch.Width = 1800
                    
                    For Each n In dom.documentElement.childNodes
                        Set li = .ListItems.Add(, "ID=" & n.getAttribute("doc_id"), n.getAttribute("name"))
                        Set li.Tag = n
                        li.Checked = GSelected.Checked(GSelected.DocumentXPath(fld, n))

                        li.ListSubItems.Add , , n.getAttribute("control")
                        If n.getAttribute("scriptbeforerender") <> "" Then li.Bold = True
                        li.ListSubItems.Add , , n.getAttribute("created")
                        li.ListSubItems.Add , , n.getAttribute("modified")
                    Next
                End With
            
            Else
                bHasCode = False
                If fld.Form.Properties.Exists("DCE_HasCode") Then
                    bHasCode = fld.Form.Properties("DCE_HasCode").Value = "1"
                End If
                
                If bHasCode Then
                    ' DCE_HasCode
                    
                    Set oFList = New clsFieldList
                    sCols = fld.Form.Properties("DCE_ListColumns").Value
                    oFList.Add sCols
                    If fld.Form.Properties.Exists("DCE_ListColumnsWidth") Then
                        arrWidths = fld.Form.Properties("DCE_ListColumnsWidth").Value
                    Else
                        arrWidths = Split("", ",")
                    End If
                    arr = oFList.Keys
                    For i = 0 To UBound(arr)
                        Set oFie = fld.Form.Fields(arr(i))
                        Set ch = lstDocuments.ColumnHeaders.Add(, , IIf(oFie.Description <> "", oFie.Description, LCase(oFie.Name)))
                        If i <= UBound(arrWidths) Then
                            ch.Width = CLng(arrWidths(i))
                        Else
                            ch.Width = 2000
                        End If
                    Next
                    
                    sCodeCol = fld.Form.Properties("DCE_CodeColumn").Value
                    oFList.Add sCodeCol
                    
                    oFList.Add "created"
                    Set ch = lstDocuments.ColumnHeaders.Add(, , "Created")
                    ch.Width = 1800
                    oFList.Add "modified"
                    Set ch = lstDocuments.ColumnHeaders.Add(, , "Modified")
                    ch.Width = 1800
                    
                    oFList.Add FormPK(fld.Form)
                    oFList.Add "doc_id"
                    
                    Set dom = fld.Search(oFList.ToString, , sCols, GlngMaxDocs)
                    CheckMaxDocs dom

                    For Each n In dom.documentElement.childNodes
                        Set li = lstDocuments.ListItems.Add(, "ID=" & n.getAttribute("doc_id"), n.getAttribute(LCase(arr(0))))
                        Set li.Tag = n
                        li.Checked = GSelected.Checked(GSelected.DocumentXPath(fld, n))
                        
                        For i = 1 To UBound(arr)
                            li.ListSubItems.Add , , n.getAttribute(arr(i))
                        Next
                        li.ListSubItems.Add , , n.getAttribute("created")
                        li.ListSubItems.Add , , n.getAttribute("modified")
                        If n.getAttribute(LCase(sCodeCol)) <> "" Then li.Bold = True
                    Next
                    
                Else
                
                    ' Documentos comunes
                    'todo: leer DCE_ListColumns y armar la lista con eso
                    
                    Set oFList = New clsFieldList
                    oFList.Add "doc_id,subject,created,modified,accessed"
                    oFList.Add FormPK(fld.Form)
                    
                    GSession.HttpRequestTimeout = 60 ' 60 segs
                    Set dom = fld.Search(oFList.ToString, , "modified desc", GlngMaxDocs)
                    CheckMaxDocs dom

                    GSession.HttpRequestTimeout = tOut
                    
                    With lstDocuments
                        Set ch = .ColumnHeaders.Add(, , "DOC_ID")
                        ch.Width = 1000
                        Set ch = .ColumnHeaders.Add(, , "Subject")
                        ch.Width = 4000
                        Set ch = .ColumnHeaders.Add(, , "Created")
                        ch.Width = 1800
                        Set ch = .ColumnHeaders.Add(, , "Modified")
                        ch.Width = 1800
                        Set ch = .ColumnHeaders.Add(, , "Accessed")
                        ch.Width = 1800
                        
                        For Each n In dom.documentElement.childNodes
                            Set li = .ListItems.Add(, "ID=" & n.getAttribute("doc_id"), n.getAttribute("doc_id"))
                            Set li.Tag = n
                            li.Checked = GSelected.Checked(GSelected.DocumentXPath(fld, n))

                            li.ListSubItems.Add , , n.getAttribute("subject")
                            li.ListSubItems.Add , , n.getAttribute("created")
                            li.ListSubItems.Add , , n.getAttribute("modified")
                            li.ListSubItems.Add , , n.getAttribute("accessed")
                        Next
                    End With
                
                End If
                
            End If
        
           ' Vistas
            Set dom = fld.ViewsList
            With lstViews
                Set ch = .ColumnHeaders.Add(, , "ID")
                ch.Width = 800
                Set ch = .ColumnHeaders.Add(, , "Name")
                ch.Width = 3500
                Set ch = .ColumnHeaders.Add(, , "Description")
                ch.Width = 3500
                Set ch = .ColumnHeaders.Add(, , "Type")
                ch.Width = 800
                Set ch = .ColumnHeaders.Add(, , "Created")
                ch.Width = 1800
                Set ch = .ColumnHeaders.Add(, , "Modified")
                ch.Width = 1800
                
                For Each n In dom.documentElement.selectNodes("/d:root/d:item[@private='0']")
                    Set li = .ListItems.Add(, "ID=" & n.getAttribute("id"), n.getAttribute("id"))
                    li.Checked = GSelected.Checked(GSelected.ViewXPath(fld, n.getAttribute("name")))
                    
                    li.ListSubItems.Add , , n.getAttribute("name")
                    li.ListSubItems.Add , , n.getAttribute("description")
                    If (CLng(Left(GSession.Version, 1)) < 7) Then
                        li.ListSubItems.Add , , "1"
                    Else
                        li.ListSubItems.Add , , n.getAttribute("viewType")
                    End If
                    li.ListSubItems.Add , , n.getAttribute("created")
                    li.ListSubItems.Add , , n.getAttribute("modified")
                Next
            End With
            
        
        ElseIf node.Key = "FLD-3" Then ' Directory
            chkAcl.Enabled = False
            chkAcl.Tag = False ' Flag para evitar el evento Click
            chkAcl.Value = 0
            chkAcl.Tag = True
            
            Set dom = GSession.Directory.AccountsList
            
            With lstDocuments
                Set ch = .ColumnHeaders.Add(, , "ACC_ID")
                ch.Width = 1000
                Set ch = .ColumnHeaders.Add(, , "Type")
                ch.Width = 600
                Set ch = .ColumnHeaders.Add(, , "System")
                ch.Width = 750
                Set ch = .ColumnHeaders.Add(, , "Name")
                ch.Width = 3500
                Set ch = .ColumnHeaders.Add(, , "Email")
                ch.Width = 2500
                Set ch = .ColumnHeaders.Add(, , "Login")
                ch.Width = 1800
                Set ch = .ColumnHeaders.Add(, , "Disabled")
                ch.Width = 900
                
                For Each n In dom.documentElement.childNodes
                    Set li = .ListItems.Add(, "ID=" & n.getAttribute("id"), n.getAttribute("id"))
                    li.Checked = GSelected.Checked(GSelected.AccountXPath(n.getAttribute("name")))

                    li.ListSubItems.Add , , n.getAttribute("type")
                    li.ListSubItems.Add , , n.getAttribute("system")
                    li.ListSubItems.Add , , n.getAttribute("name")
                    li.ListSubItems.Add , , n.getAttribute("email")
                    li.ListSubItems.Add , , n.getAttribute("login") & ""
                    li.ListSubItems.Add , , n.getAttribute("disabled") & ""
                Next
            End With
            
        End If
        
    ElseIf Left(node.Key, 4) = "FRM-" Then ' Form
        chkAcl.Enabled = False
        chkAcl.Tag = False ' Flag para evitar el evento Click
        chkAcl.Value = 0
        chkAcl.Tag = True
        
        Set frm = FormCache(Mid(node.Key, 5))
        Caption = "EXPLORE //Forms/" & frm.Name
            
        With lstSyncEvents
            Set ch = .ColumnHeaders.Add(, , "Event")
            ch.Width = 2500
            Set ch = .ColumnHeaders.Add(, , "Extensible")
            ch.Width = 1000
            Set ch = .ColumnHeaders.Add(, , "Overridable")
            ch.Width = 1000
            Set ch = .ColumnHeaders.Add(, , "Created")
            ch.Width = 1800
            Set ch = .ColumnHeaders.Add(, , "Modified")
            ch.Width = 1800
    
            For Each n In frm.EventsList.documentElement.childNodes
                Set li = .ListItems.Add(, "ID=" & n.getAttribute("id"), n.getAttribute("name"))
                li.Checked = GSelected.Checked(GSelected.FormEventXPath(frm, n.getAttribute("id")))
                
                If n.getAttribute("code") = "1" Then li.Bold = True
                li.ListSubItems.Add , , n.getAttribute("extensible")
                li.ListSubItems.Add , , n.getAttribute("overridable")
                li.ListSubItems.Add , , n.getAttribute("created")
                li.ListSubItems.Add , , n.getAttribute("modified")
            Next
        End With
        
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
Error:
    GSession.HttpRequestTimeout = tOut
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Public Sub RefreshList()
    TreeView1_NodeClick TreeView1.SelectedItem
End Sub

Public Sub mnuPopupEditClick()
    If ActiveControl.Name = "lstSyncEvents" Then
        lstSyncEvents_DblClick
    ElseIf ActiveControl.Name = "lstDocuments" Then
        lstDocuments_DblClick
    ElseIf ActiveControl.Name = "lstAsyncEvents" Then
        lstAsyncEvents_DblClick
    End If
End Sub

Public Sub mnuPopupSelectAllClick()
    SelectAll True
End Sub

Public Sub mnuPopupTreeSelectAllClick()
    SelectItemsRecursive Empty, Empty
End Sub

Public Sub mnuPopupTreeSelectByModifClick()
    Dim datFrom As Date
    Dim datTo As Date
    
    frmGetDates.Show vbModal
        
    If frmGetDates.Cancel Then
        Exit Sub
    Else
        With frmGetDates
            If IsDate(.txtFrom.Text) Then
                datFrom = CDate(.txtFrom.Text)
            Else
                datFrom = Empty
            End If
        
            If IsDate(.txtTo.Text) Then
                datTo = CDate(.txtTo.Text)
            Else
                datTo = Empty
            End If
        End With
        
    End If
    
    SelectItemsRecursive datFrom, datTo
End Sub

Public Sub mnuPopupTreeSelectForms()
    MsgBox "TO DO"
End Sub

Public Sub SelectItemsRecursive(pModifFrom As Date, pModifTo As Date)
    Dim node As MSComctlLib.node
    Dim obj As Object
    Dim n As Object
    
    Set node = TreeView1.SelectedItem
    If node.Key = "FLD-3" Then ' Directory
        MsgBox "No soportado en el Directory", vbInformation
        Exit Sub
    End If
        
    On Error GoTo Error
    Screen.MousePointer = vbHourglass
        
    If Left(node.Key, 4) = "FLD-" Then ' Folder
        Set obj = FolderCache(Mid(node.Key, 5))
        If BetweenDates(obj.Modified, pModifFrom, pModifTo) Then
            Set n = GSelected.AddFolder(obj)
            If Not obj.System Then n.setAttribute ("acl"), "1"
        End If
        SelectItemsRecursive2 obj, pModifFrom, pModifTo
    
    ElseIf Left(node.Key, 4) = "FRM-" Then ' Form
        Set obj = FormCache(Mid(node.Key, 5))
        If BetweenDates(obj.Modified, pModifFrom, pModifTo) Then
            GSelected.AddForm obj
        End If
        SelectItemsRecursive2 obj, pModifFrom, pModifTo
    
    End If
    
    mnuPopupTreeRefreshClick
    Screen.MousePointer = vbNormal

    StatusBar1.SimpleText = "Listo"
    DoEvents
    
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Public Sub SelectItemsRecursive2(pObject As Object, pModifFrom As Date, pModifTo As Date)
    Dim node As Object, n
    Dim sFormula As String
    Dim frm As Object, fld As Object
    Dim sFields As String
    Dim datAux As Date
    
    If TypeName(pObject) = "Folder" Then
        
        If pObject.id = 1 Then ' System Folders
            SelectItemsRecursive2 FolderCache(5), pModifFrom, pModifTo ' Forms
            SelectItemsRecursive2 FolderCache(11), pModifFrom, pModifTo ' Codelib
            
        ElseIf pObject.id = 5 Then ' Forms
            For Each node In GSession.FormsList.documentElement.childNodes
                If BetweenDates(node.getAttribute("modified"), pModifFrom, pModifTo) Then
                    Set frm = FormCache(CLng(node.getAttribute("id")))
                    GSelected.AddForm frm
                    SelectItemsRecursive2 frm, pModifFrom, pModifTo
                End If
                
            Next
            
        Else ' Otro folder
        
            StatusBar1.SimpleText = "Seleccionando carpeta " & pObject.Name & "..."
            DoEvents
            
            If pObject.FolderType = 1 Then ' Documents
                ' SyncEvents
                For Each node In pObject.EventsList.documentElement.childNodes
                    If IsDate(node.getAttribute("modified")) Then
                        datAux = node.getAttribute("modified")
                    Else
                        datAux = pObject.Created
                    End If
                    If BetweenDates(datAux, pModifFrom, pModifTo) Then
                        GSelected.AddFolderItem "SyncEvent", pObject, node.getAttribute("id")
                    End If
                Next
                
                ' Documents
                sFormula = ""
                If pModifFrom <> Empty Then
                    sFormula = "modified >= " & GSession.Db.SqlEncode(pModifFrom, 2)
                End If
                If pModifTo <> Empty Then
                    If sFormula <> "" Then sFormula = sFormula & " and "
                    sFormula = sFormula & "modified <= " & GSession.Db.SqlEncode(pModifTo, 2)
                End If
                
                sFields = FormPK(pObject.Form)
                If sFields <> "" Then sFields = "," & sFields
                sFields = "doc_id" & sFields
                
                For Each node In pObject.Search(sFields, sFormula).documentElement.childNodes
                    GSelected.AddFolderItem "Document", pObject, node
                Next
                
                ' Views
                For Each node In pObject.ViewsList.documentElement.childNodes
                    If BetweenDates(node.getAttribute("modified"), pModifFrom, pModifTo) Then
                        GSelected.AddFolderItem "View", pObject, node.getAttribute("name")
                    End If
                Next
            End If
            
            ' AsyncEvents
            For Each node In pObject.AsyncEventsList.documentElement.childNodes
                If BetweenDates(node.getAttribute("modified"), pModifFrom, pModifTo) Then
                    GSelected.AddFolderItem "AsyncEvent", pObject, node.getAttribute("id")
                End If
            Next
            
            'Subfolders y recursivo
            For Each node In pObject.FoldersList.documentElement.childNodes
                Set fld = FolderCache(CLng(node.getAttribute("id")))
                If BetweenDates(node.getAttribute("modified"), pModifFrom, pModifTo) Then
                    Set n = GSelected.AddFolder(fld)
                    If Not fld.System Then n.setAttribute "acl", "1"
                End If
                SelectItemsRecursive2 fld, pModifFrom, pModifTo
            Next
            
        End If
    
    ElseIf TypeName(pObject) = "CustomForm" Then
        ' SyncEvents
        StatusBar1.SimpleText = "Seleccionando form " & pObject.Name & "..."
        DoEvents
        
        For Each node In pObject.EventsList.documentElement.childNodes
            If BetweenDates(node.getAttribute("modified"), pModifFrom, pModifTo) Then
                GSelected.AddFormItem "SyncEvent", pObject, node.getAttribute("id")
            End If
        Next
        
    End If
End Sub

Private Sub CheckMaxDocs(ByRef pDom)
    If pDom.documentElement.childNodes.length >= GlngMaxDocs Then
        StatusBar1.SimpleText = "Se muestran solo " & GlngMaxDocs & " documentos"
    Else
        StatusBar1.SimpleText = pDom.documentElement.childNodes.length & " documentos"
    End If
End Sub

Private Function BetweenDates(pDate As Date, pFrom As Date, pTo As Date) As Boolean
    BetweenDates = True
    If pFrom <> Empty Then
        If pDate < pFrom Then
            BetweenDates = False
        End If
    End If
    If pTo <> Empty Then
        If pDate > pTo Then
            BetweenDates = False
        End If
    End If
End Function

Public Sub mnuPopupUnselectAllClick()
    SelectAll False
End Sub

Public Sub mnuPopupTreeUnselectAllClick()
    Dim treeNode As MSComctlLib.node
    Dim sPath As String
    Dim node, childNode
    
    Set treeNode = TreeView1.SelectedItem
    sPath = GSelected.FolderXPath(Mid(treeNode.Key, 5))
    Set node = GSelected.dom.selectSingleNode(sPath)
    
    If sPath = "/root/system" Then
        For Each childNode In node.childNodes
            RemoveChildNodes childNode
        Next
    ElseIf sPath = "/root/system/codelib" Or _
           sPath = "/root/system/directory" Or _
           sPath = "/root/system/forms" Then
        RemoveChildNodes node
    Else
        node.parentNode.RemoveChild node
    End If
    
    mnuPopupTreeRefreshClick
End Sub

Private Sub RemoveChildNodes(pNode)
    Dim childNode
    
    For Each childNode In pNode.childNodes
        pNode.RemoveChild childNode
    Next
End Sub

Private Sub SelectAll(pSelected As Boolean)
    If ActiveControl.Name = "lstSyncEvents" Then
        SelectAll2 lstSyncEvents, pSelected
    ElseIf ActiveControl.Name = "lstAsyncEvents" Then
        SelectAll2 lstAsyncEvents, pSelected
    ElseIf ActiveControl.Name = "lstDocuments" Then
        SelectAll2 lstDocuments, pSelected
    ElseIf ActiveControl.Name = "lstViews" Then
        SelectAll2 lstViews, pSelected
    End If

End Sub

Private Sub SelectAll2(pListView As MSComctlLib.ListView, pSelected As Boolean)
    Dim li As MSComctlLib.ListItem
    Dim treeNode As MSComctlLib.node
    Dim sPath As String
    Dim node, childNode
    
    If pListView.Name = "lstDocuments" And Not pSelected Then
        ' Deseleccionar todos los documentos lo proceso de otra forma
        ' sino quedan docs seleccionados por el limite de registros que se muestran
        Set treeNode = TreeView1.SelectedItem
        sPath = GSelected.FolderXPath(Mid(treeNode.Key, 5))
        Set node = GSelected.dom.selectSingleNode(sPath)
        
        If sPath = "/root/system/directory" Then
            RemoveChildNodes node
        Else
            For Each childNode In node.childNodes
                If childNode.nodeName = "documents" Then node.RemoveChild childNode
            Next
        End If
        
        For Each li In pListView.ListItems
            li.Checked = pSelected
        Next
        
    Else
        For Each li In pListView.ListItems
            If li.Checked <> pSelected Then
                li.Checked = pSelected
                CallByName Me, pListView.Name & "_ItemCheck", VbMethod, li
            End If
        Next
    End If
End Sub

Public Sub mnuPopupHistClick()
    Dim TreeKey As String
    Dim li As MSComctlLib.ListItem
    Dim fld As Object
    Dim frm As Object
    Dim hist As frmHistory
    Dim id As Long
    
    Screen.MousePointer = vbHourglass
    
    If Me.ActiveControl.Name = "lstSyncEvents" Then
        TreeKey = lstSyncEvents.Tag & ""
        Set li = lstSyncEvents.SelectedItem
        If Not li Is Nothing Then
            If Left(TreeKey, 4) = "FLD-" Then ' Folder
                Set fld = FolderCache(Mid(TreeKey, 5))
                Set hist = New frmHistory
                With hist
                    .Caption = "HISTORY of /" & fld.Path & "/" & li.Text
                    .CodeType = 1
                    Set .Folder = fld
                    .EventKey = li.Key
                    .Show
                End With
            ElseIf Left(TreeKey, 4) = "FRM-" Then ' Form
                Set frm = FormCache(Mid(TreeKey, 5))
                Set hist = New frmHistory
                With hist
                    .Caption = "HISTORY of //Forms/" & frm.Name & "/" & li.Text
                    .CodeType = 2
                    Set .dForm = frm
                    .EventKey = li.Key
                    .Show
                End With
            End If
        End If
    
    ElseIf Me.ActiveControl.Name = "lstDocuments" Then
        TreeKey = lstDocuments.Tag & ""
        Set li = lstDocuments.SelectedItem
        If Not li Is Nothing Then
            id = Mid(li.Key, 4)
            Set fld = FolderCache(Mid(TreeKey, 5))
            Set hist = New frmHistory
            With hist
                .Caption = "HISTORY of /" & fld.Path & "/" & li.Text
                .CodeType = 3
                Set .Folder = fld
                .DocId = id
                .Field = FormCode(fld.Form)
                .Show
            End With
        End If
    
    ElseIf Me.ActiveControl.Name = "lstAsyncEvents" Then
        TreeKey = lstAsyncEvents.Tag & ""
        Set li = lstAsyncEvents.SelectedItem
        If Not li Is Nothing Then
            Set fld = FolderCache(Mid(TreeKey, 5))
            Set hist = New frmHistory
            With hist
                .Caption = "HISTORY of /" & fld.Path & "/AsyncEvent-" & li.Text
                .CodeType = 4
                Set .Folder = fld
                .EventKey = li.Key
                .Show
            End With
        End If
    
    End If

    Screen.MousePointer = vbHourglass
End Sub

Public Sub mnuPopupTreeRefreshClick()
    LoadTree
End Sub

