VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileExplorer 
   Caption         =   "Form1"
   ClientHeight    =   2904
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4920
   Icon            =   "frmFileExplorer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2904
   ScaleWidth      =   4920
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Ir"
      Default         =   -1  'True
      Height          =   336
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   492
   End
   Begin VB.TextBox txtFolder 
      Height          =   288
      Left            =   360
      TabIndex        =   1
      Text            =   "txtFolder"
      Top             =   240
      Width           =   2772
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   732
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1932
      _ExtentX        =   3408
      _ExtentY        =   1291
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmFileExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastFocus As Object
Public RefreshOnFocus As Boolean

Private Sub cmdGo_Click()
    Dim dom As Object
    Dim node As Object
    Dim li As ListItem
    
    If txtFolder.Text = "" Then
        MsgBox "Especifique una carpeta", vbExclamation
        txtFolder.SetFocus
        Exit Sub
    End If
    
    On Error GoTo Error
    Screen.MousePointer = vbHourglass
    
    With ListView1
        .ListItems.Clear
        .Sorted = False
    End With
    
    DoEvents
    
    Set dom = FolderDir(txtFolder.Text)
    
    Set node = dom.documentElement
    txtFolder.Text = node.getAttribute("folder")
    If node.getAttribute("isroot") = "0" Then
        Set li = ListView1.ListItems.Add(, , "..")
        li.Tag = "folder"
    End If
    
    For Each node In dom.documentElement.childNodes
        If node.getAttribute("type") = "folder" Then
            Set li = ListView1.ListItems.Add(, , UCase(node.Text))
        Else
            Set li = ListView1.ListItems.Add(, , node.Text)
        End If
        li.Tag = node.getAttribute("type")
        li.ListSubItems.Add , , node.getAttribute("modified") & ""
        li.ListSubItems.Add , , node.getAttribute("size") & ""
    Next
    
    If ListView1.ListItems.Count > 0 Then ListView1.SetFocus
    Screen.MousePointer = vbNormal
    Exit Sub
    
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Private Sub cmdGo_GotFocus()
    Set LastFocus = cmdGo
End Sub

Private Sub Form_Activate()
    If txtFolder.Text = "txtFolder" Then
        txtFolder.Text = DefaultFolder
    End If

    If Not LastFocus Is Nothing Then
        If LastFocus.Enabled Then LastFocus.SetFocus
    Else
        txtFolder.SetFocus
    End If
    
    If RefreshOnFocus Then
        cmdGo_Click
        RefreshOnFocus = False
    End If
End Sub

Private Sub Form_Load()
    Dim ch As ColumnHeader
    
    Caption = "Explorador de archivos"
    RefreshOnFocus = True
    Set LastFocus = Nothing

    With ListView1
        .View = lvwReport
        .LabelEdit = lvwManual
        .FullRowSelect = True
        
        Set ch = .ColumnHeaders.Add(, , "Name")
        ch.Width = 4000
        Set ch = .ColumnHeaders.Add(, , "Modified")
        ch.Width = 2200
        Set ch = .ColumnHeaders.Add(, , "Size")
        ch.Width = 1300
    End With
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        With txtFolder
            .Top = 250
            .Left = 250
            If ScaleWidth > cmdGo.Width + 750 Then .Width = ScaleWidth - (cmdGo.Width + 750)
        End With
        
        With cmdGo
            .Top = 225
            If ScaleWidth > .Width Then .Left = ScaleWidth - (.Width + 250)
        End With
        
        With ListView1
            .Left = 1
            .Top = 800
            .Width = ScaleWidth
            If ScaleHeight > 800 Then .Height = ScaleHeight - 800
        End With
    End If
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo Error
    SaveFile txtFolder.Text & "\" & NewString, ""
    Exit Sub
Error:
    ErrDisplay Err
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick ListView1, ColumnHeader
End Sub

Private Sub ListView1_DblClick()
    Dim li As ListItem
    Dim frmCode As frmEditor
    Dim oStream As Object
    Dim errN As Long
    
    On Error GoTo Error
    Screen.MousePointer = vbHourglass
    
    Set li = ListView1.SelectedItem
    
    If Not li Is Nothing Then
        If li.Tag = "folder" Then
            If li.Text = ".." Then
                txtFolder.Text = Mid(txtFolder.Text, 1, InStrRev(txtFolder.Text, "\"))
            Else
                txtFolder.Text = txtFolder.Text & "\" & LCase(li.Text)
            End If
            cmdGo_Click
            
        ElseIf li.Tag = "file" Then
            Set oStream = GetFile(txtFolder.Text & "\" & li.Text)
            
            If oStream.Type = 1 Then ' adTypeBinary
                With CommonDialog1
                    .FileName = li.Text
                    .CancelError = True
                    On Error Resume Next
                    .ShowSave
                    errN = Err.Number
                    On Error GoTo 0
                    If errN = 0 Then
                        oStream.SaveToFile .FileName, 2 ' adSaveCreateOverWrite
                        oStream.Close
                    End If
                End With
                ListView1.SetFocus
            
            ElseIf oStream.Type = 2 Then ' adTypeText
            
                Set frmCode = New frmEditor
                With frmCode
                    .FilePath = txtFolder.Text & "\" & li.Text
                    .Caption = "EDIT " & .FilePath
                    .CodeMax1.Text = oStream.ReadText
                    .Charset = oStream.Charset
                    .CodeType = 5
                    .Show
                End With
            
                oStream.Close
            End If
        End If
    End If

    Screen.MousePointer = vbNormal
    Exit Sub
    
Error:
    Screen.MousePointer = vbNormal
    ErrDisplay Err
    ListView1.SetFocus
End Sub

Private Sub ListView1_GotFocus()
    cmdGo.Default = False
    Set LastFocus = ListView1
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        cmdGo_Click
    ElseIf KeyCode = vbKeyDelete Then
        mnuPopupFileExpDeleteClick
    End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListView1_DblClick
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If Not ListView1.SelectedItem Is Nothing Then
            Me.PopupMenu MDIForm1.mnuPopupFileExp
        End If
    End If
End Sub

Private Sub txtFolder_GotFocus()
    cmdGo.Default = True
    Set LastFocus = txtFolder
End Sub

Private Function DefaultFolder() As String
    Dim sCode As String
    
    sCode = "Return Server.MapPath(Application(""VirtualRoot""))"
    DefaultFolder = GSession.HttpCallCode(sCode).responseXml.documentElement.Text
End Function

Public Function FolderDir(pFolder As String) As Object
    Dim sCode As String
    
    sCode = "Set rdceDom = dSession.Xml.NewDom" & vbCrLf & _
        "rdceDom.loadXml ""<root />""" & vbCrLf & _
        "Set rdceFso = Server.CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
        "Set rdceFld = rdceFso.GetFolder(Arg(1))" & vbCrLf & _
        "rdceDom.documentElement.setAttribute ""folder"", rdceFld.Path" & vbCrLf & _
        "rdceDom.documentElement.setAttribute ""isroot"", IIf(rdceFld.IsRootFolder, 1, 0)" & vbCrLf & _
        "For Each rdceF In rdceFld.SubFolders" & vbCrLf & _
        "   Set rdceNode = rdceDom.createNode(""element"", ""item"", """")" & vbCrLf & _
        "   rdceNode.Text = rdceF.Name" & vbCrLf & _
        "   rdceNode.setAttribute ""type"", ""folder""" & vbCrLf & _
        "   rdceDom.documentElement.AppendChild rdceNode" & vbCrLf & _
        "Next" & vbCrLf
    sCode = sCode & _
        "For Each rdceF In rdceFld.Files" & vbCrLf & _
        "   Set rdceNode = rdceDom.createNode(""element"", ""item"", """")" & vbCrLf & _
        "   rdceNode.Text = rdceF.Name" & vbCrLf & _
        "   rdceNode.setAttribute ""type"", ""file""" & vbCrLf & _
        "   rdceNode.setAttribute ""modified"", dSession.Xml.XmlEncode(rdceF.DateLastModified, 2)" & vbCrLf & _
        "   rdceNode.setAttribute ""size"", dSession.Xml.XmlEncode(rdceF.Size, 3)" & vbCrLf & _
        "   rdceDom.documentElement.AppendChild rdceNode" & vbCrLf & _
        "Next" & vbCrLf & _
        "rdceDom.save Response"
    
    Set FolderDir = GSession.HttpCallCode(sCode, Array(pFolder)).responseXml
End Function

Public Function GetFile(pFilePath As String) As Object
    Dim sCode As String
    Dim node As Object
    Dim oStream As Object
    Dim arr
    Dim blnBinary As Boolean
    Dim i As Long
    
    sCode = "Set rdceStream = Server.CreateObject(""ADODB.Stream"")" & vbCrLf & _
        "rdceStream.Open" & vbCrLf & _
        "rdceStream.Type = 1 ' adTypeBinary" & vbCrLf & _
        "rdceStream.LoadFromFile Arg(1)" & vbCrLf & _
        "Set rdceDom = dSession.Xml.NewDom" & vbCrLf & _
        "rdceDom.loadXml ""<root />""" & vbCrLf & _
        "Set rdceNode = rdceDom.createNode(""element"", ""item"", """")" & vbCrLf & _
        "' Utilizamos bin encoded en Base64" & vbCrLf & _
        "rdceNode.dataType = ""bin.base64""" & vbCrLf & _
        "If Not rdceStream.EOS Then rdceNode.nodeTypedValue = rdceStream.Read" & vbCrLf & _
        "rdceDom.documentElement.AppendChild rdceNode" & vbCrLf & _
        "rdceStream.Close" & vbCrLf & _
        "rdceDom.save Response"
    
    Set node = GSession.HttpCallCode(sCode, Array(pFilePath)).responseXml.documentElement.firstChild
    
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 1 ' adTypeBinary
    oStream.Open
    If Not IsNull(node.nodeTypedValue) Then oStream.Write node.nodeTypedValue
    oStream.Position = 0
    
    arr = oStream.Read(100)
    oStream.Position = 0
    oStream.Type = 2 ' adTypeText
    oStream.Charset = "iso-8859-1"
    
    If oStream.Size >= 3 Then
        If arr(0) = 239 And arr(1) = 187 And arr(2) = 191 Then
            ' UTF-8
            oStream.Charset = "utf-8"
        End If
    End If
    
    If oStream.Size > 0 Then
        blnBinary = False
        For i = 0 To UBound(arr)
            If arr(i) = 0 Then
                blnBinary = True: Exit For
            End If
        Next
        If blnBinary Then oStream.Type = 1 ' adTypeBinary
    End If
    
    Set GetFile = oStream
End Function

Public Sub DeleteFile(pFilePath As String)
    Dim sCode As String
    
    sCode = "rdceFile = Arg(1)" & vbCrLf & _
        "Set rdceFso = CreateObject(""doorsbiz.FileSystem"")" & vbCrLf & _
        "rdceFso.DeleteFile rdceFile, True"
    
    GSession.HttpCallCode sCode, Array(pFilePath)
End Sub

Public Sub SaveFile(pFilePath As String, pText As String, Optional pCharset As String = "iso-8859-1")
    Dim sCode As String
    
    sCode = "rdceFile = Arg(1)" & vbCrLf & _
        "Set rdceFso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
        "Set rdceFs = CreateObject(""doorsbiz.FileSystem"")" & vbCrLf & _
        "Set rdceArrConv = CreateObject(""doorsbiz.ArrayConvert"")" & vbCrLf & _
        "If rdceFso.FileExists(rdceFile) Then" & vbCrLf & _
        "   Set rdceStream = rdceFs.FileToStream(rdceFile, 1) ' adTypeBinary" & vbCrLf & _
        "   rdceFs.StreamToFile (rdceStream), rdceFile & "".bak"", 2 ' adSaveCreateOverWrite" & vbCrLf & _
        "   rdceStream.Close" & vbCrLf & _
        "End If" & vbCrLf & _
        "Set rdceStream = rdceFs.NewStream" & vbCrLf & _
        "rdceStream.Type = 2 ' adTypeText" & vbCrLf & _
        "rdceStream.Charset = Arg(3)" & vbCrLf & _
        "rdceStream.Open" & vbCrLf & _
        "Set rdceNode = oReq.documentElement.childNodes(2)" & vbCrLf & _
        "rdceStream.WriteText Replace(Arg(2), Chr(10), vbCrLf)" & vbCrLf & _
        "rdceFs.StreamToFile (rdceStream), rdceFile, 2 ' adSaveCreateOverWrite" & vbCrLf & _
        "rdceStream.Close"
    
    GSession.HttpCallCode sCode, Array(pFilePath, pText, pCharset)
End Sub

Public Sub mnuPopupFileExpNewClick()
    Dim li As ListItem
      
    Set li = ListView1.ListItems.Add(, , "new file.txt")
    li.Tag = "file"
    ListView1.SelectedItem = li
    ListView1.StartLabelEdit
End Sub

Public Sub mnuPopupFileExpDeleteClick()
    On Error GoTo Error
    
    If Not ListView1.SelectedItem Is Nothing Then
        If ListView1.SelectedItem.Tag = "file" Then
            If MsgBox("Borrar " & txtFolder.Text & "\" & ListView1.SelectedItem.Text & " ?", vbOKCancel + vbExclamation) = vbOK Then
                DeleteFile txtFolder.Text & "\" & ListView1.SelectedItem.Text
                cmdGo_Click
            End If
        Else
            MsgBox "No esta soportado en carpetas", vbExclamation
        End If
    End If

    ListView1.SetFocus
    Exit Sub
Error:
    ErrDisplay Err
    ListView1.SetFocus
End Sub

