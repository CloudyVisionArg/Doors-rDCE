VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistory 
   Caption         =   "Form1"
   ClientHeight    =   4704
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6492
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4704
   ScaleWidth      =   6492
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      _ExtentX        =   7641
      _ExtentY        =   5525
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
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CodeType As Long '1-Folder, 2-Form, 3-Document, 4-AsyncEvent
Public dForm As Object
Public Folder As Object
Public DocId As Long
Public EventKey As String
Public Field As String
Public blnLoaded As Boolean

Private Sub Form_Activate()
    Dim oRcs As Object
    Dim strSQL As String
    Dim li As MSComctlLib.ListItem
    Dim si As MSComctlLib.ListSubItem
    
    On Error GoTo Error
    
    If Not blnLoaded Then
        strSQL = "select * from DCE_HISTORY where CODETYPE = " & CodeType
        If CodeType = 1 Then
            strSQL = strSQL & " and FLD_ID = " & Folder.id & " and SEV_ID = " & Mid(EventKey, 4)
        ElseIf CodeType = 2 Then
            strSQL = strSQL & " and FRM_ID = " & dForm.id & " and SEV_ID = " & Mid(EventKey, 4)
        ElseIf CodeType = 3 Then
            strSQL = strSQL & " and FLD_ID = " & Folder.id & " and DOC_ID = " & DocId
        End If
        strSQL = strSQL & " order by TIMESTAMP desc"
        
        'todo: soporte para asyncevents
        
        Set oRcs = GSession.Db.OpenRecordset(strSQL)
        Do While Not oRcs.EOF
            With ListView1
                Set li = .ListItems.Add(, , oRcs("TIMESTAMP").Value)
                li.Tag = oRcs("TIMESTAMP").Value
                li.ToolTipText = li.Text
                ' subitem 1
                Set si = li.ListSubItems.Add(, , oRcs("ACC_NAME").Value)
                si.ToolTipText = si.Text
                si.Tag = oRcs("ACC_ID").Value
                ' subitem 2
                Set si = li.ListSubItems.Add(, , Len(oRcs("CODE").Value & ""))
                si.ToolTipText = si.Text
                si.Tag = oRcs("CODE").Value & ""
            End With
            oRcs.MoveNext
        Loop
        
        oRcs.Close
        blnLoaded = True
    End If

    Screen.MousePointer = vbNormal
    ListView1.SetFocus
    If ListView1.ListItems.Count = 0 Then MsgBox "No hay historial", vbInformation
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Or _
           KeyCode = vbKeyF4 And Shift = vbCtrlMask Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim ch As MSComctlLib.ColumnHeader
    
    blnLoaded = False
    KeyPreview = True
    
    With ListView1
        .View = lvwReport
        .LabelEdit = lvwManual
        .LabelWrap = True
        .HideSelection = False
        .FullRowSelect = True
        .BorderStyle = ccFixedSingle
        .Appearance = cc3D
        
        Set ch = .ColumnHeaders.Add(, , "Date")
        ch.Width = 2500
        Set ch = .ColumnHeaders.Add(, , "User")
        ch.Width = 2500
        Set ch = .ColumnHeaders.Add(, , "Size")
        ch.Width = 1000
    End With
End Sub

Private Sub Form_Resize()
    With ListView1
        .Top = 1
        .Left = 1
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick ListView1, ColumnHeader
End Sub

Private Sub ListView1_DblClick()
    Dim li As MSComctlLib.ListItem
    Dim frmCode As frmEditor
    
    Set li = ListView1.SelectedItem
    If Not li Is Nothing Then
        Set frmCode = New frmEditor
        With frmCode
            .Caption = Caption & " (solo lectura)"
            .CodeMax1.Text = li.ListSubItems(2).Tag
            .CodeMax1.ReadOnly = True
            .Show
        End With
    End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListView1_DblClick
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If Not ListView1.SelectedItem Is Nothing Then
            Me.PopupMenu MDIForm1.mnuPopupHistory, , , , MDIForm1.mnuPopupHistoryView
        End If
    End If
End Sub

Public Sub mnuPopupHistoryDiffNextClick()
    Dim li As MSComctlLib.ListItem
    Dim newer As String
    Dim older As String
    
    Set li = ListView1.SelectedItem
    If li.Index = 1 Then
        mnuPopupHistoryDiffLastClick
    Else
        older = li.ListSubItems(2).Tag
        newer = ListView1.ListItems(li.Index - 1).ListSubItems(2).Tag
        ShowDiff older, newer
    End If
End Sub

Public Sub mnuPopupHistoryDiffLastClick()
    Dim li As MSComctlLib.ListItem
    Dim dom As Object
    Dim newer As String
    Dim older As String
    
    If CodeType = 1 Then
        newer = Folder.Events(EventKey).Code
    ElseIf CodeType = 2 Then
        newer = dForm.Events(EventKey).Code
    ElseIf CodeType = 3 Then
        Set dom = Folder.Search(Field, "doc_id = " & DocId)
        newer = dom.documentElement.firstChild.getAttribute(LCase(Field))
    End If
    
    Set li = ListView1.SelectedItem
    If Not li Is Nothing Then
        older = li.ListSubItems(2).Tag
    End If
    
    ShowDiff older, newer
End Sub

Public Sub ShowDiff(ByRef older As String, ByRef newer As String)
    Dim fso As Scripting.FileSystemObject
    Dim fldTemp As Scripting.Folder
    Dim tsNewer As Scripting.TextStream
    Dim tsOlder As Scripting.TextStream
    Dim DiffTool As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fldTemp = fso.GetSpecialFolder(2)
    
    Set tsNewer = fldTemp.CreateTextFile("newer.vbs", True)
    Set tsOlder = fldTemp.CreateTextFile("older.vbs", True)
    
    tsNewer.Write newer
    tsOlder.Write older
    
    tsNewer.Close
    tsOlder.Close
    
    DiffTool = ReadIni("Misc", "DiffTool")
    If DiffTool <> "" Then
        DiffTool = Replace(DiffTool, "{newer}", fldTemp.Path & "\newer.vbs")
        DiffTool = Replace(DiffTool, "{older}", fldTemp.Path & "\older.vbs")
        Shell DiffTool, vbNormalFocus
    Else
        MsgBox "Configure DiffTool en el INI de la sig forma:" & vbCrLf & _
            "[Misc]" & vbCrLf & _
            "DiffTool=C:\path\diff.exe {older} {newer}"
        Exit Sub
    End If
End Sub

Public Sub mnuPopupHistoryViewClick()
    ListView1_DblClick
End Sub

