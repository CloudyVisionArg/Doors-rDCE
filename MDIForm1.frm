VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3510
   ClientLeft      =   208
   ClientTop       =   832
   ClientWidth     =   7046
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   2160
      _ExtentX        =   854
      _ExtentY        =   854
      _Version        =   393216
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Instaladores"
      Begin VB.Menu mnuSetupCheckboxes 
         Caption         =   "&Modo seleccion"
      End
      Begin VB.Menu mnuSetupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetupLoad 
         Caption         =   "&Cargar seleccion"
      End
      Begin VB.Menu mnuSetupSave 
         Caption         =   "&Guardar seleccion"
      End
      Begin VB.Menu mnuSetupClear 
         Caption         =   "&Borrar seleccion"
      End
      Begin VB.Menu mnuSetupSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetupGenerate 
         Caption         =   "Generar &scripts"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowConnection 
         Caption         =   "&Conexion"
      End
      Begin VB.Menu mnuWindowExplorer 
         Caption         =   "&Explorador de codigo"
      End
      Begin VB.Menu mnuWindowFileExplorer 
         Caption         =   "&Explorador de archivos"
      End
      Begin VB.Menu mnuWindowSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowTile 
         Caption         =   "&Mosaico"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuPopupEdit 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuPopupHist 
         Caption         =   "Historial de cambios"
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupSelectAll 
         Caption         =   "Seleccionar todos"
      End
      Begin VB.Menu mnuPopupUnselectAll 
         Caption         =   "Deseleccionar todos"
      End
   End
   Begin VB.Menu mnuPopupTree 
      Caption         =   "PopupTree"
      Begin VB.Menu mnuPopupTreeRefresh 
         Caption         =   "Actualizar"
      End
      Begin VB.Menu mnuPopupTreeSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupTreeSelectAll 
         Caption         =   "Seleccionar todo"
      End
      Begin VB.Menu mnuPopupTreeSelectByModif 
         Caption         =   "Seleccionar por fecha de modificacion"
      End
      Begin VB.Menu mnuPopupTreeSelectForms 
         Caption         =   "Seleccionar Forms que usan"
      End
      Begin VB.Menu mnuPopupTreeUnselectAll 
         Caption         =   "Deseleccionar todo"
      End
   End
   Begin VB.Menu mnuPopupFileExp 
      Caption         =   "PopupFileExp"
      Begin VB.Menu mnuPopupFileExpNew 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnuPopupFileExpDelete 
         Caption         =   "Borrar"
      End
   End
   Begin VB.Menu mnuPopupHistory 
      Caption         =   "PopupHistory"
      Begin VB.Menu mnuPopupHistoryView 
         Caption         =   "Ver"
      End
      Begin VB.Menu mnuPopupHistoryDiffLast 
         Caption         =   "Diff con la ult version"
      End
      Begin VB.Menu mnuPopupHistoryDiffNext 
         Caption         =   "Diff con la sig version"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dicFld As Object

Private Sub MDIForm_Load()
    Caption = "Remote DCE v" & App.Major & "." & App.Minor & "." & App.Revision
    
    mnuPopup.Visible = False
    mnuPopupTree.Visible = False
    mnuPopupFileExp.Visible = False
    mnuPopupHistory.Visible = False
    
    frmLogon.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim i As Long
    Dim sAux As String
    Dim ses As Object
    
    If Not Debugging Then
        If GSession.IsLogged Then GSession.Logoff
    End If
    
    i = 0
    Do
        sAux = GdicURLs.Keys(i)
        If sAux <> "" Then WriteIni "Session", "ServerURL" & i, sAux
        i = i + 1
    Loop Until i = GdicURLs.Count Or i = 10 Or sAux = ""
    Do While i < 10
        WriteIni "Session", "ServerURL" & i, ""
        i = i + 1
    Loop
End Sub

Private Sub mnuPopupFileExpDelete_Click()
    frmFileExplorer.mnuPopupFileExpDeleteClick
End Sub

Private Sub mnuPopupFileExpNew_Click()
    frmFileExplorer.mnuPopupFileExpNewClick
End Sub

Private Sub mnuPopupHistoryDiffLast_Click()
    ActiveForm.mnuPopupHistoryDiffLastClick
End Sub

Private Sub mnuPopupHistoryDiffNext_Click()
    ActiveForm.mnuPopupHistoryDiffNextClick
End Sub

Private Sub mnuPopupHistoryView_Click()
    ActiveForm.mnuPopupHistoryViewClick
End Sub

Private Sub mnuPopupSelectAll_Click()
    If Not mnuSetupCheckboxes.Checked Then mnuSetupCheckboxes_Click
    frmExplorer.mnuPopupSelectAllClick
End Sub

Private Sub mnuPopupTreeSelectAll_Click()
    If Not mnuSetupCheckboxes.Checked Then mnuSetupCheckboxes_Click
    frmExplorer.mnuPopupTreeSelectAllClick
End Sub

Private Sub mnuPopupTreeSelectByModif_Click()
    If Not mnuSetupCheckboxes.Checked Then mnuSetupCheckboxes_Click
    frmExplorer.mnuPopupTreeSelectByModifClick
End Sub

Private Sub mnuPopupTreeSelectForms_Click()
    If Not mnuSetupCheckboxes.Checked Then mnuSetupCheckboxes_Click
    frmExplorer.mnuPopupTreeSelectForms
End Sub

Private Sub mnuPopupTreeUnselectAll_Click()
    If Not mnuSetupCheckboxes.Checked Then mnuSetupCheckboxes_Click
    frmExplorer.mnuPopupTreeUnselectAllClick
End Sub

Private Sub mnuPopupUnselectAll_Click()
    If Not mnuSetupCheckboxes.Checked Then mnuSetupCheckboxes_Click
    frmExplorer.mnuPopupUnselectAllClick
End Sub

Private Sub mnuSetupClear_Click()
    GSelected.Clear
    If Not mnuSetupCheckboxes.Checked Then mnuSetupCheckboxes_Click
    frmExplorer.mnuPopupTreeRefreshClick
End Sub

Private Sub mnuSetupGenerate_Click()
    Dim sCode As String
    Dim dom
    Dim errN As Long
    Dim fso As Object, txt As Object
    Dim tOut As Long
    
    With CommonDialog1
        .Filter = "VbScript Files|*.vbs"
        .FileName = ""
        .CancelError = True
        On Error Resume Next
        .ShowSave
        errN = Err.Number
        On Error GoTo 0
        If errN <> 0 Then Exit Sub
    End With

    On Error GoTo Error
    tOut = GSession.HttpRequestTimeout
    Screen.MousePointer = vbHourglass
    
    frmExplorer.StatusBar1.SimpleText = "Esperando respuesta del server (puede tardar varios minutos)..."
    DoEvents
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txt = fso.OpenTextFile(CommonDialog1.FileName, 2, True) ' ForWriting
    
    sCode = "Return ScriptObjects(Arg(1))"
    
    GSession.HttpRequestTimeout = 7200 ' 2 horas
    Set dom = GSession.HttpCallCode(sCode, Array(GSelected.dom.Xml)).responseXml
    GSession.HttpRequestTimeout = tOut

    txt.Write Replace(dom.documentElement.firstChild.Text, Chr(10), vbCrLf)
    txt.Close
    Shell "notepad.exe """ & CommonDialog1.FileName & """", vbNormalFocus

    frmExplorer.StatusBar1.SimpleText = "Listo"
    DoEvents

    Screen.MousePointer = vbNormal
    Exit Sub
Error:
    txt.Close
    GSession.HttpRequestTimeout = tOut
    Screen.MousePointer = vbNormal
    ErrDisplay Err
End Sub

Private Sub mnuSetupLoad_Click()
    Dim dom, node
    Dim errN As Long
    
    frmExplorer.StatusBar1.SimpleText = ""
    
    With CommonDialog1
        .Filter = "XML Files|*.xml"
        .FileName = GstrFileName
        .CancelError = True
        On Error Resume Next
        .ShowOpen
        errN = Err.Number
        On Error GoTo 0
        If errN <> 0 Then
            Exit Sub
        Else
            GstrFileName = .FileName
        End If
    End With
    
    GSelected.Load GstrFileName
    
    If Not mnuSetupCheckboxes.Checked Then mnuSetupCheckboxes_Click
    frmExplorer.mnuPopupTreeRefreshClick
End Sub

Private Sub mnuSetupSave_Click()
    Dim dom, node, s
    Dim errN As Long
    
    frmExplorer.StatusBar1.SimpleText = ""
    
    With CommonDialog1
        .Filter = "XML Files|*.xml"
        .FileName = GstrFileName
        .CancelError = True
        On Error Resume Next
        .ShowSave
        errN = Err.Number
        On Error GoTo 0
        If errN <> 0 Then
            Exit Sub
        Else
            GstrFileName = .FileName
        End If
    End With
    
    GSelected.Save GstrFileName
    
    frmExplorer.StatusBar1.SimpleText = "Seleccion guardada"
End Sub

Private Sub mnuWindowConnection_Click()
    frmLogon.Show
    frmLogon.SetFocus
End Sub

Private Sub mnuWindowExplorer_Click()
    frmExplorer.Show
    frmExplorer.SetFocus
End Sub

Private Sub mnuSetupCheckboxes_Click()
    Dim chk As Boolean
    
    chk = Not mnuSetupCheckboxes.Checked
    mnuSetupCheckboxes.Checked = chk
    
    With frmExplorer
        .TreeView1.Checkboxes = chk
        .lstSyncEvents.Checkboxes = chk
        .lstAsyncEvents.Checkboxes = chk
        .chkAcl.Visible = chk
        .lstDocuments.Checkboxes = chk
        .lstViews.Checkboxes = chk
    End With
End Sub

Private Sub mnuWindowFileExplorer_Click()
    On Error GoTo Error
    
    If Not GSession.LoggedUser.IsAdmin Then
        MsgBox "Solo para administradores", vbInformation
    Else
        frmFileExplorer.Show
        frmFileExplorer.SetFocus
    End If

    Exit Sub
Error:
    ErrDisplay Err
End Sub

Private Sub mnuWindowTile_Click()
    MDIForm1.Arrange vbTileHorizontal
End Sub

Private Sub mnuPopupEdit_Click()
    ActiveForm.mnuPopupEditClick
End Sub

Private Sub mnuPopupHist_Click()
    ActiveForm.mnuPopupHistClick
End Sub

Private Sub mnuPopupTreeRefresh_Click()
    ActiveForm.mnuPopupTreeRefreshClick
End Sub

