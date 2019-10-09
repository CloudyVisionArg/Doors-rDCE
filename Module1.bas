Attribute VB_Name = "Module1"
Option Explicit

Public CodeMaxGlobals As CodeMax4Ctl.Globals

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Long

Public GstrIniFile As String
Public GblnNormalizeCase As Boolean
Public GSession As Object
Public GdicURLs As Scripting.Dictionary
Public GSelected As clsSelected
Public GstrFileName As String
Public GlngMaxDocs As Long

Public SyncEventName(1 To 12) As String
Public dicFolderCache As Scripting.Dictionary
Public dicFormCache As Scripting.Dictionary

Sub Main()
    Dim lCMaxRevision As Long
    Dim i As Long
    Dim sAux As String
    
    SyncEventName(1) = "Document_Open"
    SyncEventName(2) = "Document_BeforeSave"
    SyncEventName(3) = "Document_AfterSave"
    SyncEventName(4) = "Document_BeforeCopy"
    SyncEventName(5) = "Document_AfterCopy"
    SyncEventName(6) = "Document_BeforeDelete"
    SyncEventName(7) = "Document_AfterDelete"
    SyncEventName(8) = "Document_BeforeMove"
    SyncEventName(9) = "Document_AfterMove"
    SyncEventName(10) = "Document_BeforeFieldChange"
    SyncEventName(11) = "Document_AfterFieldChange"
    SyncEventName(12) = "Document_Terminate"
    
    GstrIniFile = App.Path & "\" & App.EXEName & ".ini"
    GblnNormalizeCase = True
    Set CodeMaxGlobals = New CodeMax4Ctl.Globals
    
    Set GdicURLs = CreateObject("Scripting.Dictionary")
    GdicURLs.CompareMode = 1 ' VBTextCompare
        
    GstrFileName = ""
    
    For i = 0 To 19
        sAux = ReadIni("Session", "ServerURL" & i)
        If sAux <> "" Then GdicURLs.Add sAux, Empty
    Next
    
    Set GSession = CreateObject("dapihttp.Session")
    GSession.ShowHtmlErrPage = True
    Set GSelected = New clsSelected
    Set dicFolderCache = New Scripting.Dictionary
    Set dicFormCache = New Scripting.Dictionary
    
    sAux = ReadIni("Session", "MaxDocs")
    If IsNumeric(sAux) Then
        GlngMaxDocs = CLng(sAux)
    Else
        GlngMaxDocs = 500
    End If
    
    On Error Resume Next
    lCMaxRevision = CLng(Split(CodeMaxVersion, ".")(3))
    If lCMaxRevision < 9 Then GblnNormalizeCase = False
    On Error GoTo 0
    
    MDIForm1.Show
End Sub

Public Function ReadIni(ByRef Section As String, ByRef Key As String) As String
    ReadIni = GetProfileString(GstrIniFile, Section, Key)
End Function

Public Function WriteIni(ByRef Section As String, ByRef Key As String, ByRef Value As String) As Long
    WriteIni = WriteProfileString(GstrIniFile, Section, Key, Value)
End Function

Private Function GetProfileString(ByRef IniFile As String, ByRef Application As String, ByRef Key As String) As String
  Dim strAux As String
  
  strAux = String(1024, vbNullChar)
  If GetPrivateProfileString(Application, Key, "", strAux, 1024, IniFile) = 0 Then
    GetProfileString = ""
  Else
    GetProfileString = Left(strAux, InStr(strAux, vbNullChar) - 1)
  End If
End Function

Private Function WriteProfileString(ByRef IniFile As String, ByRef Application As String, ByRef Key As String, ByRef Value As String) As Long
  WriteProfileString = WritePrivateProfileString(Application, Key, Value, IniFile)
End Function

Private Function CodeMaxVersion() As String
    Dim oShell As Object
    Dim strDll As String
    Dim oFso As Object
    
    Set oShell = CreateObject("WScript.Shell")
    strDll = oShell.RegRead("HKEY_CLASSES_ROOT\CLSID\{BCA00001-18B1-43E0-BB89-FECDDBF0472E}\InprocServer32\")
    Set oFso = CreateObject("Scripting.FileSystemObject")
    CodeMaxVersion = oFso.GetFileVersion(strDll)
End Function

Public Sub ErrDisplay(ByRef Err As Object)
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation
End Sub

Public Sub ListViewColumnClick(ByRef pListView As MSComctlLib.ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With pListView
        If Not .Sorted Then
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            .Sorted = True
        Else
            If .SortKey = ColumnHeader.Index - 1 Then
                ' Invertir
                If .SortOrder = lvwAscending Then
                    .SortOrder = lvwDescending
                Else
                    .SortOrder = lvwAscending
                End If
            Else
                .SortKey = ColumnHeader.Index - 1
                .SortOrder = lvwAscending
            End If
        End If
    End With
End Sub

Public Function Debugging() As Boolean
    On Error Resume Next
    Debug.Assert 1 / 0
    Debugging = (Err <> 0)
End Function

Public Sub ValidateDateTextbox(pTextbox)
    With pTextbox
        If IsDate(.Text) Then
            .Text = CDate(.Text)
        Else
            .Text = ""
        End If
    End With
End Sub

Public Function FolderCache(FolderId As Long) As Object
    Dim arr, fld
    
    If dicFolderCache.Exists("ID=" & FolderId) Then
        arr = dicFolderCache("ID=" & FolderId)
        If (Now - arr(1)) * 86400 > 60 Then ' Un minuto
            Set arr(0) = GSession.FoldersGetFromId(FolderId)
            arr(1) = Now
            dicFolderCache("ID=" & FolderId) = arr
        End If
        Set FolderCache = arr(0)
    Else
        Set fld = GSession.FoldersGetFromId(FolderId)
        dicFolderCache.Add "ID=" & FolderId, Array(fld, Now)
        Set FolderCache = fld
    End If
End Function

Public Function FormCache(FormId As Variant) As Object
    Dim arr, frm
    
    If dicFormCache.Exists("ID=" & FormId) Then
        arr = dicFormCache("ID=" & FormId)
        If (Now - arr(1)) * 86400 > 60 Then ' Un minuto
            Set arr(0) = GSession.Forms(FormId)
            arr(1) = Now
            dicFormCache("ID=" & FormId) = arr
        End If
        Set FormCache = arr(0)
    Else
        Set frm = GSession.Forms(FormId)
        dicFormCache.Add "ID=" & FormId, Array(frm, Now)
        Set FormCache = frm
    End If
End Function

Public Function FormPK(Form As Object) As String
    Dim sPK As String
    
    sPK = LCase(Form.PK)
    If sPK = "" Then
        If UCase(Form.Guid) = "F89ECD42FAFF48FDA229E4D5C5F433ED" Then ' CodeLib
            sPK = "name"
        ElseIf UCase(Form.Guid) = "EAC99A4211204E1D8EEFEB8273174AC4" Then ' Controls
            sPK = "name"
        ElseIf UCase(Form.Guid) = "B87B1CB5EFB94B03BA6B1F18DBE5F5D4" Then ' Keywords3
            sPK = "id"
        ElseIf UCase(Form.Guid) = "B89302DBBE45498EA03A495B53D3F50C" Then ' Secuences3
            sPK = "sequence"
        ElseIf UCase(Form.Guid) = "5C0D6DBF72CF42608989A862ED9E7444" Then ' Settings3
            sPK = "setting"
        End If
    End If

    FormPK = sPK
End Function

Public Function FormCode(Form As Object) As String
    Dim sCode As String
    
    If Form.Properties.Exists("DCE_CodeColumn") Then
        sCode = Form.Properties("DCE_CodeColumn").Value
    Else
        If UCase(Form.Guid) = "F89ECD42FAFF48FDA229E4D5C5F433ED" Then ' CodeLib
            sCode = "code"
        ElseIf UCase(Form.Guid) = "EAC99A4211204E1D8EEFEB8273174AC4" Then ' Controls
            sCode = "scriptbeforerender"
        End If
    End If

    FormCode = sCode
End Function

