VERSION 5.00
Object = "{BCA00000-0F85-414C-A938-5526E9F1E56A}#4.0#0"; "cmax40.dll"
Begin VB.Form frmEditor 
   Caption         =   "Form1"
   ClientHeight    =   2899
   ClientLeft      =   65
   ClientTop       =   351
   ClientWidth     =   3874
   Icon            =   "frmEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2899
   ScaleWidth      =   3874
   WindowState     =   2  'Maximized
   Begin CodeMax4Ctl.CodeMax CodeMax1 
      Height          =   1695
      Left            =   240
      OleObjectBlob   =   "frmEditor.frx":058A
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CodeType As Long '1-Folder, 2-Form, 3-Document, 4-AsyncEvent, 5-File
Public dForm As Object
Public Folder As Object
Public DocId As Long
Public EventKey As String
Public Field As String
Public CodeChanged As Boolean
Public FilePath As String
Public Charset As String

Private Sub CodeMax1_Change()
    If Not CodeChanged Then
        CodeChanged = True
        Caption = "* " & Caption
    End If
End Sub

Private Sub Form_Activate()
    Dim sLang As String
    
    If TypeName(Folder) = "Folder" Then
        If Folder.Properties.Exists("DCE_Language") Then
            sLang = Folder.Properties("DCE_Language").Value
        End If
    End If

    If sLang <> "" Then
        CodeMax1.Language = CMaxLang(sLang)
    Else
        CodeMax1.Language = CMaxLang("VBScript")
    End If
    
    CodeMax1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyS And Shift = vbCtrlMask Then
        KeyCode = 0
        Save
    ElseIf KeyCode = vbKeyEscape And Shift = 0 Or _
           KeyCode = vbKeyF4 And Shift = vbCtrlMask Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim sCMaxVersion As String
    Dim lCMaxRevision As Long
    
    With CodeMax1
        .NormalizeCase = False
        .DisplayLeftMargin = False
        .Font.Size = 10
        .LineNumbering = True
    End With
    
    'CodeMaxGlobals.Languages.RemoveAll
    CMaxLoadLang "VBScript", "vbscript.lng"
    CMaxLoadLang "JavaScript", "js.lng"
    
    CodeChanged = False
End Sub

Private Sub CMaxLoadLang(pLangName As String, pLangFile As String)
    Dim bAdd As Boolean
    Dim lang As CodeMax4Ctl.Language
    Dim i As Long
    Dim sCMaxVersion As String
    Dim lCMaxRevision As Long
    
    With CodeMax1
        bAdd = True
        For i = 0 To CodeMaxGlobals.Languages.Count - 1
            Set lang = CodeMaxGlobals.Languages(i)
            If lang.Name = pLangName Then
                bAdd = False
                Exit For
            End If
        Next
        
        If bAdd Then
            Set lang = New CodeMax4Ctl.Language
            lang.LoadXmlDefinition App.Path & "\" & pLangFile
            lang.Register
        End If
        
        CodeMax1.Language = lang
    End With
End Sub

Private Function CMaxLang(pLangName As String) As CodeMax4Ctl.Language
    Dim lang As CodeMax4Ctl.Language
    Dim i As Long
    
    With CodeMax1
        For i = 0 To CodeMaxGlobals.Languages.Count - 1
            Set lang = CodeMaxGlobals.Languages(i)
            If lang.Name = pLangName Then
                Set CMaxLang = lang
                Exit For
            End If
        Next
    End With
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        With CodeMax1
            .Top = 1
            .Left = 1
            .Height = ScaleHeight
            .Width = ScaleWidth
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim resp As VbMsgBoxResult
    
    If CodeChanged Then
        resp = MsgBox("Guardar cambios?", vbYesNoCancel + vbQuestion)
        If resp = vbYes Then
            Save
            If CodeChanged Then Cancel = 1
        ElseIf resp = vbCancel Then
            Cancel = 1
        End If
    End If
    
    If Cancel = 1 Then CodeMax1.SetFocus
End Sub

Sub Save()
    Dim sbH As Object
    Dim doc As Object
    Dim fld As Object
    Dim frm As Object
    Dim evn As Object
    Dim sCode As String
    Dim Args As Variant
    
    On Error GoTo Error
    
    If CodeChanged Then
        
        If CodeType <> 5 Then
            Set sbH = GSession.ConstructNewSqlBuilder
            sbH.Add "TIMESTAMP", Now, 2
            sbH.Add "ACC_ID", GSession.LoggedUser.id, 3
            sbH.Add "ACC_NAME", GSession.LoggedUser.Name, 1
            sbH.Add "CODETYPE", CodeType, 3
            sbH.Add "CODE", "?", 0
        End If
        
        Select Case CodeType
            
            Case 1
                'FolderEvents
                Set fld = FolderCache(Folder.id)
                Set evn = fld.Events("ID=" & Mid(EventKey, 4))
                evn.Code = CodeMax1.Text
                fld.Save
                sbH.Add "FLD_ID", Folder.id, 3
                sbH.Add "SEV_ID", evn.id, 3
            
            Case 2
                'FormEvents
                Set frm = FormCache(dForm.id)
                Set evn = frm.Events("ID=" & Mid(EventKey, 4))
                evn.Code = CodeMax1.Text
                frm.Save
                sbH.Add "FRM_ID", dForm.id, 3
                sbH.Add "SEV_ID", Mid(EventKey, 4), 3
            
            Case 3
                'Document
                Set doc = GSession.DocumentsGetFromId(DocId)
                doc.Fields(Field).Value = CodeMax1.Text
                doc.Save
                If (CLng(Left(GSession.Version, 1)) >= 7) Then
                    GSession.ClearAllCustomCache
                    GSession.ClearObjectModelCache "ComCodeLibCache"
                End If
                
                sbH.Add "FLD_ID", Folder.id, 3
                sbH.Add "DOC_ID", DocId, 3
        
            Case 4
                'AsyncEvents
                Set fld = FolderCache(Folder.id)
                Set evn = fld.AsyncEvents("ID=" & Mid(EventKey, 4))
                evn.Code = CodeMax1.Text
                fld.Save
                sbH.Add "FLD_ID", Folder.id, 3
                sbH.Add "SEV_ID", evn.id, 3
        
            Case 5
                ' File
                frmFileExplorer.SaveFile FilePath, CodeMax1.Text, Charset
                frmFileExplorer.RefreshOnFocus = True
                
        End Select
        
        CodeChanged = False
        Caption = Mid(Caption, 3)
        
        ' Inserta en DCE_HISTORY
        
        If CodeType <> 5 Then
            sCode = ""
            sCode = sCode & "Set oCmd = dSession.ConstructNewADODBCommand" & vbCrLf
            sCode = sCode & "oCmd.CommandType = 1" & vbCrLf
            sCode = sCode & "Set oPar = oCmd.CreateParameter(""CODE_VALUE"", 201, 1)" & vbCrLf
            sCode = sCode & "oPar.Size = Len(CStr(Arg(2)))" & vbCrLf
            sCode = sCode & "oPar.Value = CStr(Arg(2))" & vbCrLf
            sCode = sCode & "oCmd.Parameters.Append oPar" & vbCrLf
            sCode = sCode & "oCmd.CommandText = CStr(Arg(1))" & vbCrLf
            sCode = sCode & "dSession.Db.ExecuteCommand oCmd"
        
            Args = Array(Empty, Empty)
            Args(0) = "insert into DCE_HISTORY " & sbH.InsertString
            Args(1) = CodeMax1.Text
        
            On Error Resume Next
            GSession.HttpCallCode sCode, Args
            On Error GoTo Error
        End If
    End If
    
    Exit Sub
Error:
    MsgBox Err.Description, vbExclamation
End Sub

