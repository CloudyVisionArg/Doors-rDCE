VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim oDom As Object
Dim oDomOld As Object

Public Function AddFolder(Folder As Variant, Optional Checked As Boolean = True) As Object
    Dim fld As Object
    Dim arr, node, nodeChild
    Dim i As Long
    
    If TypeName(Folder) = "Folder" Then
        Set fld = Folder
    Else
        Set fld = FolderCache(CLng(Folder))
    End If
    
    Set node = Nothing
    
    If fld.id = 11 Then ' Codelib
        Set node = oDom.selectSingleNode("/root/system/codelib")
        
    ElseIf Not fld.System Then
        arr = Split(Mid(fld.Path(True), 2), "/")
        Set node = oDom.documentElement
        For i = 0 To UBound(arr)
            Set nodeChild = node.selectSingleNode("folder[@name='" & arr(i) & "']")
            If nodeChild Is Nothing Then
                Set nodeChild = oDom.createNode("element", "folder", "")
                nodeChild.setAttribute "name", arr(i)
                nodeChild.setAttribute "checked", "0"
                node.appendChild nodeChild
            End If
            Set node = nodeChild
        Next
        If Checked Then nodeChild.setAttribute "checked", "1"
    
    End If
    
    Set AddFolder = node
End Function

Public Function AddFolderItem(ItemType As String, Folder As Variant, item As Variant) As Object
    Dim fldNode, contNode, node, arr, DocId, fie, fld
    Dim sPK As String
    Dim i As Long
    Dim doc As Object
    
    Set fldNode = AddFolder(Folder, False)
    
    If ItemType = "Document" Then
        
        If TypeName(Folder) = "Folder" Then
            Set fld = Folder
        Else
            Set fld = FolderCache(CLng(Folder))
        End If
        
        sPK = FormPK(fld.Form)
        If sPK = "" Then sPK = "doc_id"
        arr = Split(sPK, ",")
        For i = 0 To UBound(arr)
            arr(i) = Trim(LCase(arr(i)))
        Next
        
        Set contNode = GetContainer(fldNode, "documents")
        Set node = contNode.selectSingleNode(DocumentXPath(Folder, item, True, doc))
        
        If node Is Nothing Then
            Set node = oDom.createNode("element", "item", "")
            If TypeName(item) = "IXMLDOMElement" Then
                For i = 0 To UBound(arr)
                    node.setAttribute arr(i), item.getAttribute(arr(i))
                Next
            Else
                If TypeName(item) = "Document" Then
                    DocId = item.id
                Else
                    DocId = item
                End If
                
                For i = 0 To UBound(arr)
                    If arr(i) = "doc_id" Then
                        node.setAttribute "doc_id", DocId
                    Else
                        Set fie = doc.Fields(arr(i))
                        node.setAttribute arr(i), GSession.Xml.XmlEncode(fie.Value, fie.DataType)
                    End If
                Next
            End If
            
            contNode.appendChild node
        End If
        
        
    ElseIf ItemType = "View" Then
    
        Set contNode = GetContainer(fldNode, "views")
        Set node = contNode.selectSingleNode("item[@name='" & item & "']")
        If node Is Nothing Then
            Set node = oDom.createNode("element", "item", "")
            node.setAttribute "name", item
            contNode.appendChild node
        End If
            
    ElseIf ItemType = "SyncEvent" Then
    
        Set contNode = GetContainer(fldNode, "syncevents")
        Set node = contNode.selectSingleNode("item[@id='" & item & "']")
        If node Is Nothing Then
            Set node = oDom.createNode("element", "item", "")
            node.setAttribute "id", item
            node.setAttribute "name", SyncEventName(item)
            contNode.appendChild node
        End If
        
    ElseIf ItemType = "AsyncEvent" Then
    
        Set contNode = GetContainer(fldNode, "asyncevents")
        Set node = contNode.selectSingleNode("item[@id='" & item & "']")
        If node Is Nothing Then
            Set node = oDom.createNode("element", "item", "")
            node.setAttribute "id", item
            contNode.appendChild node
        End If
    
    End If
    
    Set AddFolderItem = node
End Function

Public Function AddForm(Form As Variant) As Object
    Dim frm As Object
    Dim node, nodeChild
    
    If TypeName(Form) = "CustomForm" Then
        Set frm = Form
    Else
        Set frm = FormCache(CLng(Form))
    End If
    
    Set node = oDom.selectSingleNode(FormXPath(frm))
    If node Is Nothing Then
        Set node = oDom.selectSingleNode("/root/system/forms")
        Set nodeChild = oDom.createNode("element", "item", "")
        nodeChild.setAttribute "name", frm.Name
        nodeChild.setAttribute "guid", LCase(frm.Guid)
        node.appendChild nodeChild
        Set node = nodeChild
    End If
    
    Set AddForm = node
End Function

Public Function AddFormItem(ItemType As String, Form As Variant, item As Variant) As Object
    Dim frmNode, contNode, node
    
    Set frmNode = AddForm(Form)
    
    If ItemType = "SyncEvent" Then
    
        Set contNode = frmNode.selectSingleNode("syncevents")
        If contNode Is Nothing Then
            Set contNode = oDom.createNode("element", "syncevents", "")
            frmNode.appendChild contNode
        End If
        
        Set node = contNode.selectSingleNode("item[@id='" & item & "']")
        If node Is Nothing Then
            Set node = oDom.createNode("element", "item", "")
            node.setAttribute "id", item
            node.setAttribute "name", SyncEventName(item)
            contNode.appendChild node
        End If
    
    End If
    
    Set AddFormItem = node
End Function

Public Function AddAccount(AccountName As String) As Object
    Dim node, nodeChild
        
    Set node = oDom.selectSingleNode(AccountXPath(AccountName))
    If node Is Nothing Then
        Set node = oDom.selectSingleNode("/root/system/directory")
        Set nodeChild = oDom.createNode("element", "item", "")
        nodeChild.setAttribute "name", AccountName
        node.appendChild nodeChild
        Set node = nodeChild
    End If
    
    Set AddAccount = node
End Function

Public Sub Remove(XPath As String)
    Dim node
    
    Set node = oDom.selectSingleNode(XPath)
    If Not node Is Nothing Then
        If node.baseName = "folder" Then
            node.setAttribute "checked", "0"
        Else
            node.parentNode.RemoveChild node
        End If
    End If
End Sub

Private Sub Class_Initialize()
    Clear
End Sub

Public Sub Clear()
    Set oDom = GSession.Xml.NewDom
    oDom.loadXml "<root><system><codelib /><directory /><forms /></system></root>"
End Sub

Public Sub Load(FilePath As String)
    oDom.Load FilePath
End Sub

Public Sub Save(FilePath As String)
    oDom.Save FilePath
End Sub

Public Function Checked(XPath As String) As Boolean
    Dim node As Object
    
    Checked = False
    Set node = oDom.selectSingleNode(XPath)
    If Not node Is Nothing Then
        If node.baseName = "folder" Then
            Checked = (node.getAttribute("checked") = "1")
        Else
            Checked = True
        End If
    End If
End Function

Friend Function dom() As Object
    Set dom = oDom
End Function

Public Function FolderXPath(Folder As Variant) As String
    Dim sAux As String
    Dim node, fld, arr
    Dim i As Long
    
    If TypeName(Folder) = "IXMLDOMElement" Then
        
        sAux = ""
        Set node = Folder
        Do While node.baseName = "folder"
            sAux = "/folder[@name='" & node.getAttribute("name") & "']" & sAux
            Set node = node.parentNode
        Loop
        FolderXPath = "/root" & sAux
        
    Else
        
        If TypeName(Folder) = "Folder" Then
            Set fld = Folder
        Else
            Set fld = FolderCache(CLng(Folder))
        End If
        
        If fld.id = 1 Then
            FolderXPath = "/root/system"
        ElseIf fld.id = 3 Then
            FolderXPath = "/root/system/directory"
        ElseIf fld.id = 5 Then
            FolderXPath = "/root/system/forms"
        ElseIf fld.id = 11 Then
            FolderXPath = "/root/system/codelib"
        Else
            sAux = "/root"
            arr = Split(Mid(fld.Path(True), 2), "/")
            For i = 0 To UBound(arr)
                sAux = sAux & "/folder[@name='" & arr(i) & "']"
            Next
            FolderXPath = sAux
        End If
    End If
End Function

Public Function FormXPath(Form As Variant) As String
    Dim sAux As String
    
    sAux = "/root/system/forms/item[@guid='"
    If TypeName(Form) = "IXMLDOMElement" Then
        sAux = sAux & LCase(Form.getAttribute("guid"))
    ElseIf TypeName(Form) = "CustomForm" Then
        sAux = sAux & LCase(Form.Guid)
    Else
        sAux = sAux & LCase(FormCache(CLng(Form)).Guid)
    End If
    FormXPath = sAux & "']"
End Function

Public Function AsyncEventXPath(Folder As Variant, AsyncEventId As Long) As String
    AsyncEventXPath = FolderXPath(Folder) & "/asyncevents/item[@id='" & AsyncEventId & "']"
End Function

Public Function FolderEventXPath(Folder As Variant, SyncEventId As Long) As String
    FolderEventXPath = FolderXPath(Folder) & "/syncevents/item[@id='" & SyncEventId & "']"
End Function

Public Function FormEventXPath(Form As Variant, SyncEventId As Long) As String
    FormEventXPath = FormXPath(Form) & "/syncevents/item[@id='" & SyncEventId & "']"
End Function

Public Function DocumentXPath(Folder As Variant, Document As Variant, Optional WithoutPath As Boolean = False, Optional ByRef RetDoc As Object) As String
    Dim fld, arr, DocId
    Dim sPK As String
    Dim i As Long
    Dim fie As Object
    
    If TypeName(Folder) = "Folder" Then
        Set fld = Folder
    Else
        Set fld = FolderCache(CLng(Folder))
    End If
    
    sPK = FormPK(fld.Form)
    If sPK = "" Then sPK = "doc_id"
    arr = Split(sPK, ",")
    For i = 0 To UBound(arr)
        arr(i) = Trim(LCase(arr(i)))
    Next
    
    If TypeName(Document) = "IXMLDOMElement" Then
        For i = 0 To UBound(arr)
            sPK = "@" & arr(i) & "='" & Document.getAttribute(arr(i)) & "'"
        Next
    Else
        If TypeName(Document) = "Document" Then
            DocId = Document.id
            If sPK <> "doc_id" Then Set RetDoc = Document
        Else
            DocId = Document
            If sPK <> "doc_id" Then Set RetDoc = fld.Documents(Document)
        End If
        
        sPK = ""
        For i = 0 To UBound(arr)
            If arr(i) = "doc_id" Then
                sPK = "@" & arr(i) & "='" & DocId & "'"
            Else
                Set fie = RetDoc.Fields(arr(i))
                sPK = "@" & arr(i) & "='" & GSession.Xml.XmlEncode(fie.Value, fie.DataType) & "'"
            End If
        Next
    End If
    
    If WithoutPath Then
        DocumentXPath = "item[" & sPK & "]"
    Else
        DocumentXPath = FolderXPath(Folder) & "/documents/item[" & sPK & "]"
    End If
End Function

Public Function ViewXPath(Folder As Variant, ViewName As String) As String
    ViewXPath = FolderXPath(Folder) & "/views/item[@name='" & ViewName & "']"
End Function

Public Function AccountXPath(AccountName As String) As String
    AccountXPath = "/root/system/directory/item[@name='" & AccountName & "']"
End Function

Public Function GetContainer(node, container) As Object
    Dim n
    
    Set n = node.selectSingleNode(container)
    If n Is Nothing Then
        Set n = node.ownerDocument.createNode("element", container, "")
        node.appendChild n
    End If

    Set GetContainer = n
End Function

