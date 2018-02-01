VERSION 5.00
Begin VB.Form frmScript 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3984
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3768
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3984
   ScaleWidth      =   3768
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Scripting options"
      Height          =   3012
      Left            =   180
      TabIndex        =   10
      Top             =   120
      Width           =   3372
      Begin VB.CheckBox chkFormsEvents 
         Caption         =   "Script form events"
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   2460
         Width           =   1812
      End
      Begin VB.CheckBox chkFormsActions 
         Caption         =   "Script form actions"
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox chkFormsFields 
         Caption         =   "Script form fields"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   1860
         Width           =   2055
      End
      Begin VB.CheckBox chkForms 
         Caption         =   "Script forms"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chkRecursive 
         Caption         =   "Recursive"
         Height          =   195
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkViews 
         Caption         =   "Script views"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   1260
         Width           =   1815
      End
      Begin VB.CheckBox chkAsyncEvents 
         Caption         =   "Script async events"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox chkEvents 
         Caption         =   "Script events"
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   660
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   3360
      Width           =   1212
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3360
      Width           =   1212
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Folder As Object
Public CustomForm As Object
Private stm As Object
Private arrForms

Private Sub chkForms_Click()
    Dim bEnabled As Boolean
    
    bEnabled = (chkForms.Value = 1)
    chkFormsFields.Enabled = bEnabled
    chkFormsActions.Enabled = bEnabled
    chkFormsEvents.Enabled = bEnabled
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo Error
    
    If CustomForm Is Nothing Then
        ScriptFolder
    Else
        ScriptForm
    End If
    
    Exit Sub
Error:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub ScriptFolder()
    Dim fso As Object
    
    On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stm = fso.OpenTextFile(App.Path & "\" & App.EXEName & ".vbs", 2, True) ' ForWriting
    
    stm.Write "'----------------------------------------------------------" & vbCrLf
    stm.Write "' To execute this script, uncomment the following line" & vbCrLf
    stm.Write "' and check the start-up folder" & vbCrLf
    stm.Write "'----------------------------------------------------------" & vbCrLf
    stm.Write vbCrLf
    stm.Write "' Set curFolder = dSession.FoldersGetFromId(1001)" & vbCrLf
    stm.Write vbCrLf & vbCrLf

    arrForms = Empty
    ScriptFolder2 Folder
    stm.Close
       
    Shell "notepad.exe """ & App.Path & "\" & App.EXEName & ".vbs""", vbNormalFocus
    
    Screen.MousePointer = vbNormal
    Exit Sub
Error:
    stm.Close
    Screen.MousePointer = vbNormal
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub ScriptFolder2(pFolder As Object)
    Dim oEvn As Object
    Dim node As Object
    Dim subFld As Object
    Dim oView As Object
    Dim oDoc As Object
    Dim oForm As Object
    Dim i As Long
    Dim bEsta As Boolean
    Dim sKey As String
    
    If pFolder.FolderType = 3 Then ' VirtualFolder
        Exit Sub 'todo: ver como hacemos con las subcarpetas de estas
    End If
    
    If chkForms.Value = 1 And pFolder.FolderType = 1 Then
        Set oForm = pFolder.Form
        
        bEsta = False
        If Not IsEmpty(arrForms) Then
            For i = 0 To UBound(arrForms)
                If arrForms(i) = oForm.Guid Then
                    bEsta = True
                    Exit For
                End If
            Next
        End If
            
        If Not bEsta Then
            If IsEmpty(arrForms) Then
                i = 0
                ReDim arrForms(0)
            Else
                i = UBound(arrForms) + 1
                ReDim Preserve arrForms(i)
            End If
            arrForms(i) = oForm.Guid
            
            ScriptForm2 oForm
        End If
    End If
    
    
    
    
    stm.Write vbCrLf

    If chkRecursive.Value = 1 Then
        For Each node In pFolder.FoldersList.documentElement.childNodes
            Set subFld = pFolder.Folders(node.getAttribute("id"))
            ScriptFolder2 subFld
        Next
    End If

    stm.Write "Set curFolder = curFolder.Parent" & vbCrLf
End Sub

Private Function VbStringFormat(pString As String) As String
    Dim strRet As String
    
    strRet = pString
    strRet = Replace(strRet, """", """""")
    strRet = Replace(strRet, vbCrLf, """ & vbCrLf & """)
    strRet = Replace(strRet, vbCr, """ & vbCr & """)
    strRet = Replace(strRet, vbLf, """ & vbLf & """)
    strRet = """" & strRet & """"
    If Left(strRet, 5) = """"" & " Then strRet = Mid(strRet, 6)
    If Right(strRet, 5) = " & """"" Then strRet = Left(strRet, Len(strRet) - 5)
    
    VbStringFormat = strRet
End Function

Private Function VbStringMultiLine(pVarName As String, pString As String) As String
    Dim strRet As String
    Dim arrLines
    Dim strAux As String
    Dim i As Long
    
    strRet = pVarName & " = """""
    VbStringMultiLine = strRet
    strAux = pString
    If strAux = "" Then Exit Function
    strRet = strRet & vbCrLf
    arrLines = Split(strAux, vbCrLf)
   
    For i = 0 To UBound(arrLines)
        strRet = strRet & pVarName & " = " & pVarName & " & " & _
            VbStringFormat(arrLines(i) & "")
        If i < UBound(arrLines) Then strRet = strRet & " & vbCrLf" & vbCrLf
    Next
    
    VbStringMultiLine = strRet
End Function

Private Function VbDateFormat(pDate As Date) As String
    Dim strRet As String
    
    strRet = "#" & Month(pDate) & "/" & Day(pDate) & "/" & Year(pDate)
    strRet = strRet & " " & Hour(pDate) & ":" & Minute(pDate) & ":" & Second(pDate) & "#"
    
    VbDateFormat = strRet
End Function

Private Sub Form_Load()
    chkRecursive.Value = 1
    chkEvents.Value = 1
    chkAsyncEvents.Value = 1
    chkViews.Value = 1
    chkForms.Value = 1
    chkFormsFields.Value = 1
    chkFormsActions.Value = 1
    chkFormsEvents.Value = 1
End Sub

Private Sub ScriptForm2(pForm As Object)
    Dim oField As Object
    Dim oEvn As Object
    
    
    
    stm.Write "newForm.Save" & vbCrLf
    stm.Write vbCrLf
End Sub

Public Sub ScriptForm()
    Dim fso As Object
    Dim node As Object
    Dim oDoc As Object
    
    On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stm = fso.OpenTextFile(App.Path & "\" & App.EXEName & ".vbs", 2, True) ' ForWriting
    
    ScriptForm2 CustomForm
    
    stm.Close
       
    Shell "notepad.exe """ & App.Path & "\" & App.EXEName & ".vbs""", vbNormalFocus
    
    Screen.MousePointer = vbNormal
    Exit Sub
Error:
    Screen.MousePointer = vbNormal
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


