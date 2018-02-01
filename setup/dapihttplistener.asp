<%
' Compatible con G5.x y G7
Const Module = "dapihttplistener"
Const Version = "1.2.51"

Dim oReq, blnAux, strAux, lngAux, arrAux
Dim dSession, strMethod, node
Dim obj, dom, strGuid, destObj, newObj, rcs
Dim IsG7, vErr
Dim accList

Response.Expires = 0

strAux = LCase(Request.ServerVariables("APPL_MD_PATH"))
If Right(strAux, 1) = "/" Then
	strAux = Left(strAux, Len(strAux) - 1)
End If
SetAppVar "VirtualRoot", Mid(strAux, InStr(strAux, "root") + 4)

blnAux = False
If IsObject(Session("dSession")) Then
	blnAux = (TypeName(Session("dSession")) = "Session")
End If

If Not blnAux Then
	Set dSession = Server.CreateObject("doorsapi.Session")

	IsG7 = (CLng(Split(dSession.Version, ".")(0)) >= 7)
	'dSession.TokensAdd "APPVIRTUALROOT", Application("VirtualRoot")

	strAux = ReadIni("Session", "LogDir")
	If strAux <> "" Then
		dSession.LogDir = strAux
	Else
		dSession.LogDir = Server.MapPath(Application("VirtualRoot")) & "\..\log\"
	End If
	
	dSession.CdoConfiguration = ReadIni("Session", "CdoConfiguration")
	dSession.LogErrorsInLogFiles = ReadIniBoolean("Session", "LogErrorsInLogFiles")
	dSession.LogErrorsInEventLog = ReadIniBoolean("Session", "LogErrorsInEventLog")
	dSession.DebugMode = Val(ReadIni("Session", "DebugMode"))
	
	If Not IsG7 Then
		strAux = ReadIni("Session", "MasterConnectionCrypted")
		
		If strAux <> "" Then
			dSession.MasterConnectionCrypted = strAux
		Else
			dSession.MasterConnection = ReadIni("Session", "MasterConnection")
		End If
	End If
	
	' Timeout en segundos para las operaciones largas
	lngAux = Val(ReadIni("ASP", "LongScriptTimeout"))
	If lngAux = 0 Then
		SetAppVar "LongScriptTimeout", 3600
	Else
		SetAppVar "LongScriptTimeout", lngAux
	End If
	
	Set Session("dSession") = dSession
	
Else
	Set dSession = Session("dSession")
	IsG7 = (CLng(Split(dSession.Version, ".")(0)) >= 7)
End If

strGuid = Request.QueryString("disposeguid")
If strGuid = "version" Then
	Response.Write Version
	Response.End
ElseIf strGuid <> "" Then
	Set Session(strGuid) = Nothing
	Session.Contents.Remove strGuid
Else
'TODO: revisar esto
	Set oReq = Server.CreateObject("MSXML2.DOMDocument")
    oReq.preserveWhiteSpace = True
    oReq.setProperty "SelectionLanguage", "XPath"
    oReq.async = False

	oReq.load Request
	
	If oReq.parseError.errorCode <> 0 Then
		Response.Write Module & "<br>"
		Response.Write "Invalid request"
		Response.End
	End If
	
	strMethod = Ucase(oReq.documentElement.getAttribute("method") & "")

	Response.ContentType = "text/xml"
	
	On Error Resume Next
	TryCatch
	vErr = Array(Err.Number, Err.Source, Err.Description)
	On Error Goto 0
	If vErr(0) <> 0 Then ReturnErr vErr
End If

Sub TryCatch
	Select Case strMethod

'--------------------
' Objeto Account
	
		Case "ACCOUNT.ACCOUNTTYPE_GET"
			Set obj = Session(Arg(0))
			Return obj.AccountType

		Case "ACCOUNT.EMAIL_GET"
			Set obj = Session(Arg(0))
			Return obj.Email

		Case "ACCOUNT.ID_GET"
			Set obj = Session(Arg(0))
			Return obj.Id

		Case "ACCOUNT.ISADMIN_GET"
			Set obj = Session(Arg(0))
			Return obj.IsAdmin

		Case "ACCOUNT.NAME_GET"
			Set obj = Session(Arg(0))
			Return obj.Name

		Case "ACCOUNT.SCRIPT"
			Set obj = Session(Arg(0))
			Return ScriptAccount(obj)

		Case "ACCOUNT.SCRIPTPARENTS"
			Set obj = Session(Arg(0))
			Return ScriptAccountParents(obj)

		Case "ACCOUNT.SYSTEM_GET"
			Set obj = Session(Arg(0))
			Return obj.System

'--------------------
' Objeto Application

		Case "APPLICATION.CODELIB"
			Set obj = Session(Arg(0))
			Return obj.CodeLib(Arg(1))

		Case "APPLICATION.CODELIBPARSED"
			Set obj = Session(Arg(0))
			Return obj.ParseCodeIncludes(obj.CodeLib(Arg(1)))

		Case "APPLICATION.FOLDERS"
			Set obj = Session(Arg(0))
			Return obj.Folders(Arg(1))

		Case "APPLICATION.PARSECODEINCLUDES"
			Set obj = Session(Arg(0))
			Return obj.ParseCodeIncludes(Arg(1))

		Case "APPLICATION.ROOTFOLDERID_GET"
			Set obj = Session(Arg(0))
			Return obj.RootFolderId

		Case "APPLICATION.SETTINGSGET"
			Set obj = Session(Arg(0))
			strAux = ArgOpt(2, "")
			If strAux <> "" Then
				Return obj.SettingsGet(Arg(1), strAux)
			Else
				Return obj.SettingsGet(Arg(1))
			End If

		Case "APPLICATION.SETTINGSSET"
			Set obj = Session(Arg(0))
			strAux = ArgOpt(3, "")
			If strAux <> "" Then
				Return obj.SettingsSet(Arg(1), Arg(2), strAux)
			Else
				Return obj.SettingsGet(Arg(1), Arg(2))
			End If

'-------------------
' Objeto AsyncEvent

		Case "ASYNCEVENT.CLASS_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Class

		Case "ASYNCEVENT.CODE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Code

		Case "ASYNCEVENT.CODE_LET"
			Set obj = Session(Arg(0))
			obj.Item("ID=" & Arg(2)).Code = CStr(ArgOpt(1, ""))

		Case "ASYNCEVENT.CODETIMEOUT_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).CodeTimeout

		Case "ASYNCEVENT.DISABLED_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Disabled

		Case "ASYNCEVENT.EVENTTYPE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).EventType

		Case "ASYNCEVENT.ISCOM_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).IsCom

		Case "ASYNCEVENT.LOGIN_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Login

		Case "ASYNCEVENT.METHOD_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Method

		Case "ASYNCEVENT.PASSWORD_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Password

		Case "ASYNCEVENT.RECURSIVE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Recursive

		Case "ASYNCEVENT.SCRIPT"
			Set obj = Session(Arg(0))
			Return ScriptAsyncEvent(obj.Item("ID=" & Arg(1)))

		Case "ASYNCEVENT.TIMERFREQUENCE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).TimerFrequence

		Case "ASYNCEVENT.TIMERMODE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).TimerMode

		Case "ASYNCEVENT.TIMERNEXTRUN_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).TimerNextRun

		Case "ASYNCEVENT.TIMERTIME_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).TimerTime

		Case "ASYNCEVENT.TRIGGEREVENT_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).TriggerEvent

'-------------------		
' Objeto Attachment

		Case "ATTACHMENT.FILESTREAM"
			Set obj = Session(Arg(0)).Item(CStr(Arg(1))).FileStream
			Response.AddHeader "content-length", obj.Size
			Response.ContentType = "application/octet-stream"
			If obj.Size > 0 Then Response.BinaryWrite obj.Read
			Response.Flush

'-------------------
' Objeto CustomForm

		Case "CUSTOMFORM.ACTIONS"
			Set obj = Session(Arg(0))
			obj.Actions.save Response

		Case "CUSTOMFORM.APPLICATION_GET"
			Set obj = Session(Arg(0))
			Return obj.Application

		Case "CUSTOMFORM.CREATED_GET"
			Set obj = Session(Arg(0))
			Return obj.Created

		Case "CUSTOMFORM.DESCRIPTION_GET"
			Set obj = Session(Arg(0))
			Return obj.Description

		Case "CUSTOMFORM.EVENTS"
			Set obj = Session(Arg(0))
			ReturnCollection obj.Events

		Case "CUSTOMFORM.EVENTSLIST"
			Set obj = Session(Arg(0))
			obj.EventsList.save Response

		Case "CUSTOMFORM.FIELDS"
			Set obj = Session(Arg(0))
			ReturnCollection obj.Fields

		Case "CUSTOMFORM.FIELDSLIST"
			Set obj = Session(Arg(0))
			obj.FieldsList.save Response

		Case "CUSTOMFORM.GUID_GET"
			Set obj = Session(Arg(0))
			Return obj.Guid

		Case "CUSTOMFORM.ID_GET"
			Set obj = Session(Arg(0))
			Return obj.Id

		Case "CUSTOMFORM.MODIFIED_GET"
			Set obj = Session(Arg(0))
			Return obj.Modified

		Case "CUSTOMFORM.NAME_GET"
			Set obj = Session(Arg(0))
			Return obj.Name

		Case "CUSTOMFORM.PK_GET"
			Set obj = Session(Arg(0))
			If obj.Properties.Exists("PK") Then
				Return LCase(obj.Properties("PK").Value)
			Else
				Return ""
			End If

		Case "CUSTOMFORM.PROPERTIES"
			Set obj = Session(Arg(0))
			ReturnCollection obj.Properties

		Case "CUSTOMFORM.SAVE"
			Set obj = Session(Arg(0))
			obj.Save

		Case "CUSTOMFORM.SCRIPT"
			Set obj = Session(Arg(0))
			Return ScriptForm(obj)

		Case "CUSTOMFORM.SEARCH"
			Set obj = Session(Arg(0))
			obj.Search(CStr(Arg(1)), CStr(ArgOpt(2, "")), CStr(ArgOpt(3, "")), CStr(ArgOpt(4, "")), CLng(ArgOpt(5, 0)), CBool(ArgOpt(6, True)), CLng(ArgOpt(7, -1))).save Response

		Case "CUSTOMFORM.URLRAW_GET"
			Set obj = Session(Arg(0))
			Return obj.URLRaw

'-----------
' Objeto Db

		Case "DB.DBTYPE"
			Return dSession.Db.DbType

		Case "DB.EXECUTE"
			If IsEmpty(ArgOpt(1, Empty)) And IsEmpty(ArgOpt(2, Empty)) And IsEmpty(ArgOpt(3, Empty)) And _
			   IsEmpty(ArgOpt(4, Empty)) And IsEmpty(ArgOpt(5, Empty)) Then
				Set rcs = dSession.Db.Execute(CStr(Arg(0)), lngAux)
			Else
				arrAux = Array(ArgOpt(1, Empty), ArgOpt(2, Empty), ArgOpt(3, Empty), ArgOpt(4, Empty), ArgOpt(5, Empty))
				Set rcs = dSession.Db.Execute(CStr(Arg(0)), lngAux, arrAux)
			End If
			Session("RecordsAffected") = lngAux
			If rcs.State = 1 Then
				rcs.save Response, 1
			Else
				Return Nothing
			End If

		Case "DB.EXECUTERECORDSAFFECTED"
			Return Session("RecordsAffected")

		Case "DB.OPENRECORDSET"
			If IsEmpty(ArgOpt(1, Empty)) And IsEmpty(ArgOpt(2, Empty)) And IsEmpty(ArgOpt(3, Empty)) And _
			   IsEmpty(ArgOpt(4, Empty)) And IsEmpty(ArgOpt(5, Empty)) Then
				dSession.Db.OpenRecordset(CStr(Arg(0))).save Response, 1
			Else
				arrAux = Array(ArgOpt(1, Empty), ArgOpt(2, Empty), ArgOpt(3, Empty), ArgOpt(4, Empty), ArgOpt(5, Empty))
				dSession.Db.OpenRecordset(CStr(Arg(0)), arrAux).save Response, 1
			End If

'-----------------
' Objeto Directory

		Case "DIRECTORY.ACCOUNTS"
			Return dSession.Directory.Accounts(Arg(0))

		Case "DIRECTORY.ACCOUNTSSEARCH"
			Set dom = dSession.Directory.AccountsSearch(CStr(ArgOpt(0, "")), CStr(ArgOpt(1, "")))
			dom.save Response
			
'-----------------
' Objeto Document

		Case "DOCUMENT.ATTACHMENTS"
			Set obj = Session(Arg(0))
			ReturnCollection obj.Attachments

		Case "DOCUMENT.DELETE"
			Set obj = Session(Arg(0))
			obj.Delete

		Case "DOCUMENT.FIELDS"
			Set obj = Session(Arg(0))
			ReturnCollection obj.Fields

		Case "DOCUMENT.FORM"
			Set obj = Session(Arg(0))
			Return obj.Form

		Case "DOCUMENT.ID_GET"
			Set obj = Session(Arg(0))
			Return obj.Id
			
		Case "DOCUMENT.ISNEW_GET"
			Set obj = Session(Arg(0))
			Return obj.IsNew

		Case "DOCUMENT.PARENT"
			Set obj = Session(Arg(0))
			Return obj.Parent

		Case "DOCUMENT.PROPERTIES"
			Set obj = Session(Arg(0))
			ReturnCollection obj.Properties

		Case "DOCUMENT.SAVE"
			Set obj = Session(Arg(0))
			obj.Save

		Case "DOCUMENT.SCRIPT"
			Set obj = Session(Arg(0))
			Return ScriptDocument(obj)

		Case "DOCUMENT.SUBJECT_GET"
			Set obj = Session(Arg(0))
			Return obj.Subject
			
		Case "DOCUMENT.SUBJECT_LET"
			Set obj = Session(Arg(0))
			obj.Subject = CStr(Arg(1))
			
		Case "DOCUMENT.TAGS_GET"
			Set obj = Session(Arg(0))
			Return obj.Tags(Arg(1)).value

		Case "DOCUMENT.TAGS_LET"
			Set obj = Session(Arg(0))
			obj.Tags(Arg(1)).value = Arg(2)

'--------------
' Objeto Field

		Case "FIELD.COMPUTED_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).Computed
				
		Case "FIELD.CUSTOM_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).Custom
				
		Case "FIELD.DATALENGTH_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).DataLength

		Case "FIELD.DATAPRECISION_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).DataPrecision

		Case "FIELD.DATASCALE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).DataScale
				
		Case "FIELD.DATATYPE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).DataType

		Case "FIELD.DESCRIPTION_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).Description

		Case "FIELD.DESCRIPTIONRAW_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).DescriptionRaw

		Case "FIELD.NULLABLE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).Nullable
				
		Case "FIELD.PROPERTIES"
			Set obj = Session(Arg(0))
			ReturnCollection obj.Item(CStr(Arg(1))).Properties

		Case "FIELD.VALUE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).Value

		Case "FIELD.VALUE_LET"
			Set obj = Session(Arg(0))
			'dSession.DebugPrint Arg(2) & "=" & Arg(1)
			obj.Item(CStr(Arg(2))).Value = Arg(1)

'---------------
' Objeto Folder

		Case "FOLDER.ACL"
			Set obj = Session(Arg(0))
			obj.Acl.save Response

		Case "FOLDER.ACLOWN"
			Set obj = Session(Arg(0))
			obj.AclOwn.save Response

		Case "FOLDER.ACLINHERITED"
			Set obj = Session(Arg(0))
			obj.AclInherited.save Response

		Case "FOLDER.ANCESTORS"
			Set obj = Session(Arg(0))
			obj.Ancestors.save Response

		Case "FOLDER.APP"
			Set obj = Session(Arg(0))
			Return obj.App

		Case "FOLDER.ASYNCEVENTS"
			Set obj = Session(Arg(0))
			ReturnCollection obj.AsyncEvents

		Case "FOLDER.ASYNCEVENTSLIST"
			Set obj = Session(Arg(0))
			obj.AsyncEventsList.save Response

		Case "FOLDER.CHARDATA_GET"
			Set obj = Session(Arg(0))
			Return obj.CharData

		Case "FOLDER.CHARDATA_LET"
			Set obj = Session(Arg(0))
			obj.CharData = CStr(Arg(1))

		Case "FOLDER.COMMENTS_GET"
			Set obj = Session(Arg(0))
			Return obj.Comments

		Case "FOLDER.COMMENTS_LET"
			Set obj = Session(Arg(0))
			obj.Comments = CStr(Arg(1))

		Case "FOLDER.COPY"
			Set obj = Session(Arg(0))
			If IsNumeric(Arg(1)) Then
				Set destObj = dSession.FoldersGetFromId(CLng(Arg(1)))
			Else
				Set destObj = Session(CStr(Arg(1)))
			End If
			Set newObj = obj.Copy(destObj, CStr(ArgOpt(2, "")))
			Return newObj

		Case "FOLDER.CREATED_GET"
			Set obj = Session(Arg(0))
			Return obj.Created

		Case "FOLDER.DELETE"
			Set obj = Session(Arg(0))
			obj.Delete

		Case "FOLDER.DESCENDANTS"
			Set obj = Session(Arg(0))
			obj.Descendants.save Response

		Case "FOLDER.DESCRIPTION_GET"
			Set obj = Session(Arg(0))
			Return obj.Description

		Case "FOLDER.DESCRIPTIONRAW_GET"
			Set obj = Session(Arg(0))
			Return obj.DescriptionRaw

		Case "FOLDER.DESCRIPTIONRAW_LET"
			Set obj = Session(Arg(0))
			obj.DescriptionRaw = CStr(Arg(1))

		Case "FOLDER.DOCUMENTS"
			Set obj = Session(Arg(0))
			Return obj.Documents(Arg(1))

		Case "FOLDER.DOCUMENTSCOUNT"
			Set obj = Session(Arg(0))
			Return obj.DocumentsCount

		Case "FOLDER.DOCUMENTSDELETE"
			Set obj = Session(Arg(0))
			Return obj.DocumentsDelete(CStr(ArgOpt(1, "")))

		Case "FOLDER.DOCUMENTSNEW"
			Set obj = Session(Arg(0))
			Return obj.DocumentsNew

		Case "FOLDER.EVENTS"
			Set obj = Session(Arg(0))
			ReturnCollection obj.Events

		Case "FOLDER.EVENTSLIST"
			Set obj = Session(Arg(0))
			obj.EventsList.save Response

		Case "FOLDER.HREFRAW_GET"
			Set obj = Session(Arg(0))
			Return obj.HrefRaw

		Case "FOLDER.HREFRAW_LET"
			Set obj = Session(Arg(0))
			obj.HrefRaw = CStr(Arg(1))

		Case "FOLDER.ICONRAW_GET"
			Set obj = Session(Arg(0))
			Return obj.IconRaw

		Case "FOLDER.ICONRAW_LET"
			Set obj = Session(Arg(0))
			obj.IconRaw = CStr(Arg(1))

		Case "FOLDER.ID_GET"
			Set obj = Session(Arg(0))
			Return obj.Id

		Case "FOLDER.ISNEW_GET"
			Set obj = Session(Arg(0))
			Return obj.IsNew

		Case "FOLDER.FOLDERS"
			Set obj = Session(Arg(0))
			Return obj.Folders(Arg(1))

		Case "FOLDER.FOLDERSLIST"
			Set obj = Session(Arg(0))
			obj.FoldersList.save Response

		Case "FOLDER.FOLDERTYPE_GET"
			Set obj = Session(Arg(0))
			Return obj.FolderType

		Case "FOLDER.FORMID_GET"
			Set obj = Session(Arg(0))
			Return obj.FormId

		Case "FOLDER.FORM"
			Set obj = Session(Arg(0))
			Return obj.Form

		Case "FOLDER.MODIFIED_GET"
			Set obj = Session(Arg(0))
			Return obj.Modified

		Case "FOLDER.NAME_GET"
			Set obj = Session(Arg(0))
			Return obj.Name

		Case "FOLDER.LOGCONF"
			Set obj = Session(Arg(0))
			obj.LogConf.save Response

		Case "FOLDER.PARENT"
			Set obj = Session(Arg(0))
			Return obj.Parent

		Case "FOLDER.PATH_GET"
			Set obj = Session(Arg(0))
			'TODO: Cdo se arregle Folder.Path reemplazar con la siguiente linea
			'Return obj.Path(CLng(Arg(1)))
			Return FolderPath(obj, Arg(1))

		Case "FOLDER.SAVE"
			Set obj = Session(Arg(0))
			obj.Save

		Case "FOLDER.SCRIPT"
			Set obj = Session(Arg(0))
			Return ScriptFolder(obj)

		Case "FOLDER.SCRIPTACL"
			Set obj = Session(Arg(0))
			Return ScriptFolderAcl(obj)

		Case "FOLDER.SEARCH"
			Set obj = Session(Arg(0))
			Set dom = obj.Search(CStr(ArgOpt(1, "")), CStr(ArgOpt(2, "")), CStr(ArgOpt(3, "")), CLng(ArgOpt(4, 0)), CBool(ArgOpt(5, False)))
			'dom.save "C:\Program Files (x86)\Gestar\log\dom.xml"
			dom.save Response

		Case "FOLDER.SEARCHGROUPS"
			Set obj = Session(Arg(0))
			Set dom = obj.SearchGroups(CStr(ArgOpt(1, "")), CStr(ArgOpt(2, "count(*) as TOTAL")), CStr(ArgOpt(3, "")), CStr(ArgOpt(4, "")), CLng(ArgOpt(5, 0)), CBool(ArgOpt(6, False)))
			'dom.save "C:\Program Files (x86)\Gestar\log\dom.xml"
			dom.save Response

		Case "FOLDER.SYSTEM_GET"
			Set obj = Session(Arg(0))
			Return obj.System

		Case "FOLDER.TAGS_GET"
			Set obj = Session(Arg(0))
			Return obj.Tags(Arg(1)).value

		Case "FOLDER.TAGS_LET"
			Set obj = Session(Arg(0))
			obj.Tags(Arg(1)).value = Arg(2)

		Case "FOLDER.TARGET_GET"
			Set obj = Session(Arg(0))
			Return obj.Target

		Case "FOLDER.TARGET_LET"
			Set obj = Session(Arg(0))
			obj.Target = CStr(Arg(1))

		Case "FOLDER.USERPROPERTIES"
			Set obj = Session(Arg(0))
			ReturnCollection obj.UserProperties

		Case "FOLDER.VIEWS"
			Set obj = Session(Arg(0))
			Return obj.Views(Arg(1))

		Case "FOLDER.VIEWSLIST"
			Set obj = Session(Arg(0))
			obj.ViewsList.save Response

'--------------------
' Objeto FolderEvent

		Case "FOLDEREVENT.CODE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Code

		Case "FOLDEREVENT.CODE_LET"
			Set obj = Session(Arg(0))
			obj.Item("ID=" & Arg(2)).Code = CStr(ArgOpt(1, ""))

		Case "FOLDEREVENT.OVERRIDES_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Overrides

		Case "FOLDEREVENT.SCRIPT"
			Set obj = Session(Arg(0))
			Return ScriptFolderEvent(obj.Item("ID=" & Arg(1)))

'------------------
' Objeto FormEvent

		Case "FORMEVENT.CODE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Code

		Case "FORMEVENT.CODE_LET"
			Set obj = Session(Arg(0))
			obj.Item("ID=" & Arg(2)).Code = CStr(ArgOpt(1, ""))

		Case "FORMEVENT.EXTENSIBLE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Extensible

		Case "FORMEVENT.OVERRIDABLE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item("ID=" & Arg(1)).Overridable
		
		Case "FORMEVENT.SCRIPT"
			Set obj = Session(Arg(0))
			Return ScriptFormEvent(obj.Item("ID=" & Arg(1)))

'-----------------
' Objeto MasterDb

		Case "MASTERDB.OPENRECORDSET"
			If IsEmpty(ArgOpt(1, Empty)) And IsEmpty(ArgOpt(2, Empty)) And IsEmpty(ArgOpt(3, Empty)) And _
			   IsEmpty(ArgOpt(4, Empty)) And IsEmpty(ArgOpt(5, Empty)) Then
				dSession.MasterDb.OpenRecordset(CStr(Arg(0))).save Response, 1
			Else
				arrAux = Array(ArgOpt(1, Empty), ArgOpt(2, Empty), ArgOpt(3, Empty), ArgOpt(4, Empty), ArgOpt(5, Empty))
				dSession.MasterDb.OpenRecordset(CStr(Arg(0)), arrAux).save Response, 1
			End If

		Case "MASTERDB.DBTYPE"
			Return dSession.MasterDb.DbType

'-----------------
' Objeto Property

		Case "PROPERTY.VALUE_GET"
			Set obj = Session(Arg(0))
			Return obj.Item(CStr(Arg(1))).Value

		Case "PROPERTY.VALUE_LET"
			Set obj = Session(Arg(0))
			'dSession.DebugPrint Arg(2) & "=" & Arg(1)
			obj.Item(CStr(Arg(2))).Value = Arg(1)

'---------------------------
' Objeto PropertyCollection

		Case "PROPERTYCOLLECTION.ADD"
			Set obj = Session(Arg(0))
			obj.Add CStr(Arg(1))

'----------------
' Objeto Session

		Case "SESSION.CLEARALLCUSTOMCACHE"
			dSession.ClearAllCustomCache

		Case "SESSION.CLEAROBJECTMODELCACHE"
			dSession.ClearObjectModelCache CStr(Arg(0))

		Case "SESSION.DOCUMENTSGETFROMID"
			Set obj = dSession.DocumentsGetFromId(CLng(Arg(0)))
			Return obj

		Case "SESSION.FOLDERSGETFROMID"
			Set obj = dSession.FoldersGetFromId(CLng(Arg(0)))
			Return obj

		Case "SESSION.FOLDERSLIST"
			dSession.FoldersList.save Response

		Case "SESSION.FOLDERSTREE"
			dSession.FoldersTree.save Response

		Case "SESSION.FORMSLIST"
			dSession.FormsList.save Response

		Case "SESSION.FORMS"
			Return dSession.Forms(Arg(0))

		Case "SESSION.INSTANCEGUID"
			Return dSession.InstanceGuid
	
		Case "SESSION.INSTANCEID"
			Return dSession.InstanceId

		Case "SESSION.INSTANCELIST"
			dSession.InstanceList.save Response
	
		Case "SESSION.INSTANCENAME"
			Return dSession.InstanceName
	
		Case "SESSION.ISLOGGED"
			Return CLng(dSession.IsLogged)
	
		Case "SESSION.LOGGEDUSER"
			Return dSession.LoggedUser

		Case "SESSION.LOGOFF"
			dSession.Logoff
	
		Case "SESSION.LOGON"
			dSession.Logon CStr(Arg(0)), CStr(Arg(1)), Arg(2), Arg(3)
			Session.LCID = dSession.LoggedUser.Language

		Case "SESSION.VERSION"
			Return dSession.Version
	
		Case "SESSION.WINLOGON"
			If Request.ServerVariables("LOGON_USER") = "" Then
				Response.Status = "401 Access denied"
				Response.End
			End If
			
			If Arg(0) & "" = "" Then
				blnAux = dSession.WinLogon(, obj, CBool(Arg(1)))
			Else
				blnAux = dSession.WinLogon(CStr(Arg(0)), obj, CBool(Arg(1)))
			End If
			If Not blnAux Then
				obj.save Response
			Else
				Session.LCID = dSession.LoggedUser.Language
			End If

'-------------
' Objeto User

		Case "USER.ID_GET"
			Set obj = Session(Arg(0))
			Return obj.Id

		Case "USER.ISADMIN_GET"
			Set obj = Session(Arg(0))
			Return obj.IsAdmin

		Case "USER.ISNEW_GET"
			Set obj = Session(Arg(0))
			Return obj.IsNew
			
		Case "USER.LANGUAGE_GET"
			Set obj = Session(Arg(0))
			Return obj.Language

		Case "USER.LOGIN_GET"
			Set obj = Session(Arg(0))
			Return obj.Login

		Case "USER.NAME_GET"
			Set obj = Session(Arg(0))
			Return obj.Name

'-------------
' Objeto View

		Case "VIEW.COMMENTS_GET"
			Set obj = Session(Arg(0))
			Return obj.Comments

		Case "VIEW.COPY"
			Set obj = Session(Arg(0))
			Set destObj = Session(CStr(Arg(1)))
			Set newObj = obj.Copy((destObj), CBool(ArgOpt(2, obj.PrivateView)))
			Return newObj

		Case "VIEW.DEFINITION"
			Set obj = Session(Arg(0))
			obj.Definition.save Response

		Case "VIEW.DELETE"
			Set obj = Session(Arg(0))
			obj.Delete

		Case "VIEW.DESCRIPTIONRAW_GET"
			Set obj = Session(Arg(0))
			Return obj.DescriptionRaw

		Case "VIEW.ID_GET"
			Set obj = Session(Arg(0))
			Return obj.Id

		Case "VIEW.NAME_GET"
			Set obj = Session(Arg(0))
			Return obj.Name

		Case "VIEW.SCRIPT"
			Set obj = Session(Arg(0))
			Return ScriptView(obj)

'--------------
' CustomScript

		Case "CUSTOMSCRIPT"
			strAux = CStr(Arg(0))
			If strAux <> "" Then
				Execute strAux
			End If
	
		Case Else
			Err.Raise vbObjectError + 1, Module, "Invalid method: " & strMethod
	End Select
End Sub

' Asigna una variable de aplicacion solo si el valor es distinto
Sub SetAppVar(ByRef Variable, ByRef Value)
    If Application(Variable) <> Value Then
        Application(Variable) = Value
    End If
End Sub

Function ReadIni(ByRef sApplication, ByRef sKey)
    Dim IniFile, strAux
	IniFile = Server.MapPath(Application("VirtualRoot")) & "\..\bin\doors.ini"
	strAux = dSession.ReadIni(CStr(IniFile), CStr(sApplication), CStr(sKey))
	ReadIni = strAux
End Function

Function ReadIniBoolean(ByRef sApplication, ByRef sKey)
	If ReadIni(sApplication, sKey) = "1" Then
		ReadIniBoolean = True
	Else
		ReadIniBoolean = False
	End If	
End Function

' Funcion Val de VBA
Function Val(ByRef pString)
	Val = dSession.Dispatch("Val", pString & "")
End Function

Function Arg(ByRef ArgIndex)
	Dim oNode
	
	Set oNode = oReq.documentElement.childNodes(ArgIndex)
	Arg = VbTypeDecode(oNode.text, oNode.getAttribute("type"))
End Function

Function ArgOpt(ByRef ArgIndex, ByRef DefValue)
	Dim oNode
	
	Set oNode = oReq.documentElement.childNodes(ArgIndex)
	If oNode Is Nothing Then
		ArgOpt = DefValue
	Else
		ArgOpt = Arg(ArgIndex)
	End If
End Function

Sub Return(ByRef Value)
	Dim oDom, oNode, i, lType
	
	Set oDom = dSession.XML.NewDom
	Set oDom.documentElement = oDom.createNode("element", "root", "")
	Set oNode = oDom.documentElement
	
	If IsArray(Value) Then
		For i = 0 To UBound(Values)
	        oNode.appendChild TypedNode(oDom, Value(i))
		Next
	Else
        oNode.appendChild TypedNode(oDom, Value)
	End If

	oDom.save Response
End Sub

Function TypedNode(ByRef pDom, ByRef pValue)
	Dim oNode, strGuid
	
	Set oNode = pDom.createNode("element", "item", "")
	If IsObject(pValue) Then
		If pValue Is Nothing Then
			oNode.setAttribute "type", "Nothing"
		Else
			strGuid = dSession.CreateGuid
			Set Session(strGuid) = pValue
			oNode.setAttribute "type", "Object"
			oNode.text = strGuid
		End If
	Else
		oNode.setAttribute "type", TypeName(pValue)
		oNode.text = VbTypeEncode(pValue)
	End If
	Set TypedNode = oNode
End Function

Function VbTypeEncode(ByRef pValue)
	Dim sType
	
	sType = TypeName(pValue)
	If sType = "Empty" Or sType = "Null" Then
		VbTypeEncode = ""
	ElseIf sType = "String" Then
		VbTypeEncode = pValue
	ElseIf sType = "Date" Then
		VbTypeEncode = dSession.XML.XMLEncode(pValue, 2)
	ElseIf sType = "Byte" Or sType = "Integer" Or sType = "Long" Or _
			sType = "Single" Or sType = "Double" Or sType = "Currency" Or _
			sType = "Decimal" Then
		VbTypeEncode = dSession.XML.XMLEncode(pValue, 3)
	ElseIf sType = "Boolean" Then
		VbTypeEncode = IIf(pValue, "1", "0")
	Else
		Err.Raise vbObjectError + 1, Module, "Not serializable: " & sType
	End If
End Function

Function VbTypeDecode(ByRef pValue, ByRef pType)
	If pType = "Empty" Then
		VbTypeDecode = Empty
	ElseIf pType = "Null" Then
		VbTypeDecode = Null
	ElseIf pType = "String" Or pType & "" = "" Then
		VbTypeDecode = pValue & ""
	ElseIf pType = "Date" Then
		VbTypeDecode = dSession.XML.XMLDecode(CStr(pValue), 2)
	ElseIf pType = "Byte" Or pType = "Integer" Or pType = "Long" Or _
			pType = "Single" Or pType = "Double" Or pType = "Currency" Or _
			pType = "Decimal" Then
		VbTypeDecode = dSession.XML.XMLDecode(CStr(pValue), 3)
	ElseIf pType = "Boolean" Then
		VbTypeDecode = IIf(pValue & "" = "1", True, False)
	Else
		Err.Raise vbObjectError + 1, Module, "Not serializable: " & pType
	End If
End Function

Sub ReturnErr(ByRef pErr)
	Dim oDom, oNode
	
	Set oDom = dSession.XML.NewDom
	Set oDom.documentElement = oDom.createNode("element", "root", "")
	Set oNode = oDom.documentElement
	With oNode
		.setAttribute "type", "ErrObject"
		.setAttribute "number", pErr(0)
		.setAttribute "source", pErr(1)
		.setAttribute "description", pErr(2)
	End With
	oDom.save Response
End Sub

Sub ReturnCollection(ByRef Obj)
	Dim strGuid
	Dim oDom, oNode, oNode2, it
	Dim blnEventCol

	blnEventCol = (TypeName(obj) = "FormEventCollection" Or TypeName(obj) = "FolderEventCollection" Or TypeName(obj) = "AsyncEventCollection")
	
	Set oDom = dSession.XML.NewDom
	Set oNode = oDom.createNode("element", "root", "")
	Set oDom.documentElement = oNode

	If Not Obj Is Nothing Then
		strGuid = dSession.CreateGuid
		Set Session(strGuid) = Obj
        oNode.setAttribute "type", "Collection"
        oNode.setAttribute "guid", strGuid
        
        For Each it In obj
        	Set oNode2 = oDom.createNode("element", "item", "")
        	If blnEventCol Then
        		oNode2.setAttribute "id", it.Id
        	Else
        		oNode2.setAttribute "name", it.Name
        	End If
        	oNode.appendChild oNode2
        Next
	Else
        oNode.setAttribute "type", "Nothing"
	End If
	oDom.save Response
End Sub

Function IIf(ByRef Expression, ByRef TruePart, ByRef FalsePart)
	If Expression Then
		IIf = TruePart
	Else
		IIf = FalsePart
	End If
End Function

Function ScriptObjects(ByRef pSelection)
	Dim oldTO, sb, vErr

	oldTO = Server.ScriptTimeout
	Server.ScriptTimeout = 3600 ' 1 hora
	dSession.Dispatch "SyncEventsDisabled", True
	Set sb = dSession.ConstructNewStringBuilder
	
	On Error Resume Next
	ScriptObjects2 pSelection, sb
	vErr = Array(Err.Number, Err.Source, Err.Description)
	On Error Goto 0

	Server.ScriptTimeout = oldTO
	dSession.Dispatch "SyncEventsDisabled", False
	
	If vErr(0) = 0 Then
		ScriptObjects = sb.ToString
	Else
		Err.Raise vErr(0), vErr(1), vErr(2)
	End If
End Function

Sub ScriptObjects2(ByRef pSelection, ByRef pSb)
	Dim domSel, defPwd
	Dim obj, vErr, fld
	Dim node, childNode
	Dim SyncEventName(12)
	Dim rootFolder, sPath, sFormula, sParentPath
	
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

	Set domSel = dSession.Xml.NewDom
	domSel.loadXml pSelection

	pSb.Append "'----------------------------------------" & vbCrLf
	pSb.Append "' Archivo de instalacion de Cloudy Doors" & vbCrLf
	pSb.Append "' Creado: " & Now & vbCrLf
	pSb.Append "'----------------------------------------" & vbCrLf
	pSb.Append vbCrLf
	pSb.Append "' Descomentar esta linea para desactivar la ejecucion de eventos sincronos" & vbCrLf
	pSb.Append "' mientras se ejecuta el instalador" & vbCrLf
	pSb.Append "' dSession.Dispatch ""SyncEventsDisabled"", True" & vbCrLf
	pSb.Append vbCrLf
	pSb.Append "IsG7 = (CLng(Split(dSession.Version, ""."")(0)) >= 7)" & vbCrLf
	pSb.Append vbCrLf

	'----------
	' Accounts
	
	defPwd = False

	For Each node In domSel.selectNodes("/root/system/directory/item")
		On Error Resume Next
		Set obj = dSession.Directory.Accounts(node.getAttribute("name"))
		vErr = Array(Err.Number, Err.Description)
		On Error Goto 0
		
		If vErr(0) <> 0 Then
			pSb.Append vbCrLf & "' Error instantiating account " & UCase(node.getAttribute("name")) & ": " & WithoutCrLf(vErr(1)) & vbCrLf
		Else
			If obj.AccountType = 1 Or obj.AccountType = 2 Then
				If Not defPwd Then
					pSb.Append "' Default password for new users" & vbCrLf
					pSb.Append "DefaultPassword = ""12345""" & vbCrLf
					pSb.Append vbCrLf
					defPwd = True
				End If
			
				pSb.Append vbCrLf
				pSb.Append "'---" & vbCrLf
				pSb.Append "'--- Scripts for account " & UCase(obj.Name) & vbCrLf
				pSb.Append "'---" & vbCrLf
				pSb.Append vbCrLf
				pSb.Append ScriptAccount(obj) & vbCrLf
			End If
		End If
	Next
		
    ' Una 2da pasada para las relaciones
	For Each node In domSel.selectNodes("/root/system/directory/item")
		On Error Resume Next
		Set obj = dSession.Directory.Accounts(node.getAttribute("name"))
		vErr = Array(Err.Number, Err.Description)
		On Error Goto 0

		If vErr(0) <> 0 Then
			pSb.Append vbCrLf & "' Error instantiating account " & UCase(node.getAttribute("name")) & ": " & WithoutCrLf(vErr(1)) & vbCrLf
		Else
            If obj.AccountType = 1 Or obj.AccountType = 2 Then
                pSb.Append vbCrLf
                pSb.Append "'---" & vbCrLf
                pSb.Append "'--- Scripts for " & UCase(obj.Name) & " parents" & vbCrLf
                pSb.Append "'---" & vbCrLf
				pSb.Append vbCrLf
                pSb.Append ScriptAccountParents(obj) & vbCrLf
            End If
		End If
    Next

	'-------
	' Forms
	
	For Each node In domSel.selectNodes("/root/system/forms/item")
		On Error Resume Next
		Set obj = dSession.Forms(node.getAttribute("guid"))
		vErr = Array(Err.Number, Err.Description)
		On Error Goto 0

		If vErr(0) <> 0 Then
			pSb.Append vbCrLf & "' Error instantiating form " & UCase(node.getAttribute("name")) & ": " & WithoutCrLf(vErr(1)) & vbCrLf
		Else
			If obj.Id <> 0 Then
				pSb.Append vbCrLf
				pSb.Append "'---" & vbCrLf
				pSb.Append "'--- Scripts for form " & UCase(obj.Name) & vbCrLf
				pSb.Append "'---" & vbCrLf
				pSb.Append vbCrLf
				pSb.Append ScriptForm(obj) & vbCrLf
				
				For Each childNode In node.selectNodes("syncevents/item")
					pSb.Append vbCrLf
					pSb.Append "'--- Event " & SyncEventName(childNode.getAttribute("id")) & vbCrLf
					pSb.Append vbCrLf
					pSb.Append ScriptFormEvent(obj.Events("ID=" & childNode.getAttribute("id"))) & vbCrLf
				Next
			End If
		End If
	Next

	'-----------------
	' Global Codelibs
	
	Set fld = Nothing
	For Each node In domSel.selectNodes("/root/system/codelib/documents/item")
		If fld Is Nothing Then
			Set fld = dSession.FoldersGetFromId(11)
			pSb.Append vbCrLf
			pSb.Append "'---" & vbCrLf
			pSb.Append "'--- Scripts for Global Codelibs" & vbCrLf
			pSb.Append "'---" & vbCrLf
			pSb.Append vbCrLf
			pSb.Append "Set curFolder = dSession.FoldersGetFromId(11)" & vbCrLf
			pSb.Append vbCrLf
		End If

		On Error Resume Next
		Set obj = fld.Documents("NAME = '" & node.getAttribute("name") & "'")
		vErr = Array(Err.Number, Err.Description)
		On Error Goto 0

		If vErr(0) <> 0 Then
			pSb.Append vbCrLf & "' Error instantiating document " & UCase(node.getAttribute("name")) & ": " & WithoutCrLf(vErr(1)) & vbCrLf
		Else
			pSb.Append "'--- Scripts for codelib " & UCase(obj("NAME").Value) & vbCrLf
			pSb.Append vbCrLf
			pSb.Append ScriptDocument(obj) & vbCrLf
			pSb.Append vbCrLf
		End If
	Next
	
	'-----------------
	' Folders

	pSb.Append "' Root folder" & vbCrLf
	pSb.Append "Set rootFolder = dSession.FoldersGetFromId(1001)" & vbCrLf
	pSb.Append vbCrLf
	Set rootFolder = dSession.FoldersGetFromId(1001)
	
	For Each node In domSel.selectNodes("/root/folder[@name='PublicFolders']//folder")
		sPath = FolderNodePath(node)
		On Error Resume Next
		Set fld = rootFolder.App.Folders(CStr(sPath))
		vErr = Array(Err.Number, Err.Description)
		On Error Goto 0

		If vErr(0) <> 0 Then
			pSb.Append vbCrLf & "' Error instantiating folder " & sPath & ": " & WithoutCrLf(vErr(1)) & vbCrLf
		Else
			pSb.Append vbCrLf
			pSb.Append "'---" & vbCrLf
			pSb.Append "'--- Scripts for folder " & sPath & vbCrLf
			pSb.Append "'---" & vbCrLf
			
			' Folder
			If node.getAttribute("checked") & "" = "1" Then
				pSb.Append vbCrLf
				pSb.Append "Set curFolder = rootFolder"
				sParentPath = FolderNodePath(node.parentNode)
				If sParentPath <> "" Then
					pSb.Append ".App.Folders(""" & sParentPath & """)"
				End If
				pSb.Append vbCrLf & vbCrLf
				pSb.Append ScriptFolder(fld) & vbCrLf
			End If
			
			pSb.Append vbCrLf
			pSb.Append "Set curFolder = rootFolder.App.Folders(""" & sPath & """)" & vbCrLf
			pSb.Append vbCrLf
			
			' Acl
			If node.getAttribute("acl") & "" = "1" Then
				pSb.Append vbCrLf
				pSb.Append "'--- Folder Acl" & vbCrLf
				pSb.Append vbCrLf
				pSb.Append ScriptFolderAcl(fld) & vbCrLf
			End If			
			
			' SyncEvents
			For Each childNode In node.selectNodes("syncevents/item")
				pSb.Append vbCrLf
				pSb.Append "'--- Event " & SyncEventName(childNode.getAttribute("id")) & vbCrLf
				pSb.Append vbCrLf
				pSb.Append ScriptFolderEvent(fld.Events("ID=" & childNode.getAttribute("id"))) & vbCrLf
			Next

			' AsyncEvents
			For Each childNode In node.selectNodes("asyncevents/item")
				On Error Resume Next
				Set obj = fld.AsyncEvents("ID=" & childNode.getAttribute("id"))
				aux = obj.EventType
				vErr = Array(Err.Number, Err.Description)
				On Error Goto 0

				If vErr(0) <> 0 Then
					pSb.Append vbCrLf & "' Error instantiating async event " & childNode.getAttribute("id") & ": " & WithoutCrLf(vErr(1)) & vbCrLf
				Else
					pSb.Append vbCrLf
					pSb.Append "'--- Async Event " & childNode.getAttribute("id") & vbCrLf
					pSb.Append vbCrLf
					pSb.Append ScriptAsyncEvent(obj) & vbCrLf
				End If
			Next

			' Documents
			For Each childNode In node.selectNodes("documents/item")
				On Error Resume Next
				sFormula = PKFormula(fld.Form, childNode)
				Set obj = fld.Documents(sFormula)
				vErr = Array(Err.Number, Err.Description)
				On Error Goto 0

				If vErr(0) <> 0 Then
					pSb.Append vbCrLf & "' Error instantiating document " & sFormula & ": " & WithoutCrLf(vErr(1)) & vbCrLf
				Else
					pSb.Append vbCrLf
					pSb.Append "'--- Document " & sFormula & vbCrLf
					pSb.Append vbCrLf
					pSb.Append ScriptDocument(obj) & vbCrLf
				End If
			Next

			' Views
			For Each childNode In node.selectNodes("views/item")
				On Error Resume Next
				Set obj = fld.Views(childNode.getAttribute("name"))
				vErr = Array(Err.Number, Err.Description)
				On Error Goto 0

				If vErr(0) <> 0 Then
					pSb.Append vbCrLf & "' Error instantiating view " & childNode.getAttribute("name") & ": " & WithoutCrLf(vErr(1)) & vbCrLf
				Else
					pSb.Append vbCrLf
					pSb.Append "'--- View " & UCase(obj.Name) & vbCrLf
					pSb.Append vbCrLf
					pSb.Append ScriptView(obj) & vbCrLf
				End If
			Next
		End If
	Next

	pSb.Append vbCrLf
	pSb.Append "dSession.Dispatch ""SyncEventsDisabled"", False" & vbCrLf
	pSb.Append vbCrLf
	sb.Append "If IsG7 Then" & vbCrLf
	pSb.Append vbTab & "dSession.ClearAllCustomCache" & vbCrLf
	pSb.Append vbTab & "dSession.ClearObjectModelCache ""ComCodeLibCache""" & vbCrLf
	sb.Append "End If" &  vbCrLf
	pSb.Append vbCrLf
	pSb.Append vbCrLf
	pSb.Append "Dim domAccounts" & vbCrLf
	pSb.Append "Function AccountId(pName)" & vbCrLf
	pSb.Append vbTab & "If IsEmpty(domAccounts) Then Set domAccounts = dSession.Directory.AccountsList" & vbCrLf
	pSb.Append vbTab & "AccountId = domAccounts.selectSingleNode(""/d:root/d:item[@name='"" & pName & ""']"").getAttribute(""id"")" & vbCrLf
	pSb.Append "End Function" & vbCrLf
	pSb.Append vbCrLf	
	pSb.Append vbCrLf
	pSb.Append "'--------------------------------------" & vbCrLf
	pSb.Append "' Fin del archivo de instalacion" & vbCrLf
	pSb.Append "' Finalizado: " & Now & vbCrLf
	pSb.Append "'--------------------------------------" & vbCrLf
End Sub

Function ScriptFolder(pFolder)
	Dim sAux
	
	Set sb = dSession.ConstructNewStringBuilder
	
    sb.Append "On Error Resume Next" & vbCrLf
    sb.Append "Set newFolder = curFolder.Folders(" & VbStringFormat(pFolder.Name) & ")" & vbCrLf
    sb.Append "ErrNumber = Err.Number" & vbCrLf
    sb.Append "On Error GoTo 0" & vbCrLf
    sb.Append "If ErrNumber <> 0 Then" & vbCrLf
    sb.Append vbTab & "Set newFolder = curFolder.FoldersNew" & vbCrLf
    sb.Append vbTab & "newFolder.Name = " & VbStringFormat(pFolder.Name) & vbCrLf
    sb.Append vbTab & "newFolder.FolderType = " & pFolder.FolderType & vbCrLf
    sb.Append "End If" & vbCrLf
    sb.Append "newFolder.DescriptionRaw = " & VbStringFormat(pFolder.DescriptionRaw) & vbCrLf
    sb.Append "newFolder.Comments = " & VbStringFormat(pFolder.Comments) & vbCrLf
    sb.Append "newFolder.CharData = " & VbStringFormat(pFolder.CharData) & vbCrLf
    sb.Append "newFolder.IconRaw = " & VbStringFormat(pFolder.IconRaw) & vbCrLf
    
    If pFolder.FolderType = 1 Then
		sb.Append "sAux = " & VbStringFormat(pFolder.LogConf.Xml) & vbCrLf
		sb.Append "If IsG7 Then" & vbCrLf
		sb.Append vbTab & "Set oDom = dSession.Xml.NewDom()" & vbCrLf
		sb.Append vbTab & "oDom.loadXML sAux" & vbCrLf
		sb.Append vbTab & "Set newFolder.LogConf = oDom" & vbCrLf
		sb.Append "Else" & vbCrLf
		sb.Append vbTab & "newFolder.LogConf.loadXML sAux" & vbCrLf
		sb.Append "End If" &  vbCrLf
		sb.Append "newFolder.Form = dSession.Forms(" & VbStringFormat(pFolder.Form.Guid) & ")" & vbCrLf    
        
    ElseIf pFolder.FolderType = 2 Then
		sb.Append "newFolder.HrefRaw = " & VbStringFormat(pFolder.HrefRaw) & vbCrLf
		sb.Append "newFolder.Target = " & VbStringFormat(pFolder.Target) & vbCrLf
	End If
    
    sb.Append "newFolder.Save" & vbCrLf
    
	ScriptFolder = sb.ToString
End Function

Function ScriptFolderAcl(pFolder)
	Dim node, accName
	
	Set sb = dSession.ConstructNewStringBuilder
	
    sb.Append "curFolder.AclInherits = " & IIf(pFolder.AclInherits, "True", "False") & vbCrLf
    
    For Each node In pFolder.AclOwn.documentElement.childNodes
		sb.Append vbCrLf
    	If CLng(node.getAttribute("id")) <= 0 Then
    		sb.Append "lngAccId = " & node.getAttribute("id") & vbCrLf
    	Else
	    	accName = AccountsList.selectSingleNode("/d:root/d:item[@id='" & node.getAttribute("id") & "']").getAttribute("name")
	    	sb.Append "lngAccId = AccountId(""" & accName & """)" & vbCrLf
    	End If
    	sb.Append AclGrantRevoke(node, "fld_create") & vbCrLf
    	sb.Append AclGrantRevoke(node, "fld_read") & vbCrLf
    	sb.Append AclGrantRevoke(node, "fld_view") & vbCrLf
    	sb.Append AclGrantRevoke(node, "fld_admin") & vbCrLf
    	sb.Append AclGrantRevoke(node, "fld_create") & vbCrLf
    	sb.Append AclGrantRevoke(node, "doc_create") & vbCrLf
    	sb.Append AclGrantRevoke(node, "doc_read") & vbCrLf
    	sb.Append AclGrantRevoke(node, "doc_modify") & vbCrLf
    	sb.Append AclGrantRevoke(node, "doc_delete") & vbCrLf
    	sb.Append AclGrantRevoke(node, "doc_admin") & vbCrLf
    	sb.Append AclGrantRevoke(node, "doc_admin") & vbCrLf
    	sb.Append AclGrantRevoke(node, "vie_create") & vbCrLf
    	sb.Append AclGrantRevoke(node, "vie_read") & vbCrLf
    	sb.Append AclGrantRevoke(node, "vie_modify") & vbCrLf
    	sb.Append AclGrantRevoke(node, "vie_delete") & vbCrLf
    	sb.Append AclGrantRevoke(node, "vie_admin") & vbCrLf
    	sb.Append AclGrantRevoke(node, "vie_create_priv") & vbCrLf
    Next

	ScriptFolderAcl = sb.ToString
End Function

Function AclGrantRevoke(pNode, pAccess)
	AclGrantRevoke = "curFolder.Acl" & IIf(pNode.getAttribute(pAccess) & "" = "1", "Grant", "Revoke") & " CLng(lngAccId), """ & pAccess & """"
End Function

Function ScriptFolderEvent(pEvn)
	Set sb = dSession.ConstructNewStringBuilder

	sb.Append "Set oEvn = curFolder.Events(""ID=" & pEvn.id & """)" & vbCrLf
	If pEvn.code = "" Then
		sb.Append "oEvn.Code = """"" & vbCrLf
	Else
		sb.Append VbStringMultiLine("sb", pEvn.code) & vbCrLf
		sb.Append "oEvn.Code = sb.ToString" & vbCrLf
	End If
	sb.Append "oEvn.Overrides = " & IIf(pEvn.Overrides, "True", "False") & vbCrLf
	sb.Append "curFolder.Save" & vbCrLf
	
	ScriptFolderEvent = sb.ToString
End Function

Function ScriptAsyncEvent(pEvn)
	Set sb = dSession.ConstructNewStringBuilder

	sb.Append "Set oEvn = curFolder.AsyncEvents.Add" & vbCrLf
	sb.Append "oEvn.EventType = " & pEvn.EventType & vbCrLf
	sb.Append "oEvn.Login = " & VbStringFormat(pEvn.Login) & vbCrLf
	sb.Append "oEvn.Password = " & VbStringFormat(pEvn.Password) & vbCrLf
	sb.Append "oEvn.Disabled = " & IIf(pEvn.Disabled, "True", "False") & vbCrLf
	sb.Append "oEvn.IsCom = " & IIf(pEvn.IsCom, "True", "False") & vbCrLf
	
	If pEvn.IsCom Then
		sb.Append "oEvn.Class = " & VbStringFormat(pEvn.Class) & vbCrLf
		sb.Append "oEvn.Method = " & VbStringFormat(pEvn.Method) & vbCrLf
	Else
		sb.Append VbStringMultiLine("sb", pEvn.Code) & vbCrLf
		sb.Append "oEvn.Code = sb.ToString" & vbCrLf
		sb.Append "oEvn.CodeTimeout = " & pEvn.CodeTimeout & vbCrLf
	End If
	
	If pEvn.EventType = 0 Then ' TimerEvent
		sb.Append "oEvn.TimerNextRun = " & VbDateFormat(pEvn.TimerNextRun) & vbCrLf 'todo: este es fecha
		sb.Append "oEvn.TimerMode = " & pEvn.TimerMode & vbCrLf 'todo: que pasa con los null?
		sb.Append "oEvn.TimerFrequence = """ & pEvn.TimerFrequence & """" & vbCrLf
		sb.Append "oEvn.TimerTime = " & VbStringFormat(pEvn.TimerTime) & vbCrLf

	ElseIf pEvn.EventType = 1 Then ' TriggerEvent
		sb.Append "oEvn.TriggerEvent = " & pEvn.TriggerEvent & vbCrLf
		sb.Append "oEvn.Recursive = " & IIf(pEvn.Recursive, "True", "False") & vbCrLf
	End If
	sb.Append "curFolder.save " & vbCrLf
	ScriptAsyncEvent = sb.ToString
End Function

Function ScriptForm(pForm)
	Set sb = dSession.ConstructNewStringBuilder

    sb.Append "On Error Resume Next" & vbCrLf
    sb.Append "Set newForm = dSession.Forms(" & VbStringFormat(pForm.Guid) & ")" & vbCrLf
    sb.Append "ErrNumber = Err.Number" & vbCrLf
    sb.Append "On Error GoTo 0" & vbCrLf
    sb.Append "If ErrNumber <> 0 Then" & vbCrLf
    sb.Append vbTab & "Set newForm = dSession.FormsNew" & vbCrLf
    sb.Append vbTab & "newForm.Guid = " & VbStringFormat(pForm.Guid) & vbCrLf
    sb.Append "End If" & vbCrLf
    sb.Append "newForm.Name = " & VbStringFormat(pForm.Name) & vbCrLf
    sb.Append "newForm.Description = " & VbStringFormat(pForm.Description) & vbCrLf
    sb.Append "newForm.URLRaw = " & VbStringFormat(pForm.URLRaw) & vbCrLf
    sb.Append "newForm.Application = " & VbStringFormat(pForm.Application) & vbCrLf
    
	For Each oField In pForm.Fields
		If oField.Custom And Not oField.Computed Then
			sb.Append vbCrLf
			sb.Append "On Error Resume Next" & vbCrLf
			sb.Append "Set newField = newForm.Fields(" & VbStringFormat(oField.Name) & ")" & vbCrLf
			sb.Append "ErrNumber = Err.Number" & vbCrLf
			sb.Append "On Error GoTo 0" & vbCrLf
			sb.Append "If ErrNumber <> 0 Then" & vbCrLf
			sb.Append vbTab & "Set newField = newForm.Fields.Add(" & VbStringFormat(oField.Name) & ")" & vbCrLf
			sb.Append vbTab & "newField.DataType = " & oField.DataType & vbCrLf
			If oField.DataType = 1 Then
				sb.Append vbTab & "newField.DataLength = " & oField.DataLength & vbCrLf
			ElseIf oField.DataType = 3 Then
				sb.Append vbTab & "newField.DataPrecision = " & oField.DataPrecision & vbCrLf
				sb.Append vbTab & "newField.DataScale = " & oField.DataScale & vbCrLf
			End If
			sb.Append vbTab & "newField.Nullable = " & IIf(oField.Nullable, "True", "False") & vbCrLf
			sb.Append "End If" & vbCrLf
			sb.Append "newField.DescriptionRaw = " & VbStringFormat(oField.DescriptionRaw & "") & vbCrLf
		End If
	Next

	sb.Append vbCrLf
	sb.Append "sAux = " & VbStringFormat(pForm.Actions.Xml) & vbCrLf
	sb.Append "If IsG7 Then" &  vbCrLf
	sb.Append vbTab & "Set oDom = dSession.Xml.NewDom()" & vbCrLf
	sb.Append vbTab & "oDom.loadXML sAux" & vbCrLf
	sb.Append vbTab & "Set newForm.Actions = oDom" & vbCrLf
	sb.Append "Else" & vbCrLf
    sb.Append vbTab & "newForm.Actions.loadXml sAux" & vbCrLf
	sb.Append "End If" & vbCrLf
    sb.Append vbCrLf
    sb.Append "newForm.Save" & vbCrLf
	sb.Append "Set curForm = newForm" & vbCrLf

	sb.Append ScriptProperties(pForm, "curForm")
	
	ScriptForm = sb.ToString
End Function

Function ScriptFormEvent(pEvn)
	Set sb = dSession.ConstructNewStringBuilder

	sb.Append "Set oEvn = curForm.Events(""ID=" & pEvn.id & """)" & vbCrLf
	If pEvn.code = "" Then
		sb.Append "oEvn.Code = """"" & vbCrLf
	Else
		sb.Append VbStringMultiLine("sb", pEvn.code) & vbCrLf
		sb.Append "oEvn.Code = sb.ToString" & vbCrLf
	End If
	sb.Append "oEvn.Overridable = " & IIf(pEvn.Overridable, "True", "False") & vbCrLf
	sb.Append "oEvn.Extensible = " & IIf(pEvn.Extensible, "True", "False") & vbCrLf
	sb.Append "curForm.Save" & vbCrLf
	
	ScriptFormEvent = sb.ToString
End Function

Function ScriptView(pView)
	Set sb = dSession.ConstructNewStringBuilder

	sb.Append "On Error Resume Next" & vbCrLf
	sb.Append "Set newView = curFolder.Views(" & VbStringFormat(pView.Name) & ")" & vbCrLf
	sb.Append "ErrNumber = Err.Number" & vbCrLf
	sb.Append "On Error GoTo 0" & vbCrLf
	sb.Append "If ErrNumber <> 0 Then" & vbCrLf
	sb.Append vbTab & "Set newView = curFolder.ViewsNew" & vbCrLf
	sb.Append vbTab & "newView.Name = " & VbStringFormat(pView.Name) & vbCrLf
	sb.Append "End If" & vbCrLf
	sb.Append "newView.DescriptionRaw = " & VbStringFormat(pView.DescriptionRaw) & vbCrLf
	sb.Append "newView.Comments = " & VbStringFormat(pView.Comments) & vbCrLf
	sb.Append "Set oDom = dSession.Xml.NewDom()" & vbCrLf
	sb.Append "oDom.loadXML" & VbStringFormat(pView.Definition.Xml) & vbCrLf
	sb.Append "oDom.setProperty ""SelectionNamespaces"", ""xmlns:d=""""viewDefinition""""""" & vbCrLf
	sb.Append "If IsG7 Then" &  vbCrLf
	sb.Append vbTab & "Set newView.Definition = oDom" & vbCrLf
	sb.Append "Else" & vbCrLf
	sb.Append vbTab & "oDom.documentElement.removeAttribute(""viewtype"")" & vbCrLf
	sb.Append vbTab & "For Each node In oDom.selectNodes(""/d:root/d:fields/d:item"")" & vbCrLf
	sb.Append vbTab & vbTab & "node.removeAttribute(""isimage"")" & vbCrLf
	sb.Append vbTab & "Next" & vbCrLf
	sb.Append vbTab & "For Each node in oDom.selectNodes(""/d:root/d:groups/d:item"")" & vbCrLf
	sb.Append vbTab & vbTab & "node.removeAttribute(""direction"")" & vbCrLf
	sb.Append vbTab & vbTab & "node.removeAttribute(""orderby"")" & vbCrLf
	sb.Append vbTab & "Next"& vbCrLf
	sb.Append vbTab & "newView.Definition.loadXml oDom.Xml" & vbCrLf
	sb.Append "End If" & vbCrLf
	sb.Append "newView.Save" & vbCrLf

	ScriptView = sb.ToString
End Function

Function ScriptDocument(pDoc)
	Dim formula, sb, sbHeader, sbAux
	Dim oField
	
	Set sb = dSession.ConstructNewStringBuilder

	formula = PKFormula(pDoc.Form, pDoc)
	
	If formula <> "" And UCase(Left(formula, 6)) <> "DOC_ID" Then
		sb.Append "On Error Resume Next" & vbCrLf
		sb.Append "Set newDoc = curFolder.Documents(""" & formula & """)" & vbCrLf
		sb.Append "ErrNumber = Err.Number" & vbCrLf
		sb.Append "On Error GoTo 0" & vbCrLf
		sb.Append "If ErrNumber <> 0 Then" & vbCrLf
		sb.Append vbTab & "Set newDoc = curFolder.DocumentsNew" & vbCrLf
		sb.Append "End If" & vbCrLf
	Else
		sb.Append "Set newDoc = curFolder.DocumentsNew" & vbCrLf
	End If

	Set sbHeader = dSession.ConstructNewStringBuilder
	
	For Each oField In pDoc.Fields
		If Updatable(oField) Then ' TODO: cdo este arreglado reemplazar por oField.Updatable
			If UCase(oField.Name) = "DOC_ID" Or Not oField.Custom Then
				Set sbAux = sbHeader
			Else
				Set sbAux = sb
			End If
			
			If IsNull(oField.Value) Then
				sbAux.Append "newDoc(""" & oField.Name & """).Value = Null" & vbCrLf
			ElseIf oField.DataType = 1 Then ' Char
				If oField.Value & "" = "" Then
					sbAux.Append "newDoc(""" & oField.Name & """).Value = """"" & vbCrLf
				Else
					If oField.DataLength = 0 or oField.DataLength > 500 Then
						sbAux.Append VbStringMultiLine("sb", oField.Value) & vbCrLf
						sbAux.Append "newDoc(""" & oField.Name & """).Value = sb.ToString" & vbCrLf
					Else
						sbAux.Append "newDoc(""" & oField.Name & """).Value = " & VbStringFormat(oField.Value) & vbCrLf
					End If
				End If
			ElseIf oField.DataType = 2 Then ' DateTime
				sbAux.Append "newDoc(""" & oField.Name & """).Value = " & VbDateFormat(oField.Value) & vbCrLf
			ElseIf oField.DataType = 3 Then ' Numeric
				sbAux.Append "newDoc(""" & oField.Name & """).Value = " & VbNumberFormat(oField.Value) & vbCrLf
			End If
		End If
	Next
	
	sb.Append "On Error Resume Next" & vbCrLf
	sb.Append sbHeader.ToString
	sb.Append "On Error Goto 0" & vbCrLf
	sb.Append "newDoc.Save" & vbCrLf
	
	ScriptDocument = sb.ToString
End Function

Function PKFormula(pForm, pDoc)
	Dim sKey, formula, arr, sGuid
	
	sKey = ""
	If pForm.Properties.Exists("PK") Then sKey = LCase(pForm.Properties("PK").Value)
	If sKey = "" Then
		sGuid = UCase(pForm.Guid)
		If sGuid = "F89ECD42FAFF48FDA229E4D5C5F433ED" Then ' CodeLib
			sKey = "name"
		ElseIf sGuid = "EAC99A4211204E1D8EEFEB8273174AC4" Then ' Controls
			sKey = "name"
		ElseIf sGuid = "B87B1CB5EFB94B03BA6B1F18DBE5F5D4" Then ' Keywords3
			sKey = "id"
		ElseIf sGuid = "B89302DBBE45498EA03A495B53D3F50C" Then ' Secuences3
			sKey = "sequence"
		ElseIf sGuid = "5C0D6DBF72CF42608989A862ED9E7444" Then ' Settings3
			sKey = "setting"
		Else
			sKey = "doc_id"
		End If
	End If

	arr = Split(sKey, ",")
	formula = ""
	For i = 0 To UBound(arr)
		field = LCase(Trim(arr(i)))
		If TypeName(pDoc) = "Document" Then
			If pDoc(field).Value & "" <> "" Then
				formula = formula & " and " & field & " = " & dSession.Db.SqlEncode(pDoc(field).Value, pDoc(field).DataType)
			End If
		Else
			' Nodo
			If pDoc.getAttribute(field) & "" <> "" Then
				formula = formula & " and " & field & " = " & dSession.Db.SqlEncode(dSession.Xml.XmlDecode(pDoc.getAttribute(field), pForm(field).DataType), pForm(field).DataType)
			End If
		End If
	Next
	If formula <> "" Then formula = Mid(formula, 6)
	
	PKFormula = formula
End Function

Function ScriptAccount(pAcc)
	Set sb = dSession.ConstructNewStringBuilder

	If pAcc.AccountType <> 1 And pAcc.AccountType <> 2 Then
		ScriptAccount = "' Account not scriptable"
		Exit Function
	End If
	
	sb.Append "On Error Resume Next" & vbCrLf
	sb.Append "Set newAcc = dSession.Directory.Accounts(" & VbStringFormat(pAcc.Name) & ")" & vbCrLf
	sb.Append "ErrNumber = Err.Number" & vbCrLf
	sb.Append "On Error GoTo 0" & vbCrLf
	sb.Append "If ErrNumber <> 0 Then" & vbCrLf
	sb.Append vbTab & "Set newAcc = dSession.Directory.AccountsNew(" & pAcc.AccountType & ")" & vbCrLf
	sb.Append "End If" & vbCrLf
	
	If pAcc.AccountType = 1 Then ' User

		Set oUsr = pAcc.Cast2User
		sb.Append "newAcc.FullName = " & VbStringFormat(oUsr.FullName) & vbCrLf
		sb.Append "newAcc.Login = " & VbStringFormat(oUsr.Login) & vbCrLf
		sb.Append "newAcc.Name = " & VbStringFormat(oUsr.Name) & vbCrLf
		sb.Append "newAcc.Description = " & VbStringFormat(oUsr.Description) & vbCrLf
		sb.Append "newAcc.Email = " & VbStringFormat(oUsr.Email) & vbCrLf
		sb.Append "newAcc.Password = DefaultPassword" & vbCrLf
		sb.Append "newAcc.CannotChangePwd = " & IIf(oUsr.CannotChangePwd, "True", "False") & vbCrLf
		sb.Append "newAcc.ChangePwdNextLogon = True" & vbCrLf
		sb.Append "newAcc.PwdNeverExpires = " & IIf(oUsr.PwdNeverExpires, "True", "False") & vbCrLf
		sb.Append "newAcc.WinLogon = " & IIf(oUsr.WinLogon, "True", "False") & vbCrLf
		sb.Append "newAcc.GestarLogon = " & IIf(oUsr.GestarLogon, "True", "False") & vbCrLf
		sb.Append "newAcc.Language = " & oUsr.Language & vbCrLf
		sb.Append "newAcc.Theme = " & VbStringFormat(oUsr.Theme) & vbCrLf
		sb.Append "newAcc.Save" & vbCrLf

	ElseIf pAcc.AccountType = 2 Then ' Group
        
		sb.Append "newAcc.Name = " & VbStringFormat(pAcc.Name) & vbCrLf
		sb.Append "newAcc.Description = " & VbStringFormat(pAcc.Description) & vbCrLf
		sb.Append "newAcc.Email = " & VbStringFormat(pAcc.Email) & vbCrLf
		sb.Append "newAcc.Save" & vbCrLf

	End If
	
	ScriptAccount = sb.ToString
End Function

Function ScriptAccountParents(pAcc)
	Set sb = dSession.ConstructNewStringBuilder

	If pAcc.AccountType <> 1 And pAcc.AccountType <> 2 Then
		ScriptAccountParents = "' Account not scriptable"
		Exit Function
	End If

	Set dom = pAcc.ParentAccountsList

	If dom.documentElement.childNodes.length > 0 Then
		sb.Append "Set acc = dSession.Directory.Accounts(" & VbStringFormat(pAcc.Name) & ")" & vbCrLf
	Else
		sb.Append "' No parents" & vbCrLf
	End If
	
	For Each node In dom.documentElement.childNodes
		sb.Append "acc.ParentAccountsAdd " & VbStringFormat(node.getAttribute("name")) & vbCrLf
	Next
	
	ScriptAccountParents = sb.ToString
End Function

Function ScriptProperties(pObj, pObjName)
	Dim prop, sAux, sName, sValue
	
	sAux = ""
	For Each prop In pObj.Properties
		sName = VbStringFormat(prop.Name)
		sValue = VbStringFormat(prop.Value)
		sAux = sAux & vbCrLf
		sAux = sAux & "If " & pObjName & ".Properties.Exists(" & sName & ") Then" & vbCrLf
		sAux = sAux & vbTab & pObjName & ".Properties(" & sName & ").Value = " & sValue & vbCrLf
		sAux = sAux & "Else" & vbCrLf
		sAux = sAux & vbTab & pObjName & ".Properties.Add " & sName & ", " & sValue & vbCrLf
		sAux = sAux & "End If" & vbCrLf
	Next
	ScriptProperties = sAux
End Function

Function VbStringFormat(pString)
    Dim strRet
    
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

Private Function VbStringMultiLine(pStringBuilderName, pString)
    Dim strRet, arrLines
    Dim strAux, i
    
	Set sb = dSession.ConstructNewStringBuilder
	sb.Append "Set " & pStringBuilderName & " = dSession.ConstructNewStringBuilder" & vbCrLf
    If pString <> "" Then
		arrLines = Split(Replace(pString, vbCrLf, VbLf), vbLf)
		For i = 0 To UBound(arrLines)
			sb.Append pStringBuilderName & ".Append " & VbStringFormat(arrLines(i) & "")
			If i < UBound(arrLines) Then sb.Append " & vbCrLf" & vbCrLf
		Next
    End If

    VbStringMultiLine = sb.ToString
End Function

Function VbDateFormat(pDate)
    Dim strRet
    
	If IsNull(pDate) or CLng(pDate) = 0 Then
		strRet = "Null"
	Else
		strRet = "#" & Month(pDate) & "/" & Day(pDate) & "/" & Year(pDate)
		strRet = strRet & " " & Hour(pDate) & ":" & Minute(pDate) & ":" & Second(pDate) & "#"
	End If
	
    VbDateFormat = strRet
End Function

Function VbNumberFormat(pNumber)
	If pNumber & "" = "" Then
		VbNumberFormat = "Null"
	Else
		VbNumberFormat = dSession.Xml.XmlEncode(pNumber, 3)
	End If
End Function

Function Updatable(pField)
	Updatable = True
	sName = UCase(pField.Name)
	If sName = "DOC_ID" Or sName = "FRM_ID" Or sName = "FLD_ID" Or _
		sName = "ACC_ID" Or sName = "CREATED" Or sName = "MODIFIED" Or _
		sName = "ACCESSED" Or sName = "INHERITS" Or sName = "FLD_ID_OLD" Or _
		pField.Computed Then
			Updatable = False
	End If
End Function

'TODO: Funcion de parche hasta que se resuelva Ancestors
'cdo este dejar la de abajo que va mas rapido
Function FolderPath(pFolder, pOnlyNames)
	Dim strAux

	Set fld = pFolder
	strAux = ""
	
	Do While Not fld Is Nothing
		If fld.Description <> "" And Not pOnlyNames Then
			strAux = "/" & fld.Description & strAux
		Else
			strAux = "/" & fld.Name & strAux
		End If
		If fld.Id = 1 Or fld.Id = 1001 Then
			Set fld = Nothing
		Else
			Set fld = fld.Parent
		End If
	Loop
	
	FolderPath = strAux
End Function

'Function FolderPath(pFolder, pOnlyNames)
'	Dim strAux, dom, node, desc
'	
'	strAux = "/"
'	If pFolder.Description <> "" And Not pOnlyNames Then
'		strAux = strAux & pFolder.Description
'	Else
'		strAux = strAux & pFolder.Name
'	End If
'
'	Set dom = pFolder.Ancestors
'	For Each node In dom.documentElement.childNodes
'		desc = node.getAttribute("description") & ""
'		If desc <> "" And Not pOnlyNames Then
'			strAux = desc & strAux
'		Else
'			strAux = node.getAttribute("name") & strAux
'		End If
'		strAux = "/" & strAux
'	Next
'	
'	FolderPath = strAux
'End Function

Function FolderNodePath(pNode)
	Dim sAux, node
	
	sAux = ""
	Set node = pNode
	Do While node.getAttribute("name") <> "PublicFolders"
		sAux = "/" & node.getAttribute("name") & sAux
		Set node = node.parentNode
	Loop
	
	FolderNodePath = sAux
End Function

Function AccountsList()
	If IsEmpty(accList) Then
		Set accList = dSession.Directory.AccountsList
	End If
	Set AccountsList = accList
End Function

Function WithoutCrLf(pString)
	Dim sAux
	
    sAux = Replace(pString, vbCr, " ")
    sAux = Replace(sAux, vbLf, " ")
	WithoutCrLf = sAux
End Function
%>
