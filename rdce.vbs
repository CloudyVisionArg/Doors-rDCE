'----------------------------------------
' Archivo de instalacion de Cloudy Doors
' Creado: 10/11/2015 15:40:37
'----------------------------------------

IsG7 = (CLng(Split(dSession.Version, ".")(0)) >= 7)


'---
'--- Scripts for form MAILS
'---

On Error Resume Next
Set newForm = dSession.Forms("b1ac88af09484dd7bc9efc88912fa358")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newForm = dSession.FormsNew
	newForm.Guid = "b1ac88af09484dd7bc9efc88912fa358"
End If
newForm.Name = "mails"
newForm.Description = "CRM Mails"
newForm.URLRaw = "[APPVIRTUALROOT]/forms/generic.asp"
newForm.Application = "CRM"

On Error Resume Next
Set newField = newForm.Fields("NAME")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("NAME")
	newField.DataType = 1
	newField.DataLength = 100
	newField.Nullable = True
End If
newField.DescriptionRaw = "[LANGSTRING(319)]"

On Error Resume Next
Set newField = newForm.Fields("TEMPLATEBODY")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("TEMPLATEBODY")
	newField.DataType = 1
	newField.DataLength = 0
	newField.Nullable = True
End If
newField.DescriptionRaw = "[LANGSTRING(1024)]"

On Error Resume Next
Set newField = newForm.Fields("TEMPLATESUBJECT")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("TEMPLATESUBJECT")
	newField.DataType = 1
	newField.DataLength = 255
	newField.Nullable = True
End If
newField.DescriptionRaw = "[LANGSTRING(431)]"

On Error Resume Next
Set newField = newForm.Fields("LINKTODOC")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("LINKTODOC")
	newField.DataType = 3
	newField.DataPrecision = 3
	newField.DataScale = 0
	newField.Nullable = True
End If
newField.DescriptionRaw = ""

On Error Resume Next
Set newField = newForm.Fields("MODULO")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("MODULO")
	newField.DataType = 1
	newField.DataLength = 50
	newField.Nullable = True
End If
newField.DescriptionRaw = "[LANGSTRING(1025)]"

On Error Resume Next
Set newField = newForm.Fields("DESCRIPTION")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("DESCRIPTION")
	newField.DataType = 1
	newField.DataLength = 150
	newField.Nullable = True
End If
newField.DescriptionRaw = "[LANGSTRING(48)]"

On Error Resume Next
Set newField = newForm.Fields("QUICKMAIL")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("QUICKMAIL")
	newField.DataType = 3
	newField.DataPrecision = 1
	newField.DataScale = 0
	newField.Nullable = True
End If
newField.DescriptionRaw = ""

On Error Resume Next
Set newField = newForm.Fields("CODE")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("CODE")
	newField.DataType = 1
	newField.DataLength = 0
	newField.Nullable = True
End If
newField.DescriptionRaw = ""

On Error Resume Next
Set newField = newForm.Fields("ATTACHMENTS")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("ATTACHMENTS")
	newField.DataType = 3
	newField.DataPrecision = 1
	newField.DataScale = 0
	newField.Nullable = True
End If
newField.DescriptionRaw = ""

On Error Resume Next
Set newField = newForm.Fields("TEMPLATEFORMAT")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("TEMPLATEFORMAT")
	newField.DataType = 3
	newField.DataPrecision = 1
	newField.DataScale = 0
	newField.Nullable = True
End If
newField.DescriptionRaw = "[LANGSTRING(1968)]"

On Error Resume Next
Set newField = newForm.Fields("BODY")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newField = newForm.Fields.Add("BODY")
	newField.DataType = 1
	newField.DataLength = 0
	newField.Nullable = True
End If
newField.DescriptionRaw = ""

sAux = "<?xml version=""1.0""?>" & vbCrLf & "<root xmlns=""actionList""/>" & vbCrLf
If IsG7 Then
	Set oDom = dSession.Xml.NewDom()
	oDom.loadXML sAux
	Set newForm.Actions = oDom
Else 
	newForm.Actions.loadXml sAux
End If

newForm.Save
Set curForm = newForm

If curForm.Properties.Exists("DCE_CodeColumn") Then
	curForm.Properties("DCE_CodeColumn").Value = "code"
Else
	curForm.Properties.Add "DCE_CodeColumn", "code"
End If

If curForm.Properties.Exists("DCE_HasCode") Then
	curForm.Properties("DCE_HasCode").Value = "1"
Else
	curForm.Properties.Add "DCE_HasCode", "1"
End If

If curForm.Properties.Exists("DCE_ListColumns") Then
	curForm.Properties("DCE_ListColumns").Value = "name"
Else
	curForm.Properties.Add "DCE_ListColumns", "name"
End If


'--- Event Document_BeforeMove

Set oEvn = curForm.Events("ID=8")
oEvn.Code = ""
oEvn.Overridable = True
oEvn.Extensible = True
curForm.Save


'--- Event Document_BeforeSave

Set oEvn = curForm.Events("ID=2")
oEvn.Code = ""
oEvn.Overridable = True
oEvn.Extensible = True
curForm.Save


'---
'--- Scripts for Global Codelibs
'---

Set curFolder = dSession.FoldersGetFromId(11)

'--- Scripts for codelib SESSION_ONSTART

On Error Resume Next
Set newDoc = curFolder.Documents("name = 'Session_OnStart'")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newDoc = curFolder.DocumentsNew
End If
newDoc("NAME").Value = "Session_OnStart"
newDoc("DESCRIPTION").Value = ""
Set sb = dSession.ConstructNewStringBuilder
sb.Append "dSession.DebugPrint ""Session_OnStart""" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "' Agrega favoritos predeterminados solo una vez" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "Dim favs" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "'If dSession.LoggedUser.Settings(""FLD_FAVORITES_EXECUTED"") <> ""1"" Then" & vbCrLf
sb.Append "	AgregarFavorito ""//crm_root/contactos""" & vbCrLf
sb.Append "	AgregarFavorito ""//crm_root/oportunidades""" & vbCrLf
sb.Append "	AgregarFavorito ""//crm_root/leads""" & vbCrLf
sb.Append "	AgregarFavorito 4 ' Cambiar contraseña" & vbCrLf
sb.Append "	dSession.LoggedUser.Settings(""FLD_FAVORITES_EXECUTED"") = ""1""" & vbCrLf
sb.Append "'End If" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "' Establece como inicial la carpeta Oportunidades, vista Mis pendientes" & vbCrLf
sb.Append "dSession.LoggedUser.Settings(""FLD_ID"") = ""5107""" & vbCrLf
sb.Append "dSession.LoggedUser.Settings(""5107"") = ""4698""" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "Sub AgregarFavorito(fav)" & vbCrLf
sb.Append "	If Not IsNumeric(fav) Then" & vbCrLf
sb.Append "		fav = dSession.FoldersGetFromId(1001).App.Folders(CStr(fav)).Id" & vbCrLf
sb.Append "	End If" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "	favs = dSession.LoggedUser.Settings(""FLD_FAVORITES"")" & vbCrLf
sb.Append "	If InStr("","" & favs & "","", "","" & fav & "","") = 0 Then" & vbCrLf
sb.Append "		If favs <> """" Then favs = favs & "",""" & vbCrLf
sb.Append "		favs = favs & fav" & vbCrLf
sb.Append "	End If " & vbCrLf
sb.Append "	dSession.LoggedUser.Settings(""FLD_FAVORITES"") = favs" & vbCrLf
sb.Append "End Sub" & vbCrLf
sb.Append ""
newDoc("CODE").Value = sb.ToString
On Error Resume Next
newDoc("TASK_ID").Value = Null
newDoc("TASK_RESPONSIBLE").Value = ""
newDoc("TASK_RESPONSIBLEID").Value = Null
newDoc("TASK_TYPE").Value = ""
newDoc("TASK_TYPEID").Value = Null
newDoc("TASK_STATE").Value = ""
newDoc("TASK_STATEID").Value = Null
newDoc("TASK_DUEDATE").Value = Null
newDoc("TASK_CUSTOMER").Value = ""
newDoc("TASK_CUSTOMERID").Value = Null
newDoc("TASK_SOURCE").Value = ""
newDoc("WORKFLOW_ID").Value = ""
On Error Goto 0
newDoc.Save


'--- Scripts for codelib USERQUERY

On Error Resume Next
Set newDoc = curFolder.Documents("name = 'UserQuery'")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newDoc = curFolder.DocumentsNew
End If
newDoc("NAME").Value = "UserQuery"
newDoc("DESCRIPTION").Value = ""
Set sb = dSession.ConstructNewStringBuilder
sb.Append "'Copad" & vbCrLf
sb.Append "'----------------------------" & vbCrLf
sb.Append "' UserQuery 1.1" & vbCrLf
sb.Append "'----------------------------" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "Class clsToken" & vbCrLf
sb.Append "	Public ttype" & vbCrLf
sb.Append "	Public value" & vbCrLf
sb.Append "End Class" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "Class clsUserQuery" & vbCrLf
sb.Append "	Private NoiseWords" & vbCrLf
sb.Append "	Private ClassInit" & vbCrLf
sb.Append "	Private USERITEM" & vbCrLf
sb.Append "	Private ANDOP" & vbCrLf
sb.Append "	Private OROP" & vbCrLf
sb.Append "	Private NOTOP" & vbCrLf
sb.Append "	Private LEFTPAREN" & vbCrLf
sb.Append "	Private RIGHTPAREN" & vbCrLf
sb.Append "	Private NEAROP" & vbCrLf
sb.Append "	Private NOISEWORD" & vbCrLf
sb.Append "	Private OPERATOR" & vbCrLf
sb.Append "	Private BINARYOP" & vbCrLf
sb.Append "	Private EXPRESSION" & vbCrLf
sb.Append "	Private builtIn" & vbCrLf
sb.Append "	Private sErr" & vbCrLf
sb.Append "	Private tokens" & vbCrLf
sb.Append "	" & vbCrLf
sb.Append " 	Private Sub Class_Initialize()" & vbCrLf
sb.Append "		If IsEmpty(ClassInit) Then" & vbCrLf
sb.Append "			ClassInit = True" & vbCrLf
sb.Append "			USERITEM = 1" & vbCrLf
sb.Append "			ANDOP = 2" & vbCrLf
sb.Append "			OROP = 4" & vbCrLf
sb.Append "			NOTOP = 8" & vbCrLf
sb.Append "			LEFTPAREN = 16" & vbCrLf
sb.Append "			RIGHTPAREN = 32" & vbCrLf
sb.Append "			NEAROP = 64" & vbCrLf
sb.Append "			NOISEWORD = 128" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "			OPERATOR = ANDOP Or OROP Or NOTOP Or NEAROP" & vbCrLf
sb.Append "			BINARYOP = ANDOP Or OROP Or NEAROP" & vbCrLf
sb.Append "			EXPRESSION = RIGHTPAREN Or USERITEM" & vbCrLf
sb.Append "		End If" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		Set builtIn = CreateObject(""Scripting.Dictionary"")" & vbCrLf
sb.Append "		sErr = """"" & vbCrLf
sb.Append "		AddBuiltIn ""and"", ""AND"", ANDOP" & vbCrLf
sb.Append "		AddBuiltIn ""or"", ""OR"", OROP" & vbCrLf
sb.Append "		AddBuiltIn ""near"", ""NEAR"", NEAROP" & vbCrLf
sb.Append "		AddBuiltIn ""not"", ""NOT"", NOTOP" & vbCrLf
sb.Append "		AddBuiltIn ""("", ""("", LEFTPAREN" & vbCrLf
sb.Append "		AddBuiltIn "")"", "")"", RIGHTPAREN" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		NoiseWords = ""C,1,D,2,I,3,II,4,III,5,IV,6,IX,7,L,8,M,9,Ud.,0,Uds.,$,V,VI,VII,VIII,Vd.,Vds.,X,XI,XII,XIII,XIV,XIX,XL,XLI,XLII,XLIII,XLIV,XLIX,XLV,XLVI,XLVII,XLVIII,XV,XVI,XVII,XVIII,XX,XXI,XXII,XXIII,XXIV,XXIX,XXV,XXVI,XXVII,XXVIII,XXX,XXXI,XXXII,XXXIII,XXXIV,XXXIX,XXXV,XXXVI,XXXVII,XXXVIII,a,ab,abajo,acerca,adelante,además,adentro,adiós,adonde,adónde,afuera,ah,ahora,ajá,albricias,ale,aleluya,algo,alguien,alguna,algunas,alguno,algunos,algún,alta,alto,ambas,ambos,andando,ante,antes,aquel,aquella,aquellas,aquello,aquellos,aquél,aquélla,aquéllas,aquéllos,arre,arriba,así,atiza,atrás,aun,aunque,ay,aúpa,bah,bajo,basta,bastante,bastantes,bien,billonésima,billonésimas,billonésimo,billonésimos,billón,cabe,cada,caramba,caray,cataplum,cataplún,catorce,catorcena,catorcenas,catorceno,catorcenos,ce,centena,centenas,centeno,centenos,centésima,centésimas,centésimo,centésimos,chau,che,chis,chist,chito,chitón,cien,cienmilésima,cienmilésimas,cienmilésimo,cienmilésimos,ciento,cierta,ciertas,cierto,ciertos,cinco,cincuenta,cincuentena,cincuentenas,cincuenteno,cincuentenos,como,con,conmigo,conque,consigo,contigo,contra,cuadragésima,cuadragésimas,cuadragésimo,cuadragésimos,cuadringentésima,cuadringentésimas,cuadringentésimo,cuadringentésimos,cual,cuales,cualesquiera,cualquier,cualquiera,cuando,cuanta,cuantas,cuanto,cuantos,cuarenta,cuarta,cuartas,cuarto,cuartos,cuatrillón,cuatro,cuatrocientos,cuya,cuyas,cuyo,cuyos,cuál,cuáles,cuándo,cuánta,cuántas,cuánto,cuántos,cáspita,cómo,córcholis,daca,de,deba,debamos,deban,debas,debe,debed,debemos,deben,debes,debido,debiendo,debiera,debierais,debieran,debieras,debiere,debiereis,debieren,debieres,debieron,debiese,debieseis,debiesen,debieses,debimos,debiste,debisteis,debiéramos,debiéremos,debiésemos,debió,debo,debáis,debéis,debí,debía,debíais,debíamos,debían,debías,decena,decenas,deceno,decenos,decimacuarta,decimacuartas,decimanovena,decimanovenas,decimaoctava,decimaoctavas,decimaquinta,decimaquintas,decimasexta,decimasextas,decimaséptima,decimaséptimas,decimatercera,decimaterceras,decimoctava,decimoctavas,decimoctavo,decimoctavos,decimocuarta,decimocuartas,decimocuarto,decimocuartos,decimonona,decimononas,decimonono,decimononos,decimonovena,decimonovenas,decimonoveno,decimonovenos,decimoquinta,decimoquintas,decimoquinto,decimoquintos,decimosexta,decimosextas,decimosexto,decimosextos,decimoséptima,decimoséptimas,decimoséptimo,decimoséptimos,decimotercia,decimotercias,decimotercio,decimotercios,demasiada,demasiadas,demasiado,demasiados,demás,desde,después,diecinueve,dieciochena,dieciochenas,dieciocheno,dieciochenos,dieciocho,diecisiete,dieciséis,diez,diezmillonésima,diezmillonésimas,diezmillonésimo,diezmillonésimos,diezmilésima,diezmilésimas,diezmilésimo,diezmilésimos,doce,docena,docenas,doceno,docenos,donde,dos,doscientos,ducentésima,ducentésimas,ducentésimo,ducentésimos,duodécima,duodécimas,duodécimo,duodécimos,durante,décima,décimas,décimo,décimos,dónde,e,ea,eh,ejem,el,ella,ellas,ello,ellos,empero,en,entre,enésima,enésimas,enésimo,enésimos,epa,era,erais,eran,eras,eres,ergo,es,esa,esas,ese,eso,esos,esotra,esotras,esotro,esotros,esta,estaba,estabais,estaban,estabas,estad,estado,estamos,estando,estar,estaremos,estará,estarán,estarás,estaré,estaréis,estaría,estaríais,estaríamos,estarían,estarías,estas,este,estemos,esto,estos,estotra,estotras,estotro,estotros,estoy,estuve,estuviera,estuvierais,estuvieran,estuvieras,estuviere,estuviereis,estuvieren,estuvieres,estuvieron,estuviese,estuvieseis,estuviesen,estuvieses,estuvimos,estuviste,estuvisteis,estuviéramos,estuviéremos,estuviésemos,estuvo,está,estábamos,estáis,están,estás,esté,estéis,estén,estés,excepto,extra,forte,fu,fue,fuera,fuerais,fueran,fueras,fuere,fuereis,fueren,fueres,fueron,fuese,fueseis,fuesen,fueses,fui,fuimos,fuiste,fuisteis,fuéramos,fuéremos,fuésemos,gua,guapa,guapo,ha,habed,haber,habiendo,habremos,habrá,habrán,habrás,habré,habréis,habría,habríais,habríamos,habrían,habrías,habéis,había,habíais,habíamos,habían,habías,hacia,hala,han,has,hasta,hay,haya,hayamos,hayan,hayas,hayáis,he,hemos,hola,hopo,hube,hubiera,hubierais,hubieran,hubieras,hubiere,hubiereis,hubieren,hubieres,hubieron,hubiese,hubieseis,hubiesen,hubieses,hubimos,hubiste,hubisteis,hubiéramos,hubiéremos,hubiésemos,hubo,hurra,huy,ja,jamás,je,ji,jo,la,larga,largo,las,le,les,lo,los,luego,maldición,malhaya,mas,me,mecachis,mediante,menos,mi,mientras,mil,millonésima,millonésimas,millonésimo,millonésimos,millón,milmillonésima,milmillonésimas,milmillonésimo,milmillonésimos,milésima,milésimas,milésimo,milésimos,mis,mucha,muchas,mucho,muchos,mí,mía,mías,mío,míos,nada,nadie,ni,ninguna,ningunas,ninguno,ningunos,ningún,no,nonagésima,nonagésimas,nonagésimo,nonagésimos,noningentésima,noningentésimas,noningentésimo,noningentésimos,nos,nosotras,nosotros,novecientos,novena,novenas,noveno,novenos,noventa,nra.,nro.,ntra.,ntro.,nuestra,nuestras,nuestro,nuestros,nueve,nunca,o,ochenta,ochentena,ochentenas,ochenteno,ochentenos,ocho,ochocientos,octava,octavas,octavo,octavos,octingentésima,octingentésimas,octingentésimo,octingentésimos,octogésima,octogésimas,octogésimo,octogésimos,oh,ojalá,ole,olé,once,oncena,oncenas,onceno,oncenos,ora,os,otra,otras,otro,otros,otrosí,paf,para,pardiez,pataplún,pche,pchs,pero,pf,poca,pocas,poco,pocos,podamos,poded,podemos,poder,podido,podremos,podrá,podrán,podrás,podré,podréis,podría,podríais,podríamos,podrían,podrías,podáis,podéis,podía,podíais,podíamos,podían,podías,poquita,poquitas,poquito,poquitos,por,porque,primer,primera,primeras,primero,primeros,pude,pudiendo,pudiera,pudierais,pudieran,pudieras,pudiere,pudiereis,pudieren,pudieres,pudieron,pudiese,pudieseis,pudiesen,pudieses,pudimos,pudiste,pudisteis,pudiéramos,pudiéremos,pudiésemos,pudo,pueda,puedan,puedas,puede,pueden,puedes,puedo,pues,puf,pum,que,queramos,quered,queremos,querer,querido,queriendo,querremos,querrá,querrán,querrás,querré,querréis,querría,querríais,querríamos,querrían,querrías,queráis,queréis,quería,queríais,queríamos,querían,querías,quia,quien,quienes,quienesquiera,quienquiera,quiera,quieran,quieras,quiere,quieren,quieres,quiero,quince,quincena,quincenas,quinceno,quincenos,quincuagésima,quincuagésimas,quincuagésimo,quincuagésimos,quingentésima,quingentésimas,quingentésimo,quingentésimos,quinientos,quinta,quintas,quinto,quintos,quise,quisiera,quisierais,quisieran,quisieras,quisiere,quisiereis,quisieren,quisieres,quisieron,quisiese,quisieseis,quisiesen,quisieses,quisimos,quisiste,quisisteis,quisiéramos,quisiéremos,quisiésemos,quiso,quién,quiénes,qué,ro,salud,salve,salvo,se,sea,seamos,sean,seas,sed,segunda,segundas,segundo,segundos,según,seis,seiscientos,seisena,seisenas,seiseno,seisenos,sendas,sendos,septena,septenas,septeno,septenos,septingentésima,septingentésimas,septingentésimo,septingentésimos,septuagésima,septuagésimas,septuagésimo,septuagésimos,ser,seremos,será,serán,serás,seré,seréis,sería,seríais,seríamos,serían,serías,sesenta,setecientos,setenta,sexagésima,sexagésimas,sexagésimo,sexagésimos,sexcentésima,sexcentésimas,sexcentésimo,sexcentésimos,sexta,sextas,sexto,sextos,seáis,si,sido,siendo,siete,sin,sino,siquiera,so,sobre,socorro,sois,solamos,solemos,soler,soliera,solierais,solieran,solieras,soliese,solieseis,soliesen,solieses,soliéramos,soliésemos,soláis,soléis,solía,solíais,solíamos,solían,solías,somos,son,soy,su,suelan,suelas,suele,suelen,sueles,sus,suya,suyas,suyo,suyos,sé,séptima,séptimas,séptimo,séptimos,sétima,sétimas,sétimo,sétimos,sí,ta,tal,tampoco,tanta,tantas,tanto,tantos,tate,te,tercer,tercera,terceras,tercero,terceros,tercia,terciaria,terciarias,terciario,terciarios,tercias,tercio,tercios,ti,toda,todas,todavía,todo,todos,tras,trece,trecena,trecenas,treceno,trecenos,tredécima,tredécimas,tredécimo,tredécimos,treinta,treintaidosena,treintaidosenas,treintaidoseno,treintaidosenos,treintena,treintenas,treinteno,treintenos,tres,trescientos,tricentésima,tricentésimas,tricentésimo,tricentésimos,trigésima,trigésimas,trigésimo,trigésimos,trillón,tu,tus,tuya,tuyas,tuyo,tuyos,tú,u,uf,un,una,unas,undécima,undécimas,undécimo,undécimos,uno,unos,upa,usted,ustedes,uy,vale,varias,varios,veinte,veintena,veintenas,veinteno,veintenos,veinteochena,veinteochenas,veinteocheno,veinteochenos,veinticinco,veinticuatrena,veinticuatrenas,veinticuatreno,veinticuatrenos,veinticuatro,veintidosena,veintidosenas,veintidoseno,veintidosenos,veintidós,veintinueve,veintiochena,veintiochenas,veintiocheno,veintiochenos,veintiocho,veintiseisena,veintiseisenas,veintiseiseno,veintiseisenos,veintisiete,veintiséis,veintitrés,veintiuno,veintésima,veintésimas,veintésimo,veintésimos,verbigracia,vigésima,vigésimas,vigésimo,vigésimos,viva,vos,vosotras,vosotros,vra.,vras.,vro.,vros.,vtra.,vtras.,vtro.,vtros.,vuestra,vuestras,vuestro,vuestros,vía,y,ya,yo,zape,zas,zis,zuzo,él,éramos,ésa,ésas,ése,ésos,ésta,éstas,éste,éstos,about,1,after,2,all,also,3,an,4,and,5,another,6,any,7,are,8,as,9,at,0,be,$,because,been,before,being,between,both,but,by,came,can,come,could,did,do,each,for,from,get,got,has,had,he,have,her,here,him,himself,his,how,if,in,into,is,it,like,make,many,me,might,more,most,much,must,my,never,now,of,on,only,or,other,our,out,over,said,same,see,should,since,some,still,such,take,than,that,the,their,them,then,there,these,they,this,those,through,to,too,under,up,very,was,way,we,well,were,what,where,which,while,who,with,would,you,your,a b c d e f g h i j k l m n o p q r s t u v w x y z""" & vbCrLf
sb.Append "	End Sub" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "	Public Property Get StopWords()" & vbCrLf
sb.Append "		StopWords = NoiseWords" & vbCrLf
sb.Append "	End Property" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "	Public Property Let StopWords(pValue)" & vbCrLf
sb.Append "		NoiseWords = pValue & """"" & vbCrLf
sb.Append "	End Property" & vbCrLf
sb.Append "	" & vbCrLf
sb.Append "	Public Function Error()" & vbCrLf
sb.Append "		Error = sErr" & vbCrLf
sb.Append "	End Function" & vbCrLf
sb.Append "	" & vbCrLf
sb.Append "	Public Sub AddBuiltIn(key, value, ttype)" & vbCrLf
sb.Append "		Dim token" & vbCrLf
sb.Append "		Set token = New clsToken" & vbCrLf
sb.Append "		token.ttype = ttype" & vbCrLf
sb.Append "		token.value = value" & vbCrLf
sb.Append "		builtIn.Add key, token" & vbCrLf
sb.Append "	End Sub" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "	Public Function ParseTokens(ByVal userSearch)" & vbCrLf
sb.Append "		Dim pos, parseOK, lastParsedToken" & vbCrLf
sb.Append "		Dim re, matchs, tokenStr, sAux" & vbCrLf
sb.Append "		Dim token, bAux" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		pos = 0" & vbCrLf
sb.Append "		parseOK = True" & vbCrLf
sb.Append "		lastParsedToken = Empty" & vbCrLf
sb.Append "		Set re = New RegExp" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "		userSearch = LCase(userSearch)" & vbCrLf
sb.Append "		userSearch = Replace(userSearch, ""'"", """""""")" & vbCrLf
sb.Append "" & vbCrLf
sb.Append " 		Do While userSearch & """" <> """"" & vbCrLf
sb.Append "	 		' The first search line does not handle internation characters, wheres the second one does" & vbCrLf
sb.Append "	 		' pos = userSearch.search(/\s*(\w+\*)|(\w+)|([""][^""]*[""])|([\(\)])\s*/);" & vbCrLf
sb.Append "           If dSession.Db.DbType = 6 Then ' SqlServer" & vbCrLf
sb.Append "                 re.Pattern = ""\s*([A-Za-z0-9_\u00C0-\u00FF]+\*)|([A-Za-z0-9_\u00C0-\u00FF]+)|([""""][^""""]*[""""])|([\(\)])\s*""" & vbCrLf
sb.Append "           ElseIf dSession.Db.DbType = 5 Then ' Oracle" & vbCrLf
sb.Append "                 re.Pattern = ""\s*([\?A-Za-z0-9_\u00C0-\u00FF]+\*)|([\?A-Za-z0-9_\u00C0-\u00FF]+)|([""""][^""""]*[""""])|([\(\)])\s*""" & vbCrLf
sb.Append "           Else" & vbCrLf
sb.Append "                 re.Pattern = ""\s*([\?A-Za-z0-9_\u00C0-\u00FF]+\*)|([\?A-Za-z0-9_\u00C0-\u00FF]+)|([""""][^""""]*[""""])|([\(\)])\s*""" & vbCrLf
sb.Append "           End If    	 		" & vbCrLf
sb.Append "           " & vbCrLf
sb.Append "	 		Set matchs = re.Execute(userSearch)" & vbCrLf
sb.Append "	 		If matchs.Count > 0 Then" & vbCrLf
sb.Append "	 			pos = matchs(0).FirstIndex" & vbCrLf
sb.Append "	 		Else" & vbCrLf
sb.Append "	 			pos = -1" & vbCrLf
sb.Append "	 		End If" & vbCrLf
sb.Append "	 		" & vbCrLf
sb.Append "	 		tokenStr = """"" & vbCrLf
sb.Append "	 		If pos >= 0 Then" & vbCrLf
sb.Append "				sAux = matchs(0).SubMatches(0)" & vbCrLf
sb.Append "				If sAux & """" <> """" Then" & vbCrLf
sb.Append "					tokenStr = """""""" & sAux & """"""""" & vbCrLf
sb.Append "				Else" & vbCrLf
sb.Append "					sAux = matchs(0).SubMatches(1)" & vbCrLf
sb.Append "					If sAux & """" <> """" Then" & vbCrLf
sb.Append "						tokenStr = sAux" & vbCrLf
sb.Append "					Else" & vbCrLf
sb.Append "						sAux = matchs(0).SubMatches(2)" & vbCrLf
sb.Append "						If sAux & """" <> """" Then" & vbCrLf
sb.Append "							tokenStr = sAux" & vbCrLf
sb.Append "						Else" & vbCrLf
sb.Append "							sAux = matchs(0).SubMatches(3)" & vbCrLf
sb.Append "							If sAux & """" <> """" Then" & vbCrLf
sb.Append "								tokenStr = sAux" & vbCrLf
sb.Append "							End If" & vbCrLf
sb.Append "						End If" & vbCrLf
sb.Append "					End If" & vbCrLf
sb.Append "				End If" & vbCrLf
sb.Append "	 			userSearch = Mid(userSearch, pos + 1 + matchs(0).Length)" & vbCrLf
sb.Append "	 		End If" & vbCrLf
sb.Append "	 		If tokenStr & """" <> """" Then" & vbCrLf
sb.Append "	 			Set token = GetToken(tokenStr)" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "				If IsObject(lastParsedToken) Then" & vbCrLf
sb.Append "					bAux = (lastParsedToken.ttype = NOISEWORD)" & vbCrLf
sb.Append "				Else" & vbCrLf
sb.Append "					bAux = False" & vbCrLf
sb.Append "				End If" & vbCrLf
sb.Append "				" & vbCrLf
sb.Append "				If IsObject(lastParsedToken) And bAux And ((token.ttype And BINARYOP) <> 0) Then" & vbCrLf
sb.Append "					' skip this token since it joins a noise word" & vbCrLf
sb.Append "				ElseIf token.ttype = NOISEWORD Then" & vbCrLf
sb.Append "					UnrollExpression(OPERATOR)" & vbCrLf
sb.Append "					Set lastParsedToken = token" & vbCrLf
sb.Append "				Else" & vbCrLf
sb.Append "					If token.ttype = USERITEM Then" & vbCrLf
sb.Append "						If IsObject(GetLastToken) Then" & vbCrLf
sb.Append "							bAux = ((GetLastToken.ttype And EXPRESSION) <> 0)" & vbCrLf
sb.Append "						Else" & vbCrLf
sb.Append "							bAux = False" & vbCrLf
sb.Append "						End If" & vbCrLf
sb.Append "						If IsObject(GetLastToken) And bAux Then" & vbCrLf
sb.Append "							InsertDefaultOperator" & vbCrLf
sb.Append "						End If" & vbCrLf
sb.Append "					Else" & vbCrLf
sb.Append "						If IsObject(GetLastToken) Then" & vbCrLf
sb.Append "							bAux = (GetLastToken.ttype And EXPRESSION)" & vbCrLf
sb.Append "						Else" & vbCrLf
sb.Append "							bAux = False" & vbCrLf
sb.Append "						End If" & vbCrLf
sb.Append "						If token.ttype = NOTOP And IsObject(GetLastToken) And bAux Then" & vbCrLf
sb.Append "							AddToken builtIn(""and"")" & vbCrLf
sb.Append "						End If" & vbCrLf
sb.Append "					End If" & vbCrLf
sb.Append "					" & vbCrLf
sb.Append "					AddToken token" & vbCrLf
sb.Append " 					Set lastParsedToken = token" & vbCrLf
sb.Append " 				End If" & vbCrLf
sb.Append " 			Else" & vbCrLf
sb.Append " 				re.Pattern = ""\S""" & vbCrLf
sb.Append " 				If re.Test(userSearch) Then" & vbCrLf
sb.Append " 					parseOK = False" & vbCrLf
sb.Append "					sErr = ""Illegal character found in search string: '"" + userSearch + ""'""" & vbCrLf
sb.Append " 					Exit Do" & vbCrLf
sb.Append " 				Else" & vbCrLf
sb.Append " 					userSearch = """"" & vbCrLf
sb.Append " 				End If" & vbCrLf
sb.Append "			End If" & vbCrLf
sb.Append " 		Loop" & vbCrLf
sb.Append "		ParseTokens = parseOK" & vbCrLf
sb.Append "	End Function" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "	Function Validate()" & vbCrLf
sb.Append "		Dim valid, lastItemOK, nextItem, balance" & vbCrLf
sb.Append "		Dim tokIndex" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		valid = True" & vbCrLf
sb.Append "		lastItemOK = False" & vbCrLf
sb.Append "		nextItem = USERITEM Or LEFTPAREN Or NOTOP" & vbCrLf
sb.Append "		balance = 0" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		If Not IsArray(tokens) Then" & vbCrLf
sb.Append "			valid = False" & vbCrLf
sb.Append "			sErr = ""Empty search string, probably because all input words were noise words ('a','the', etc.).""" & vbCrLf
sb.Append "		Else" & vbCrLf
sb.Append "			For tokIndex = 0 To UBound(tokens)" & vbCrLf
sb.Append "				If (tokens(tokIndex).ttype And nextItem) <> 0 Then" & vbCrLf
sb.Append "					Select Case tokens(tokIndex).ttype" & vbCrLf
sb.Append "						Case USERITEM" & vbCrLf
sb.Append "							nextItem = BINARYOP Or RIGHTPAREN" & vbCrLf
sb.Append "							lastItemOK = True" & vbCrLf
sb.Append "						Case ANDOP" & vbCrLf
sb.Append "							nextItem = USERITEM Or NOTOP Or LEFTPAREN" & vbCrLf
sb.Append "							lastItemOK = False" & vbCrLf
sb.Append "						Case NEAROP" & vbCrLf
sb.Append "							nextItem = USERITEM" & vbCrLf
sb.Append "							lastItemOK = False" & vbCrLf
sb.Append "						Case OROP" & vbCrLf
sb.Append "							nextItem = USERITEM Or LEFTPAREN" & vbCrLf
sb.Append "							lastItemOK = False" & vbCrLf
sb.Append "						Case NOTOP" & vbCrLf
sb.Append "							nextItem = USERITEM Or LEFTPAREN" & vbCrLf
sb.Append "							lastItemOK = False" & vbCrLf
sb.Append "						Case LEFTPAREN" & vbCrLf
sb.Append "							balance = balance + 1" & vbCrLf
sb.Append "							nextItem = USERITEM" & vbCrLf
sb.Append "							lastItemOK = False" & vbCrLf
sb.Append "						Case RIGHTPAREN" & vbCrLf
sb.Append "							balance = balance - 1" & vbCrLf
sb.Append "							nextItem = OROP Or ANDOP" & vbCrLf
sb.Append "							If balance > 0 Then" & vbCrLf
sb.Append "								lastItemOK = False" & vbCrLf
sb.Append "							Else" & vbCrLf
sb.Append "								lastItemOK = True" & vbCrLf
sb.Append "							End If" & vbCrLf
sb.Append "					End Select" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "					If balance < 0 Then" & vbCrLf
sb.Append "						valid = False" & vbCrLf
sb.Append "						sErr = ""Mismatched parenthesis""" & vbCrLf
sb.Append "						Exit For" & vbCrLf
sb.Append "					End If" & vbCrLf
sb.Append "					" & vbCrLf
sb.Append "				Else" & vbCrLf
sb.Append "					valid = False" & vbCrLf
sb.Append "					sErr = ""Unexpected word or character found: "" & tokens(tokIndex).value" & vbCrLf
sb.Append "					Exit For" & vbCrLf
sb.Append "				End If" & vbCrLf
sb.Append "			Next" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "			If balance <> 0 Then" & vbCrLf
sb.Append "				valid = False" & vbCrLf
sb.Append "				sErr = ""Mismatched parenthesis""" & vbCrLf
sb.Append "			ElseIf valid And Not lastItemOK Then" & vbCrLf
sb.Append "				valid = False" & vbCrLf
sb.Append "				sErr = ""Unexpected end of search string after: "" & tokens(UBound(tokens)).value" & vbCrLf
sb.Append "			End If" & vbCrLf
sb.Append "		End If" & vbCrLf
sb.Append "		Validate = valid" & vbCrLf
sb.Append "	End Function" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "	Public Function GetToken(tokenStr)" & vbCrLf
sb.Append "		Dim token" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		If IsObject(builtIn(tokenStr)) Then" & vbCrLf
sb.Append "			Set token = builtIn(tokenStr)" & vbCrLf
sb.Append "		Else" & vbCrLf
sb.Append "			Set token = New clsToken" & vbCrLf
sb.Append "			If InStr(1, NoiseWords, "","" & tokenStr & "","", vbTextCompare) > 0 Then" & vbCrLf
sb.Append "				token.ttype = NOISEWORD" & vbCrLf
sb.Append "			Else" & vbCrLf
sb.Append "				token.ttype = USERITEM" & vbCrLf
sb.Append "			End If" & vbCrLf
sb.Append "			token.value = tokenStr" & vbCrLf
sb.Append "		End If" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		Set GetToken = token" & vbCrLf
sb.Append "	End Function" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "	Public Function GetLastToken()" & vbCrLf
sb.Append "		If IsArray(tokens) Then" & vbCrLf
sb.Append "			Set GetLastToken = tokens(UBound(tokens))" & vbCrLf
sb.Append "		Else" & vbCrLf
sb.Append "			GetLastToken = Empty" & vbCrLf
sb.Append "		End If" & vbCrLf
sb.Append "	End Function" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "	Public Sub AddToken(token)" & vbCrLf
sb.Append "		Dim i" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		If IsArray(tokens) Then" & vbCrLf
sb.Append "			i = UBound(tokens) + 1" & vbCrLf
sb.Append "			ReDim Preserve tokens(i)" & vbCrLf
sb.Append "		Else" & vbCrLf
sb.Append "			i = 0" & vbCrLf
sb.Append "			ReDim tokens(i)" & vbCrLf
sb.Append "		End If" & vbCrLf
sb.Append "		Set tokens(i) = token" & vbCrLf
sb.Append "	End Sub" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "	Public Function GetMSSQLSearchStr()" & vbCrLf
sb.Append "		Dim searchStr, i" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		searchStr = tokens(0).value" & vbCrLf
sb.Append "		For i = 1 To UBound(tokens)" & vbCrLf
sb.Append "			searchStr = searchStr & "" "" & tokens(i).value" & vbCrLf
sb.Append "		Next" & vbCrLf
sb.Append "		GetMSSQLSearchStr = searchStr" & vbCrLf
sb.Append "	End Function" & vbCrLf
sb.Append "	" & vbCrLf
sb.Append "	Public Sub InsertDefaultOperator()" & vbCrLf
sb.Append "		AddToken builtIn(""and"")" & vbCrLf
sb.Append "	End Sub" & vbCrLf
sb.Append "	" & vbCrLf
sb.Append "	' This function searches back until it reaches the first uType token truncating any" & vbCrLf
sb.Append "	' tokens that follow it" & vbCrLf
sb.Append "	Public Sub UnrollExpression(uType)" & vbCrLf
sb.Append "		Dim newLength, i" & vbCrLf
sb.Append "		" & vbCrLf
sb.Append "		If IsArray(tokens) Then" & vbCrLf
sb.Append "			newLength = UBound(tokens)" & vbCrLf
sb.Append "			For i = UBound(tokens) To 0" & vbCrLf
sb.Append "				If (tokens(i).ttype And uType) <> 0 Then" & vbCrLf
sb.Append "					newLength = i" & vbCrLf
sb.Append "				Else" & vbCrLf
sb.Append "					Exit For" & vbCrLf
sb.Append "				End If" & vbCrLf
sb.Append "			Next" & vbCrLf
sb.Append "			" & vbCrLf
sb.Append "			ReDim Preserve tokens(newLength)" & vbCrLf
sb.Append "		End If" & vbCrLf
sb.Append "	End Sub" & vbCrLf
sb.Append "End Class"
newDoc("CODE").Value = sb.ToString
On Error Resume Next
newDoc("TASK_ID").Value = Null
newDoc("TASK_RESPONSIBLE").Value = ""
newDoc("TASK_RESPONSIBLEID").Value = Null
newDoc("TASK_TYPE").Value = ""
newDoc("TASK_TYPEID").Value = Null
newDoc("TASK_STATE").Value = ""
newDoc("TASK_STATEID").Value = Null
newDoc("TASK_DUEDATE").Value = Null
newDoc("TASK_CUSTOMER").Value = ""
newDoc("TASK_CUSTOMERID").Value = Null
newDoc("TASK_SOURCE").Value = ""
newDoc("WORKFLOW_ID").Value = ""
On Error Goto 0
newDoc.Save


' Root folder
Set rootFolder = dSession.FoldersGetFromId(1001)


'---
'--- Scripts for folder /crm_root
'---

Set curFolder = rootFolder.App.Folders("/crm_root")


'---
'--- Scripts for folder /crm_root/config
'---

Set curFolder = rootFolder.App.Folders("/crm_root/config")


'---
'--- Scripts for folder /crm_root/config/mails
'---

Set curFolder = rootFolder.App.Folders("/crm_root/config")

On Error Resume Next
Set newFolder = curFolder.Folders("mails")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newFolder = curFolder.FoldersNew
	newFolder.Name = "mails"
	newFolder.FolderType = 1
End If
newFolder.DescriptionRaw = ""
newFolder.Comments = ""
newFolder.CharData = ""
newFolder.IconRaw = "email-edit"
sAux = "<?xml version=""1.0""?>" & vbCrLf & "<root xmlns=""LogConf""/>" & vbCrLf
If IsG7 Then 
	Set oDom = dSession.Xml.NewDom()
	oDom.loadXML sAux
	Set newFolder.LogConf = oDom
Else
	newFolder.LogConf.loadXML sAux
End If 
newFolder.Form = dSession.Forms("b1ac88af09484dd7bc9efc88912fa358")
newFolder.Save


Set curFolder = rootFolder.App.Folders("/crm_root/config/mails")


'--- Folder Acl

curFolder.AclInherits = True


'--- Event Document_Open

Set oEvn = curFolder.Events("ID=1")
oEvn.Code = ""
oEvn.Overrides = False
curFolder.Save


'--- Event Document_BeforeSave

Set oEvn = curFolder.Events("ID=2")
oEvn.Code = ""
oEvn.Overrides = False
curFolder.Save


'--- Document doc_id = 453639

Set newDoc = curFolder.DocumentsNew
newDoc("NAME").Value = "AnexoTejidoSoloMadre"
newDoc("TEMPLATEBODY").Value = ""
newDoc("TEMPLATESUBJECT").Value = ""
newDoc("LINKTODOC").Value = Null
newDoc("MODULO").Value = ""
newDoc("DESCRIPTION").Value = ""
newDoc("QUICKMAIL").Value = Null
Set sb = dSession.ConstructNewStringBuilder
sb.Append "#include /config/codigo/functions" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "Tokens.Add ""[@MAMA]"", NombreContrato(Doc(""REPRESENTANTE"").Value & """")" & vbCrLf
sb.Append "Tokens.Add ""[@RUTMAMA]"", Doc(""REPRES_RUT"").Value & """"	" & vbCrLf
sb.Append "Tokens.Add ""[@Contrato]"", Doc(""CONTRATO"").Value & """"	" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "Tokens.Add ""[@javascript]"", <#" & vbCrLf
sb.Append "<style type=""text/css"">" & vbCrLf
sb.Append "@page {" & vbCrLf
sb.Append "	margin-top: 4cm;" & vbCrLf
sb.Append "	margin-bottom: 3cm;" & vbCrLf
sb.Append "	margin-right: 2cm;" & vbCrLf
sb.Append "	margin-left: 2cm;" & vbCrLf
sb.Append "    " & vbCrLf
sb.Append "    counter-increment: page;" & vbCrLf
sb.Append "	@bottom-right {" & vbCrLf
sb.Append "		padding-right:20px;" & vbCrLf
sb.Append "		content: ""Page "" counter(page);" & vbCrLf
sb.Append "	}" & vbCrLf
sb.Append "}" & vbCrLf
sb.Append "</style>" & vbCrLf
sb.Append "#>" & vbCrLf
sb.Append ""
newDoc("CODE").Value = sb.ToString
newDoc("ATTACHMENTS").Value = Null
newDoc("TEMPLATEFORMAT").Value = Null
Set sb = dSession.ConstructNewStringBuilder
sb.Append "<p style=""text-align:center""><strong><span style=""font-size:16px""><span style=""font-family:arial,helvetica,sans-serif"">ALMACENAMIENTO DE TEJIDO DE CORD&Oacute;N UMBILICAL</span></span></strong></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p style=""text-align:right""><span style=""font-family:arial,helvetica,sans-serif""><span style=""font-size:12px"">ANEXO CONTRATO: [@contrato]</span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">En conjunto con la Crio preservacion de la sangre del cord&oacute;n umbilical /placenta, se almacenar&aacute; el tejido de cord&oacute;n umbilical tambi&eacute;n llamado Gelatina de Wharton.</span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">VALORES:</span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<ul>" & vbCrLf
sb.Append "	<li><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">20,5 UF + IVA</span></span></li>" & vbCrLf
sb.Append "	<li><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">Almacenamiento Anual : 2 UF + IVA a partir del 2&deg; a&ntilde;o de mantenci&oacute;n</span></span></li>" & vbCrLf
sb.Append "	<li><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">Los precios son en Pesos al valor de la UF del d&iacute;a en que se efect&uacute;e el pago.</span></span></li>" & vbCrLf
sb.Append "</ul>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p><strong>EJEMPLARES </strong></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">El&nbsp; presente&nbsp; Anexo&nbsp; de&nbsp;&nbsp; Contrato&nbsp; es&nbsp; otorgado&nbsp; en&nbsp; dos&nbsp; ejemplares&nbsp; de&nbsp; igual&nbsp; tenor&nbsp; y&nbsp; fecha, quedando un ejemplar en poder de cada una de las partes.</span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>______________________________ &nbsp;&nbsp;&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">Nombre Madre Ni&ntilde;o por nacer: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">[@mama]</span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">RUT: [@rutmama]</span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <img alt="""" src=""data:image/png;base64,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"" style=""height:130px; width:130px"" /></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>[@javascript]</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p style=""text-align:right"">&nbsp;</p>" & vbCrLf
sb.Append ""
newDoc("BODY").Value = sb.ToString
On Error Resume Next
newDoc("TASK_ID").Value = Null
newDoc("TASK_RESPONSIBLE").Value = ""
newDoc("TASK_RESPONSIBLEID").Value = Null
newDoc("TASK_TYPE").Value = ""
newDoc("TASK_TYPEID").Value = Null
newDoc("TASK_STATE").Value = ""
newDoc("TASK_STATEID").Value = Null
newDoc("TASK_DUEDATE").Value = Null
newDoc("TASK_CUSTOMER").Value = ""
newDoc("TASK_CUSTOMERID").Value = Null
newDoc("TASK_SOURCE").Value = ""
newDoc("WORKFLOW_ID").Value = ""
On Error Goto 0
newDoc.Save


'--- Document doc_id = 562368

Set newDoc = curFolder.DocumentsNew
newDoc("NAME").Value = "ConsentimientoPulpa"
newDoc("TEMPLATEBODY").Value = ""
newDoc("TEMPLATESUBJECT").Value = ""
newDoc("LINKTODOC").Value = Null
newDoc("MODULO").Value = ""
newDoc("DESCRIPTION").Value = ""
newDoc("QUICKMAIL").Value = Null
Set sb = dSession.ConstructNewStringBuilder
sb.Append "#include /config/codigo/functions" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "Tokens.Add ""[@FECHACONTRATO]"", FechaTexto(Doc(""FECHACIERRE"").Value)" & vbCrLf
sb.Append "Tokens.Add ""[@PACIENTE]"", NombreContrato(Doc(""PACIENTE"").Value & """")" & vbCrLf
sb.Append "Tokens.Add ""[@RutPaciente]"", Doc(""PACIENTE_RUT"").Value & """"" & vbCrLf
sb.Append "Tokens.Add ""[@DireccionPaciente]"", Doc(""PACIENTE_DIRECCION"").Value & """"" & vbCrLf
sb.Append "Tokens.Add ""[@CONTRATO]"", Doc(""Contrato"").Value & """"" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "Tokens.Add ""[@javascript]"", <#" & vbCrLf
sb.Append "<style type=""text/css"">" & vbCrLf
sb.Append "@page {" & vbCrLf
sb.Append "	margin-top: 4cm;" & vbCrLf
sb.Append "	margin-bottom: 3cm;" & vbCrLf
sb.Append "	margin-right: 2cm;" & vbCrLf
sb.Append "	margin-left: 2cm;" & vbCrLf
sb.Append "    " & vbCrLf
sb.Append "    counter-increment: page;" & vbCrLf
sb.Append "	@bottom-right {" & vbCrLf
sb.Append "		padding-right:20px;" & vbCrLf
sb.Append "		content: ""Page "" counter(page);" & vbCrLf
sb.Append "	}" & vbCrLf
sb.Append "}" & vbCrLf
sb.Append "</style>" & vbCrLf
sb.Append "#>"
newDoc("CODE").Value = sb.ToString
newDoc("ATTACHMENTS").Value = Null
newDoc("TEMPLATEFORMAT").Value = Null
Set sb = dSession.ConstructNewStringBuilder
sb.Append "<p style=""text-align:center""><span style=""font-size:16px""><strong>CARTA CONSENTIMIENTO Y COMPROMISO PACIENTE</strong></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p style=""text-align:right""><strong>Contrato: [@contrato]</strong></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>Yo [@PACIENTE], RUT [@RutPaciente], domiciliado en&nbsp; [@DireccionPaciente] comuna &nbsp;de [@comuna], declaro:</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>Por la presente DECLARO, que me han informado sobre el procedimiento para la recolecci&oacute;n de las c&eacute;lulas madre de la pulpa dental, y damos nuestro consentimiento para permitir que nuestro odont&oacute;logo recoja la pulpa dental que contiene dichas c&eacute;lulas madre y sean enviadas a VIDACEL para su evaluaci&oacute;n, procesamiento, criopreservaci&oacute;n y almacenamiento. Las muestras ser&aacute;n conservadas para su potencial uso en el futuro, siendo el due&ntilde;o de la muestra el paciente, o su representante legal en caso de ser menor de edad hasta su mayor&iacute;a de edad.</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>A su vez me comprometo a:</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>1. Realizar el contrato de criopreservaci&oacute;n y almacenamiento de c&eacute;lulas madre con la empresa VidaCel.&nbsp; Para lo cual, firmar&eacute; el contrato y pagar&eacute; el servicio correspondiente a UF 29 + IVA, de acuerdo a las condiciones y formas de pago vigentes. El servicio de VidaCel no incluye los honorarios del odont&oacute;logo.</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>2. En el caso que la muestra recolectada sea inadecuada para ser criopreservada, se dar&aacute; aviso al Cliente y los valores cancelados quedar&aacute;n abonados para una nueva muestra, o bien, ser&aacute;n restituidos<a name=""_GoBack""></a>.</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<div>&nbsp;</div>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p style=""text-align:right"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style=""font-family:arial,helvetica,sans-serif""><span style=""font-size:12px"">&nbsp;&nbsp;&nbsp;<span style=""font-size:11px"">&nbsp;&nbsp;&nbsp; Santiago de Chile, [@fechacierre]</span></span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<table cellpadding=""15"" style=""height:188px; width:595px"">" & vbCrLf
sb.Append "	<tbody>" & vbCrLf
sb.Append "		<tr>" & vbCrLf
sb.Append "			<td style=""text-align:center"">" & vbCrLf
sb.Append "			<p style=""text-align:left""><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">_____________________________</span></span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "			<p style=""text-align:left""><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">Nombre Cliente </span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "			<p style=""text-align:left""><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">[@MAMA]</span></span></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "			<p style=""text-align:left""><span style=""font-size:12px""><span style=""font-family:arial,helvetica,sans-serif"">RUT: [@RUTMAMA]</span></span></p>" & vbCrLf
sb.Append "			</td>" & vbCrLf
sb.Append "			<td style=""text-align:center"">" & vbCrLf
sb.Append "			<p style=""text-align:left"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "			<p style=""text-align:left"">&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "			<p style=""text-align:left"">&nbsp;</p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "			<p style=""text-align:left"">&nbsp;</p>" & vbCrLf
sb.Append "			</td>" & vbCrLf
sb.Append "		</tr>" & vbCrLf
sb.Append "	</tbody>" & vbCrLf
sb.Append "</table>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; <img alt="""" src=""data:image/png;base64,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"" style=""height:132px; width:132px"" /></p>" & vbCrLf
sb.Append "" & vbCrLf
sb.Append "<p>[@javascript]</p>" & vbCrLf
sb.Append ""
newDoc("BODY").Value = sb.ToString
On Error Resume Next
newDoc("TASK_ID").Value = Null
newDoc("TASK_RESPONSIBLE").Value = ""
newDoc("TASK_RESPONSIBLEID").Value = Null
newDoc("TASK_TYPE").Value = ""
newDoc("TASK_TYPEID").Value = Null
newDoc("TASK_STATE").Value = ""
newDoc("TASK_STATEID").Value = Null
newDoc("TASK_DUEDATE").Value = Null
newDoc("TASK_CUSTOMER").Value = ""
newDoc("TASK_CUSTOMERID").Value = Null
newDoc("TASK_SOURCE").Value = ""
newDoc("WORKFLOW_ID").Value = ""
On Error Goto 0
newDoc.Save


'--- View RECIENTES

On Error Resume Next
Set newView = curFolder.Views("Recientes")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newView = curFolder.ViewsNew
	newView.Name = "Recientes"
End If
newView.DescriptionRaw = ""
newView.Comments = ""
Set oDom = dSession.Xml.NewDom()
oDom.loadXML"<root customizedformula=""0"" viewtype=""1"" xmlns=""viewDefinition""><fields><item field=""name"" width=""0"" description="""" format="""" isimage=""0""/><item field=""subject"" width=""0"" description="""" format="""" isimage=""0""/><item field=""body"" width=""0"" maxlength=""300"" description="""" format=""RemoveHtml"" isimage=""0""/></fields><groups/><orders><item field=""accessed"" direction=""1""/></orders><filters/></root>" & vbCrLf
oDom.setProperty "SelectionNamespaces", "xmlns:d=""viewDefinition"""
If IsG7 Then
	Set newView.Definition = oDom
Else
	oDom.documentElement.removeAttribute("viewtype")
	For Each node In oDom.selectNodes("/d:root/d:fields/d:item")
		node.removeAttribute("isimage")
	Next
	For Each node in oDom.selectNodes("/d:root/d:groups/d:item")
		node.removeAttribute("direction")
		node.removeAttribute("orderby")
	Next
	newView.Definition.loadXml oDom.Xml
End If
newView.Save


'---
'--- Scripts for folder /crm_root/oportunidades
'---

Set curFolder = rootFolder.App.Folders("/crm_root")

On Error Resume Next
Set newFolder = curFolder.Folders("oportunidades")
ErrNumber = Err.Number
On Error GoTo 0
If ErrNumber <> 0 Then
	Set newFolder = curFolder.FoldersNew
	newFolder.Name = "oportunidades"
	newFolder.FolderType = 1
End If
newFolder.DescriptionRaw = "Oportunidades"
newFolder.Comments = ""
newFolder.CharData = "<root xmlns=""CharData""><item name=""default_view"" id=""0""/></root>" & vbCrLf
newFolder.IconRaw = "folder-page"
sAux = "<?xml version=""1.0""?>" & vbCrLf & "<root xmlns=""LogConf""><item field=""id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""producto"" log=""1"" old_value=""1"" new_value=""1""/><item field=""subproductos"" log=""1"" old_value=""1"" new_value=""1""/><item field=""cliente"" log=""1"" old_value=""1"" new_value=""1""/><item field=""cliente_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""paciente"" log=""1"" old_value=""1"" new_value=""1""/><item field=""paciente_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""representante"" log=""1"" old_value=""1"" new_value=""1""/><item field=""representante_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""representante2"" log=""1"" old_value=""1"" new_value=""1""/><item field=""representante2_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""clinica"" log=""1"" old_value=""1"" new_value=""1""/><item field=""clinica_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""isapre"" log=""1"" old_value=""1"" new_value=""1""/><item field=""isapre_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""medico"" log=""1"" old_value=""1"" new_value=""1""/><item field=""medico_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""estado"" log=""1"" old_value=""1"" new_value=""1""/><item field=""referencia"" log=""1"" old_value=""1"" new_value=""1""/><item field=""referencia_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""cajacompensacion"" log=""1"" old_value=""1"" new_value=""1""/><item field=""cajacompensacion_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""laboratorio"" log=""1"" old_value=""1"" new_value=""1""/><item field=""laboratorio_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""preciolista"" log=""1"" old_value=""1"" new_value=""1""/><item field=""preciofinal"" log=""1"" old_value=""1"" new_value=""1""/><item field=""fechacreacion"" log=""1"" old_value=""1"" new_value=""1""/><item field=""fechapago"" log=""1"" old_value=""1"" new_value=""1""/><item field=""fechamuestra"" log=""1"" old_value=""1"" new_value=""1""/><item field=""fechaprobmuestra"" log=""1"" old_value=""1"" new_value=""1""/><item field=""region"" log=""1"" old_value=""1"" new_value=""1""/><item field=""origen"" log=""1"" old_value=""1"" new_value=""1""/><item field=""proxaccion"" log=""1"" old_value=""1"" new_value=""1""/><item field=""fechaproxaccion"" log=""1"" old_value=""1"" new_value=""1""/><item field=""contrato"" log=""1"" old_value=""1"" new_value=""1""/><item field=""observaciones"" log=""1"" old_value=""1"" new_value=""1""/><item field=""motivonoventa"" log=""1"" old_value=""1"" new_value=""1""/><item field=""formapago"" log=""1"" old_value=""1"" new_value=""1""/><item field=""embarazonro"" log=""1"" old_value=""1"" new_value=""1""/><item field=""fechacierre"" log=""1"" old_value=""1"" new_value=""1""/><item field=""cliente_telefono"" log=""1"" old_value=""1"" new_value=""1""/><item field=""repres_telefono"" log=""1"" old_value=""1"" new_value=""1""/><item field=""repres_email"" log=""1"" old_value=""1"" new_value=""1""/><item field=""repres_direccion"" log=""1"" old_value=""1"" new_value=""1""/><item field=""cliente_email"" log=""1"" old_value=""1"" new_value=""1""/><item field=""cliente_direccion"" log=""1"" old_value=""1"" new_value=""1""/><item field=""cliente_comuna"" log=""1"" old_value=""1"" new_value=""1""/><item field=""repres_comuna"" log=""1"" old_value=""1"" new_value=""1""/><item field=""repres_rut"" log=""1"" old_value=""1"" new_value=""1""/><item field=""cliente_rut"" log=""1"" old_value=""1"" new_value=""1""/><item field=""origen_detalle"" log=""1"" old_value=""1"" new_value=""1""/><item field=""interes"" log=""1"" old_value=""1"" new_value=""1""/><item field=""comentarios"" log=""1"" old_value=""1"" new_value=""1""/><item field=""lead"" log=""1"" old_value=""1"" new_value=""1""/><item field=""fechatomadato"" log=""1"" old_value=""1"" new_value=""1""/><item field=""nrokit"" log=""1"" old_value=""1"" new_value=""1""/><item field=""llamadoutm1"" log=""1"" old_value=""1"" new_value=""1""/><item field=""llamadoutm2"" log=""1"" old_value=""1"" new_value=""1""/><item field=""ejecutivo"" log=""1"" old_value=""1"" new_value=""1""/><item field=""ejecutivo_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""team"" log=""1"" old_value=""1"" new_value=""1""/><item field=""team_id"" log=""1"" old_value=""1"" new_value=""1""/><item field=""subproductos_xml"" log=""1"" old_value=""1"" new_value=""1""/><item field=""descuentoaprobado"" log=""1"" old_value=""1"" new_value=""1""/><item field=""reqaprobacion"" log=""1"" old_value=""1"" new_value=""1""/><item field=""referencia_contrato"" log=""1"" old_value=""1"" new_value=""1""/><item field=""subestado"" log=""1"" old_value=""1"" new_value=""1""/><item field=""fechaproxvisitaclinica"" log=""1"" old_value=""1"" new_value=""1""/><item field=""repres2_rut"" log=""1"" old_value=""1"" new_value=""1""/></root>" & vbCrLf
If IsG7 Then 
	Set oDom = dSession.Xml.NewDom()
	oDom.loadXML sAux
	Set newFolder.LogConf = oDom
Else
	newFolder.LogConf.loadXML sAux
End If 
newFolder.Form = dSession.Forms("7f8cfebc9bed4464a7004c2451ccc9e2")
newFolder.Save


Set curFolder = rootFolder.App.Folders("/crm_root/oportunidades")


dSession.ClearAllCustomCache
dSession.ClearObjectModelCache "ComCodeLibCache"


Dim domAccounts
Function AccountId(pName)
	If IsEmpty(domAccounts) Then Set domAccounts = dSession.Directory.AccountsList
	AccountId = domAccounts.selectSingleNode("/d:root/d:item[@name='" & pName & "']").getAttribute("id")
End Function


'--------------------------------------
' Fin del archivo de instalacion
' Finalizado: 10/11/2015 15:40:37
'--------------------------------------