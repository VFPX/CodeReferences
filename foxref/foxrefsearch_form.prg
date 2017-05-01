* Abstract...:
*	Search / Replace functionality for a Form or Class Libary (SCX, VCX).
*
* Changes....:
*
#include "foxref.h"
#include "foxpro.h"

Define Class RefSearchForm As RefSearch Of FoxRefSearch.prg
	Name = "RefSearchForm"

	Norefresh        = .F.
	BackupExtensions = "SCX,SCT"

	Function OpenFile(lReadWrite)
		Local oError

		m.oError = .Null.

		If Used(TABLEFIND_ALIAS)
			Use In (TABLEFIND_ALIAS)
		Endif

		Try
			If m.lReadWrite
				Use (This.Filename) Alias TABLEFIND_ALIAS In 0 Shared Again
			Else
				Use (This.Filename) Alias TABLEFIND_ALIAS In 0 Shared Again Noupdate
			Endif

		Catch To oError
		Endtry

		If Isnull(m.oError) And Used(TABLEFIND_ALIAS)
			Try
				If ;
						TYPE(TABLEFIND_ALIAS + ".Methods") == 'M' And ;
						TYPE(TABLEFIND_ALIAS + ".Platform") == 'C' And ;
						TYPE(TABLEFIND_ALIAS + ".UniqueID") == 'C'
					Select TABLEFIND_ALIAS
				Else
					Use In (TABLEFIND_ALIAS)
					Throw ERROR_NOTFORM_LOC
				Endif

			Catch To oError When oError.ErrorNo == 2071
			Catch To oError
			Endtry
		Endif

		Return m.oError
	Endfunc

	Function CloseFile()
		If Used(TABLEFIND_ALIAS)
			Use In (TABLEFIND_ALIAS)
		Endif

		Return .T.
	Endfunc

	Function DoDefinitions()
		Local nSelect
		Local cObjName
		Local cClassName
		Local cRootClass

		nSelect = Select()

		cRootClass = .Null.

		Select TABLEFIND_ALIAS

		Scan All For PlatForm = "WINDOWS" And !Empty(Methods)
			If Reserved1 == "Class"
				cRootClass = ObjName
			Endif

			If Empty(Parent)
				If BaseClass == "dataenvironment"
					cObjName = ObjName
				Else
					cObjName = ''
				Endif
				cClassName = Class
			Else
				If At_c('.', Parent) == 0
					cClassName = Parent
					cObjName = ObjName
				Else
					cClassName = Leftc(Parent, At_c('.', Parent) - 1)
					cObjName = Substrc(Parent, At_c('.', Parent) + 1) + '.' + ObjName
				Endif
			Endif
			This.FindDefinitions(Methods, Nvl(cRootClass, cClassName), cObjName, SEARCHTYPE_NORMAL)
		Endscan

		Select (nSelect)
	Endfunc

	Function DoSearch()
		Local nSelect
		Local cObjName
		Local cClassName
		Local cRootClass
		Local lSuccess

		nSelect = Select()

		cRootClass = .Null.

		m.lSuccess = .T.

		Select TABLEFIND_ALIAS

		If Reccount() > 0 And !Empty(Reserved8)
			&& this is where global include file is stored for a form
			This.AddFileToProcess( ;
				DEFTYPE_INCLUDEFILE, ;
				Reserved8, ;
				'', ;
				'', ;
				1, ;
				1, ;
				Reserved8 ;
				)
			m.lSuccess = This.FindInText(Reserved8, FINDTYPE_TEXT, '', INCLUDENAME_LOC, SEARCHTYPE_NORMAL, UniqueID, "RESERVED8")
		Endif

		Scan All For PlatForm = "WINDOWS"
			*** JRN 2010-03-20 : Save timestamp for this obect
			This.ObjectTimeStamp = Timestamp
			If Reserved1 == "Class"
				cRootClass = ObjName

				If !Empty(Reserved8) && this is where global include file is stored for a class
					This.AddFileToProcess( ;
						DEFTYPE_INCLUDEFILE, ;
						Reserved8, ;
						NVL(cRootClass, cClassName), ;
						'', ;
						1, ;
						1, ;
						Reserved8 ;
						)

					m.lSuccess = This.FindInText(Reserved8, FINDTYPE_TEXT, Nvl(cRootClass, cClassName), INCLUDENAME_LOC, SEARCHTYPE_NORMAL, UniqueID, "RESERVED8")
				Endif
			Endif


			If Empty(Parent)
				If BaseClass == "dataenvironment"
					cObjName = ObjName
				Else
					cObjName = ''
				Endif
				cClassName = Class
			Else
				If At_c('.', Parent) == 0
					cClassName = Parent
					cObjName = ObjName
				Else
					cClassName = Leftc(Parent, At_c('.', Parent) - 1)
					cObjName = Substrc(Parent, At_c('.', Parent) + 1) + '.' + ObjName
				Endif
			Endif

			* defined properties and methods
			*!*				IF !EMPTY(Reserved3)
			*!*					m.lSuccess = THIS.FindDefinedPEMS(Reserved3, NVL(cRootClass, cClassName), cObjName, SEARCHTYPE_EXPR, UniqueID)
			*!*				ENDIF

			If !Empty(Reserved3)
				m.lSuccess = This.FindDefinedMethods(Reserved3, Nvl(cRootClass, cClassName), cObjName, SEARCHTYPE_EXPR, UniqueID)
			Endif

			If !Empty(Methods)
				m.lSuccess = This.FindInCode(Methods, FINDTYPE_CODE, Nvl(cRootClass, cClassName), cObjName, SEARCHTYPE_METHOD, UniqueID, "METHODS")
			Endif

			If !This.CodeOnly
				If !Empty(ObjName)
					If Reserved1 == "Class"
						m.lSuccess = This.FindInText('Name = ' + ObjName, FINDTYPE_NAME, Nvl(cRootClass, cClassName), CLASSDEF_LOC, SEARCHTYPE_EXPR, UniqueID, "OBJNAME", .T.)
					Else
						*	m.lSuccess = This.FindInText('Name = ' + ObjName, FINDTYPE_NAME, Nvl(cRootClass, cClassName), IIF('.' $ Parent, Substr (Parent, 1 + At('.', Parent)) + '.', '') + OBJECTDEF_LOC, SEARCHTYPE_EXPR, UniqueID, "OBJNAME", .T.)
						m.lSuccess = This.FindInText('Name = ' + ObjName, FINDTYPE_NAME, Nvl(cRootClass, cClassName), IIF('.' $ Parent, Substr (Parent, 1 + At('.', Parent)) + '.', '') + ObjName + "." + OBJECTDEF_LOC, SEARCHTYPE_EXPR, UniqueID, "OBJNAME", .T.)
					Endif
				Endif
			Endif

			*** JRN 2010-03-12 : search in class
			If !Empty(Class)
				*	m.lSuccess = This.FindInText(Class, FINDTYPE_CLASS, Nvl(cRootClass, cClassName), cObjName, SEARCHTYPE_EXPR, UniqueID, "Class", .T.)
				*	m.lSuccess = THIS.FindInLine(Class, FINDTYPE_CLASS, Nvl(cRootClass, cClassName), cObjName, SEARCHTYPE_EXPR, UniqueID, "Class", .T.)
				m.lSuccess = This.FindInLine(Class, FINDTYPE_CLASS, Nvl(cRootClass, cClassName), cObjName + Iif(Empty(cObjName), '', '.') + CLASSNAME_LOC, SEARCHTYPE_EXPR, UniqueID, "Class", .T.)
			Endif

			If This.FormProperties And !Empty(Properties)
				m.lSuccess = This.FindInProperties(Properties, Reserved3, Nvl(cRootClass, cClassName), cObjName, SEARCHTYPE_EXPR, UniqueID)
			Endif
		Endscan
		Select (nSelect)

		Return m.lSuccess
	Endfunc


	Function DoReplace(cReplaceText As String, oReplaceCollection)
		Local nSelect
		Local cObjCode
		Local oFoxRefRecord
		Local cRefCode
		Local cColumn
		Local cTable
		Local cUpdField
		Local cRecordID
		Local cNewText
		Local oError
		Local i
		Local oMethodCollection

		m.oError = .Null.

		nSelect = Select()

		oFoxRefRecord = oReplaceCollection.Item(1)
		cRefCode  = oFoxRefRecord.Abstract
		cColumn   = Rtrim(oFoxRefRecord.ClassName)
		cTable    = Juststem(This.Filename)
		cUpdField = Rtrim(oFoxRefRecord.UpdField)
		cRecordID = oFoxRefRecord.RecordID
		cNewText  = cRefCode

		Select TABLEFIND_ALIAS
		Locate For UniqueID == cRecordID
		If Found()
			Try
				Do Case
					Case cUpdField = "METHODS"
						cNewText = Methods
						For m.i = oReplaceCollection.Count To 1 Step -1
							oFoxRefRecord = oReplaceCollection.Item(m.i)
							m.cNewText = This.ReplaceText(m.cNewText, oFoxRefRecord.Lineno, oFoxRefRecord.ColPos, oFoxRefRecord.MatchLen, m.cReplaceText, oFoxRefRecord.Abstract)
						Endfor

						*!*						FOR EACH oFoxRefRecord IN oReplaceCollection
						*!*							cNewText = THIS.ReplaceText(cNewText, oFoxRefRecord.LineNo, oFoxRefRecord.ColPos, oFoxRefRecord.MatchLen, cReplaceText, oFoxRefRecord.Abstract)
						*!*						ENDFOR

						If Isnull(cNewText)
							Throw ERROR_REPLACE_LOC
						Else
							cObjCode = This.CompileCode(cNewText)
							If !Isnull(cObjCode)
								* update reserved3 and protected field which correspeonds to our list of properties and methods
								Replace ;
									Methods With cNewText, ;
									ObjCode With cObjCode ;
									IN TABLEFIND_ALIAS

								If !Empty(Timestamp)
									Replace Timestamp With This.RowTimeStamp() In TABLEFIND_ALIAS
								Endif
							Endif
						Endif

					Case cUpdField = "RESERVED8"
						cNewText = Reserved8
						For m.i = oReplaceCollection.Count To 1 Step -1
							oFoxRefRecord = oReplaceCollection.Item(m.i)
							m.cNewText = This.ReplaceText(m.cNewText, oFoxRefRecord.Lineno, oFoxRefRecord.ColPos, oFoxRefRecord.MatchLen, m.cReplaceText, oFoxRefRecord.Abstract)
						Endfor
						*!*						FOR EACH oFoxRefRecord IN oReplaceCollection
						*!*							m.cNewText = THIS.ReplaceText(m.cNewText, oFoxRefRecord.LineNo, oFoxRefRecord.ColPos, oFoxRefRecord.MatchLen, m.cReplaceText, m.oFoxRefRecord.Abstract)
						*!*						ENDFOR
						If Isnull(cNewText)
							Throw ERROR_REPLACE_LOC
						Else
							Replace ;
								Reserved8 With cNewText ;
								IN TABLEFIND_ALIAS

							If !Empty(Timestamp)
								Replace Timestamp With This.RowTimeStamp() In TABLEFIND_ALIAS
							Endif
						Endif

					Case cUpdField = "METHODNAME"
						* replacement not supported

					Case cUpdField = "PROPERTYNAME"
						* replacement not supported
						*!*						cNewText = Properties
						*!*						FOR EACH oFoxRefRecord IN oReplaceCollection
						*!*							cNewText = THIS.ReplaceText(cNewText, oFoxRefRecord.LineNo, oFoxRefRecord.ColPos, oFoxRefRecord.MatchLen, cReplaceText, oFoxRefRecord.Abstract)
						*!*						ENDFOR

					Case cUpdField = "PROPERTYVALUE"
						* replacement not supported
						*!*						cNewText = Properties
						*!*						FOR EACH oFoxRefRecord IN oReplaceCollection
						*!*							cNewText = THIS.ReplaceText(cNewText, oFoxRefRecord.LineNo, oFoxRefRecord.ColPos, oFoxRefRecord.MatchLen, cReplaceText, oFoxRefRecord.Abstract)
						*!*						ENDFOR

					Case cUpdField = "OBJNAME"
				Endcase

			Catch To oError When oError.ErrorNo == 2071
			Catch To oError When oError.ErrorNo == 111  && read-only
				Throw REPLACE_READONLY_LOC
			Catch To oError
				Throw oError.Message
			Endtry
		Endif

		Select (nSelect)

		Return m.oError
	Endfunc


	Function DoReplaceLog(cReplaceText, oReplaceCollection As Collection)
		Local cRefCode
		Local cNewText
		Local cColumn
		Local cTable
		Local cRuleExpr
		Local cUpdField
		Local cLog
		Local oError
		Local oFoxRefRecord
		Local cDatabase

		m.cLog = ''
		m.oError = .Null.

		oFoxRefRecord = oReplaceCollection.Item(1)
		cRefCode  = oFoxRefRecord.Abstract
		cColumn   = Rtrim(oFoxRefRecord.ClassName)
		cTable    = Juststem(This.Filename)
		cUpdField = Rtrim(oFoxRefRecord.UpdField)
		cNewText  = cRefCode

		For Each oFoxRefRecord In oReplaceCollection
			cNewText  = Leftc(cNewText, oFoxRefRecord.ColPos - 1) + cReplaceText + Substrc(cNewText, oFoxRefRecord.ColPos + oFoxRefRecord.MatchLen)
		Endfor

		Try
			Do Case
				Case cUpdField = "PROPERTYNAME"
					m.cLog = LOG_PREFIX + REPLACE_NOTSUPPORTED_LOC

				Case cUpdField = "PROPERTYVALUE"
					m.cLog = LOG_PREFIX + REPLACE_NOTSUPPORTED_LOC

				Case cUpdField = "OBJNAME"
					m.cLog = LOG_PREFIX + REPLACE_NOTSUPPORTED_LOC

				Case cUpdField = "METHODNAME"
					m.cLog = LOG_PREFIX + REPLACE_NOTSUPPORTED_LOC

				Case cUpdField = "CLASS"
					m.cLog = LOG_PREFIX + REPLACE_NOTSUPPORTED_LOC

			Endcase

		Catch To oError
		Endtry

		Return m.cLog
	Endfunc
Enddefine
