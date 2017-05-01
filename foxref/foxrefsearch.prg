* Abstract...:
*	Main engine for searching files.
*	Many of these methods are overridden
*	for specific file types.
*
* Changes....:
*
#include "foxref.h"
#include "foxpro.h"

Define Class RefSearch As Custom
	Name = "RefSearch"

	* this is a reference to the engine object (see FoxMatch.prg) that
	* searches a text string -- for example, it could be replaced
	* with a regular expression search engine.
	oSearchEngine  = .Null.
	oEngineController = .Null.

	* set to true to not allow this filetype to refreshed (must always re-search)
	Norefresh       = .F.

	* comma-delimited list of file extensions to backup -- override in subclasses
	* set to empty string to not backup at all
	BackupExtensions = .Null.

	IncludeDefTable = .T.
	FormProperties  = .T.
	CodeOnly        = .F.
	PreserveCase    = .F. && TRUE to preserve case during a Replace operation
	CheckOutSCC     = .F. && True if items must be checked out; only if using SCC\
	TreeViewFolderNames = .F.

	Filename        = ''
	FName           = ''  && filename w/o path
	Folder          = ''  && folder file is in
	Pattern         = ''
	Comments        = COMMENTS_INCLUDE
	FileTimeStamp   = .Null.
	ObjectTimeStamp   = 0

	FileID			= ''

	SetID          = ''
	RefID          = ''

	ReplaceLog     = ''

	* if an error occurs in a TRY/CATCH, then store it to this property
	oErr           = .Null.

	* used internally to track when we need to create a new
	* unique RefID.  Each reference that appears on the
	* same line gets the same RefID.
	nLastLineNo    = -1
	cLastProcName  = ''
	cLastClassName = ''
	cLastFilename  = ''

	nMatchCnt      = 0

	* text of the file we opened
	cFileText      = ''


	lDefinitionsOnly = .F.

	cExtendedPropertyText = Replicate(Chr(1), 517)


	Procedure Init()
		Set Talk Off
		Set Deleted On

		This.SetID = Sys(2015)  && unique ID
		This.FileTimeStamp = Datetime()
		*** JRN 2010-03-19 : timestamp for each line in VCX / SCX
		This.ObjectTimeStamp = 0
	Endproc

	Procedure Destroy()
	Endproc


	Function OpenFile(lReadWrite)
		Local oException

		m.oException = .Null.
		If File(This.Filename)
			Try
				This.cFileText = Filetostr(This.Filename)
			Catch To oException
			Endtry
		Endif

		Return m.oException
	Endfunc

	Function CloseFile
		This.cFileText = ''
	Endfunc

	* Create a backup of the file
	* nBackupStyle: 1 = filename.ext.bak  2 = Backup of filename.ext
	Function BackupFile(cFilename, nBackupStyle)
		Local cBackupFile
		Local cFileToBackup
		Local lSuccess
		Local cExt
		Local i
		Local nCnt
		Local cSafety
		Local oErr
		Local Array aExtensions[1]

		If Empty(This.BackupExtensions)
			Return .T.
		Endif

		m.lSuccess = .T.

		If Vartype(m.cFilename) <> 'C' Or Empty(m.cFilename)
			m.cFilename = This.Filename
		Endif

		m.cSafety = Set("SAFETY")
		Set Safety Off
		If Isnull(This.BackupExtensions)
			* assume it's a text file and we only have one file to backup
			m.cBackupFile = This.CreateBackupFilename(m.cFilename, m.nBackupStyle)
			Try
				Copy File (m.cFilename) To (m.cBackupFile)
			Catch To oErr
				m.lSuccess = .F.
			Endtry
		Else
			m.nCnt = Alines(m.aExtensions, This.BackupExtensions, .T., ',')
			For m.i = 1 To m.nCnt
				If !Empty(m.aExtensions[i])
					m.cFileToBackup = Forceext(m.cFilename, m.aExtensions[i])
					m.cBackupFile = This.CreateBackupFilename(m.cFileToBackup, m.nBackupStyle)
					If File(m.cFileToBackup) And !Empty(m.cBackupFile)
						Try
							Copy File (m.cFileToBackup) To (m.cBackupFile)
						Catch To oErr
							m.lSuccess = .F.
						Endtry
					Endif
				Endif
			Endfor
		Endif
		Set Safety &cSafety


		Return m.lSuccess
	Endfunc

	* nBackupStyle: 1 = filename.ext.bak  2 = Backup of filename.ext
	Protected Function CreateBackupFilename(cFilename, nBackupStyle)
		Do Case
			Case m.nBackupStyle == 1 && filename.ext.bak
				m.cFilename = m.cFilename + '.' + BACKUP_EXTENSION

			Case m.nBackupStyle == 2 && "Backup of filename.ext"
				m.cFilename = Addbs(Justpath(m.cFilename)) + BACKUP_PREFIX_LOC + ' ' + Justfname(m.cFilename)

			Otherwise && same as style = 1
				m.cFilename = m.cFilename + '.' + BACKUP_EXTENSION
		Endcase

		Return m.cFilename
	Endfunc


	Procedure SetTimeStamp()
		Try
			This.FileTimeStamp = Fdate(This.Filename, 1)
		Catch
			This.FileTimeStamp = Datetime()
		Endtry
		This.ObjectTimeStamp = 0
	Endfunc

	* Generate a FoxPro 3.0-style row timestamp
	Protected Function RowTimeStamp(tDateTime)
		Local cTimeValue
		If Vartype(m.tDateTime) <> 'T'
			m.tDateTime = Datetime()
			m.cTimeValue = Time()
		Else
			m.cTimeValue = Ttoc(m.tDateTime, 2)
		Endif

		Return ((Year(m.tDateTime) - 1980) * 2 ** 25);
			+ (Month(m.tDateTime) * 2 ** 21);
			+ (Day(m.tDateTime) * 2 ** 16);
			+ (Val(Leftc(m.cTimeValue, 2)) * 2 ** 11);
			+ (Val(Substrc(m.cTimeValue, 4, 2)) * 2 ** 5);
			+  Val(Rightc(m.cTimeValue, 2))
	Endfunc


	* -- This is the method that should get overridden
	Function DoSearch()
		Return This.FindInText(This.cFileText, FINDTYPE_TEXT, '', '', SEARCHTYPE_NORMAL)
	Endfunc

	* -- Abstract method for parsing symbol definitions
	* -- out of a file
	Function DoDefinitions()
	Endfunc

	* do a replacement on text
	Function DoReplace(cReplaceText, oReplaceCollection)
		Local cCodeBlock
		Local oException
		Local oFoxRefRecord
		Local i

		m.oException = .Null.

		m.cCodeBlock = This.cFileText
		For m.i = oReplaceCollection.Count To 1 Step -1
			oFoxRefRecord = oReplaceCollection.Item(m.i)
			m.cCodeBlock = This.ReplaceText(m.cCodeBlock, oFoxRefRecord.Lineno, oFoxRefRecord.ColPos, oFoxRefRecord.MatchLen, m.cReplaceText, oFoxRefRecord.Abstract)
		Endfor
		Try
			If Isnull(m.cCodeBlock)
				Throw ERROR_REPLACE_LOC
			Else
				If (Strtofile(m.cCodeBlock, This.Filename) == 0)
					Throw ERROR_WRITE_LOC
				Endif
			Endif

		Catch To oException
		Endtry

		Return m.oException
	Endfunc

	* return what to add to the replace log
	Function DoReplaceLog(cReplaceText, oReplaceCollection)
		Return ''
	Endfunc

	* add an error in the FoxRef cursor
	Function AddError(cErrorMsg)
		If Seek(REFTYPE_INACTIVE, "FoxRefCursor", "RefType")
			Replace ;
				SetID With This.SetID, ;
				RefID With Sys(2015), ;
				RefType With REFTYPE_ERROR, ;
				FileID With This.FileID, ;
				Symbol With '', ;
				ClassName With '', ;
				ProcName With '', ;
				ProcLineNo With 0, ;
				LineNo With 0, ;
				ColPos With 0, ;
				MatchLen With 0, ;
				Abstract With m.cErrorMsg, ;
				RecordID With '', ;
				UpdField With '', ;
				Checked With .F., ;
				NoReplace With .F., ;
				TimeStamp With Datetime(), ;
				Inactive With .F. ;
				IN FoxRefCursor
		Else
			Insert Into FoxRefCursor ( ;
				UniqueID, ;
				SetID, ;
				RefID, ;
				RefType, ;
				FileID, ;
				Symbol, ;
				ClassName, ;
				ProcName, ;
				ProcLineNo, ;
				LineNo, ;
				ColPos, ;
				MatchLen, ;
				Abstract, ;
				RecordID, ;
				UpdField, ;
				Checked, ;
				NoReplace, ;
				Timestamp, ;
				Inactive ;
				) Values ( ;
				SYS(2015), ;
				THIS.SetID, ;
				SYS(2015), ;
				REFTYPE_ERROR, ;
				THIS.FileID, ;
				'', ;
				'', ;
				'', ;
				0, ;
				0, ;
				0, ;
				0, ;
				m.cErrorMsg, ;
				'', ;
				'', ;
				.F., ;
				.F., ;
				DATETIME(), ;
				.F. ;
				)
		Endif
	Endfunc

	Function AddMatch(cFindType, cClassName, cProcName, nProcLineNo, nLineNo, nColPos, nMatchLen, cCode, cRecordID, cUpdField, lNoReplace)
		Local cSymbol
		Local Array aPropInfo[1]

		This.nMatchCnt = This.nMatchCnt + 1

		If This.nLastLineNo <> nLineNo Or !(This.cLastProcName == cProcName) Or !(This.cLastClassName == cClassName) Or !(This.cLastFilename == This.Filename)
			This.RefID = Sys(2015)
		Endif


		* if it's a property and occured in an extended property (multi-line property), then get the symbol
		If m.cFindType == FINDTYPE_PROPERTYVALUE And m.nProcLineNo > 1
			If Alines(aPropInfo, m.cCode) >= m.nProcLineNo
				m.cSymbol = Substrc(aPropInfo[m.nProcLineNo], m.nColPos, m.nMatchLen)
			Else
				m.cSymbol = Substrc(m.cCode, m.nColPos, m.nMatchLen)
			Endif
		Else
			m.cSymbol = Substrc(m.cCode, m.nColPos, m.nMatchLen)
		Endif



		Select FoxRefCursor
		If Seek(REFTYPE_INACTIVE, "FoxRefCursor", "RefType")
			Replace ;
				SetID With This.SetID, ;
				RefID With This.RefID, ;
				RefType With REFTYPE_RESULT, ;
				FindType With m.cFindType, ;
				FileID With This.FileID, ;
				Symbol With cSymbol, ;
				ClassName With m.cClassName, ;
				ProcName With m.cProcName, ;
				ProcLineNo With m.nProcLineNo, ;
				LineNo With m.nLineNo, ;
				ColPos With m.nColPos, ;
				MatchLen With m.nMatchLen, ;
				Abstract With m.cCode, ;
				RecordID With m.cRecordID, ;
				UpdField With m.cUpdField, ;
				Checked With .F., ;
				NoReplace With m.lNoReplace, ;
				TimeStamp With This.FileTimeStamp, ;
				Inactive With .F. ;
				IN FoxRefCursor
		Else
			Insert Into FoxRefCursor ( ;
				UniqueID, ;
				SetID, ;
				RefID, ;
				RefType, ;
				FindType, ;
				FileID, ;
				Symbol, ;
				ClassName, ;
				ProcName, ;
				ProcLineNo, ;
				LineNo, ;
				ColPos, ;
				MatchLen, ;
				Abstract, ;
				RecordID, ;
				UpdField, ;
				Checked, ;
				NoReplace, ;
				Timestamp, ;
				oTimestamp, ;
				Inactive ;
				) Values ( ;
				SYS(2015), ;
				THIS.SetID, ;
				THIS.RefID, ;
				REFTYPE_RESULT, ;
				m.cFindType, ;
				THIS.FileID, ;
				m.cSymbol, ;
				m.cClassName, ;
				m.cProcName, ;
				m.nProcLineNo, ;
				m.nLineNo, ;
				m.nColPos, ;
				m.nMatchLen, ;
				m.cCode, ;
				m.cRecordID, ;
				m.cUpdField, ;
				.F., ;
				m.lNoReplace, ;
				THIS.FileTimeStamp, ;
				THIS.ObjectTimeStamp, ;
				.F. ;
				)
		Endif

		This.nLastLineNo    = m.nLineNo
		This.cLastProcName  = m.cProcName
		This.cLastClassName = m.cClassName
		This.cLastFilename  = This.Filename
	Endfunc


	Function AddDefinition(cSymbol, cDefType, cClassName, cProcName, nProcLineNo, nLineNo, cCode, lIncludeFile)
		m.cSymbol = Upper(m.cSymbol)

		If !m.lIncludeFile
			* if definition name doesn't begin with an underscore or Alpha character
			* then assume it's not a valid symbol name
			If !Isalpha(m.cSymbol) And Leftc(m.cSymbol, 1) <> '_'
				Return .F.
			Endif

			* remove any memory variable designations
			If Leftc(m.cSymbol, 2) == "M."
				m.cSymbol = Substrc(m.cSymbol, 3)
				If Empty(m.cSymbol)
					Return .F.
				Endif
			Endif
		Endif

		This.nMatchCnt = This.nMatchCnt + 1

		If This.nLastLineNo <> nLineNo Or !(This.cLastProcName == cProcName) Or !(This.cLastClassName == cClassName) Or !(This.cLastFilename == This.Filename)
			This.RefID = Sys(2015)
		Endif

		If Seek(.T., "FoxDefCursor", "Inactive")
			Replace ;
				DefType With m.cDefType, ;
				FileID With This.FileID, ;
				Symbol With m.cSymbol, ;
				ClassName With m.cClassName, ;
				ProcName With m.cProcName, ;
				ProcLineNo With m.nProcLineNo, ;
				LineNo With m.nLineNo, ;
				Abstract With This.StripTabs(m.cCode), ;
				Inactive With .F. ;
				IN FoxDefCursor
		Else
			Insert Into FoxDefCursor ( ;
				UniqueID, ;
				DefType, ;
				FileID, ;
				Symbol, ;
				ClassName, ;
				ProcName, ;
				ProcLineNo, ;
				LineNo, ;
				Abstract, ;
				Inactive ;
				) Values ( ;
				SYS(2015), ;
				m.cDefType, ;
				THIS.FileID, ;
				m.cSymbol, ;
				m.cClassName, ;
				m.cProcName, ;
				m.nProcLineNo, ;
				m.nLineNo, ;
				THIS.StripTabs(m.cCode), ;
				.F. ;
				)
		Endif

		This.nLastLineNo    = m.nLineNo
		This.cLastProcName  = m.cProcName
		This.cLastClassName = m.cClassName
		This.cLastFilename  = This.Filename

		Return .T.
	Endfunc

	* Add a file to process -- usually when we encounter
	* #include files while we're processing definitions
	Function AddFileToProcess(cDefType, cFilename, cClassName, cProcName, nProcLineNo, nLineNo, cCode)
		If Leftc(m.cFilename, 1) == '"'
			m.cFilename = Substrc(m.cFilename, 2)
		Endif
		If Rightc(m.cFilename, 1) == '"'
			m.cFilename = Leftc(m.cFilename, Lenc(m.cFilename) - 1)
		Endif

		If Vartype(This.oEngineController) == 'O'
			This.oEngineController.AddFileToProcess(m.cFilename)
		Endif

		This.AddDefinition(m.cFilename, m.cDefType, m.cClassName, m.cProcName, m.nProcLineNo, m.nLineNo, m.cCode, .T.)
	Endfunc

	Function AddFileToSearch(cFilename)
		If Vartype(This.oEngineController) == 'O'
			This.oEngineController.AddFileToSearch(m.cFilename)
		Endif
	Endfunc


	Function ProcessDefinitions(oEngineController)
		If Vartype(m.oEngineController) == 'O'
			This.oEngineController = m.oEngineController
		Endif
		This.DoDefinitions()
		This.oEngineController = .Null.

	Endfunc

	* Returns the FileAction ('N' = Processed w/o Definitions, 'D' = Processed w/definitions, 'E' = Error, 'S' = Stop immediately)
	Function SearchFor(cPattern, lSearch, lDefinitions, oEngineController)
		Local cFileAction && file action to return
		Local oException
		Local nSelect

		m.nSelect = Select()

		If Vartype(m.oEngineController) == 'O'
			This.oEngineController = m.oEngineController
		Endif

		If Vartype(m.cPattern) == 'C'
			This.Pattern = m.cPattern
		Endif

		If Pcount() < 3
			m.lDefinitions = This.IncludeDefTable
		Endif

		m.cFileAction = ''

		This.FName  = Justfname(This.Filename)
		This.Folder = Justpath(This.Filename)
		This.nMatchCnt = 0

		If This.oSearchEngine.SetPattern(This.Pattern)
			m.oException = This.OpenFile()
			If Isnull(m.oException)
				Try
					If m.lDefinitions
						This.nLastLineNo = -1
						This.DoDefinitions()

						m.cFileAction = FILEACTION_DEFINITIONS
					Endif

					If lSearch And !Empty(m.cPattern) && if no search pattern passed, then only do definitions
						This.nLastLineNo = -1
						If !This.DoSearch()
							m.cFileAction = FILEACTION_STOP
						Endif
					Endif

					This.CloseFile()

					If This.nMatchCnt == 0 And !(m.cFileAction == FILEACTION_STOP)
						If Seek(REFTYPE_INACTIVE, "FoxRefCursor", "RefType")
							Replace ;
								SetID With This.SetID, ;
								RefID With Sys(2015), ;
								RefType With REFTYPE_NOMATCH, ;
								FileID With This.FileID, ;
								Symbol With '', ;
								ClassName With '', ;
								ProcName With '', ;
								ProcLineNo With 0, ;
								LineNo With 0, ;
								ColPos With 0, ;
								MatchLen With 0, ;
								Abstract With '', ;
								RecordID With '', ;
								UpdField With '', ;
								Checked With .F., ;
								NoReplace With .F., ;
								TimeStamp With This.FileTimeStamp, ;
								Inactive With .F. ;
								IN FoxRefCursor
						Else
							* no matches, but still record that we searched this file
							Insert Into FoxRefCursor ( ;
								UniqueID, ;
								SetID, ;
								RefID, ;
								RefType, ;
								FileID, ;
								Symbol, ;
								ClassName, ;
								ProcName, ;
								ProcLineNo, ;
								LineNo, ;
								ColPos, ;
								MatchLen, ;
								Abstract, ;
								RecordID, ;
								UpdField, ;
								Checked, ;
								NoReplace, ;
								Timestamp, ;
								Inactive ;
								) Values ( ;
								SYS(2015), ;
								THIS.SetID, ;
								SYS(2015), ;
								REFTYPE_NOMATCH, ;
								THIS.FileID, ;
								'', ;
								'', ;
								'', ;
								0, ;
								0, ;
								0, ;
								0, ;
								'', ;
								'', ;
								'', ;
								.F., ;
								.F., ;
								THIS.FileTimeStamp, ;
								.F. ;
								)
						Endif
					Endif
				Catch To oException
					* error is recorded below
				Endtry
			Endif

			If Vartype(oException) == 'O'
				This.AddError(Iif(Empty(m.oException.UserValue), m.oException.Message, m.oException.UserValue))
				m.cFileAction = FILEACTION_ERROR
			Endif
		Else
			* unable to set the pattern
			m.cFileAction = FILEACTION_STOP
		Endif

		This.oEngineController = .Null.

		Select (m.nSelect)

		Return m.cFileAction
	Endfunc

	* this is what FoxRefEngine calls to do the replacements
	Function ReplaceWith(cReplaceText, oReplaceCollection, cFilename) As Boolean
		Local lSuccess
		Local nSelect
		Local oException

		This.Filename = m.cFilename

		m.nSelect = Select()

		m.oException = This.OpenFile(.T.)

		If Isnull(m.oException)
			Try
				m.oException = This.DoReplace(m.cReplaceText, m.oReplaceCollection)
			Catch To m.oException
			Endtry

			This.ReplaceLog = This.DoReplaceLog(m.cReplaceText, m.oReplaceCollection)

			This.CloseFile()
		Endif

		Select (m.nSelect)

		Return m.oException
	Endfunc

	* take the original text, determine the case,
	* and apply it the new text
	Function SetCasePreservation(cOriginalText, cReplacementText)
		Local i
		Local ch
		Local lLower
		Local lUpper

		If Isupper(Leftc(m.cOriginalText, 1)) And Islower(Substrc(m.cOriginalText, 2, 1))
			* proper case
			m.cReplacementText = Proper(m.cReplacementText)
		Else
			m.lLower = .F.
			m.lUpper = .F.
			For m.i = 1 To Lenc(m.cOriginalText)
				ch = Substrc(m.cOriginalText, m.i, 1)
				Do Case
					Case Between(ch, 'A', 'Z')
						m.lUpper = .T.
					Case Between(ch, 'a', 'z')
						m.lLower = .T.
				Endcase

				If m.lLower And m.lUpper
					Exit
				Endif
			Endfor

			Do Case
				Case m.lLower And !m.lUpper
					m.cReplacementText = Lower(m.cReplacementText)
				Case !m.lLower And m.lUpper
					m.cReplacementText = Upper(m.cReplacementText)
			Endcase
		Endif

		Return m.cReplacementText
	Endfunc

	* perform a replacement on text
	*  NOTE: cOldCode contains the original text as it was when we original searched, but we don't use it here anymore
	Function ReplaceText(cCodeBlock, nLineNo, nColPos, nMatchLen, cNewText, cOldCode)
		Local nLineCnt
		Local i
		Local lSuccess
		Local cOriginalText
		Local cReplacementText
		Local Array aCodeList[1]

		If Isnull(cCodeBlock)
			Return .Null.
		Endif

		m.lSuccess = .F.

		m.cReplacementText = m.cNewText
		If m.nLineNo == 0
			If This.PreserveCase
				m.cOriginalText = Substrc(m.cCodeBlock, m.nColPos, m.nMatchLen)
				m.cNewText = This.SetCasePreservation(m.cOriginalText, m.cReplacementText)
			Endif

			m.cCodeBlock = Leftc(m.cCodeBlock, m.nColPos - 1) + m.cNewText + Iif(m.nColPos + m.nMatchLen > Lenc(m.cCodeBlock), '', Substrc(m.cCodeBlock, m.nColPos + m.nMatchLen))
			m.lSuccess = .T.
		Else
			m.nLineCnt = Alines(m.aCodeList, m.cCodeBlock, .F.)
			If m.nLineNo <= m.nLineCnt
				If This.PreserveCase
					m.cOriginalText = Substrc(m.aCodeList[m.nLineNo], m.nColPos, m.nMatchLen)
					m.cNewText = This.SetCasePreservation(m.cOriginalText, m.cReplacementText)
				Endif

				m.aCodeList[m.nLineNo] = Leftc(m.aCodeList[m.nLineNo], m.nColPos - 1) + m.cNewText + Iif(m.nColPos + m.nMatchLen > Lenc(m.aCodeList[m.nLineNo]), '', Substrc(m.aCodeList[m.nLineNo], m.nColPos + m.nMatchLen))

				m.cCodeBlock = ''
				For m.i = 1 To nLineCnt
					m.cCodeBlock = m.cCodeBlock + Iif(m.i == 1, '', Chr(13) + Chr(10)) + m.aCodeList[m.i]
				Endfor

				m.lSuccess = .T.
			Endif
		Endif

		If !m.lSuccess
			m.cCodeBlock = .Null.
		Endif

		Return m.cCodeBlock
	Endfunc



	* -- This is the meat of our find.  Searches through text or code
	Function FindInCode(cTextBlock, cFindType, cClassName, cObjName, nSearchType, cRecordID, cUpdField, lNoReplace)
		Local i
		Local nSelect
		Local nLineCnt
		Local cProcName
		Local nProcLineNo
		Local cFirstChar
		Local nOffset
		Local lFound
		Local nCommentPos
		Local lExitLine
		Local nWordNum
		Local cFirstWord
		Local lComment
		Local nMatchCnt
		Local nLastLineNo
		Local nMatchIndex
		Local oMatch
		Local lMethodName
		Local lUseMemLines
		Local cCodeLine
		Local nMemoWidth
		Local lCommentAndContinue
		Local cSpecialMethodName
		Local Array aCodeList[1]

		m.nMatchCnt = This.oSearchEngine.Execute(m.cTextBlock)
		If m.nMatchCnt <= 0
			Return (m.nMatchCnt >= 0)
		Endif

		If Vartype(m.cClassName) <> 'C'
			m.cClassName = ''
		Endif

		If Vartype(m.cObjName) <> 'C'
			m.cObjName = ''
		Endif

		If m.nSearchType == SEARCHTYPE_METHOD  && search in method code of .vcx or .scx
			m.nOffset = -1
		Else
			m.nOffset = 0
		Endif

		* this is the UNIQUEID field from a form or class library
		If Vartype(m.cRecordID) <> 'C'
			m.cRecordID = ''
		Endif
		If Vartype(m.cUpdField) <> 'C'
			m.cUpdField = ''
		Endif

		m.cProcName   = ''
		m.nProcLineNo = 0
		m.lComment    = .F.
		m.lCommentAndContinue = .F.

		m.nSelect = Select()

		nMatchIndex    = 1


		* ALINES should be faster, but we can't use that if
		* there are more than 65k lines in the code block
		m.lUseMemLines = .F.
		Try
			m.nLineCnt = Alines(aCodeList, m.cTextBlock, .F.)
		Catch
			m.lUseMemLines = .T.
		Endtry

		If m.lUseMemLines
			m.nMemoWidth = Set("MEMOWIDTH")
			Set Memowidth To 8192
			m.nLineCnt = Memlines(m.cTextBlock)
			_Mline = 0
		Endif


		m.nLastLineNo  = Min(This.oSearchEngine.oMatches(m.nMatchCnt).MatchLineNo, m.nLineCnt)
		For m.i = 1 To m.nLastLineNo
			If m.lUseMemLines
				m.cCodeLine = Mline(m.cTextBlock, 1, _Mline)
			Else
				m.cCodeLine = aCodeList[m.i]
			Endif

			m.lMethodName = .F.

			* get first word
			m.nWordNum = 1
			m.cFirstWord = Upper(Getwordnum(m.cCodeLine, m.nWordNum, WORD_DELIMITERS))

			If m.lCommentAndContinue
				* since the last line was a comment line and also
				* was a continuation line, then this is a comment line
				m.lComment = .T.
			Else
				m.lComment = m.cFirstWord = '*' Or (m.cFirstWord = ("&" + "&")) Or m.cFirstWord == "NOTE"
			Endif

			cSpecialMethodName = ''
			If (m.lComment And This.Comments == COMMENTS_EXCLUDE)
				m.nProcLineNo = m.nProcLineNo + 1

				* if this line was a comment line and also
				* was a continuation line, then the next line
				* is also a comment
				m.lCommentAndContinue = m.lComment And Rightc(Rtrim(m.cCodeLine), 1) == ';'

				Loop
			Endif

			If !m.lComment And Lenc(m.cFirstWord) >= 4 And Inlist(m.cFirstWord, 'E', 'H', 'P', 'F', 'L', 'D')
				If "PROTECTED" = m.cFirstWord Or "HIDDEN" = m.cFirstWord
					m.nWordNum = m.nWordNum + 1
					m.cFirstWord = Upper(Getwordnum(m.cCodeLine, m.nWordNum, WORD_DELIMITERS))
				Endif

				Do Case
					Case "PROCEDURE" = m.cFirstWord
						m.cProcName = Getwordnum(m.cCodeLine, m.nWordNum + 1, METHOD_DELIMITERS)
						m.nProcLineNo = 0
						m.lMethodName = .T.
						
						If m.nSearchType == SEARCHTYPE_METHOD  && search in method code of .vcx or .scx
							cSpecialMethodName = "." + METHOD_LOC
						Else
							cSpecialMethodName = "." + PROCEDURE_LOC
						Endif

					Case "FUNCTION" = cFirstWord
						m.cProcName = Getwordnum(m.cCodeLine, m.nWordNum + 1, METHOD_DELIMITERS)
						m.nProcLineNo = 0
						cSpecialMethodName = "." + FUNCTION_LOC

					Case "ENDFUNC" = m.cFirstWord
						m.cProcName = ''
						m.nProcLineNo = 0

					Case "ENDPROC" = m.cFirstWord
						m.cProcName = ''
						m.nProcLineNo = 0

					Case "DEFINE" = m.cFirstWord
						m.cSecondWord = Upper(Getwordnum(m.cCodeLine, m.nWordNum + 1, WORD_DELIMITERS))
						If Lenc(m.cSecondWord) >= 4 And "CLASS" = m.cSecondWord
							m.cClassName = Getwordnum(m.cCodeLine, m.nWordNum + 2, WORD_DELIMITERS)
							m.cProcName = ''
							m.nProcLineNo = 0
						Endif

					Case "ENDDEFINE" = m.cFirstWord
						m.cClassName  = ''
						m.cProcName   = ''
						m.nProcLineNo = 0

				Endcase
			Endif

			m.nProcLineNo = m.nProcLineNo + 1

			If This.Comments == COMMENTS_EXCLUDE Or This.Comments == COMMENTS_ONLY
				m.nCommentPos = At_c('&' + '&', m.cCodeLine)

				* if this line was a comment line and also
				* was a continuation line, then the next line
				* is also a comment
				m.lCommentAndContinue = m.nCommentPos > 0 And Rightc(Rtrim(m.cCodeLine), 1) == ';'
			Else
				m.nCommentPos = 0
			Endif


			For m.j = m.nMatchIndex To m.nMatchCnt
				oMatch = This.oSearchEngine.oMatches(m.j)

				If oMatch.MatchLineNo == m.i
					If (This.Comments == COMMENTS_EXCLUDE And m.nCommentPos > 0 And oMatch.MatchPos >= m.nCommentPos)
						Exit
					Endif

					If This.Comments <> COMMENTS_ONLY Or m.lComment Or (m.nCommentPos > 0 And oMatch.MatchPos >= m.nCommentPos)
						This.AddMatch( ;
							m.cFindType, ;
							m.cClassName, ;
							m.cObjName + Iif(Empty(m.cObjName) Or Empty(m.cProcName), '', '.') + m.cProcName + cSpecialMethodName, ;
							IIF(nSearchType == SEARCHTYPE_EXPR, 0, m.nProcLineNo + m.nOffset), ;
							IIF(nSearchType == SEARCHTYPE_EXPR, 0, m.i), ;
							oMatch.MatchPos, ;
							oMatch.MatchLen, ;
							m.cCodeLine, ;
							m.cRecordID, ;
							IIF(m.lMethodName, "METHODNAME", m.cUpdField), ;
							oMatch.NoReplace Or Iif(m.lMethodName, .T., m.lNoReplace) ;
							)
					Endif

				Else
					If oMatch.MatchLineNo > m.i
						Exit
					Endif
				Endif

				m.nMatchIndex = m.nMatchIndex + 1
			Endfor

		Endfor

		If m.lUseMemLines
			Set Memowidth To (m.nMemoWidth)
		Endif

		Select (m.nSelect)
	Endfunc

	* -- This searches a text block, assuming it's NOT code so there's
	* -- no reason to locate comments, entering & exiting procedures, etc
	Function FindInText(cTextBlock, cFindType, cClassName, cObjName, nSearchType, cRecordID, cUpdField, lNoReplace)
		Local i
		Local nSelect
		Local nLineCnt
		Local lUseMemLines
		Local nMemoWidth
		Local Array aCodeList[1]

		m.nSelect = Select()

		If Vartype(m.cClassName) <> 'C'
			m.cClassName = ''
		Endif

		If Vartype(m.cObjName) <> 'C'
			m.cObjName = ''
		Endif

		* this is the UNIQUEID field from a form or class library
		If Vartype(m.cRecordID) <> 'C'
			m.cRecordID = ''
		Endif
		If Vartype(m.cUpdField) <> 'C'
			m.cUpdField = ''
		Endif

		m.nMatchCnt = This.oSearchEngine.Execute(m.cTextBlock)
		If m.nMatchCnt > 0
			* ALINES should be faster, but we can't use that if
			* there are more than 65k lines in the code block
			m.lUseMemLines = .F.
			Try
				m.nLineCnt = Alines(aCodeList, m.cTextBlock, .F.)
			Catch
				m.lUseMemLines = .T.
			Endtry

			If m.lUseMemLines
				m.nMemoWidth = Set("MEMOWIDTH")
				Set Memowidth To 8192
				m.nLineCnt = Memlines(m.cTextBlock)
			Endif

			For Each oMatch In This.oSearchEngine.oMatches
				This.AddMatch( ;
					m.cFindType, ;
					m.cClassName, ;
					m.cObjName, ;
					oMatch.MatchLineNo, ;
					oMatch.MatchLineNo, ;
					oMatch.MatchPos, ;
					oMatch.MatchLen, ;
					IIF(m.lUseMemLines, Mline(m.cTextBlock, oMatch.MatchLineNo), aCodeList[oMatch.MatchLineNo]), ;
					m.cRecordID, ;
					m.cUpdField, ;
					oMatch.NoReplace Or m.lNoReplace ;
					)

			Endfor
		Endif

		If m.lUseMemLines
			Set Memowidth To (m.nMemoWidth)
		Endif

		Select (m.nSelect)
	Endfunc

	* perform search on a single line of text (no line # recorded)
	Function FindInLine(cTextBlock, cFindType, cClassName, cObjName, nSearchType, cRecordID, cUpdField, lNoReplace, cAbstract, nColOffset, nLineNo)
		Local nSelect
		Local nMatchLen
		Local nMatchCnt

		m.nSelect = Select()

		If Vartype(m.cClassName) <> 'C'
			m.cClassName = ''
		Endif

		If Vartype(m.cObjName) <> 'C'
			m.cObjName = ''
		Endif

		* this is the UNIQUEID field from a form or class library
		If Vartype(m.cRecordID) <> 'C'
			m.cRecordID = ''
		Endif
		If Vartype(m.cUpdField) <> 'C'
			m.cUpdField = ''
		Endif
		If Vartype(m.nColOffset) <> 'N'
			m.nColOffset = 0
		Endif
		If Vartype(m.nLineNo) <> 'N'
			m.nLineNo = 0
		Endif

		This.RefID = Sys(2015)


		* see if reference occurs on this line
		m.nMatchCnt = This.oSearchEngine.Execute(m.cTextBlock)
		If m.nMatchCnt > 0

			For Each oMatch In This.oSearchEngine.oMatches
				This.AddMatch( ;
					m.cFindType, ;
					m.cClassName, ;
					m.cObjName, ;
					oMatch.MatchLineNo, ;
					m.nLineNo, ;
					oMatch.MatchPos + m.nColOffset, ;
					oMatch.MatchLen, ;
					IIF(Vartype(m.cAbstract) == 'C', m.cAbstract, m.cTextBlock), ;
					m.cRecordID, ;
					m.cUpdField, ;
					oMatch.NoReplace Or m.lNoReplace ;
					)

			Endfor
		Endif

		Select (m.nSelect)

		Return m.nMatchCnt >= 0 && -1 returned on error
	Endfunc


	* -- This searches a property list
	Function FindInProperties(cTextBlock, cPEMList, cClassName, cObjName, nSearchType, cRecordID)
		Local i
		Local j
		Local nSelect
		Local nLineCnt
		Local nMatchLen
		Local cPropertyName
		Local nEqualPos
		Local cPropertyValue
		Local lSuccess
		Local nPEMCnt
		Local cProperty
		Local nLen
		Local nPos
		Local Array aCodeList[1]
		Local Array aPEMList[1]

		If (This.Comments == COMMENTS_ONLY)
			Return
		Endif

		m.nSelect = Select()

		m.lSuccess = .T.

		If Vartype(m.cClassName) <> 'C'
			m.cClassName = ''
		Endif

		If Vartype(m.cObjName) <> 'C'
			m.cObjName = ''
		Endif

		* this is the UNIQUEID field from a form or class library
		If Vartype(m.cRecordID) <> 'C'
			m.cRecordID = ''
		Endif

		If Empty(m.cPEMList)
			m.nPEMCnt = 0
		Else
			m.nPEMCnt = Alines(aPEMList, m.cPEMList, .T.)
		Endif

		m.nLineCnt = 0
		m.cProperty = Rtrim(Getwordnum(m.cTextBlock, 1, Chr(10)), 0, Chr(13))
		Do While !Empty(m.cProperty)
			m.nEqualPos = At_c(' = ', m.cProperty)
			If m.nEqualPos > 3
				If Substr(m.cProperty, m.nEqualPos + 3, 517) == This.cExtendedPropertyText
					* we have an extended propererty
					m.cPropertyName = Leftc(m.cProperty, m.nEqualPos - 1)

					m.nLen = Val(Substr(m.cProperty, m.nEqualPos + 3 + 517, 8))

					m.cProperty = m.cPropertyName + " = " + Substr(m.cProperty, m.nEqualPos + 3 + 517 + 8, nLen)
					m.cTextBlock = Substr(m.cTextBlock, m.nEqualPos + 5 + 517 + 8 + nLen)
				Else
					m.nPos = At(Chr(10), m.cTextBlock)
					If m.nPos = 0
						m.nPos = At(Chr(13), m.cTextBlock)
					Endif
					If m.nPos > 0
						m.cTextBlock = Substr(m.cTextBlock, m.nPos + 1)
					Else
						m.cTextBlock = ''
					Endif
				Endif

				m.nLineCnt = m.nLineCnt + 1
				Dimension aCodeList[m.nLineCnt]
				aCodeList[m.nLineCnt] = m.cProperty

				m.cProperty = Rtrim(Getwordnum(m.cTextBlock, 1, Chr(10)), 0, Chr(13))
			Else
				m.cProperty = ''
			Endif
		Enddo

		For m.i = 1 To m.nLineCnt
			* First see if we have a match for PropertName = Propertyvalue (e.g. Width = 300)
			m.nEqualPos = At_c(' = ', aCodeList[m.i])
			If m.nEqualPos > 3
				m.cPropertyName = Leftc(aCodeList[m.i], m.nEqualPos - 1)

				*** JRN 2010-03-17 : following removed ... I'd LIKE to see it again!
				*!*	* if the property name occurs here, zero it out from the PEMList
				*!*	* so we don't find it again
				*!*	FOR m.j = 1 TO m.nPEMCnt
				*!*		* this won't find method names because they start with '*'
				*!*		IF aPEMList[m.j] == m.cPropertyName
				*!*			aPEMList[m.j] = ''
				*!*		ENDIF
				*!*	ENDFOR

				m.cProperty = aCodeList[m.i]
				*** JRN 2010-03-21 :
				* m.lSuccess = This.FindInLine(m.cProperty, FINDTYPE_PROPERTYVALUE, m.cClassName, m.cObjName, SEARCHTYPE_NORMAL, m.cRecordID, "PROPERTYNAME", .T., m.cProperty,, m.i)
				m.lSuccess = This.FindInLine(m.cProperty, FINDTYPE_PROPERTYVALUE, m.cClassName, cObjName + Iif(Empty(cObjName), '', '.') + PROPERTY_LOC, SEARCHTYPE_NORMAL, m.cRecordID, "PROPERTYNAME", .T., '.' + m.cProperty,, m.i)
				If !m.lSuccess
					* see if reference occurs in the value portion of the property name
					m.cPropertyValue = Substrc(m.cProperty, m.nEqualPos + 3)

					m.lSuccess = This.FindInLine(m.cPropertyName, FINDTYPE_PROPERTYNAME, m.cClassName, cObjName + Iif(Empty(cObjName), '', '.') + PROPERTYNAME_LOC, SEARCHTYPE_NORMAL, m.cRecordID, "PROPERTYNAME", .T., m.cProperty,, m.i) Or m.lSuccess
					m.lSuccess = This.FindInLine(m.cPropertyValue, FINDTYPE_PROPERTYVALUE, m.cClassName, m.cObjName + Iif(Empty(m.cObjName) Or Empty(m.cPropertyName), '', '.') + m.cPropertyName, SEARCHTYPE_NORMAL, m.cRecordID, "PROPERTYVALUE", .T., m.cProperty, m.nEqualPos + 2, m.i) Or m.lSuccess
				Endif
			Endif
		Endfor

		* see if match occurs in a defined property that has a default value
		* (if it has a default value, it won't show up in the Properties field)
		For m.i = 1 To m.nPEMCnt
			If !Empty(aPEMList[m.i]) And Leftc(aPEMList[m.i], 1) <> '*'
				m.lSuccess = This.FindInLine(aPEMList[m.i], FINDTYPE_PROPERTYNAME, m.cClassName, cObjName + Iif(Empty(cObjName), '', '.') + PROPERTYNAME_LOC, SEARCHTYPE_NORMAL, m.cRecordID, "PROPERTYNAME", .T., aPEMList[m.i],, m.i) Or m.lSuccess
			Endif
		Endfor


		Select (m.nSelect)

		Return m.lSuccess
	Endfunc

	* find any defined Methods
	Function FindDefinedMethods(cTextBlock, cClassName, cObjName, nSearchType, cRecordID)
		Local i
		Local nLineCnt
		Local nMatchLen
		Local cPEMName
		Local lSuccess
		Local Array aCodeList[1]

		If (This.Comments == COMMENTS_ONLY)
			Return
		Endif

		m.lSuccess = .T.

		If Vartype(m.cClassName) <> 'C'
			m.cClassName = ''
		Endif

		If Vartype(m.cObjName) <> 'C'
			m.cObjName = ''
		Endif

		* this is the UNIQUEID field from a form or class library
		If Vartype(m.cRecordID) <> 'C'
			m.cRecordID = ''
		Endif

		m.nLineCnt = Alines(aCodeList, m.cTextBlock, .F.)
		For m.i = 1 To m.nLineCnt
			If Leftc(aCodeList[m.i], 1) == '*' && methods begin with '*' in reserved3 field
				m.cPEMName = Substrc(aCodeList[m.i], 2)
				m.lSuccess = This.FindInLine(m.cPEMName, FINDTYPE_METHODNAME, m.cClassName, cObjName + Iif(Empty(cObjName), '', '.') + METHODNAME_LOC, SEARCHTYPE_NORMAL, m.cRecordID, "METHOD", .T., m.cPEMName,, m.i)
			Endif
		Endfor

		Return m.lSuccess
	Endfunc

	Function ParseLine(cLine, aTokens, nTokenCnt)
		Local cEndQuote
		Local cWord
		Local cLastCh
		Local ch
		Local nTerminal
		Local lInQuote
		Local lInSymbol
		Local nTokenCnt
		Local nLen
		Local i


		nTerminal = 0
		cWord     = ''
		nLen      = Lenc(cLine)
		cLastCh   = ''
		For i = 1 To nLen
			ch = Substrc(cLine, i, 1)

			If lInQuote
				If ch == cEndQuote
					nTerminal = 1
					cEndQuote = ''
					cWord = cWord + ch
					ch = ''
					lInQuote = .F.
				Else
					nTerminal = 0
				Endif
			Else
				If ch == '_' Or Isalpha(ch) Or ch == '.' Or ch $ "0123456789"
					If lInSymbol
						nTerminal = 0
					Else
						lInSymbol = .T.
						nTerminal = 1
					Endif
				Else
					lInSymbol = .F.

					Do Case
						Case ch == ' ' Or ch == Tab
							nTerminal = 2

						Case ch == '"' Or ch == '[' Or ch == "'"
							nTerminal = 1
							cEndQuote = Iif(ch == '[', ']', ch)
							lInQuote  = .T.

						Case ch == '&' And cLastCh == '&'
							nTerminal = 2
							Exit

						Case ch == ';'
							nTerminal = 1

						Otherwise
							nTerminal = 1

					Endcase
				Endif
			Endif

			If nTerminal <> 0
				If !Empty(cWord) And nTokenCnt < MAX_TOKENS
					nTokenCnt = nTokenCnt + 1
					aTokens[nTokenCnt] = cWord
				Endif
				If nTerminal <> 2
					cWord = ch
				Else
					cWord = ''
				Endif
				nTerminal = 0
			Else
				cWord = cWord + ch
			Endif
			cLastCh = ch
		Endfor

		If nTerminal <> 2 And !Empty(cWord) And nTokenCnt < MAX_TOKENS
			nTokenCnt = nTokenCnt + 1
			aTokens[nTokenCnt] = cWord
		Endif

		Return nTokenCnt
	Endfunc

	* -- Given a code block, create list of definitions
	* -- Definitions are:
	* --   Class Definitions (DEFINE CLASS)
	* --   Procedures/Functions (FUNCTION/PROCEDURE)
	* --   Parameters (PARAMETERS/LPARAMETERS/Inline)
	* --   LOCAL/PUBLIC/PRIVATE/HIDDEN/PROTECTED declarations
	* --   #define
	Function FindDefinitions(cTextBlock, cClassName, cObjName, nSearchType)
		Local i, j
		Local nLineCnt
		Local cProcName
		Local nProcLineNo
		Local nOffset
		Local cDefType
		Local cDefinition
		Local nDefWordNum
		Local cUpperDefinition
		Local lInClassDef
		Local nTokenCnt
		Local nToken
		Local cStopToken
		Local lUseMemLines
		Local nMemoWidth
		Local cCodeLine
		Local Array aCodeList[1]
		Local Array aTokens[MAX_TOKENS]

		If Vartype(m.cClassName) <> 'C'
			m.cClassName = ''
		Endif

		If Vartype(m.cObjName) <> 'C'
			m.cObjName = ''
		Endif

		If m.nSearchType == SEARCHTYPE_METHOD  && search in method code of .vcx or .scx
			m.nOffset = -1
		Else
			m.nOffset = 0
		Endif

		m.cProcName     = ''   && name of procedure/function we're in
		m.nProcLineNo   = 1
		m.lInClassDef   = .F.  && in a class definition?
		m.nTokenCnt     = 0
		m.cStopToken    = ''


		* ALINES should be faster, but we can't use that if
		* there are more than 65k lines in the code block
		m.lUseMemLines = .F.
		Try
			m.nLineCnt = Alines(aCodeList, m.cTextBlock, .F.)
		Catch
			m.lUseMemLines = .T.
		Endtry

		If m.lUseMemLines
			m.nMemoWidth = Set("MEMOWIDTH")
			Set Memowidth To 8192
			m.nLineCnt = Memlines(m.cTextBlock)
			_Mline = 0
		Endif


		For m.i = 1 To m.nLineCnt
			If m.lUseMemLines
				m.cCodeLine = Mline(m.cTextBlock, 1, _Mline)
			Else
				m.cCodeLine = aCodeList[m.i]
			Endif


			m.nTokenCnt = This.ParseLine(m.cCodeLine, @aTokens, m.nTokenCnt)
			If m.nTokenCnt > 0 And aTokens[m.nTokenCnt] == ';'
				m.nProcLineNo = m.nProcLineNo + 1
				m.nTokenCnt = m.nTokenCnt - 1
				Loop
			Endif

			m.nToken = 0
			If m.nTokenCnt > 1
				If Lenc(aTokens[1]) >= 4
					m.nStopToken = m.nTokenCnt
					m.cWord1 = Upper(aTokens[1])
					m.cWord2 = Upper(aTokens[2])

					Do Case
						Case m.nTokenCnt > 2 And ("PROTECTED" = m.cWord1 Or "HIDDEN" = m.cWord1) And (Lenc(m.cWord2) >= 4 And ("PROCEDURE" = m.cWord2 Or "FUNCTION" = m.cWord2))
							m.cProcName   = aTokens[3]
							m.nProcLineNo = 0
							m.lInClassDef = .F.
							m.cStopToken = ')'

							This.AddDefinition( ;
								m.cProcName, ;
								DEFTYPE_PROCEDURE, ;
								m.cClassName, ;
								m.cObjName + Iif(Empty(m.cObjName) Or Empty(m.cProcName), '', '.') + m.cProcName, ;
								0, ;
								m.i, ;
								m.cCodeLine ;
								)

							m.cDefType = DEFTYPE_PARAMETER
							m.nToken = 4
							m.cStopToken  = ')'

						Case "PROCEDURE" = m.cWord1 Or "FUNCTION" = m.cWord1
							m.cProcName   = aTokens[2]
							m.nProcLineNo = 0
							m.lInClassDef = .F.

							This.AddDefinition( ;
								m.cProcName, ;
								DEFTYPE_PROCEDURE, ;
								m.cClassName, ;
								m.cObjName + Iif(Empty(m.cObjName) Or Empty(m.cProcName), '', '.') + m.cProcName, ;
								0, ;
								m.i, ;
								m.cCodeLine ;
								)

							m.cDefType = DEFTYPE_PARAMETER
							m.nToken = 3
							m.cStopToken  = ')'

						Case "ENDFUNC" = m.cWord1
							m.cProcName = ''
							m.nProcLineNo = 0
							m.lInClassDef = .F.

						Case "ENDPROC" = m.cWord1
							m.cProcName   = ''
							m.nProcLineNo = 0
							m.lInClassDef = .F.

						Case "LOCAL" = m.cWord1
							m.cDefType = DEFTYPE_LOCAL
							m.nToken = 2

						Case "LPARAMETERS" = m.cWord1
							m.cDefType = DEFTYPE_PARAMETER
							m.lInClassDef = .F.
							m.nToken = 2

						Case "PARAMETERS" = m.cWord1
							m.cDefType = DEFTYPE_PARAMETER
							m.lInClassDef = .F.
							m.nToken = 2

						Case "PUBLIC" = m.cWord1
							m.cDefType = DEFTYPE_PUBLIC
							m.nToken = 2

						Case "PRIVATE" = m.cWord1
							m.cDefType = DEFTYPE_PRIVATE
							m.nToken = 2

						Case "HIDDEN" = m.cWord1
							m.cDefType = DEFTYPE_PROPERTY
							m.nToken = 2

						Case "PROTECTED" = m.cWord1
							m.cDefType = DEFTYPE_PROPERTY
							m.nToken = 2

						Case "DEFINE" = m.cWord1
							If Lenc(m.cWord2) >= 4 And "CLASS" = m.cWord2 And m.nTokenCnt > 2
								m.cClassName  = aTokens[3]
								m.cProcName   = ''
								m.nProcLineNo = 0
								m.lInClassDef = .T.

								m.cDefType = DEFTYPE_CLASS
								m.nToken = 3
								m.nStopToken = 3

								* if this is a "DEFINE CLASS <classname> AS <parentclass> OF <classlibrary>"
								* then add the class library to files to process
								If m.nTokenCnt >= 7 And Upper(aTokens[6]) == "OF"
									* add to collection of files we found to process
									This.AddFileToProcess( ;
										DEFTYPE_SETCLASSPROC, ;
										aTokens[7], ;
										m.cClassName, ;
										m.cObjName + Iif(Empty(m.cObjName) Or Empty(m.cProcName), '', '.') + m.cProcName, ;
										m.nProcLineNo, ;
										m.i, ;
										m.cCodeLine ;
										)
								Endif
							Endif

						Case "ENDDEFINE" = m.cWord1
							m.cClassName  = ''
							m.cProcName   = ''
							m.nProcLineNo = 0
							m.lInClassDef = .F.

						Case m.lInClassDef And ("DIMENSION" = m.cWord1 Or "DECLARE" = m.cWord1)
							m.cDefType = DEFTYPE_PROPERTY
							m.nToken = 2
							m.nStopToken = 2

						Case m.lInClassDef And aTokens[2] == '='
							* if we're in a class definition and we have an
							* assignment or DIMENSION/DECLARE statement,
							* then that's a Property definition
							m.cDefType = DEFTYPE_PROPERTY
							m.nToken = 1
							m.nStopToken = 1
					Endcase
				Else
					Do Case
						Case aTokens[1] == "#" And Lenc(aTokens[2]) >= 4 And "DEFINE" = Upper(aTokens[2]) && #define
							m.cDefType = DEFTYPE_DEFINE
							m.nToken = 3
							m.nStopToken = 3

						Case aTokens[1] == "#" And Lenc(aTokens[2]) >= 4 And "INCLUDE" = Upper(aTokens[2]) && #include
							* add to collection of files we found to process
							This.AddFileToProcess( ;
								DEFTYPE_INCLUDEFILE, ;
								aTokens[3], ;
								m.cClassName, ;
								m.cObjName + Iif(Empty(m.cObjName) Or Empty(m.cProcName), '', '.') + m.cProcName, ;
								m.nProcLineNo, ;
								m.i, ;
								m.cCodeLine ;
								)

						Case Upper(aTokens[1]) == "SET"
							m.cWord2 = Upper(aTokens[2])
							* SET PROCEDURE TO <program> or SET CLASSLIB TO <classlibrary>
							If Lenc(m.cWord2) >= 4 And m.nTokenCnt >= 4 And Upper(aTokens[3]) == "TO"
								Do Case
									Case "CLASSLIB" = m.cWord2
										* add to collection of files we found to process
										This.AddFileToProcess( ;
											DEFTYPE_SETCLASSPROC, ;
											DEFAULTEXT(aTokens[4], "vcx"), ;
											m.cClassName, ;
											m.cObjName + Iif(Empty(m.cObjName) Or Empty(m.cProcName), '', '.') + m.cProcName, ;
											m.nProcLineNo, ;
											m.i, ;
											m.cCodeLine ;
											)

									Case "PROCEDURE" = m.cWord2
										* SET PROCEDURE TO supports a comma-delimited list of filenames
										For m.j = 4 To m.nTokenCnt
											If Lenc(aTokens[m.j]) >= 4 And "ADDI" = Upper(aTokens[m.j])
												Exit
											Endif
											This.AddFileToProcess( ;
												DEFTYPE_SETCLASSPROC, ;
												DEFAULTEXT(aTokens[m.j], "prg"), ;
												m.cClassName, ;
												m.cObjName + Iif(Empty(m.cObjName) Or Empty(m.cProcName), '', '.') + m.cProcName, ;
												m.nProcLineNo, ;
												m.i, ;
												m.cCodeLine ;
												)
										Endfor
								Endcase
							Endif

						Case m.lInClassDef And aTokens[2] == '='
							* if we're in a class definition and we have an
							* assignment or DIMENSION/DECLARE statement,
							* then that's a Property definition
							m.cDefType = DEFTYPE_PROPERTY
							m.nToken = 1
							m.nStopToken = 1
					Endcase
				Endif
			Endif

			* Grab all definitions from this line
			If m.nToken > 0
				Do While m.nToken <= m.nStopToken
					m.cDefinition= aTokens[m.nToken]

					If m.cDefinition == m.cStopToken
						Exit
					Endif

					If Isalpha(m.cDefinition) Or m.cDefinition = '_'
						m.cUpperDefinition = Upper(m.cDefinition)

						Do Case
							Case m.cUpperDefinition == "ARRAY" Or m.cUpperDefinition == "ARRA"
								m.nToken = m.nToken + 1
								Loop

							Case m.cUpperDefinition == "AS" Or m.cUpperDefinition == "OF" Or (Lenc(m.cUpperDefinition) >= 4 And m.cUpperDefinition = "OLEPUBLIC")
								m.nToken = m.nToken + 2
								Loop

						Endcase

						This.AddDefinition( ;
							m.cDefinition, ;
							m.cDefType, ;
							IIF(m.cDefType == DEFTYPE_CLASS, '', m.cClassName), ;
							m.cObjName + Iif(Empty(m.cObjName) Or Empty(m.cProcName), '', '.') + m.cProcName, ;
							IIF(nSearchType == SEARCHTYPE_EXPR, 0, m.nProcLineNo + m.nOffset), ;
							IIF(nSearchType == SEARCHTYPE_EXPR, 0, m.i), ;
							m.cCodeLine ;
							)

						If m.cDefType == DEFTYPE_DEFINE && for a #DEFINE, only grab the word after the #DEFINE statement
							Exit
						Endif
					Endif
					m.nToken = m.nToken + 1
				Enddo
			Endif

			m.nProcLineNo = m.nProcLineNo + 1
			m.nTokenCnt = 0
		Endfor

		If m.lUseMemLines
			Set Memowidth To (m.nMemoWidth)
		Endif

	Endfunc


	* Abstract:
	*   Strip tabs, spaces from the beginning
	*	of a code line.
	*	This is necessary because LTRIM()
	*	does not handle tabs (only spaces)
	*
	Function StripTabs(cRefCode)
		Return Alltrim(Chrtran(Rtrim(m.cRefCode), Tab, ' '))
	Endfunc

	* compile a block of code and return the compiled
	* version of it
	Function CompileCode(cCodeBlock)
		Local cObjCode
		Local cTempFile
		Local cSafety

		cObjCode = .Null.
		cTempFile = Addbs(Getenv("TMP")) + Rightc(Sys(2015), 8)

		cSafety = Set("SAFETY")
		Set Safety Off

		If Strtofile(cCodeBlock, cTempFile + ".prg") > 0
			Compile (cTempFile + ".prg")
			cObjCode = Filetostr(cTempFile + ".fxp")
		Endif
		Erase (cTempFile + ".prg")
		Erase (cTempFile + ".fxp")

		Set Safety &cSafety)

		Return cObjCode
	Endfunc
Enddefine
