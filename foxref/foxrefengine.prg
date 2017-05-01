* Abstract...:
*	Primary class for Code References application.
*
* Changes....:
*
#Include "foxpro.h"
#Include "foxref.h"

#Define USEPROJECTHOOK .F.

Define Class FoxRef As Session
	Protected lIgnoreErrors As Boolean
	Protected lRefreshMode
	Protected cProgressForm
	Protected lCancel
	Protected tTimeStamp
	Protected lIgnoreErrors

	Name = 'FoxRef'

	* search match engines
	MatchClass		   = 'MatchDefault'
	MatchClassLib      = 'foxmatch.prg'
	WildMatchClass     = 'MatchWildcard'
	WildMatchClassLib  = 'foxmatch.prg'

	* default search engine for Open Window
	FindWindowClass    = 'RefSearchWindow'
	FindWindowClassLib = 'FoxRefSearch_Window.prg'

	oSearchEngine   = .Null.
	oWindowEngine   = .Null.

	WindowHandle    = -1
	WindowFilename  = ''
	WindowLineNo    = 0


	Comments        = COMMENTS_INCLUDE
	MatchCase       = .F.
	WholeWordsOnly  = .F.
	ProjectHomeDir  = .F.  && True to search only files in Project's Home Directory or below

	SubFolders      = .F.
	Wildcards       = .F.
	Quiet           = .F.  && quiet mode -- don't display search progress
	ShowProgress    = .T.  && show a progress form

	Errors          = .Null.

	FileTypes       = ''
	ReportFile      = REPORT_FILE

	XSLTemplate     = 'foxref.xsl'

	* MRU array lists
	Dimension aLookForMRU[10]
	Dimension aReplaceMRU[10]
	Dimension aFolderMRU[10]
	Dimension aFileTypesMRU[10]
	Dimension aDefaultFileTypes[1]

	aLookForMRU     = ''
	aReplaceMRU     = ''
	aFolderMRU      = ''
	aFileTypesMRU   = ''

	Pattern             = ''
	OverwritePrior      = .F.
	ConfirmReplace      = .T.  && confirm each replacement
	BackupOnReplace     = .T.  && false to not backup when doing global replace
	DisplayReplaceLog   = .T.  && create activity log for replacements
	PreserveCase        = .F.  && preserve case during a Replace operation

	FoxRefDirectory = ''
	RefTable        = ''
	DefTable        = ''
	FileTable       = ''
	AddInTable      = ''
	ProjectFile     = ''
	FileDirectory   = ''

	ActivityLog     = ''

	* The following are set by the Options dialog
	* There should be a corresponding entry in FoxRefOption.DBF
	* (except for BackupStyle & FontString)
	IncludeDefTable     = .T.  && create Definition table when searching
	CodeOnly            = .F.  && search only source code & expressions (not names and other none-code items)
	CheckOutSCC     	= .F.  && True if items must be checked out; only if using SCC
	TreeViewFolderNames	= .F.  && True if TreeView shows folder names
	FormProperties      = .T.  && search form/class property names & values
	AutoProjectHomeDir  = .F.  && True to search only files in Project's Home Directory or below when doing definitions automatically
	ShowRefsPerLine     = .F.  && display a column in search results that depicts number of references found on the line
	ShowFileTypeHistory	= .F.  && True to keep filetype history in addition to showing common filetypes in search dialog
	ShowDistinctMethodLine = .F. && True to show columns for method/line apart from Class
	SortMostRecentFirst = .F.

	BackupStyle         = 1    && 1 = "filename.ext.bak"   2 = "Backup of filename.ext"
	FontString          = FONT_DEFAULT

	* XML Export Options
	XMLFormat          = XMLFORMAT_ELEMENTS
	XMLSchema          = .T.

	* This is the SetID for the last Replacement Log after it's saved
	ReplaceLogSetID    = ''

	* properties used internally
	cSetID              = ''
	lRefreshMode        = .F.
	lIgnoreErrors       = .F.
	oProgressForm       = .Null.
	lCancel             = .F.
	tTimeStamp          = .Null.
	lDefinitionsOnly    = .F.

	* collection of files we've backed up in this session
	oBackupCollection   = .Null.

	oFileCollection     = .Null.
	oSearchCollection   = .Null.
	oProcessedCollection = .Null.

	oEngineCollection   = .Null.
	oReportCollection   = .Null.

	oFileTypeCollection = .Null.

	cTalk           = ''
	nLangOpt        = 0
	cEscapeState    = ''
	cSYS3054        = ''
	cSaveUDFParms   = ''
	cSaveLib        = ''
	cExclusive      = ''
	cCompatible     = ''


	oOptions        = .Null.
	lInitError      = .F.

	oProjectFileRef = .Null.


	Procedure Init (lRestorePrefs)
		Local nSelect
		Local oException
		Local cAddInType
		Local nMemoWidth

		This.cTalk = Set ('TALK')
		Set Talk Off
		Set Deleted On


		This.cCompatible = Set ('COMPATIBLE')
		Set Compatible Off

		This.cExclusive = Set ('EXCLUSIVE')
		Set Exclusive Off

		This.nLangOpt = _vfp.LanguageOptions
		_vfp.LanguageOptions = 0

		This.cEscapeState = Set ('ESCAPE')
		Set Escape Off

		This.cSYS3054 = Sys(3054)
		Sys(3054, 0)

		This.cSaveLib      = Set ('LIBRARY')

		This.cSaveUDFParms = Set ('UDFPARMS')
		Set Udfparms To Value

		Set Exact Off

		* Changed on 02/14/2002 06:13:18 PM by Ryan - moved to FoxRefStart.prg
		* THIS.RestorePrefs()

		* collection of errors
		This.Errors = Newobject ('CFoxRefCollection', 'FoxRefCollection.prg')

		* create a collection that contains the files we've backup up
		This.oBackupCollection = Newobject ('CFoxRefCollection', 'FoxRefCollection.prg')
		This.oFileCollection   = Newobject ('CFoxRefCollection', 'FoxRefCollection.prg')
		This.oSearchCollection = Newobject ('CFoxRefCollection', 'FoxRefCollection.prg')
		This.oProcessedCollection = Createobject ('Collection')

		This.oEngineCollection = Createobject ('Collection')
		This.oReportCollection = Createobject ('Collection')
		This.oFileTypeCollection = Createobject ('Collection') && default filetypes from FoxRefAddin

		This.oOptions = Newobject ('FoxResource', 'FoxResource.prg')

		If Not File (RESOURCE_FILE)
			This.CreateResourceFile (RESOURCE_FILE)
		Endif

		If m.lRestorePrefs
			This.RestorePrefs()
		Endif

		This.FoxRefDirectory = This.FoxRefDirectory

	Endfunc


	Procedure Destroy()
		Local cCompatible

		This.CloseProgress()

		If This.cEscapeState = 'ON'
			Set Escape On
		Endif
		If This.cTalk = 'ON'
			Set Talk On
		Endif
		If This.cExclusive = 'ON'
			Set Exclusive On
		Endif
		Sys(3054, Int (Val (This.cSYS3054)))

		_vfp.LanguageOptions = This.nLangOpt

		If This.cSaveUDFParms = 'REFERENCE'
			Set Udfparms To Reference
		Endif

		m.cCompatible = This.cCompatible
		Set Compatible &cCompatible
	Endfunc

	Function FoxRefDirectory_Assign (cFoxRefDirectory)
		m.cFoxRefDirectory = Addbs (m.cFoxRefDirectory)
		If Empty (m.cFoxRefDirectory) Or Not Directory (m.cFoxRefDirectory)
			m.cFoxRefDirectory = Addbs (Home(7))
			If Not Directory (m.cFoxRefDirectory)
				m.cFoxRefDirectory = Addbs (Home())
			Endif
		Endif
		This.FoxRefDirectory = m.cFoxRefDirectory

		This.DefTable   = This.FoxRefDirectory + DEF_TABLE
		This.FileTable  = This.FoxRefDirectory + FILE_TABLE
		This.AddInTable = This.FoxRefDirectory + ADDIN_TABLE

		This.InitAddIns()
		This.OpenTables()
	Endfunc

	* Open the Add In table which contains our filetypes to process
	Function InitAddIns (lExclusive)
		Local nSelect
		Local nMemoWidth
		Local cAddInType

		m.nSelect = Select()

		* if we don't find the AddIn table on disk, then copy out
		* our project version of it
		If Not File (Forceext (This.AddInTable, 'DBF'))
			Try
				Use FoxRefAddin In 0 Shared Again
				Select FoxRefAddin
				Copy To (This.AddInTable) With Production
			Catch
			Finally
				If Used ('FoxRefAddin')
					Use In FoxRefAddin
				Endif
			Endtry
		Endif

		Try
			If m.lExclusive
				Use (This.AddInTable) Alias AddInCursor In 0 Exclusive
			Else
				Use (This.AddInTable) Alias AddInCursor In 0 Shared Again
			Endif
		Catch
		Endtry

		If Not Used ('AddInCursor')
			* we didn't find the Add-In table on disk, so use our built-in version
			Try
				Use FoxRefAddin Alias AddInCursor In 0 Shared Again
			Catch
			Endtry
		Endif

		* process AddIn table
		This.oEngineCollection = Createobject ('Collection')
		This.oReportCollection = Createobject ('Collection')
		This.oFileTypeCollection = Createobject ('Collection') && default filetypes from FoxRefAddin

		This.AddFileType ('*.*', FILETYPE_CLASS_DEFAULT, FILETYPE_LIBRARY_DEFAULT)

		If Used ('AddInCursor')
			m.nMemoWidth = Set ('MEMOWIDTH')
			Set Memowidth To 1000

			* setup file types
			Select AddInCursor
			Scan All
				m.cAddInType = Rtrim (AddInCursor.Type)

				Do Case
					Case m.cAddInType == ADDINTYPE_FINDFILE
						This.AddFileType (AddInCursor.Data, AddInCursor.ClassName, AddInCursor.Classlib)

					Case m.cAddInType == ADDINTYPE_IGNOREFILE
						This.AddFileType (AddInCursor.Data)

					Case m.cAddInType == ADDINTYPE_FINDWINDOW
						If Not Empty (AddInCursor.ClassName)
							This.FindWindowClass    = AddInCursor.ClassName
							This.FindWindowClassLib = AddInCursor.Classlib
						Endif

					Case m.cAddInType == ADDINTYPE_MATCH
						If Not Empty (AddInCursor.ClassName)
							This.MatchClass    = AddInCursor.ClassName
							This.MatchClassLib = AddInCursor.Classlib
						Endif

					Case m.cAddInType == ADDINTYPE_WILDMATCH
						If Not Empty (AddInCursor.ClassName)
							This.WildMatchClass    = AddInCursor.ClassName
							This.WildMatchClassLib = AddInCursor.Classlib
						Endif

					Case m.cAddInType == ADDINTYPE_REPORT
						This.AddReport (AddInCursor.Data, AddInCursor.Classlib, AddInCursor.ClassName, AddInCursor.Method, AddInCursor.Filename)

					Case m.cAddInType == ADDINTYPE_FILETYPE
						This.oFileTypeCollection.Add (AddInCursor.Data)
				Endcase
			Endscan
			Set Memowidth To (m.nMemoWidth)

		Endif
		Select (m.nSelect)
	Endfunc

	Procedure FontString_Access
		If Empty (This.FontString)
			Return FONT_DEFAULT
		Else
			Return This.FontString
		Endif
	Endproc


	* Add another file to process definitions for
	* These are files that aren't in our normal
	* scope processing, but are discovered along the
	* way, such as #include
	Function AddFileToProcess (cFilename)
		If File (m.cFilename)
			This.oFileCollection.AddNoDupe (Lower (Fullpath (m.cFilename)))
		Endif
	Endfunc

	* Add a file to search -- these are files that aren't
	* in our normal scope processing, but we search anyhow,
	* such as a Table within a DBC when searching a project
	Function AddFileToSearch (cFilename)
		If File (m.cFilename)
			This.oSearchCollection.AddNoDupe (Lower (Fullpath (m.cFilename)))
		Endif
	Endfunc


	* Add a report from the Add-Ins
	Function AddReport (cReportName, cClassLib, cClassName, cMethod, cFilename)
		Local oReportAddIn

		oReportAddIn = Newobject ('ReportAddIn')
		oReportAddIn.ReportName      = m.cReportName
		oReportAddIn.RptClassLibrary = m.cClassLib
		oReportAddIn.RptClassName    = m.cClassName
		oReportAddIn.RptMethod       = m.cMethod
		oReportAddIn.RptFilename     = m.cFilename

		Try
			This.oReportCollection.Add (oReportAddIn, m.cReportName)
		Catch
		Endtry
	Endfunc

	* Add a filetype search
	Function AddFileType (cFileSkeleton, cClassName, cClassLibrary)
		Local nIndex
		Local lSuccess
		Local oEngine

		If Vartype (m.cClassName) # 'C'
			m.cClassName = ''
		Endif

		m.cFileSkeleton = Upper (Alltrim (m.cFileSkeleton))
		If Empty (m.cFileSkeleton)
			m.cFileSkeleton = '*.*'
		Endif

		For m.i = 1 To This.oEngineCollection.Count
			If This.oEngineCollection.GetKey (m.i) == m.cFileSkeleton
				This.oEngineCollection.Remove (m.i)
				Exit
			Endif
		Endfor

		If Empty (m.cClassName)
			This.oEngineCollection.Add (.Null., m.cFileSkeleton)
		Else
			Try
				m.oEngine = Newobject (m.cClassName, m.cClassLibrary)
			Catch
				m.oEngine = .Null.
			Finally
			Endtry
			If Not Isnull (oEngine)
				If m.cFileSkeleton = '*.*' And This.oEngineCollection.Count > 0
					This.oEngineCollection.Add (m.oEngine, m.cFileSkeleton, 1)
				Else
					This.oEngineCollection.Add (m.oEngine, m.cFileSkeleton)
				Endif
			Endif
		Endif

		Return m.lSuccess
	Endfunc

	* Initialize all of the search engines
	* with our search options
	Procedure SearchInit()
		Local i
		Local oException
		Local nSelect
		Local lSuccess

		nSelect = Select()

		This.ClearErrors()

		m.oException = .Null.
		m.lSuccess = .T.

		This.oProcessedCollection.Remove (-1)

		Try
			If This.Wildcards
				This.oSearchEngine = Newobject (This.WildMatchClass, This.WildMatchClassLib)
			Else
				This.oSearchEngine = Newobject (This.MatchClass, This.MatchClassLib)
			Endif
			m.lSuccess = This.oSearchEngine.InitEngine()
		Catch To oException
			m.lSuccess = .F.
			Messagebox (m.oException.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
		Endtry


		If m.lSuccess
			This.oSearchEngine.MatchCase      = This.MatchCase
			This.oSearchEngine.WholeWordsOnly = This.WholeWordsOnly

			If This.oSearchEngine.SetPattern (This.Pattern)
				* IF THIS.WindowHandle >= 0
				* this create the engine for searching open windows
				Try
					This.oWindowEngine = Newobject (This.FindWindowClass, This.FindWindowClassLib)
				Catch To oException
					Messagebox (m.oException.Message + Chr(10) + Chr(10) + This.FindWindowClass + ' (' + This.FindWindowClassLib + ')', MB_ICONSTOP, APPNAME_LOC)
				Endtry
				If Vartype (This.oWindowEngine) == 'O'
					With This.oWindowEngine
						.SetID           = This.cSetID
						.oSearchEngine   = This.oSearchEngine
						.Pattern         = This.Pattern
						.Comments        = This.Comments
						.IncludeDefTable = This.IncludeDefTable
						.FormProperties  = This.FormProperties
						.CodeOnly        = This.CodeOnly
						.CheckOutSCC	 = This.CheckOutSCC
						.TreeViewFolderNames  = This.TreeViewFolderNames
						.PreserveCase    = This.PreserveCase
					Endwith
				Endif
				* ENDIF

				* collection of engines for various filetypes
				For m.i = 1 To This.oEngineCollection.Count
					If Vartype (This.oEngineCollection.Item (m.i)) == 'O'
						With This.oEngineCollection.Item (m.i)
							.SetID             = ''
							.oSearchEngine     = This.oSearchEngine
							.Pattern           = This.Pattern
							.Comments          = This.Comments
							.IncludeDefTable   = This.IncludeDefTable
							.FormProperties    = This.FormProperties
							.CodeOnly          = This.CodeOnly
							.CheckOutSCC	   = This.CheckOutSCC
							.TreeViewFolderNames	   = This.TreeViewFolderNames
							.PreserveCase      = This.PreserveCase
						Endwith
					Endif
				Endfor

			Else
				m.lSuccess = .F.
				Messagebox (ERROR_PATTERN_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
			Endif
		Endif

		Return m.lSuccess
	Endproc

	* Global Replace on all checked files in current RefTable
	Function GlobalReplace (cReplaceText)
		Local nSelect
		Local i
		Local j
		Local nReplaceCnt
		Local nRefCnt
		Local nFileCnt
		Local cDescription
		Local cOriginalText
		Local cModifiedText
		Local cOriginalHTML
		Local cModifiedHTML
		Local nUpdateCnt
		Local cFilename
		Local nSaveSYS3099
		Local nReplaceIndex
		Local nLastPos
		Local Array aReplaceFiles[1]
		Local Array aRefList[1]
		Local Array aReplaceList[1]

		m.nSelect = Select()

		This.ClearErrors()

		This.ActivityLog = ''

		m.nUpdateCnt = 0

		* we must replace in reverse order
		* for example, by line # descending, then by column descending

		If Not Used ('RefFile')
			Use (This.FileTable) Alias RefFile In 0 Shared Again
		Endif

		* process by file so we can check to make sure
		* the file hasn't changed between when we did
		* our search and now
		Select  Distinct																		;
			RefTable.FileID,																;
			RefTable.Timestamp,																;
			FileTable.UniqueID,																;
			Padr (FileTable.Folder, 240)	As	SortFolder,									;
			FileTable.Filename																;
			From (This.RefTable) RefTable INNER Join (This.FileTable) FileTable On RefTable.FileID == FileTable.UniqueID ;
			Where Checked And Not Inactive														;
			Order By SortFolder, Filename														;
			Into Array aReplaceFiles
		m.nFileCnt = _Tally
		For m.i = 1 To nFileCnt
			If Seek (aReplaceFiles[m.i, 3], 'RefFile', 'UniqueID')
				m.cFilename = Addbs (Rtrim (RefFile.Folder)) + Rtrim (RefFile.Filename)

				If File (m.cFilename)
					If Fdate (m.cFilename, 1) == aReplaceFiles[m.i, 2] && check timestamp on file
						* this groups replacements into ones that occur on the same line

						m.nSaveSYS3099 = Sys(3099)
						Sys(3099, 70)

						Select  RefID															;
							From (This.RefTable)												;
							Where FileID == aReplaceFiles[m.i, 1] And Checked And Not Inactive	;
							Group By RefID														;
							Order By ColPos Descend												;
							Into Array aRefList
						m.nRefCnt = _Tally

						Sys(3099, m.nSaveSYS3099)

						For m.j = 1 To m.nRefCnt
							Select  UniqueID, Abstract, ColPos, MatchLen, ClassName, ProcName, ProcLineNo, FindType ;
								From (This.RefTable)											;
								Where RefID == aRefList[m.j, 1]									;
								Order By ColPos Descend											;
								Into Array aReplaceList
							m.nReplaceCnt = _Tally

							If m.nReplaceCnt > 0
								* get the original and the modified text so we
								* can show it in our confirmation dialog
								m.cOriginalText = aReplaceList[1, 2]
								m.cModifiedText = m.cOriginalText

								m.cOriginalHTML = m.cOriginalText
								m.cModifiedHTML = m.cModifiedText
								For m.nReplaceIndex = 1 To m.nReplaceCnt
									m.nLastPos = aReplaceList[m.nReplaceIndex, 3] + aReplaceList[m.nReplaceIndex, 4]
									m.cOriginalHTML = Leftc (m.cOriginalHTML, aReplaceList[m.nReplaceIndex, 3] - 1) + [<font color="blue">] + Substrc (m.cOriginalHTML, aReplaceList[m.nReplaceIndex, 3], aReplaceList[m.nReplaceIndex, 4]) + [</font>] + ;
										IIf (m.nLastPos > Lenc (m.cOriginalHTML), '', Substrc (m.cOriginalHTML, m.nLastPos))
									m.cModifiedHTML = Leftc (m.cModifiedHTML, aReplaceList[m.nReplaceIndex, 3] - 1) + [<font color="blue">] + IIf (This.PreserveCase, This.SetCasePreservation (Substrc (m.cModifiedHTML, aReplaceList[m.nReplaceIndex, 3], aReplaceList[m.nReplaceIndex, 4]), m.cReplaceText), m.cReplaceText) + [</font>] + ;
										IIf (m.nLastPos > Lenc (m.cModifiedHTML), '', Substrc (m.cModifiedHTML, m.nLastPos))
								Endfor

								* this is the description of the filename, class.method we're replacing on
								m.cDescription = m.cFilename + ', ' +							;
									IIf (Empty (aReplaceList[1, 5]) And Empty (aReplaceList[1, 6]), '', ;
									aReplaceList[1, 5] + IIf (Empty (aReplaceList[1, 5]) Or Empty (aReplaceList[1, 6]), '', '.') + aReplaceList[1, 6]) + ;
									IIf (aReplaceList[1, 7] == 0, '', IIf (Empty (aReplaceList[1, 5] + aReplaceList[1, 6]), '', ', ') + Ltrim (Str (aReplaceList[1, 7], 8, 0)))

								If This.ConfirmReplace
									Do Form FoxRefReplaceConfirm With m.cDescription, m.cOriginalHTML, m.cModifiedHTML To m.lDoReplace
									If Isnull (m.lDoReplace)
										Exit
									Endif
								Else
									m.lDoReplace = .T.
								Endif

								If m.lDoReplace
									This.AddLog (LOG_PREFIX + m.cDescription + ':')
									This.AddLog (LOG_PREFIX + m.cOriginalHTML + ' -> ' + m.cModifiedHTML)

									If This.ReplaceFile (aRefList[m.j, 1], m.cReplaceText)
										m.nUpdateCnt = m.nUpdateCnt + 1
									Endif
								Else
									This.AddLog (LOG_PREFIX + m.cDescription + ': ' + REPLACE_SKIPPED)
								Endif
								This.AddLog()
							Endif
						Endfor
					Else
						* timestamps don't match
						This.AddError (m.cFilename + ': ' + ERROR_FILEMODIFIED_LOC)
						This.AddLog (LOG_PREFIX + m.cFilename + ': ' + ERROR_FILEMODIFIED_LOC)
					Endif
				Else
					* file not found
					This.AddError (m.cFilename + ': ' + ERROR_FILENOTFOUND_LOC)
					This.AddLog (LOG_PREFIX + m.cFilename + ': ' + ERROR_FILENOTFOUND_LOC)
				Endif
			Endif
		Endfor

		This.SaveLog (m.cReplaceText)

		Select (m.nSelect)

		Return m.nUpdateCnt > 0
	Endfunc

	* take the original text, determine the case,
	* and apply it the new text
	Function SetCasePreservation (cOriginalText, cReplacementText)
		Local i
		Local ch
		Local lLower
		Local lUpper

		If Isupper (Leftc (m.cOriginalText, 1)) And Islower (Substrc (m.cOriginalText, 2, 1))
			* proper case
			m.cReplacementText = Proper (m.cReplacementText)
		Else
			m.lLower = .F.
			m.lUpper = .F.
			For m.i = 1 To Lenc (m.cOriginalText)
				ch = Substrc (m.cOriginalText, m.i, 1)
				Do Case
					Case Between (ch, 'A', 'Z')
						m.lUpper = .T.
					Case Between (ch, 'a', 'z')
						m.lLower = .T.
				Endcase

				If m.lLower And m.lUpper
					Exit
				Endif
			Endfor

			Do Case
				Case m.lLower And Not m.lUpper
					m.cReplacementText = Lower (m.cReplacementText)
				Case Not m.lLower And m.lUpper
					m.cReplacementText = Upper (m.cReplacementText)
			Endcase
		Endif

		Return m.cReplacementText
	Endfunc


	* Do a replacement on designated file
	Function ReplaceFile (cRefID, cReplaceText, lBackupOnReplace)
		Local nSelect
		Local oFoxRefRecord
		Local lSuccess
		Local oEngine
		Local lBackup
		Local cFilename
		Local oException
		Local oReplaceCollection

		m.lSuccess = .F.
		If Used ('FoxRefCursor') And Vartype (cRefID) == 'C'
			m.nSelect = Select()

			If Pcount() < 3 Or Vartype (m.lBackupOnReplace) # 'L'
				m.lBackupOnReplace = This.BackupOnReplace
			Endif

			If Seek (m.cRefID, 'FoxRefCursor', 'RefID')
				If Used ('FileCursor') And Seek (FoxRefCursor.FileID, 'FileCursor', 'UniqueID')
					m.cFilename = This.DiskFileName(Addbs (Rtrim (FileCursor.Folder)) + Rtrim (FileCursor.Filename))

					If File (m.cFilename)
						m.oEngine = This.GetEngine (Justfname (m.cFilename))

						* if we don't have an engine object defined, then assume
						* we're supposed to ignore this filetype
						If Vartype (m.oEngine) == 'O'
							With m.oEngine
								.PreserveCase = This.PreserveCase
								m.lSuccess = .T.

								* create a backup of the file if we haven't already
								* done so in this session
								If m.lBackupOnReplace
									* m.cFilename = ADDBS(RTRIM(m.oFoxRefRecord.Folder)) + RTRIM(m.oFoxRefRecord.Filename)
									m.lBackup = .T.
									For m.i = 1 To This.oBackupCollection.Count
										If This.oBackupCollection.Item (m.i) == Upper (m.cFilename)
											m.lBackup = .F.
											Exit
										Endif
									Endfor
									If m.lBackup
										If .BackupFile (m.cFilename, This.BackupStyle)
											* add to collection so we don't backup again during this session
											This.oBackupCollection.Add (Upper (m.cFilename))
										Else
											m.lSuccess = .F.
											This.AddError (m.cFilename + ': ' + ERROR_NOBACKUP_LOC)
										Endif
									Endif
								Endif


								* do the replace
								If m.lSuccess
									oReplaceCollection = Createobject ('Collection')

									Select FoxRefCursor
									Scan All For RefID == cRefID
										Scatter Memo Name oFoxRefRecord
										oReplaceCollection.Add (oFoxRefRecord)
									Endscan
									m.oException = .ReplaceWith (m.cReplaceText, m.oReplaceCollection, m.cFilename)
									If Not Isnull (m.oException)
										m.lSuccess = .F.
										This.AddError (m.cFilename + ': ' + IIf (Empty (m.oException.UserValue), m.oException.Message, m.oException.UserValue))
										This.AddLog (LOG_PREFIX + ERROR_LOC + ': ' + IIf (Empty (m.oException.UserValue), m.oException.Message, m.oException.UserValue))
									Endif
									This.AddLog (.ReplaceLog)
								Endif
							Endwith

						Else
							This.AddError (m.cFilename + ': ' + ERROR_NOENGINE_LOC)
						Endif
					Else
						This.AddError (m.cFilename + ': ' + ERROR_FILENOTFOUND_LOC)
					Endif
				Endif

			Endif
			Select (m.nSelect)
		Endif



		Return m.lSuccess
	Endfunc




	* Methods for updating Errors collection
	Procedure ClearErrors()
		This.Errors.Remove (-1)
	Endproc

	Procedure AddError (cErrorMsg)
		This.Errors.Add (cErrorMsg)
	Endproc

	Procedure AddLog (cLog)
		If Pcount() == 0
			If Not Empty (This.ActivityLog)
				This.ActivityLog = This.ActivityLog + Chr(13) + Chr(10)
			Endif
		Else
			If Not Empty (m.cLog)
				This.ActivityLog = This.ActivityLog + m.cLog + Chr(13) + Chr(10)
			Endif
		Endif
	Endproc


	* Save replacement log
	Function SaveLog (cReplaceText)
		If Vartype (m.cReplaceText) # 'C'
			m.cReplaceText = ''
		Endif

		This.ReplaceLogSetID = Sys(2015)
		Insert Into FoxRefCursor (			;
			UniqueID,						;
			SetID,						;
			RefID,						;
			RefType,						;
			FileID,						;
			Symbol,						;
			ClassName,					;
			ProcName,						;
			ProcLineNo,					;
			Lineno,						;
			ColPos,						;
			MatchLen,						;
			Abstract,						;
			RecordID,						;
			UpdField,						;
			Checked,						;
			NoReplace,					;
			Timestamp,					;
			Inactive						;
			) Values (					;
			Sys(2015),					;
			This.ReplaceLogSetID,			;
			Sys(2015),					;
			REFTYPE_LOG,					;
			'',							;
			M.cReplaceText,				;
			'',							;
			'',							;
			0,							;
			0,							;
			0,							;
			0,							;
			This.ActivityLog,				;
			'',							;
			'',							;
			.F.,							;
			.F.,							;
			Datetime(),					;
			.F.							;
			)
	Endfunc




	* Make sure the Filetypes list is delimited with spaces
	Procedure FileTypes_Assign (cFileTypes)
		*** JRN 2010-03-14 : Don't remove spaces
		This.FileTypes = cFileTypes
		*	THIS.FileTypes = CHRTRAN(cFileTypes, ',;', '  ')
	Endfunc


	* ---
	* --- Definition Table methods
	* ---
	Function CreateFileTable()
		Local lSuccess
		Local cSafety

		m.lSuccess = .T.

		m.cSafety = Set ('SAFETY')
		Set Safety Off
		If Used (Juststem (This.FileTable))
			Use In (This.FileTable)
		Endif


		Try
			Create Table (This.FileTable) Free (		;
				UniqueID C(10),						;
				Folder M,								;
				Filename C(100),						;
				FileAction C(1),						;
				Timestamp T Null						;
				)
		Catch
			m.lSuccess = .F.
			Messagebox (ERROR_CREATEFILETABLE_LOC + Chr(10) + Chr(10) + Forceext (This.FileTable, 'DBF'), MB_ICONSTOP, APPNAME_LOC)
		Endtry

		If m.lSuccess
			* insert a record that represents open windows
			Insert Into (This.FileTable) (		;
				UniqueID,						;
				Filename,						;
				Folder,						;
				Timestamp,					;
				FileAction					;
				) Values (					;
				'WINDOW',						;
				OPENWINDOW_LOC,				;
				'',							;
				Datetime(),					;
				FILEACTION_NODEFINITIONS		;
				)

			Index On UniqueID Tag UniqueID
			Index On Filename Tag Filename
			Index On FileAction Tag FileAction

			Use In (Juststem (This.FileTable))
		Endif

		Set Safety &cSafety

		Return m.lSuccess
	Endfunc

	Function CreateDefTable()
		Local lSuccess
		Local cSafety

		m.lSuccess = .T.

		m.cSafety = Set ('SAFETY')
		Set Safety Off
		If Used (Juststem (This.DefTable))
			Use In (This.DefTable)
		Endif


		Try
			Create Table (This.DefTable) Free (			;
				UniqueID C(10),						;
				DefType C(1),							;
				FileID C(10),							;
				Symbol M,								;
				ClassName M,							;
				ProcName M,							;
				ProcLineNo i,							;
				Lineno i,								;
				Abstract M,							;
				Inactive L							;
				)

		Catch
			m.lSuccess = .F.
			Messagebox (ERROR_CREATEDEFTABLE_LOC + Chr(10) + Chr(10) + Forceext (This.DefTable, 'DBF'), MB_ICONSTOP, APPNAME_LOC)
		Endtry

		If m.lSuccess
			Index On UniqueID Tag UniqueID
			Index On DefType Tag DefType
			Index On Inactive Tag Inactive
			Index On FileID Tag FileID

			Use In (Juststem (This.DefTable))
		Endif

		Set Safety &cSafety

		Return m.lSuccess
	Endfunc


	* Open a File & Definition tables
	Function OpenTables (lExclusive, lQuiet)
		Local lSuccess

		This.CloseTables()

		m.lSuccess = .T.

		If Not File (Forceext (This.FileTable, 'DBF'))
			m.lSuccess = This.CreateFileTable()
		Endif
		If m.lSuccess And Not File (Forceext (This.DefTable, 'DBF'))
			m.lSuccess = This.CreateDefTable()
		Endif

		If m.lSuccess
			* Open the table of files processed
			Try
				If m.lExclusive
					Use (This.FileTable) Alias FileCursor In 0 Exclusive
				Else
					Use (This.FileTable) Alias FileCursor In 0 Shared Again
				Endif
			Catch
			Endtry

			* Open the table of Definitions
			Try
				If m.lExclusive
					Use (This.DefTable) Alias FoxDefCursor In 0 Exclusive
				Else
					Use (This.DefTable) Alias FoxDefCursor In 0 Shared Again
				Endif
			Catch
			Endtry



			If Used ('FileCursor')
				If Type ('FileCursor.UniqueID') # 'C' Or Type ('FileCursor.Filename') # 'C'
					m.lSuccess = .F.
					If Not m.lQuiet
						Messagebox (ERROR_BADFILETABLE_LOC + Chr(10) + Chr(10) + Forceext (This.FileTable, 'DBF'), MB_ICONSTOP, APPNAME_LOC)
					Endif
				Endif
			Else
				If Not m.lQuiet
					Messagebox (ERROR_OPENFILETABLE_LOC + Chr(10) + Chr(10) + Forceext (This.FileTable, 'DBF'), MB_ICONSTOP, APPNAME_LOC)
				Endif
				m.lSuccess = .F.
			Endif

			If m.lSuccess
				If Used ('FoxDefCursor')
					If Type ('FoxDefCursor.DefType') # 'C'
						m.lSuccess = .F.
						If Not m.lQuiet
							Messagebox (ERROR_BADDEFTABLE_LOC + Chr(10) + Chr(10) + Forceext (This.DefTable, 'DBF'), MB_ICONSTOP, APPNAME_LOC)
						Endif
					Endif
				Else
					If Not m.lQuiet
						Messagebox (ERROR_OPENDEFTABLE_LOC + Chr(10) + Chr(10) + Forceext (This.DefTable, 'DBF'), MB_ICONSTOP, APPNAME_LOC)
					Endif
					m.lSuccess = .F.
				Endif
			Endif

		Endif

		Return m.lSuccess
	Endfunc

	Function GetAvailableOptions()
		Local oOptions
		Local nSelect
		Local oException
		Local oRefOption

		nSelect = Select()

		oOptions = .Null.
		oException = .Null.
		Try
			Select  * From					;
				FoxRefOption			;
				Order By DisplayOrd			;
				Into Cursor FoxRefOptionCursor
		Catch To oException
		Endtry

		If Isnull (oException)
			oOptions = Newobject ('Collection')
			Select FoxRefOptionCursor
			Scan All
				oRefOption = Newobject ('RefOption', 'foxrefengine.prg')
				With oRefOption
					.OptionName   = Rtrim (FoxRefOptionCursor.OptionName)
					.Description  = FoxRefOptionCursor.Descrip
					.PropertyName = Rtrim (FoxRefOptionCursor.OptionProp)
					Try
						.OptionValue  = Eval ('THIS.' + .PropertyName)
					Catch
					Endtry
				Endwith
				oOptions.Add (oRefOption, oRefOption.PropertyName)
			Endscan
		Else
			Messagebox (m.oException.Message, MB_ICONSTOP, APPNAME_LOC)
		Endif

		If Used ('FoxRefOption')
			Use In FoxRefOption
		Endif
		If Used ('FoxRefOptionCursor')
			Use In FoxRefOptionCursor
		Endif

		Select (nSelect)

		Return oOptions
	Endfunc


	Function SaveOptions (oOptionsCollection As Collection)
		Local oOptions
		Local oException
		Local oRefOption


		For Each oRefOption In oOptionsCollection
			If Not Isnull (oRefOption.OptionValue)
				Store (oRefOption.OptionValue) To ('THIS.' + oRefOption.PropertyName)
			Endif
		Endfor
	Endfunc


	Function CloseTables()
		If Used ('FoxDefCursor')
			Use In FoxDefCursor
		Endif
		If Used ('FileCursor')
			Use In FileCursor
		Endif
		If Used ('AddInCursor')
			Use In AddInCursor
		Endif
	Endfunc

	* Get the name of the first valid Reference table
	* We check the structure here of existing files
	* before we overwrite or open a reference table.
	* Increment a file count as necessary, so instead
	* of opening "giftrap_ref.dbf", we might end up
	* opening "giftrap_ref3.dbf" (if giftrap_ref, giftrap_ref1, and
	* giftrap_ref2 are not of the correct type)
	Function GetRefTableName (cRefTable)
		Local lTableIsOkay
		Local nFileCnt
		Local cNewRefTable
		Local oException
		Local cFilename

		m.cFilename = Chrtran (Justfname (m.cRefTable), INVALID_ALIAS_CHARS, Replicate ('_', Lenc (INVALID_ALIAS_CHARS)))
		m.cRefTable = Addbs (Justpath (m.cRefTable)) + m.cFilename

		m.cNewRefTable = Forceext (m.cRefTable, 'DBF')

		m.nFileCnt = 0
		m.lTableIsOkay = .F.
		Do While Not lTableIsOkay And m.nFileCnt < 100
			Try
				If File (m.cNewRefTable)
					Use (m.cNewRefTable) Alias CheckRefTable In 0 Shared Again
					m.lTableIsOkay = (Type ('CheckRefTable.RefType') == 'C')
					Use In CheckRefTable
				Else
					m.lTableIsOkay = .T.
				Endif
			Catch
			Endtry

			If Not m.lTableIsOkay
				m.nFileCnt = m.nFileCnt + 1
				m.cNewRefTable = Addbs (Justpath (m.cRefTable)) + Juststem (m.cRefTable) + Transform (m.nFileCnt) + '.DBF'
			Endif
		Enddo
		If m.nFileCnt > 0
			m.cRefTable = Addbs (Justpath (m.cRefTable)) + Juststem (m.cRefTable) + Transform (m.nFileCnt) + '.DBF'
		Endif

		Return m.cRefTable
	Endfunc

	* ---
	* --- Reference Table methods
	* ---


	Function CreateRefTable (cRefTable)
		Local lSuccess
		Local cSafety
		Local oException

		m.lSuccess = .T.

		This.RefTable = ''

		m.cRefTable = This.GetRefTableName (m.cRefTable)

		m.cSafety = Set ('SAFETY')
		Set Safety Off

		If Used (Juststem (m.cRefTable))
			Use In (Juststem (m.cRefTable))
		Endif

		Try
			Create Table (m.cRefTable) Free (		;
				UniqueID C(10),					;
				SetID C(10),						;
				RefID C(10),						;
				RefType C(1),						;
				FindType C(1),					;
				FileID C(10),						;
				Symbol M,							;
				ClassName M,						;
				ProcName M,						;
				ProcLineNo i,						;
				Lineno i,							;
				ColPos i,							;
				MatchLen i,						;
				Abstract M,						;
				RecordID C(10),					;
				UpdField C(15),					;
				Checked L Null,					;
				NoReplace L,						;
				Timestamp T Null,					;
				oTimeStamp N(10),					;
				Inactive L						;
				)
		Catch To oException
			m.lSuccess = .F.
			Messagebox (ERROR_CREATEREFTABLE_LOC + ' (' + m.oException.Message + '):' + Chr(10) + Chr(10) + Forceext (m.cRefTable, 'DBF'), MB_ICONSTOP, APPNAME_LOC)
		Endtry

		If m.lSuccess
			Index On RefType Tag RefType
			Index On SetID Tag SetID
			Index On RefID Tag RefID
			Index On UniqueID Tag UniqueID
			Index On FileID Tag FileID
			Index On Checked Tag Checked
			Index On Inactive Tag Inactive


			* add the record that holds our results window search position & other options
			Insert Into (m.cRefTable) (			;
				UniqueID,						;
				SetID,						;
				RefType,						;
				FindType,						;
				FileID,						;
				Symbol,						;
				ClassName,					;
				ProcName,						;
				ProcLineNo,					;
				Lineno,						;
				ColPos,						;
				MatchLen,						;
				Abstract,						;
				RecordID,						;
				UpdField,						;
				Timestamp,					;
				Checked,						;
				NoReplace,					;
				Inactive						;
				) Values (					;
				Sys(2015),					;
				'',							;
				REFTYPE_INIT,					;
				'',							;
				'',							;
				This.ProjectFile,				;
				'',							;
				'',							;
				0,							;
				0,							;
				0,							;
				0,							;
				'',							;
				'',							;
				'',							;
				Datetime(),					;
				.F.,							;
				.F.,							;
				.F.							;
				)

			This.RefTable = m.cRefTable

			Use
		Endif

		Set Safety &cSafety

		Return m.lSuccess
	Endfunc

	* Open a FoxRef table
	* Return TRUE if table exists and it's in the correct format
	* [lCreate]    = True to create table if it doesn't exist
	* [lExclusive] = True to open for exclusive use
	Function OpenRefTable (cRefTable, lExclusive)
		Local lSuccess
		Local oException

		If Used ('FoxRefCursor')
			Use In FoxRefCursor
		Endif
		This.RefTable = ''

		m.lSuccess = .T.

		m.cRefTable = This.GetRefTableName (m.cRefTable)

		If Not File (Forceext (m.cRefTable, 'DBF'))
			m.lSuccess = This.CreateRefTable (m.cRefTable)
		Endif

		If m.lSuccess
			Try
				If m.lExclusive
					Use (m.cRefTable) Alias FoxRefCursor In 0 Exclusive
				Else
					Use (m.cRefTable) Alias FoxRefCursor In 0 Shared Again
				Endif
			Catch To oException
				Messagebox (ERROR_OPENREFTABLE_LOC + ' (' + m.oException.Message + '):' + Chr(10) + Chr(10) + Forceext (m.cRefTable, 'DBF'), MB_ICONSTOP, APPNAME_LOC)
				m.lSuccess = .F.
			Endtry

			If m.lSuccess
				If Type ('FoxRefCursor.RefType') == 'C'
					This.RefTable = m.cRefTable
					*** JRN 2010-03-20 : Add field oTimeStamp
					If Type ('FoxRefCursor.oTimeStamp') # 'N'
						Local lnSelect
						lnSelect = Select()

						Select FoxRefCursor
						Use (cRefTable) Exclusive Alias FoxRefCursor
						Alter Table FoxRefCursor Add Column oTimeStamp N(10)

						Use (cRefTable) Alias FoxRefCursor Shared Again
						Select (lnSelect)
					Endif
					***
				Else
					m.lSuccess = .F.
					Messagebox (ERROR_BADREFTABLE_LOC + Chr(10) + Chr(10) + Forceext (m.cRefTable, 'DBF'), MB_ICONSTOP, APPNAME_LOC)
				Endif
			Endif
		Endif

		Return m.lSuccess
	Endfunc

	* return number of search sets in the current FoxRef table
	Function SearchCount()
		Local nSelect
		Local nSearchCnt
		Local Array aSearchCnt[1]

		m.nSelect = Select()
		Select  Cnt (*)											;
			From (This.RefTable)								;
			Where												;
			RefType == REFTYPE_SEARCH And Not Inactive		;
			Into Array aSearchCnt
		If _Tally > 0
			m.nSearchCnt = aSearchCnt[1]
		Else
			m.nSearchCnt = 0
		Endif

		Select (m.nSelect)

		Return m.nSearchCnt
	Endfunc

	* Returns TRUE if there is only 1 record in the FoxRef table,
	* indicating that no searches have been done yet
	* (the first record is initialization information)
	Function FirstSearch()
		Return Used ('FoxRefCursor') And Reccount ('FoxRefCursor') <= 1
	Endfunc

	* Abstract:
	*   Set a specific project to display result sets for.
	*	Pass an empty string or "global" to display result sets
	*	that are not associated with a project.
	*
	* Parameters:
	*   [cProject]
	Function SetProject (cProjectFile, lOverwrite)
		Local lSuccess
		Local cRefTable
		Local i
		Local lFoundProject
		Local lOpened
		Local oFileRef
		Local oErr

		lSuccess = .F.

		If Vartype (cProjectFile) # 'C'
			cProjectFile = This.ProjectFile
		Endif

		This.oProjectFileRef = .Null.
		If Empty (cProjectFile)
			* use the active project if a project name is not passsed
			If Application.Projects.Count > 0
				cProjectFile = Application.ActiveProject.Name
			Else
				cProjectFile = PROJECT_GLOBAL
			Endif

			lSuccess = .T.
		Else
			* make sure Project specified is open
			If cProjectFile == PROJECT_GLOBAL
				lSuccess = .T.
			Else
				cProjectFile = Upper (Forceext (Fullpath (cProjectFile), 'PJX'))

				For i = 1 To Application.Projects.Count
					If Upper (Application.Projects (i).Name) == cProjectFile
						cProjectFile = Application.Projects (i).Name
						lSuccess = .T.
						Exit
					Endif
				Endfor
				If Not lSuccess
					If File (cProjectFile)
						lOpened = .T.
						* open the project
						Try
							Modify Project (cProjectFile) Nowait
							This.AddMRUFile(cProjectFile)

						Catch To oErr
							Messagebox (oErr.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
							lOpened = .F.
						Endtry

						If lOpened
							* search again to find where in the Projects collection it is
							For i = 1 To Application.Projects.Count
								If Upper (Application.Projects (i).Name) == cProjectFile
									cProjectFile = Application.Projects (i).Name
									lSuccess = .T.
									Exit
								Endif
							Endfor
						Endif
					Endif
				Endif

			Endif
		Endif

		If lSuccess
			If Empty (cProjectFile) Or cProjectFile == PROJECT_GLOBAL
				This.ProjectFile = PROJECT_GLOBAL
				cRefTable        = This.FoxRefDirectory + GLOBAL_TABLE + RESULT_EXT
			Else
				This.ProjectFile = Upper (cProjectFile)
				cRefTable        = This.GetRefTableName (Addbs (Justpath (cProjectFile)) + Juststem (cProjectFile) + RESULT_EXT)

				This.oProjectFileRef = Application.ActiveProject
			Endif

			If lOverwrite
				lSuccess = This.CreateRefTable (cRefTable)
			Else
				lSuccess = This.OpenRefTable (cRefTable)
			Endif
		Endif
		Return lSuccess
	Endfunc

	* Add to cursor of available project files.
	* This cursor is all files in current project,
	* plus any files we encounter along the way
	* that are #include or SET PROCEDURE TO, SET CLASSLIB TO
	Function AddFileToProjectCursor (cFilename)
		Local nSelect
		Local i
		Local nCnt
		Local Array aFileList[1]

		m.nSelect = Select()

		m.cFilename = Lower (Fullpath (m.cFilename))
		m.cFolder   = Justpath (m.cFilename)

		m.cFilename = Padr (Justfname (m.cFilename), 100)

		Select ProjectFilesCursor
		Locate For Filename == m.cFilename And Folder == m.cFolder
		If Not Found()
			Insert Into ProjectFilesCursor (		;
				Filename,							;
				Folder							;
				) Values (						;
				M.cFilename,						;
				M.cFolder							;
				)
		Endif

		Select (m.nSelect)
	Endfunc

	* Grabs all Include files for this project
	* and adds them to our project list cursor
	Function UpdateProjectFiles()
		Local nSelect
		Local i
		Local nCnt
		Local oFileRef
		Local Array aFileList[1]

		m.nSelect = Select()

		If Used ('ProjectFilesCursor')
			Use In ProjectFilesCursor
		Endif
		Create Cursor ProjectFilesCursor (		;
			Folder M,							;
			Filename C(100)					;
			)
		Index On Filename Tag Filename

		* Add in all files that are in currently in the project
		If Vartype (This.oProjectFileRef) == 'O'
			For Each m.oFileRef In This.oProjectFileRef.Files
				This.AddFileToProjectCursor (oFileRef.Name)
			Endfor
		Endif

		* add in all dependencies
		Select																		;
			Distinct Padr (Leftc (Symbol, 254), 254)	  As  IncludeFile		;
			From (This.DefTable) DefTable											;
			Where																	;
			DefTable.DefType == DEFTYPE_INCLUDEFILE And							;
			Not DefTable.Inactive												;
			Into Array aFileList
		m.nCnt = _Tally
		For m.i = 1 To m.nCnt
			Try
				If File (Rtrim (aFileList[m.i, 1]))  && make sure we can find the file along our path somewhere
					This.AddFileToProjectCursor (aFileList[m.i, 1])
				Endif
			Catch
			Endtry
		Endfor

		Select (m.nSelect)
	Endfunc

	* return project files as a collection
	Function GetProjectFiles()
		Local nSelect
		Local oProjectFiles

		nSelect = Select()

		oProjectFiles = Createobject ('Collection')

		If Used ('ProjectFilesCursor')
			Select ProjectFilesCursor
			Scan All
				oProjectFiles.Add (Addbs (Rtrim (ProjectFilesCursor.Folder)) + Rtrim (ProjectFilesCursor.Filename))
			Endscan
		Endif
		Select (nSelect)

		Return oProjectFiles
	Endfunc

	* collect all definitions for a project/folder and currently open window
	* without doing an actual search
	* [lLocalOnly] = search open window for LOCALS & PARAMETERS
	Function CollectDefinitions (lLocalOnly)
		Local nSelect
		Local cOpenFile
		Local nCnt
		Local i
		Local lSuccess
		Local lDefinitionsOnly
		Local lOverwritePrior
		Local lCodeOnly
		Local lCheckOutSCC
		Local lTreeViewFolderNames
		Local cFileTypes
		Local lProjectHomeDir
		Local oException


		If Not This.SearchInit()
			Return .F.
		Endif

		lDefinitionsOnly  = This.lDefinitionsOnly
		lOverwritePrior   = This.OverwritePrior
		lCodeOnly         = This.CodeOnly
		lCheckOutSCC      = This.CheckOutSCC
		lTreeViewFolderNames      = This.TreeViewFolderNames
		cFileTypes        = This.FileTypes
		lProjectHomeDir   = This.ProjectHomeDir

		This.lDefinitionsOnly = .T.
		This.OverwritePrior   = .F.
		This.CodeOnly         = .F.
		This.CheckOutSCC      = .F.
		This.TreeViewFolderNames      = .F.
		This.FileTypes        = FILETYPES_DEFINITIONS
		This.ProjectHomeDir   = This.AutoProjectHomeDir

		m.nSelect = Select()

		* make sure we have the latest definitions for the open window
		Try
			If m.lLocalOnly
				Update  FoxDefCursor		;
					Set Inactive = .T.		;
					Where FileID = 'WINDOW'
				If This.WindowHandle >= 0 And Vartype (This.oWindowEngine) == 'O'
					This.oWindowEngine.WindowHandle = This.WindowHandle
					This.oWindowEngine.ProcessDefinitions (This)

					Update  FileCursor Set										;
						Filename = Justfname (This.WindowFilename),		;
						Folder = Justpath (This.WindowFilename),			;
						FileAction = FILEACTION_DEFINITIONS					;
						Where UniqueID = 'WINDOW'
				Endif
			Else
				If This.SetProject (.Null., .F.)
					If This.ProjectFile == PROJECT_GLOBAL Or Empty (This.ProjectFile)
						m.lSuccess = This.FolderSearch ('')
					Else
						m.lSuccess = This.ProjectSearch ('')
					Endif
				Endif

				* Process definitions for files that
				* were #included
				m.nCnt = This.oFileCollection.Count
				For m.i = 1 To m.nCnt
					m.lSuccess = This.FileSearch (This.oFileCollection.Item (m.i), '')
					If Not m.lSuccess
						Exit
					Endif
				Endfor
				This.oFileCollection.Remove (-1)
				This.oProcessedCollection.Remove (-1)

				This.UpdateProjectFiles()
			Endif
		Catch To oException
			Messagebox (oException.Message)
		Endtry

		This.lDefinitionsOnly   = lDefinitionsOnly
		This.OverwritePrior     = lOverwritePrior
		This.CodeOnly           = lCodeOnly
		This.CheckOutSCC        = lCheckOutSCC
		This.TreeViewFolderNames        = lTreeViewFolderNames
		This.FileTypes          = cFileTypes
		This.ProjectHomeDir     = lProjectHomeDir

		Select (m.nSelect)
	Endfunc


	Function Search (cPattern, lShowDialog)
		Local lSuccess
		Local i
		Local nCnt

		If Vartype (cPattern) # 'C' Or Empty (cPattern)
			lShowDialog = .T.
			cPattern = ''
		Endif

		lSuccess = .T.
		If lShowDialog
			Do Form FoxRefFind With This, cPattern
		Else
			This.lRefreshMode = .F.
			This.ClearErrors()
			If This.SetProject (.Null., This.OverwritePrior)
				If This.ProjectFile == PROJECT_GLOBAL Or Empty (This.ProjectFile)
					lSuccess = This.FolderSearch (cPattern)
				Else
					lSuccess = This.ProjectSearch (cPattern)
				Endif

				* Search for files that weren't necessarily in the
				* project, but were referenced by certain files in the project
				* (for example, a Table within a DBC)
				If m.lSuccess And Not This.lCancel
					m.nCnt = This.oSearchCollection.Count
					For m.i = 1 To m.nCnt
						m.lSuccess = This.FileSearch (This.oSearchCollection.Item (m.i), cPattern)
						If Not m.lSuccess
							Exit
						Endif
					Endfor
				Endif
				This.oSearchCollection.Remove (-1)

				* after we're all done, do definitions for files
				* that were #included
				If lSuccess
					m.nCnt = This.oFileCollection.Count
					For m.i = 1 To m.nCnt
						This.FileSearch (This.oFileCollection.Item (m.i), '')
					Endfor
				Endif
				This.oFileCollection.Remove (-1)
				This.oProcessedCollection.Remove (-1)

			Endif
		Endif

		Return lSuccess
	Endfunc


	* Determine if the Reference table we want to open is
	* actually one of ours.  If we're overwriting or a reference
	* table doesn't exist for this project, then create a new
	* Reference Table.
	*
	* Once we have a reference table, then we add a new record
	* that represents the search criteria for this particular
	* search.
	Function UpdateRefTable (cScope, cPattern, cProjectOrDir)
		Local nSelect
		Local cSafety
		Local cSearchOptions
		Local cRefTable
		Local i

		m.nSelect = Select()

		If Vartype (cRefTable) # 'C' Or Empty (cRefTable)
			cRefTable = This.RefTable
		Endif

		If Empty (cRefTable)
			Return .F.
		Endif

		If Used ('FoxRefCursor')
			Use In FoxRefCursor
		Endif

		If Not This.OpenRefTable (cRefTable)
			Return .F.
		Endif


		This.tTimeStamp = Datetime()


		* Since we're only doing definitions and there
		* is no search in progress, then don't create
		* a Search Set record in the FoxRef table
		If This.lDefinitionsOnly
			Return .T.
		Endif

		* build a string representing the search options that
		* we can store to the FoxRef cursor
		cSearchOptions = IIf (This.Comments == COMMENTS_EXCLUDE, 'X', '') +			;
			IIf (This.Comments == COMMENTS_ONLY, 'C', '') +							;
			IIf (This.MatchCase, 'M', '') +											;
			IIf (This.WholeWordsOnly, 'W', '') +									;
			IIf (This.ProjectHomeDir, 'H', '') +									;
			IIf (This.FormProperties, 'P', '') +									;
			IIf (This.SubFolders, 'S', '') +										;
			IIf (This.Wildcards, 'Z', '') +											;
			';' + Alltrim (This.FileTypes)


		* if we've already searched for this same exact symbol
		* with the same exact criteria in the same exact project/folder,
		* then simply update what we have
		If This.lRefreshMode
			Select FoxRefCursor
			Locate For SetID == This.cSetID And RefType == REFTYPE_SEARCH And Not Inactive
			This.lRefreshMode = Found()
		Endif

		If Not This.lRefreshMode
			Select FoxRefCursor
			Locate For RefType == REFTYPE_SEARCH And ClassName == cProjectOrDir And Symbol == cPattern And Abstract == cSearchOptions And Not Inactive
			This.lRefreshMode = Found()
		Endif

		If This.lRefreshMode
			This.tTimeStamp = FoxRefCursor.Timestamp

			This.cSetID = FoxRefCursor.SetID
			Update  FoxRefCursor					;
				Set Inactive = .T.					;
				Where								;
				SetID == This.cSetID And		;
				(RefType == REFTYPE_RESULT Or RefType == REFTYPE_ERROR Or RefType == REFTYPE_NOMATCH)
		Else
			This.cSetID = Sys(2015)

			* add the record that specifies the search criteria, etc
			Insert Into FoxRefCursor (		;
				UniqueID,					;
				SetID,					;
				RefType,					;
				FindType,					;
				Symbol,					;
				ClassName,				;
				ProcName,					;
				ProcLineNo,				;
				Lineno,					;
				ColPos,					;
				MatchLen,					;
				Abstract,					;
				RecordID,					;
				UpdField,					;
				Checked,					;
				NoReplace,				;
				Timestamp,				;
				Inactive					;
				) Values (				;
				Sys(2015),				;
				This.cSetID,				;
				REFTYPE_SEARCH,			;
				'',						;
				cPattern,					;
				cProjectOrDir,			;
				'',						;
				0,						;
				0,						;
				0,						;
				0,						;
				cSearchOptions,			;
				'',						;
				'',						;
				.F.,						;
				.F.,						;
				Datetime(),				;
				.F.						;
				)
		Endif

		* update each of the search engines with the new SetID
		If Vartype (This.oWindowEngine) == 'O'
			This.oWindowEngine.SetID = This.cSetID
		Endif
		For m.i = 1 To This.oEngineCollection.Count
			If Vartype (This.oEngineCollection.Item (m.i)) == 'O'
				With This.oEngineCollection.Item (m.i)
					.SetID = This.cSetID
				Endwith
			Endif
		Endfor

		Select  Distinct SetID, FileID, Timestamp												;
			From FoxRefCursor																	;
			Where (RefType == REFTYPE_RESULT Or RefType == REFTYPE_NOMATCH) And Not Inactive	;
			Into Cursor FoxRefSearchedCursor


		Select (m.nSelect)

		Return .T.
	Endfunc


	* -- Search a Folder
	Function FolderSearch (cPattern, cFileDir)
		Local i, j
		Local cFileDir
		Local nFileTypesCnt
		Local cFileTypes
		Local lAutoYield
		Local lSuccess
		Local Array aFileList[1]
		Local Array aFileTypes[1]

		If Vartype (cPattern) # 'C'
			cPattern = This.Pattern
		Endif

		If Vartype (cFileDir) # 'C' Or Empty (cFileDir)
			cFileDir = Addbs (This.FileDirectory)
		Else
			cFileDir = Addbs (cFileDir)
		Endif

		If Empty (cFileDir) Or Not Directory (cFileDir)
			Return .F.
		Endif

		If Not This.SearchInit()
			Return .F.
		Endif

		If Not This.UpdateRefTable (SCOPE_FOLDER, cPattern, cFileDir)
			Return .F.
		Endif

		cFileTypes = Chrtran (This.FileTypes, ';', ',')
		nFileTypesCnt = Alines (aFileTypes, Alltrim (cFileTypes), .T., ',')

		lAutoYield = _vfp.AutoYield
		_vfp.AutoYield = .T.


		lSuccess = This.ProcessFolder (cFileDir, cPattern, @aFileTypes, nFileTypesCnt)

		This.CloseProgress()

		_vfp.AutoYield = lAutoYield

		This.UpdateLookForMRU (cPattern)
		This.UpdateFolderMRU (cFileDir)
		This.UpdateFileTypesMRU (cFileTypes)

		Return lSuccess
	Endfunc

	* used in conjuction with FolderSearch() for
	* when we're searching subfolders
	Function ProcessFolder (cFileDir, cPattern, aFileTypes, nFileTypesCnt)
		Local nFolderCnt
		Local cFilename
		Local i, j
		Local nFileCnt
		Local nProgress
		Local lSuccess
		Local Array aFileList[1]
		Local Array aFolderList[1]

		cFileDir = Addbs (cFileDir)

		lSuccess = .T.

		If This.ShowProgress
			* determine how many files there are to process
			nFileCnt = 0
			For i = 1 To nFileTypesCnt
				Try
					nFileCnt = nFileCnt + Adir (aFileList, cFileDir + aFileTypes[i], '', 1)
				Catch
					* ADIR failed, so just ignore
				Endtry
			Endfor

			* Set the description on the Progress form
			This.ProgressInit (IIf (This.lRefreshMode, PROGRESS_REFRESHING_LOC, IIf (This.lDefinitionsOnly, PROGRESS_DEFINITIONS_LOC, PROGRESS_SEARCHING_LOC)) + ' ' + Displaypath (cFileDir, 40) + IIf ( Not This.lRefreshMode Or This.lDefinitionsOnly, '',  ' [' + cPattern + ']'), nFileCnt)
		Endif

		nProgress = 0
		For i = 1 To nFileTypesCnt
			If This.lCancel
				Exit
			Endif

			Try
				nFileCnt = Adir (aFileList, cFileDir + aFileTypes[i], '', 1)
			Catch
				nFileCnt = 0
			Endtry
			For j = 1 To nFileCnt
				If This.lCancel
					Exit
				Endif

				cFilename = aFileList[j, 1]
				nProgress = nProgress + 1

				This.UpdateProgress (nProgress, cFilename, .T.)
				If Not This.FileSearch (cFileDir + cFilename, cPattern)
					Exit
				Endif
			Endfor
		Endfor

		* Process any sub-directories
		If Not This.lCancel
			If This.SubFolders
				Try
					nFolderCnt = Adir (aFolderList, cFileDir + '*.*', 'D', 1)
				Catch
					nFolderCnt = 0
				Endtry
				For i = 1 To nFolderCnt
					If Not aFolderList[i, 1] == '.' And Not aFolderList[i, 1] == '..' And 'D' $ aFolderList[i, 5] And Directory (cFileDir + aFolderList[i, 1])
						This.ProcessFolder (cFileDir + aFolderList[i, 1], cPattern, @aFileTypes, nFileTypesCnt)
					Endif
					If This.lCancel
						Exit
					Endif
				Endfor
			Endif
		Endif

		Return lSuccess
	Endfunc


	* -- Search files in a Project
	* -- Pass an empty cPattern to only collect definitions
	Function ProjectSearch (cPattern, cProjectFile)
		Local nFileIndex
		Local nProjectIndex
		Local oProjectRef
		Local oFileRef
		Local cFileTypes
		Local nFileTypesCnt
		Local lAutoYield
		Local lSuccess
		Local nFileCnt
		Local lSuccess
		Local i
		Local cHomeDir
		Local oDefFileTypes
		Local cFilePath
		Local oMatchFileCollection
		Local Array aFileTypes[1]
		Local Array aFileList[1]

		If Vartype (cPattern) # 'C'
			cPattern = This.Pattern
		Endif

		If Vartype (cProjectFile) # 'C' Or Empty (cProjectFile)
			cProjectFile = This.ProjectFile
		Endif

		If Not This.SearchInit()
			Return .F.
		Endif

		If Not This.UpdateRefTable (SCOPE_PROJECT, cPattern, This.ProjectFile)
			Return .F.
		Endif


		lSuccess = .T.

		cFileTypes = This.FileTypes
		cFileTypes = Chrtran (This.FileTypes, ';', ',')

		oMatchFileCollection = Newobject ('CFoxRefCollection', 'FoxRefCollection.prg')
		*** JRN 2010-03-14 : allow spaces in file names
		oMatchFileCollection.AddList (cFileTypes, ',')
		*	oMatchFileCollection.AddList(cFileTypes, ' ')



		oDefFileTypes = Newobject ('CFoxRefCollection', 'FoxRefCollection.prg')
		oDefFileTypes.AddList (Upper (FILETYPES_DEFINITIONS), ' ')

		lAutoYield = _vfp.AutoYield
		_vfp.AutoYield = .T.

		For Each oProjectRef In Application.Projects
			If Upper (oProjectRef.Name) == This.ProjectFile
				cHomeDir = Addbs (Upper (oProjectRef.HomeDir))

				* determine how many files to process so we can update the Progress form
				nFileCnt = 0
				If This.ShowProgress
					For Each oFileRef In oProjectRef.Files
						If ( Not This.ProjectHomeDir Or Addbs (Upper (Justpath (oFileRef.Name))) = cHomeDir) And ;
								((This.IncludeDefTable And oDefFileTypes.GetIndex ('*.' + Upper (Justext (oFileRef.Name))) > 0) Or This.ProjectMatch (cFileTypes, oMatchFileCollection, oFileRef.Name))
							nFileCnt = nFileCnt + 1
						Endif
					Endfor
					This.ProgressInit (IIf (This.lRefreshMode, PROGRESS_REFRESHING_LOC, IIf (This.lDefinitionsOnly, PROGRESS_DEFINITIONS_LOC, PROGRESS_SEARCHING_LOC)) + ' ' + Justfname (This.ProjectFile) + IIf ( Not This.lRefreshMode Or This.lDefinitionsOnly, '',  ' [' + cPattern + ']'), nFileCnt)
				Endif


				* now process each file in the project that matches our filetypes
				i = 0
				For Each oFileRef In oProjectRef.Files
					If ( Not This.ProjectHomeDir Or Addbs (Upper (Justpath (oFileRef.Name))) = cHomeDir)
						If This.ProjectMatch (cFileTypes, oMatchFileCollection, oFileRef.Name)
							i = i + 1
							This.UpdateProgress (i, oFileRef.Name, .T.)
							If Not This.FileSearch (oFileRef.Name, cPattern)
								Exit
							Endif
						Else
							If This.IncludeDefTable And oDefFileTypes.GetIndex ('*.' + Upper (Justext (oFileRef.Name))) > 0
								* don't search the file because it doesn't match the filetypes
								* specified -- however, we still want to collection definitions
								* on this filetype
								i = i + 1
								This.UpdateProgress (i, oFileRef.Name, .T.)
								If Not This.FileSearch (oFileRef.Name, '')
									Exit
								Endif
							Endif
						Endif
					Endif
					If This.lCancel
						lSuccess = .F.
						Exit
					Endif
				Endfor

				Exit
			Endif
		Endfor

		oFileTypesCollection = .Null.

		This.CloseProgress()

		_vfp.AutoYield = lAutoYield

		This.UpdateLookForMRU (cPattern)
		This.UpdateFileTypesMRU (cFileTypes)

		Return lSuccess
	Endfunc


	* Return Search Engine to use based upon filetype
	Function GetEngine (m.cFilename)
		Local nEngineIndex
		Local i

		If This.oEngineCollection.Count == 0
			Return .Null.
		Endif

		m.cFilename = Upper (m.cFilename)

		* determine which search engine to use based upon the filetype
		m.nEngineIndex = 1
		For m.i = 2 To This.oEngineCollection.Count
			If Vartype (This.oEngineCollection.Item (m.i)) == 'O'
				If This.WildcardMatch (This.oEngineCollection.GetKey (m.i), m.cFilename)
					m.nEngineIndex = m.i
					Exit
				Endif
			Endif
		Endfor

		* if we're still using the default search engine,
		* then make sure this file type isn't set to
		* be excluded
		If m.nEngineIndex == 1
			For m.i = 2 To This.oEngineCollection.Count
				If Vartype (This.oEngineCollection.Item (m.i)) # 'O'
					If This.WildcardMatch (This.oEngineCollection.GetKey (m.i), m.cFilename)
						m.nEngineIndex = m.i
						Exit
					Endif
				Endif
			Endfor
		Endif

		Return This.oEngineCollection.Item (m.nEngineIndex)
	Endfunc


	* Search a file
	* -- Pass an empty cPattern to only collect definitions
	Function FileSearch (cFilename, cPattern)
		Local nSelect
		Local cFileFind
		Local cFolderFind
		Local nSelect
		Local lDefinitions
		Local lSearch
		Local cFileID
		Local cFileAction
		Local oEngine
		Local lSuccess
		Local nIndex
		Local cLowerFilename
		Local Array aFileList[1]

		If This.lCancel
			Return .F.
		Endif

		* make sure we haven't searched for this file already
		m.cLowerFilename = Lower (m.cFilename)
		For m.nIndex = 1 To This.oProcessedCollection.Count
			If This.oProcessedCollection.Item (m.nIndex) == m.cLowerFilename
				Return .T.
			Endif
		Endfor

		This.oProcessedCollection.Add (Lower (m.cFilename))

		If Vartype (m.cPattern) # 'C'
			m.cPattern = This.Pattern
		Endif

		m.nSelect = Select()

		m.lSearch      = Not Empty (m.cPattern)
		m.lDefinitions = Not m.lSearch Or This.IncludeDefTable && false to search only

		m.oEngine = This.GetEngine (Justfname (m.cFilename))

		m.lSuccess = .T.

		* if we don't have an engine object defined, then assume
		* we're supposed to ignore this filetype
		If Vartype (m.oEngine) == 'O'
			With m.oEngine
				.Filename   = m.cFilename
				.SetTimeStamp()

				m.cFileFind   = Padr (Lower (Justfname (m.cFilename)), 100)
				* m.cFolderFind = PADR(LOWER(JUSTPATH(m.cFilename)), 240)
				m.cFolderFind = Lower (Justpath (m.cFilename))

				Select FileCursor
				Locate For Filename == m.cFileFind And Folder == m.cFolderFind
				If Found()
					* This file has been previously processed
					* However, if the TimeStamp is different or we've
					* never processed definitions, then we need to reprocess
					m.cFileID  = FileCursor.UniqueID
					If m.lDefinitions
						m.lDefinitions = (FileCursor.Timestamp # .FileTimeStamp) Or FileCursor.FileAction # FILEACTION_DEFINITIONS Or m.oEngine.Norefresh
					Endif

					Replace											;
						Timestamp  With	 .FileTimeStamp			;
						In FileCursor

					If m.lDefinitions
						Update FoxDefCursor Set Inactive = .T. Where FileID == m.cFileID
					Endif
				Else
					* file has never been processed, so add a file record to
					* the RefFile table
					m.cFileID  = Sys(2015)

					Insert Into FileCursor (			;
						UniqueID,						;
						Filename,						;
						Folder,						;
						Timestamp,					;
						FileAction					;
						) Values (					;
						M.cFileID,					;
						M.cFileFind,					;
						M.cFolderFind,				;
						.FileTimeStamp,				;
						FILEACTION_NODEFINITIONS		;
						)
				Endif

				* determine whether or not we need to process
				* this file at all.  Two conditions that might
				* force us to process:
				*  1) We need to collect definitions
				*  2) We need to search file for references (only if cPattern is not empty)
				If m.lSearch
					Select FoxRefSearchedCursor
					Locate For SetID == This.cSetID And FileID == m.cFileID
					If Found()
						m.lSearch = FoxRefSearchedCursor.Timestamp # FileCursor.Timestamp
						If Not m.lSearch
							Update  FoxRefCursor												;
								Set Inactive = .F.												;
								Where															;
								SetID == This.cSetID And									;
								FileID == m.cFileID And										;
								(RefType == REFTYPE_RESULT Or RefType == REFTYPE_NOMATCH) And ;
								Inactive
						Endif
					Else
						m.lSearch = .T.
					Endif
				Endif

				If m.lSearch Or m.lDefinitions
					.FileID = m.cFileID

					m.cFileAction = .SearchFor (m.cPattern, m.lSearch, m.lDefinitions, This)

					m.lSuccess = Not (m.cFileAction == FILEACTION_STOP)

					If m.lSuccess
						If Not Empty (m.cFileAction)
							Replace FileAction With m.cFileAction In FileCursor
						Endif
					Else
						Messagebox (ERROR_SEARCHENGINE_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
					Endif
				Endif
			Endwith

			If m.lSearch
				Update  FoxRefCursor					;
					Set RefType = REFTYPE_INACTIVE		;
					Where								;
					SetID == This.cSetID And		;
					FileID == m.cFileID And			;
					Inactive
			Endif
		Endif

		Select (m.nSelect)

		Return m.lSuccess
	Endfunc



	* refresh results for all Sets in the Ref table or a single set
	Function RefreshResults (cSetID)
		Local nSelect
		Local i
		Local nCnt
		Local lSuccess
		Local Array aRefList[1]

		m.nSelect = Select()

		m.lSuccess = .T.

		If Vartype (m.cSetID) == 'C' And Not Empty (m.cSetID)
			This.RefreshResultSet (m.cSetID)
		Else
			If File (Forceext (This.RefTable, 'dbf'))
				m.nSelect = Select()

				m.lSuccess = This.OpenRefTable (This.RefTable)
				If m.lSuccess
					Select  SetID												;
						From FoxRefCursor										;
						Where RefType == REFTYPE_SEARCH And Not Inactive		;
						Into Array aRefList
					m.nCnt = _Tally

					For m.i = 1 To m.nCnt
						This.RefreshResultSet (aRefList[i])
						If This.lCancel
							Exit
						Endif
					Endfor
				Endif


				Select (m.nSelect)
			Endif
			This.cSetID = ''
		Endif

		Return m.lSuccess
	Endfunc

	* refresh an existing search set
	Function RefreshResultSet (cSetID)
		Local nSelect
		Local lSuccess
		Local cScope
		Local cFolder
		Local cProject
		Local cPattern
		Local cSearchOptions

		m.lSuccess = .F.

		If File (Forceext (This.RefTable, 'dbf'))
			m.nSelect = Select()

			If Not This.OpenRefTable (This.RefTable)
				Return .F.
			Endif

			This.cSetID = m.cSetID

			Select FoxRefCursor
			Locate For RefType == REFTYPE_SEARCH And SetID == m.cSetID And Not Inactive
			lSuccess = Found()
			If lSuccess
				cSearchOptions = Leftc (FoxRefCursor.Abstract, At_c (';', FoxRefCursor.Abstract) - 1)
				If 'X' $ cSearchOptions
					This.Comments = COMMENTS_EXCLUDE
				Endif
				If 'C' $ cSearchOptions
					This.Comments = COMMENTS_ONLY
				Endif
				This.MatchCase      = 'M' $ cSearchOptions
				This.WholeWordsOnly = 'W' $ cSearchOptions
				This.FormProperties = 'P' $ cSearchOptions
				This.ProjectHomeDir = 'H' $ cSearchOptions
				This.SubFolders     = 'S' $ cSearchOptions
				This.Wildcards      = 'Z' $ cSearchOptions

				This.OverwritePrior = .F.

				This.FileTypes = Alltrim (Substrc (FoxRefCursor.Abstract, At_c (';', FoxRefCursor.Abstract) + 1))

				cFolder  = Rtrim (FoxRefCursor.ClassName)
				cProject = ''

				If Upper (Justext (cFolder)) == 'PJX'
					cScope = SCOPE_PROJECT
					cProject = cFolder
				Else
					cScope = SCOPE_FOLDER
				Endif

				cPattern = FoxRefCursor.Symbol
			Endif


			If lSuccess
				Do Case
					Case cScope == SCOPE_FOLDER
						lSuccess = This.FolderSearch (cPattern, cFolder)
					Case cScope == SCOPE_PROJECT
						lSuccess = This.ProjectSearch (cPattern, cProject)
					Otherwise
						lSuccess = .F.
				Endcase
			Endif

			Select (m.nSelect)
		Endif

		Return m.lSuccess
	Endfunc

	Function SetChecked (cUniqueID, lChecked)
		If Pcount() < 2
			lChecked = .T.
		Endif
		If Used ('FoxRefCursor') And Seek (cUniqueID, 'FoxRefCursor', 'UniqueID')
			If Not Isnull (FoxRefCursor.Checked)
				Replace Checked With lChecked In FoxRefCursor
			Endif
		Endif
	Endfunc


	* -- Show the Results form
	Function ShowResults()
		Local i

		* first see if there is an open Results window for this
		* project and display that if there is
		For i = 1 To _Screen.FormCount
			If Pemstatus (_Screen.Forms (i), 'oFoxRef', 5) And						;
					Upper (_Screen.Forms (i).Name) == 'FRMFOXREFRESULTS' And		;
					Type ('_SCREEN.Forms(i).oFoxRef') == 'O' And Not Isnull (_Screen.Forms (i).oFoxRef) And Upper (_Screen.Forms (i).oFoxRef.ProjectFile) == This.ProjectFile
				_Screen.Forms (i).SetRefTable (This.cSetID)
				_Screen.Forms (i).Show()
				Return
			Endif
		Endfor

		Do Form FoxRefResults With This
	Endfunc

	* goto a definition
	Function GotoDefinition (cUniqueID)
		Local cFilename
		Local nSelect
		Local oException

		If Vartype (m.cUniqueID) # 'C' Or Empty (m.cUniqueID)
			Return .F.
		Endif

		If Used ('FoxDefCursor') And Seek (m.cUniqueID, 'FoxDefCursor', 'UniqueID')
			If Inlist (FoxDefCursor.DefType, DEFTYPE_INCLUDEFILE, DEFTYPE_SETCLASSPROC)
				m.cFilename = FoxDefCursor.Symbol

				* if the file doesn't exist (we may not have a full path
				* becase it's from a Window reference), then see if we
				* can find a similar reference in the REFFILE that matches
				* our project
				If Not File (m.cFilename) And Used ('ProjectFilesCursor')
					m.nSelect = Select()
					Select ProjectFilesCursor
					m.cFilename = Padr (Lower (m.cFilename), Lenc (ProjectFilesCursor.Filename))
					Locate For Filename == m.cFilename
					If Found()
						m.cFilename = Addbs (ProjectFilesCursor.Folder) + Rtrim (ProjectFilesCursor.Filename)
					Endif
					Select (m.nSelect)
				Endif
				m.cFilename = Alltrim (m.cFilename)
				Try
					m.cFilename = Locfile (m.cFilename, Justext (m.cFilename))
					Editsource (m.cFilename)
				Catch To oException
					Messagebox (oException.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
				Endtry
			Else
				This.OpenSource (					;
					FoxDefCursor.FileID,			;
					FoxDefCursor.Lineno,			;
					FoxDefCursor.ProcLineNo,		;
					FoxDefCursor.ClassName,		;
					FoxDefCursor.ProcName,		;
					'',							;
					FoxDefCursor.Symbol,			;
					0								;
					)
			Endif
		Endif
	Endfunc

	* goto a specific reference
	Function GotoReference (cUniqueID)
		If Vartype (m.cUniqueID) # 'C' Or Empty (m.cUniqueID)
			Return .F.
		Endif

		If Used ('FoxRefCursor') And Seek (m.cUniqueID, 'FoxRefCursor', 'UniqueID')
			This.OpenSource (					;
				FoxRefCursor.FileID,			;
				FoxRefCursor.Lineno,			;
				FoxRefCursor.ProcLineNo,		;
				FoxRefCursor.ClassName,		;
				FoxRefCursor.ProcName,		;
				FoxRefCursor.UpdField,		;
				FoxRefCursor.Symbol			;
				)

		Endif

	Endfunc


	* open source file for editing
	Function OpenSource (cFileID, nLineNo, nProcLineNo, cClassName, cProcName, cUpdField, cSymbol, nColPos)
		Local cFileType
		Local nSelect
		Local cAlias
		Local lLibraryOpen
		Local nStartPos
		Local nRetCode
		Local oException
		Local nOpenError
		Local Array aEdEnv[25]

		*** JRN 2010-03-16 : Correct ProcName for the new items now stuffed into it
		Do Case
			Case cProcName = '<'
				cProcName = ''
			Case '<' $ cProcName
				cProcName = Left (cProcName, At ('<', cProcName) - 2)
		Endcase

		If Vartype (m.cSymbol) # 'C'
			m.cSymbol = ''
		Endif
		If Vartype (m.nColPos) # 'N'
			m.nColPos = 0
		Endif

		m.nSelect = Select()
		If Used ('FileCursor') And Seek (cFileID, 'FileCursor', 'UniqueID')
			If FileCursor.UniqueID = 'WINDOW' And This.WindowHandle >= 0
				* special case for jumping to a specific line # in the open window
				If Atcc ('FOXTOOLS.FLL', Set ('LIBRARY')) == 0
					m.lLibraryOpen = .F.
					m.cFoxtoolsLibrary = Sys(2004) + 'FOXTOOLS.FLL'
					If Not File (m.cFoxtoolsLibrary)
						Return .F.
					Endif
					Set Library To (m.cFoxtoolsLibrary) Additive
				Else
					m.lLibraryOpen = .T.
				Endif

				* Check environment of window
				* get the length of the window
				m.nRetCode = _edgetenv (This.WindowHandle, @aEdEnv)

				If m.nRetCode == 1 && AND (aEdEnv[EDENV_LENGTH] == 0)
					If aEdEnv[EDENV_LENGTH] > 0
						m.nStartPos = _EdGetLPos (This.WindowHandle, m.nLineNo - 1)
						_EdSelect (This.WindowHandle, m.nStartPos, m.nStartPos)
						_EdSToPos (This.WindowHandle, m.nStartPos, .F.)
					Endif
					_wselect (This.WindowHandle)

					This.HighlightSymbol (m.cSymbol, m.nLineNo, m.nColPos)
				Endif

				If Not m.lLibraryOpen And Atcc (m.cFoxtoolsLibrary, Set ('LIBRARY')) > 0
					Release Library (m.cFoxtoolsLibrary)
				Endif
			Else
				* open the file for editing
				m.cFilename  = Addbs (Rtrim (FileCursor.Folder)) + Rtrim (FileCursor.Filename)
				m.cFileType  = Upper (Justext (m.cFilename))

				* make sure the file exists
				If File (m.cFilename)

					*** JRN 2010-08-26 : Use ProjectHook?
					If This.CheckOutSCC ;
							and 0 # _vfp.Projects.Count;
							And Not (This.ProjectFile == PROJECT_GLOBAL Or Empty (This.ProjectFile))
						If Not Inlist (_vfp.ActiveProject.Files (m.cFilename).SCCStatus, 0, 2)
							If Not _vfp.ActiveProject.Files (m.cFilename).CheckOut()
								*		Return
							Endif
						Endif
					Endif

					*** JRN 2010-08-18 : If Project is open with ProjectHook, invoke QueryModifyFile
					If USEPROJECTHOOK
						If _vfp.Projects.Count # 0
							If 'O' = Vartype (_vfp.Projects(1).ProjectHook)
								Local loFile
								loFile = Createobject ('Empty')
								AddProperty (loFile, 'Name', m.cFilename)
								_vfp.Projects(1).ProjectHook.QueryModifyFile (loFile, m.cClassName)
							Endif
						Endif
					Endif

					If Vartype (m.cUpdField) # 'C'
						m.cUpdField = ''
					Else
						m.cUpdField  = Upper (m.cUpdField)
					Endif

					If m.cFileType == 'CDX'
						m.cFilename = Forceext (m.cFilename, 'DBF')
						m.cFileType = 'DBF'
					Endif

					*** JRN 2010-08-21 : Maintain mixed case, add to VFP's MRU list
					m.cFilename = This.DiskFileName(m.cFilename)
					This.AddMRUFile (m.cFilename, m.cClassName)

					m.nOpenError = 0
					Try
						Do Case
							Case m.cFileType == 'SCX'
								m.nOpenError = Editsource (m.cFilename, Max (m.nProcLineNo, 1), m.cClassName, m.cProcName)
								Do Case
									Case m.nOpenError == 925 && couldn't find methodname
										This.HighlightObject (m.cClassName, m.cProcName)

									Case m.nOpenError == 0
										This.HighlightSymbol (m.cSymbol, Max (m.nProcLineNo, 1), m.nColPos)
								Endcase

							Case m.cFileType == 'VCX'
								m.nOpenError = Editsource (m.cFilename, Max (m.nProcLineNo, 1), m.cClassName, m.cProcName)

								Do Case
									Case m.nOpenError == 925 && couldn't find methodname
										This.HighlightObject (m.cClassName, m.cProcName)

									Case m.nOpenError == 0
										This.HighlightSymbol (m.cSymbol, Max (m.nProcLineNo, 1), m.nColPos)
								Endcase

							Case m.cFileType == 'DBC'
								Do Case
									Case m.cUpdField == 'STOREDPROCEDURE'
										m.nOpenError = Editsource (m.cFilename, m.nLineNo)
										If m.nOpenError == 0
											This.HighlightSymbol (m.cSymbol, m.nLineNo, m.nColPos)
										Endif

									Otherwise
										Modify Database (m.cFilename)
								Endcase


							Case m.cFileType == 'DBF'
								m.cAlias = Juststem (m.cFilename)
								If Used (m.cAlias)
									Select (m.cAlias)
								Else
									Select 0
									Try
										Use (m.cFilename) Exclusive
									Catch
									Endtry

									If Not (Upper (Alias()) == Upper (m.cAlias))
										Try
											Use (m.cFilename) Shared Again
										Catch
										Endtry
									Endif
								Endif
								If Upper (Alias()) == Upper (m.cAlias)
									Modify Structure

									Use In (m.cAlias)
								Endif

							Case m.cFileType == 'PRG' Or			;
									M.cFileType == 'MPR' Or			;
									M.cFileType == 'QPR' Or			;
									M.cFileType == 'SPR' Or			;
									M.cFileType == 'H' Or			;
									M.cFileType == 'FRX' Or			;
									M.cFileType == 'LBX' Or			;
									M.cFileType == 'MNX' Or			;
									M.cFileType == 'TXT' Or			;
									M.cFileType == 'LOG' Or			;
									M.cFileType == 'ME'
								m.nOpenError = Editsource (m.cFilename, m.nLineNo)
								If m.nOpenError == 0
									This.HighlightSymbol (m.cSymbol, m.nLineNo, m.nColPos)
								Endif


							Otherwise
								This.ShellTo (m.cFilename)
						Endcase
					Catch To oException
						Messagebox (oException.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
					Endtry
				Else
					Messagebox (ERROR_FILENOTFOUND_LOC + ':' + Chr(10) + Chr(10) + m.cFilename, MB_ICONEXCLAMATION, APPNAME_LOC)
				Endif

				If m.nOpenError > 0 And Not Inlist (m.nOpenError, 901, 925)
					Messagebox (ERROR_OPENFILE_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
				Endif


			Endif
		Endif

		Select (m.nSelect)
	Endfunc

	* -- Highlight symbol given it's line # (assume it's open and active)
	Procedure HighlightSymbol (cSymbol, nLineNo, nColPos)
		Local nWindowHandle
		Local nRetCode
		Local cLine
		Local nStartPos
		Local nEndPos
		Local lLibraryOpen
		Local cFoxtoolsLibrary
		Local lSuccess
		Local Array aEdEnv[25]

		m.lSuccess = .F.

		If m.nLineNo <= 0
			Return
		Endif

		* special case for jumping to a specific line # in the open window
		If Atcc ('FOXTOOLS.FLL', Set ('LIBRARY')) == 0
			m.lLibraryOpen = .F.
			m.cFoxtoolsLibrary = Sys(2004) + 'FOXTOOLS.FLL'
			If File (m.cFoxtoolsLibrary)
				Try
					Set Library To (m.cFoxtoolsLibrary) Additive
					m.lSuccess = .T.
				Catch
				Endtry
			Endif
		Else
			m.lLibraryOpen = .T.
			m.lSuccess = .T.
		Endif

		If m.lSuccess
			m.nWindowHandle = _wontop()
			m.nRetCode = _edgetenv (m.nWindowHandle, @aEdEnv)

			If m.nRetCode == 1
				If aEdEnv[EDENV_LENGTH] > 0
					m.nStartPos = _EdGetLPos (m.nWindowHandle, m.nLineNo - 1)
					m.nEndPos   = _EdGetLPos (m.nWindowHandle, m.nLineNo)

					If Vartype (m.nColPos) # 'N' Or m.nColPos <= 0
						m.cLine = _EdGetStr (m.nWindowHandle, m.nStartPos, m.nEndPos - 1)
						m.nColPos = Atcc (m.cSymbol, m.cLine)
					Endif
					If m.nColPos > 0
						_EdSelect (m.nWindowHandle, m.nStartPos + m.nColPos - 1, m.nStartPos + m.nColPos - 1 + Lenc (m.cSymbol))
						_EdSToPos (m.nWindowHandle, m.nStartPos, .F.)
					Endif
				Endif
			Endif

			If Not m.lLibraryOpen And Atcc (m.cFoxtoolsLibrary, Set ('LIBRARY')) > 0
				Release Library (m.cFoxtoolsLibrary)
			Endif
		Endif
	Endproc


	* -- Set focus to the control where the reference was found
	* -- This is for when our reference is in a property of an object
	* -- and not in actual code
	Procedure HighlightObject (cParentName, cObjName)
		Local oObjRef
		Local oParentRef
		Local Array aObjRef[1]

		* see if object specified exists and highlight it
		If Aselobj (aObjRef, 1) == 1
			m.oParentRef = aObjRef[1]
			Do While Vartype (m.oParentRef) == 'O' And Not (m.oParentRef.Name == m.cParentName)
				m.oParentRef = m.oParentRef.Parent
			Enddo


			If Vartype (m.oParentRef) == 'O'
				m.oObjRef = m.oParentRef.&cObjName
				If Vartype (m.oObjRef) == 'O'
					If Pemstatus (m.oObjRef, 'SetFocus', 5)
						Try
							m.oObjRef.SetFocus()
						Catch
							* just in case, don't display an error because it's not critical
						Endtry
					Else
						* the object doesn't have a SetFocus method (such as Optiongroup), so at
						* least select the container (for example, select the Page the Optiongroup is on)
						Try
							m.oObjRef.Parent.SetFocus()
						Catch
							* just in case, don't display an error because it's not critical
						Endtry
					Endif
				Endif
			Endif
		Endif
	Endproc



	* given a filename, return the FileID from RefFile
	Function GetFileID (cFilename)
		Local cFName
		Local cFolder
		Local nSelect
		Local cFileID
		Local Array aFileList[1]

		m.nSelect = Select()

		m.cFilename = Lower (m.cFilename)
		cFName  = Padr (Lower (Justfname (m.cFilename)), 100)
		* cFolder = PADR(LOWER(JUSTPATH(m.cFilename)), 240)
		cFolder = Lower (Justpath (m.cFilename))

		Select  UniqueID										;
			From (This.FileTable) FileTable						;
			Where												;
			FileTable.Filename == cFName And				;
			FileTable.Folder == cFolder And					;
			FileTable.FileAction # FILEACTION_INACTIVE		;
			Into Array aFileList
		If _Tally > 0
			m.cFileID = aFileList[1]
		Else
			m.cFileID = ''
		Endif

		Select (m.nSelect)

		Return m.cFileID
	Endfunc

	* find all matching definitions and go directly to it
	* or display ambiguous references
	Function GotoSymbol (cGotoSymbol, cFilename, cClassName, cProcName, nLineNo, lLocalOnly)
		Local nSelect
		Local cSymbol
		Local nFunctionCnt
		Local i
		Local nPos
		Local cHelpTopic
		Local cWhere
		Local cWhere2
		Local cFName
		Local cFolder
		Local lFindProperty
		Local lFindProcedure
		Local nIndex
		Local lSuccess
		Local cHelpSymbol
		Local Array aCommandList[1]
		Local Array aFunctionList[1]
		Local Array aBaseClassList[1]

		If Vartype (m.cGotoSymbol) # 'C' Or Empty (m.cGotoSymbol)
			Return .F.
		Endif

		* search only for locals within current window, and
		* return if there is no open window
		If m.lLocalOnly And This.WindowHandle < 0
			Return .F.
		Endif

		m.cSymbol = Alltrim (Upper (m.cGotoSymbol))

		m.lFindProperty = Leftc (m.cSymbol, 5) == 'THIS.' Or Leftc (m.cSymbol, 9) == 'THISFORM.'
		m.lFindProcedure = .F.

		* if this is a function (as determined by having a '(' in the symbol name),
		* then strip off everything after the paren
		m.nPos = Ratc ('(', m.cSymbol)
		If m.nPos > 1
			m.cSymbol = Leftc (m.cSymbol, m.nPos - 1)
			m.lFindProcedure = .T.
		Endif


		* remove the "M.", "THIS." or "THISFORM." from the symbol all the
		* way up to the last word after the '.'
		* For example:
		*   "THIS.Parent.cmdButton.AutoSize" = "AutoSize"
		m.nPos = Ratc ('.', m.cSymbol)
		If m.nPos > 0 And m.nPos < Lenc (m.cSymbol)
			m.cSymbol = Substrc (m.cSymbol, m.nPos + 1)
		Endif

		* remove ampersand character from beginning if it exists
		m.nPos = Ratc ('&', m.cSymbol)
		If m.nPos > 0 And m.nPos < Lenc (m.cSymbol)
			m.cSymbol = Substrc (m.cSymbol, m.nPos + 1)
		Endif


		If Empty (m.cSymbol)
			Return .F.
		Endif

		m.lSuccess = .F.

		If Vartype (m.cFilename) == 'C'
			cFName  = Padr (Lower (Justfname (m.cFilename)), 100)
			cFolder = Lower (Justpath (m.cFilename))
		Else
			cFName  = ''
			cFolder = ''
		Endif

		If Vartype (m.cProcName) # 'C'
			m.cProcName = ''
		Endif



		m.nSelect = Select()
		m.cWhere = [DefTable.Symbol == cSymbol AND !DefTable.Inactive]
		m.nCnt = 0

		If m.lFindProcedure
			* search only for procedure names
			m.cWhere = m.cWhere + [ AND (DefTable.DefType == "] + DEFTYPE_PROCEDURE + [")]
		Else
			* search only for properties and procedures
			If m.lFindProperty
				m.cWhere = m.cWhere + [ AND (DefTable.DefType == "] + DEFTYPE_PROPERTY + [" OR DefTable.DefType == "] + DEFTYPE_PROCEDURE + [")]
			Endif
		Endif
		m.cWhere2 = m.cWhere

		If This.WindowHandle >= 0
			* see if we're positioned on a #include, SET PROCEDURE TO line
			If Vartype (m.nLineNo) == 'N' And m.nLineNo > 0
				Select																			;
					DefTable.UniqueID,													;
					DefTable.FileID,														;
					DefTable.DefType,														;
					DefTable.ProcName,														;
					DefTable.ClassName,														;
					DefTable.ProcLineNo,													;
					DefTable.Lineno,														;
					DefTable.Abstract														;
					From (This.DefTable) DefTable												;
					Where																		;
					DefTable.FileID == 'WINDOW' And											;
					(DefTable.DefType == DEFTYPE_INCLUDEFILE Or DefTable.DefType == DEFTYPE_SETCLASSPROC) And ;
					DefTable.Lineno == m.nLineNo And										;
					Not DefTable.Inactive													;
					Into Cursor DefinitionCursor
				nCnt = _Tally
			Endif

			If nCnt == 0
				If m.lLocalOnly And Empty (m.cProcName) And Vartype (This.oWindowEngine) == 'O'
					m.cProcName = This.oWindowEngine.GetProcedure (m.nLineNo)

					cWhere = cWhere + [ AND DefTable.ProcName == m.cProcName]
				Endif

				* always search for the LOCAL or PRIVATE or PARAMETER first
				Select																			;
					DefTable.UniqueID,													;
					DefTable.FileID,														;
					DefTable.DefType,														;
					DefTable.ProcName,														;
					DefTable.ClassName,														;
					DefTable.ProcLineNo,													;
					DefTable.Lineno,														;
					DefTable.Abstract,														;
					''	  As  Folder,														;
					''	  As  Filename														;
					From (This.DefTable) DefTable INNER Join (This.FileTable) FileTable On DefTable.FileID == FileTable.UniqueID ;
					Where DefTable.FileID = 'WINDOW' And										;
					&cWhere And																;
					Inlist (DefTable.DefType, DEFTYPE_PARAMETER, DEFTYPE_LOCAL, DEFTYPE_PRIVATE, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE, DEFTYPE_PROPERTY) ;
					Into Cursor DefinitionCursor
				nCnt = _Tally
			Endif

			* now search in the current file
			If nCnt == 0
				Select																			;
					DefTable.UniqueID,													;
					DefTable.FileID,														;
					DefTable.DefType,														;
					DefTable.ProcName,														;
					DefTable.ClassName,														;
					DefTable.ProcLineNo,													;
					DefTable.Lineno,														;
					DefTable.Abstract,														;
					''	  As  Folder,														;
					''	  As  Filename														;
					From (This.DefTable) DefTable INNER Join (This.FileTable) FileTable On DefTable.FileID == FileTable.UniqueID ;
					Where DefTable.FileID = 'WINDOW' And										;
					&cWhere2 And															;
					Inlist (DefTable.DefType, DEFTYPE_PRIVATE, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE, DEFTYPE_PROPERTY) ;
					Into Cursor DefinitionCursor
				nCnt = _Tally
			Endif
		Else
			* always search for the LOCAL or PRIVATE or PARAMETER first
			Select																				;
				DefTable.UniqueID,														;
				DefTable.FileID,															;
				DefTable.DefType,															;
				DefTable.ProcName,															;
				DefTable.ClassName,															;
				DefTable.ProcLineNo,														;
				DefTable.Lineno,															;
				DefTable.Abstract,															;
				FileTable.Folder,															;
				FileTable.Filename															;
				From (This.DefTable) DefTable INNER Join (This.FileTable) FileTable On DefTable.FileID == FileTable.UniqueID ;
				Where FileTable.Filename == cFName And FileTable.Folder == cFolder And			;
				&cWhere And																	;
				Inlist (DefTable.DefType, DEFTYPE_PARAMETER, DEFTYPE_LOCAL, DEFTYPE_PRIVATE, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE, DEFTYPE_PROPERTY) ;
				Into Cursor DefinitionCursor
			nCnt = _Tally
		Endif


		If nCnt == 0 And Not m.lLocalOnly
			* now find all appropriate definitions
			If Used ('ProjectFilesCursor') And Reccount ('ProjectFilesCursor') > 0
				* if we have a list of files in this project, then we
				* only want to concern ourselves with symbols that
				* appear in those files
				Select																			;
					DefTable.UniqueID,													;
					DefTable.FileID,														;
					DefTable.DefType,														;
					DefTable.ProcName,														;
					DefTable.ClassName,														;
					DefTable.ProcLineNo,													;
					DefTable.Lineno,														;
					DefTable.Abstract,														;
					FileTable.Folder,														;
					FileTable.Filename														;
					From (This.DefTable) DefTable INNER Join (This.FileTable) FileTable On DefTable.FileID == FileTable.UniqueID ;
					INNER Join ProjectFilesCursor On FileTable.Filename == ProjectFilesCursor.Filename And FileTable.Folder == ProjectFilesCursor.Folder ;
					Where &cWhere And															;
					Inlist (DefTable.DefType, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE) ;
					Into Cursor DefinitionCursor
			Else
				Select																			;
					DefTable.UniqueID,													;
					DefTable.FileID,														;
					DefTable.DefType,														;
					DefTable.ProcName,														;
					DefTable.ClassName,														;
					DefTable.ProcLineNo,													;
					DefTable.Lineno,														;
					DefTable.Abstract,														;
					FileTable.Folder,														;
					FileTable.Filename														;
					From (This.DefTable) DefTable INNER Join (This.FileTable) FileTable On DefTable.FileID == FileTable.UniqueID ;
					Where &cWhere And															;
					Inlist (DefTable.DefType, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE) ;
					Into Cursor DefinitionCursor
			Endif
			nCnt = _Tally
		Endif

		* Moved higher so opens #include files faster
		*!*			IF nCnt == 0 AND VARTYPE(m.nLineNo) == 'N' AND m.nLineNo > 0 AND !m.lLocalOnly
		*!*				* see if we're positioned on a #include, SET PROCEDURE TO line
		*!*				SELECT																	;
		*!*				  DefTable.UniqueID,													;
		*!*				  DefTable.FileID,														;
		*!*				  DefTable.DefType,														;
		*!*				  DefTable.ProcName,													;
		*!*				  DefTable.ClassName,													;
		*!*				  DefTable.ProcLineNo,													;
		*!*				  DefTable.LineNo,														;
		*!*				  DefTable.Abstract														;
		*!*				FROM (THIS.DefTable) DefTable											;
		*!*				 WHERE																	;
		*!*				  DefTable.FileID == "WINDOW" AND										;
		*!*				  (DefTable.DefType == DEFTYPE_INCLUDEFILE OR DefTable.DefType == DEFTYPE_SETCLASSPROC) AND ;
		*!*				  DefTable.LineNo == m.nLineNo AND										;
		*!*				  !DefTable.Inactive													;
		*!*				 INTO CURSOR DefinitionCursor
		*!*				nCnt = _TALLY
		*!*			ENDIF

		Do Case
			Case nCnt == 0  && no matches found
				If Not m.lLocalOnly
					* if this is an internal VFP command/function, then
					* display the VFP Help on that symbol
					cHelpTopic = ''

					* look for base classes
					= Alanguage (aBaseClassList, 3)

					nIndex = Ascan (aBaseClassList, Getwordnum (cGotoSymbol, 1), -1, -1, 1, 7)
					If nIndex > 0
						cHelpTopic = aBaseClassList[nIndex]
						Help &cHelpTopic
					Else
						* create array of all internal VFP commands
						m.cHelpSymbol = Upper (Getwordnum (cGotoSymbol, 1))
						= Alanguage (aCommandList, 1)
						If Ascan (aCommandList, m.cHelpSymbol, -1, -1, 1, 14) > 0
							cHelpTopic = Upper (cGotoSymbol)
						Else
							* check if there's a matching function
							* don't use ASCAN because we have to check for abbreviated match
							nFunctionCnt = Alanguage (aFunctionList, 2)
							For i = 1 To nFunctionCnt
								If 'M' $ aFunctionList[i, 2]
									* 'M' in 2nd parameter means we need an exact match
									If aFunctionList[i, 1] == m.cHelpSymbol
										cHelpTopic = aFunctionList[i, 1] + '()'
									Endif
								Else
									* no 'M' in 2nd parameter returned by ALANGUAGE, so only need to match up to 4 characters
									If aFunctionList[i, 1] == m.cHelpSymbol Or (Lenc (m.cHelpSymbol) < Lenc (aFunctionList[i, 1]) And Lenc (m.cHelpSymbol) >= 4 And aFunctionList[i, 1] = m.cHelpSymbol)
										cHelpTopic = aFunctionList[i, 1] + '()'
									Endif
								Endif
							Endfor
						Endif
					Endif


					If Empty (cHelpTopic)
						Messagebox (NODEFINITION_LOC + Chr(10) + Chr(10) + cGotoSymbol, MB_ICONEXCLAMATION, GOTODEFINITION_LOC)
					Else
						cHelpTopic = Alltrim (cHelpTopic)
						Help &cHelpTopic
					Endif
				Endif

			Case nCnt == 1
				* only a single match, so go right to it
				This.GotoDefinition (DefinitionCursor.UniqueID)
				m.lSuccess = .T.

			Otherwise
				* more than one match found, so display a cursor of
				* the available matches.
				Do Form FoxRefGotoDef With This, cGotoSymbol
				m.lSuccess = .T.
		Endcase

		If Used ('DefinitionCursor')
			Use In DefinitionCursor
		Endif

		Select (m.nSelect)

		Return m.lSuccess
	Endfunc


	* Show a progress form while searching
	Function ProgressInit (cDescription, nMax)
		If This.ShowProgress
			If Vartype (This.oProgressForm) # 'O'
				This.UpdateProgress()
			Endif
			If Vartype (This.oProgressForm) == 'O'
				This.oProgressForm.SetMax (m.nMax)
				This.oProgressForm.SetDescription (m.cDescription)
			Endif
			DoEvents
		Endif
	Endfunc

	Function UpdateProgress (nValue, cMsg, lFilename)
		If This.ShowProgress
			If Vartype (This.oProgressForm) # 'O'
				This.lCancel = .F.
				This.oProgressForm = Newobject ('CProgressForm', 'FoxRef.vcx')
				This.oProgressForm.Show()
			Endif

			If m.lFilename And Not Empty (Justpath (m.cMsg))
				* truncate filenames so they fit
				m.cMsg = Displaypath (m.cMsg, 60)
			Endif

			If Not This.oProgressForm.SetProgress (m.nValue, m.cMsg)  && FALSE is returned if Cancel button is pressed
				If Messagebox (SEARCH_CANCEL_LOC, MB_ICONQUESTION + MB_YESNO, APPNAME_LOC) == IDYES
					This.lCancel = .T.
				Else
					This.oProgressForm.lCancel = .F.
				Endif
			Endif
			DoEvents
		Endif
	Endfunc

	Function CloseProgress()
		If Vartype (This.oProgressForm) == 'O'
			This.oProgressForm.Release()
		Endif
	Endfunc


	* Export reference table
	Function ExportReferences (cExportType, cFilename, cSetID, cFileID, cRefID, lSelectedOnly)
		Local nSelect
		Local cWhere
		Local lSuccess
		Local cXSLFilename
		Local cTempFilename
		Local cHTML
		Local xslDoc
		Local xmlDoc
		Local cSafety
		Local oException

		nSelect = Select()

		oException = .Null.

		If Vartype (cExportType) # 'C' Or Empty (cExportType)
			cExportType = EXPORTTYPE_DBF
		Endif
		If Vartype (cSetID) # 'C'
			cSetID = ''
		Endif
		If Vartype (cFileID) # 'C'
			cFileID = ''
		Endif
		If Vartype (cRefID) # 'C'
			cRefID = ''
		Endif
		If Vartype (lSelectedOnly) # 'L'
			lSelectedOnly = .F.
		Endif

		* all export types except "Clipboard" require a filename
		If cExportType # EXPORTTYPE_CLIPBOARD
			If Vartype (cFilename) # 'C'
				Return .F.
			Endif
			cFilename = Fullpath (cFilename)
			If Not Directory (Justpath (cFilename))
				Return .F.
			Endif
		Endif


		cWhere = 'RefTable.RefType == [' + REFTYPE_RESULT + '] AND !RefTable.Inactive'
		If Empty (cRefID)
			If Not Empty (cSetID)
				cWhere = cWhere + ' AND RefTable.SetID == [' + cSetID + ']'
			Endif
			If Not Empty (cFileID)
				cWhere = cWhere + ' AND RefTable.FileID == [' + cFileID + ']'
			Endif
		Else
			cWhere = cWhere + ' AND RefTable.RefID == [' + cRefID + ']'
		Endif

		If lSelectedOnly
			cWhere = cWhere + ' AND RefTable.Checked'
		Endif

		If Inlist (cExportType, EXPORTTYPE_TXT, EXPORTTYPE_XLS)
			* memo fields won't be copied so we need to convert
			* to long characters
			Select																				;
				Padr (RefTable.Symbol, 254)			  As  Symbol,						;
				FileTable.Folder,															;
				FileTable.Filename,															;
				Padr (RefTable.ClassName, 254)			  As  ClassName,					;
				Padr (RefTable.ProcName, 254)			  As  ProcName,						;
				RefTable.ProcLineNo,														;
				RefTable.Lineno,															;
				RefTable.ColPos,															;
				RefTable.MatchLen,															;
				Padr (StripTabs (RefTable.Abstract), 254) As  Abstract						;
				From (This.RefTable) RefTable INNER Join (This.FileTable) FileTable On RefTable.FileID == FileTable.UniqueID ;
				Where &cWhere																	;
				Into Cursor ExportCursor
		Else
			Select																				;
				RefTable.Symbol,															;
				FileTable.Folder,															;
				FileTable.Filename,															;
				RefTable.ClassName,															;
				RefTable.ProcName,															;
				RefTable.ProcLineNo,														;
				RefTable.Lineno,															;
				RefTable.ColPos,															;
				RefTable.Abstract															;
				From (This.RefTable) RefTable INNER Join (This.FileTable) FileTable On RefTable.FileID == FileTable.UniqueID ;
				Where &cWhere																	;
				Into Cursor ExportCursor
		Endif
		Select ExportCursor

		Do Case
			Case cExportType == EXPORTTYPE_DBF
				Try
					Copy To (cFilename)			;
						Fields					;
						Symbol,					;
						Folder,					;
						Filename,				;
						ClassName,				;
						ProcName,				;
						ProcLineNo,				;
						Lineno,					;
						ColPos,					;
						Abstract

				Catch To oException
				Endtry

			Case cExportType == EXPORTTYPE_TXT
				Try
					Copy To (cFilename)			;
						Fields					;
						Symbol,					;
						Folder,					;
						Filename,				;
						ClassName,				;
						ProcName,				;
						ProcLineNo,				;
						Lineno,					;
						ColPos,					;
						Abstract				;
						Delimited
				Catch To oException
				Endtry

			Case cExportType == EXPORTTYPE_XML
				If Empty (Justext (cFilename))
					cFilename = Forceext (cFilename, 'xml')
				Endif


				Try
					If This.XMLSchema
						Cursortoxml ('ExportCursor', cFilename, IIf (This.XMLFormat == XMLFORMAT_ELEMENTS, 1, 2), IIf (This.XMLFormat == XMLFORMAT_ELEMENTS, 520, 512), 0, '1')
					Else
						Cursortoxml ('ExportCursor', cFilename, IIf (This.XMLFormat == XMLFORMAT_ELEMENTS, 1, 2), IIf (This.XMLFormat == XMLFORMAT_ELEMENTS, 520, 512))
					Endif

				Catch To oException
				Endtry

			Case cExportType == EXPORTTYPE_HTML
				cTempFilename = Addbs (Sys(2023)) + Sys(2015) + '.xml'
				If Empty (Justext (cFilename))
					cFilename = Forceext (cFilename, 'html')
				Endif

				cHTML = ''
				lSuccess = .T.
				Try
					Cursortoxml ('ExportCursor', cTempFilename, 1, 522, 0, '1')
					cXSLFilename = Fullpath (This.XSLTemplate)
					If Not File (cXSLFilename)
						cXSLFilename = This.FoxRefDirectory + This.XSLTemplate
						If Not File (cXSLFilename)
							* still not found, so copy internal version
							* to home directory and use that
							cXSLFilename = This.FoxRefDirectory + This.XSLTemplate
							Strtofile (Filetostr ('foxreftemplate.xsl'), cXSLFilename)
						Endif
					Endif

					xmlDoc = Createobject ('MSXML2.DOMDocument')
					xslDoc = Createobject ('MSXML2.DOMDocument')

					xmlDoc.Load (cTempFilename)
					If xmlDoc.parseerror.errorcode == 0
						xslDoc.Load (cXSLFilename)
						cHTML = xmlDoc.TransformNode (xslDoc)
						Strtofile (cHTML, cFilename)
					Endif

				Catch To oException
				Endtry


				cSafety = Set ('SAFETY')
				Set Safety Off
				Erase (cTempFilename)
				Set Safety &cSafety

				xslDoc = Null
				xmlDoc = Null

			Case cExportType == EXPORTTYPE_XLS
				Try
					Copy To (cFilename)			;
						Fields					;
						Symbol,					;
						Folder,					;
						Filename,				;
						ClassName,				;
						ProcName,				;
						ProcLineNo,				;
						Lineno,					;
						ColPos,					;
						Abstract				;
						Type Xl5

				Catch To oException
				Endtry

			Case cExportType == EXPORTTYPE_CLIPBOARD
				* NOTE: we can't use DataToClip because it only works for DataSession 1
				* _VFP.DataToClip()

				_Cliptext = ''
				Scan All
					_Cliptext =														;
						_Cliptext + IIf (Recno() > 1, Chr(13) + Chr(10), '') +		;
						Symbol + ',' +												;
						Rtrim (Folder) + ',' +										;
						Rtrim (Filename) + ',' +									;
						ClassName + ',' +											;
						ProcName + ',' +											;
						Transform (ProcLineNo) + ',' +								;
						Transform (Lineno) + ',' +									;
						Transform (ColPos) + ',' +									;
						StripTabs (Abstract)
				Endscan

		Endcase

		If Used ('ExportCursor')
			Use In ExportCursor
		Endif

		If Vartype (oException) == 'O'
			Messagebox (oException.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
		Endif


		Select (nSelect)

		Return Isnull (oException)
	Endfunc


	* Print a report of found references
	Function PrintReferences (cReportName, lPreview, cSetID, lSelectedOnly, cSortColumns)
		Local oRpt As .RptClassName Of  .RptClassLibrary
		*:Global cExecute, cExt, cWhere, nRecordCnt, nSelect, oException, oReportAddIn


		nSelect = Select()
		Select 0

		nRecordCnt = -1


		If Vartype (lPreview) # 'L'
			lPreview = .F.
		Endif
		If Vartype (lSelectedOnly) # 'L'
			lSelectedOnly = .F.
		Endif
		If Vartype (cSetID) # 'C'
			cSetID = ''
		Endif

		If Vartype (cSortColumns) # 'C' Or Empty (cSortColumns)
			cSortColumns = 'FileTable.Folder, FileTable.Filename, RefTable.LineNo'
		Endif

		Try
			oReportAddIn = This.oReportCollection.Item (m.cReportName)
		Catch
			oReportAddIn = .Null.
		Endtry

		If Not Isnull (oReportAddIn)
			cWhere = 'RefTable.RefType == [' + REFTYPE_RESULT + '] AND !RefTable.Inactive'
			If Not Empty (cSetID)
				cWhere = cWhere + ' AND RefTable.SetID == [' + cSetID + ']'
			Endif
			If lSelectedOnly
				cWhere = cWhere + ' AND RefTable.Checked'
			Endif

			Select  RefTable.RefID,																;
				RefTable.SetID,																;
				RefTable.Symbol,															;
				RefTable.ClassName,															;
				RefTable.ProcName,															;
				RefTable.ProcLineNo,														;
				RefTable.Lineno,															;
				RefTable.ColPos,															;
				RefTable.MatchLen,															;
				RefTable.Abstract,															;
				RefTable.Timestamp,															;
				FileTable.Folder,															;
				FileTable.Filename,															;
				Padr (FileTable.Folder, 240)			  As  SortFolder,					;
				Lower (Padr (Justext (FileTable.Filename), 3)) As SortFileType,				;
				Padr (RefTable.ClassName, 100)			  As  ClassNameSort,				;
				Padr (RefTable.ProcName, 100)			  As  ProcNameSort					;
				;
				, Padr (Justext (FileTable.Filename), 3)	  As  SortExtension					;
				, Padr (Lower (StripTabs (RefTable.Abstract)), 100) As AbstractSort			;
				, Padr (IIf ('.' $ RefTable.ProcName, Justext (RefTable.ProcName), RefTable.ProcName), 100) As MethodNameSort ;
				, IIf (Empty (RefTable.oTimeStamp), RefTable.Timestamp, DeCodeTimeStamp (RefTable.oTimeStamp)) As TimeStampSort ;
				;
				From (This.RefTable) RefTable													;
				INNER Join (This.FileTable) FileTable On RefTable.FileID == FileTable.UniqueID ;
				Where &cWhere																	;
				Order By &cSortColumns															;
				Into Cursor RptCursor
			m.nRecordCnt = _Tally

			If m.nRecordCnt > 0
				With oReportAddIn
					Try
						Do Case
							Case Not Empty (.RptFilename)
								m.cExt = Upper (Justext (.RptFilename))
								If Empty (m.cExt)
									m.cExt = 'FRX'
								Endif

								Do Case
									Case m.cExt == 'PRG' Or m.cExt == 'FXP'
										Do (.RptFilename) With m.lPreview, This

									Case m.cExt == 'SCX'
										Do Form (.RptFilename) With m.lPreview, This

									Case m.cExt == 'FRX'
										If m.lPreview
											Report Form (.RptFilename) Preview
										Else
											Report Form (.RptFilename) Noconsole To Printer Prompt
										Endif
								Endcase

							Case Not Empty (.RptClassName) And Not Empty (.RptMethod)
								cExecute = [oRpt.] + .RptMethod + [(] + Transform (m.lPreview) + [, THIS"]
								oRpt = Newobject (.RptClassName, .RptClassLibrary)
								&cExecute
							Otherwise
								nRecordCnt = -1
						Endcase
					Catch To oException
						Messagebox (oException.Message, MB_ICONSTOP, APPNAME_LOC)
					Endtry
				Endwith
			Endif

			If Used ('RptCursor')
				Use In RptCursor
			Endif

		Endif

		Select (nSelect)

		Return nRecordCnt
	Endfunc

	* Clear a specified Result Set or all results
	* from the Reference table
	Function ClearResults (cSetID, cFileID)
		Local nSelect
		Local cAlias

		m.nSelect = Select()

		Do Case
			Case Vartype (m.cFileID) == 'C' And Not Empty (m.cFileID)
				* Clear specified file
				Delete From (This.RefTable) Where SetID == m.cSetID And FileID == m.cFileID

			Case Vartype (m.cSetID) == 'C' And Not Empty (m.cSetID)
				* Clear specified result set
				Delete From (This.RefTable) Where SetID == m.cSetID

			Otherwise
				* Clear all results
				Delete  From (This.RefTable)				;
					Where									;
					RefType == REFTYPE_RESULT Or		;
					RefType == REFTYPE_SEARCH Or		;
					RefType == REFTYPE_NOMATCH Or		;
					RefType == REFTYPE_ERROR Or			;
					RefType == REFTYPE_LOG
		Endcase


		* if we can get this table open exclusive, then
		* we should pack it
		If This.OpenRefTable (This.RefTable, .T.)
			Select FoxRefCursor
			Try
				Pack In FoxRefCursor
			Catch
				* no big deal that we can't pack the table -- just ignore the error
			Endtry
		Endif


		* open it again shared
		This.OpenRefTable (This.RefTable)


		Select (m.nSelect)

		Return .T.
	Endfunc



	* Save preferences to FoxPro Resource file
	Function SavePrefs()
		Local lSuccess
		Local FOXREF_FILETYPES_MRU[10], FOXREF_FOLDER_MRU[10], FOXREF_LOOKFOR_MRU[10], FOXREF_OPTIONS[1]
		Local aFileList[1], cData, i, nCnt, nMemoWidth, nSelect, oOptionCollection
		*:Global FOXREF_REPLACE_MRU[1]

		If Not (Set ('RESOURCE') == 'ON')
			Return .F.
		Endif

		= Acopy (This.aLookForMRU, FOXREF_LOOKFOR_MRU)
		= Acopy (This.aReplaceMRU, FOXREF_REPLACE_MRU)
		= Acopy (This.aFolderMRU, FOXREF_FOLDER_MRU)
		= Acopy (This.aFileTypesMRU, FOXREF_FILETYPES_MRU)

		oOptionCollection = Createobject ('Collection')
		* Add any properties you want to save to
		* the resource file to this collection
		With oOptionCollection
			.Add (This.Comments, 'Comments')
			.Add (This.MatchCase, 'MatchCase')
			.Add (This.WholeWordsOnly, 'WholeWordsOnly')
			.Add (This.Wildcards, 'Wildcards')
			.Add (This.ProjectHomeDir, 'ProjectHomeDir')
			.Add (This.SubFolders, 'SubFolders')
			.Add (This.OverwritePrior, 'OverwritePrior')
			.Add (This.FileTypes, 'FileTypes')
			.Add (This.IncludeDefTable, 'IncludeDefTable')
			.Add (This.CodeOnly, 'CodeOnly')
			.Add (This.CheckOutSCC, 'CheckOutSCC')
			.Add (This.TreeViewFolderNames, 'TreeViewFolderNames')
			.Add (This.FormProperties, 'FormProperties')
			.Add (This.AutoProjectHomeDir, 'AutoProjectHomeDir')
			.Add (This.ConfirmReplace, 'ConfirmReplace')
			.Add (This.BackupOnReplace, 'BackupOnReplace')
			.Add (This.DisplayReplaceLog, 'DisplayReplaceLog')
			.Add (This.PreserveCase, 'PreserveCase')
			.Add (This.BackupStyle, 'BackupStyle')
			.Add (This.ShowRefsPerLine, 'ShowRefsPerLine')
			.Add (This.ShowFileTypeHistory, 'ShowFileTypeHistory')
			.Add (This.ShowDistinctMethodLine, 'ShowDistinctMethodLine')
			.Add (This.SortMostRecentFirst, 'SortMostRecentFirst')
			.Add (This.FontString, 'FontString')
			.Add (This.FoxRefDirectory, 'FoxRefDirectory')
		Endwith

		Dimension FOXREF_OPTIONS[MAX (oOptionCollection.Count, 1), 2]
		For i = 1 To oOptionCollection.Count
			FOXREF_OPTIONS[m.i, 1] = oOptionCollection.GetKey (m.i)
			FOXREF_OPTIONS[m.i, 2] = oOptionCollection.Item (m.i)
		Endfor


		nSelect = Select()

		lSuccess = .F.

		* make sure Resource file exists and is not read-only
		Try
			nCnt = Adir (aFileList, RESOURCE_FILE)
		Catch
			nCnt = 0
		Endtry

		If nCnt > 0 And Atcc ('R', aFileList[1, 5]) == 0
			If Not Used ('FoxResource')
				Use (RESOURCE_FILE) In 0 Shared Again Alias FoxResource
			Endif
			If Used ('FoxResource') And Not Isreadonly ('FoxResource')
				nMemoWidth = Set ('MEMOWIDTH')
				Set Memowidth To 255

				Select FoxResource
				Locate For Upper (Alltrim (Type)) == 'PREFW' And Upper (Alltrim (Id)) == RESOURCE_ID And Empty (Name)
				If Not Found()
					Append Blank In FoxResource
					Replace										;
						Type	  With	'PREFW',			;
						Id		  With	RESOURCE_ID,		;
						ReadOnly  With	.F.					;
						In FoxResource
				Endif

				If Not FoxResource.ReadOnly
					Save To Memo Data All Like FOXREF_*

					Replace															;
						Updated	 With  Date(),									;
						ckval	 With  Val (Sys(2007, FoxResource.Data))		;
						In FoxResource

					lSuccess = .T.
				Endif
				Set Memowidth To (nMemoWidth)

				*** JRN 2010-04-02 : Pack
				* Curiously, even if there are no changes, savings causes file explosion, so we pack afterwards
				Try
					Use (RESOURCE_FILE) Exclusive Alias FoxResource
					Pack
				Catch

				Finally
					Use

				Endtry
			Endif
		Endif

		Select (nSelect)

		Return lSuccess
	Endfunc


	* retrieve preferences from the FoxPro Resource file
	Function RestorePrefs()
		Local nSelect
		Local lSuccess
		Local nMemoWidth
		Local i
		Local nCnt

		Local Array FOXREF_LOOKFOR_MRU[10]
		Local Array FOXREF_REPLACE_MRU[10]
		Local Array FOXREF_FOLDER_MRU[10]
		Local Array FOXREF_FILETYPES_MRU[10]
		Local Array FOXREF_OPTIONS[1]

		If Not (Set ('RESOURCE') == 'ON')
			Return .F.
		Endif


		m.nSelect = Select()

		m.lSuccess = .F.

		If File (RESOURCE_FILE)    && resource file not found.
			Use (RESOURCE_FILE) In 0 Shared Again Alias FoxResource
			If Used ('FoxResource')
				m.nMemoWidth = Set ('MEMOWIDTH')
				Set Memowidth To 255

				Select FoxResource
				Locate For Upper (Alltrim (Type)) == 'PREFW' And		;
					Upper (Alltrim (Id)) == RESOURCE_ID And				;
					Empty (Name) And									;
					Not Deleted()

				If Found() And Not Empty (Data) And ckval == Val (Sys(2007, Data))
					Restore From Memo Data Additive

					If Type ('FOXREF_LOOKFOR_MRU') == 'C'
						= Acopy (FOXREF_LOOKFOR_MRU, This.aLookForMRU)
					Endif
					If Type ('FOXREF_REPLACE_MRU') == 'C'
						= Acopy (FOXREF_REPLACE_MRU, This.aReplaceMRU)
					Endif
					If Type ('FOXREF_FOLDER_MRU') == 'C'
						= Acopy (FOXREF_FOLDER_MRU, This.aFolderMRU)
					Endif
					If Type ('FOXREF_FILETYPES_MRU') == 'C'
						= Acopy (FOXREF_FILETYPES_MRU, This.aFileTypesMRU)
					Endif

					If Type ('FOXREF_OPTIONS') == 'C'
						m.nCnt = Alen (FOXREF_OPTIONS, 1)
						For m.i = 1 To m.nCnt
							If Vartype (FOXREF_OPTIONS[m.i, 1]) == 'C' And Pemstatus (This, FOXREF_OPTIONS[m.i, 1], 5)
								Store FOXREF_OPTIONS[m.i, 2] To ('THIS.' + FOXREF_OPTIONS[m.i, 1])
							Endif
						Endfor
					Endif

					m.lSuccess = .T.
				Endif

				Set Memowidth To (m.nMemoWidth)

				Use In FoxResource
			Endif
		Endif

		Select (m.nSelect)

		Return m.lSuccess
	Endfunc


	Function UpdateLookForMRU (cPattern)
		Local nRow

		nRow = Ascan (This.aLookForMRU, cPattern, -1, -1, 1, 15)
		If nRow > 0
			= Adel (This.aLookForMRU, nRow)
		Endif
		= Ains (This.aLookForMRU, 1)
		This.aLookForMRU[1] = cPattern
	Endfunc

	Function UpdateReplaceMRU (cReplaceText)
		Local nRow

		nRow = Ascan (This.aReplaceMRU, cReplaceText, -1, -1, 1, 15)
		If nRow > 0
			= Adel (This.aReplaceMRU, nRow)
		Endif
		= Ains (This.aReplaceMRU, 1)
		This.aReplaceMRU[1] = cReplaceText
	Endfunc

	Function UpdateFolderMRU (cFolder)
		Local nRow

		nRow = Ascan (This.aFolderMRU, cFolder, -1, -1, 1, 15)
		If nRow > 0
			= Adel (This.aFolderMRU, nRow)
		Endif
		= Ains (This.aFolderMRU, 1)
		This.aFolderMRU[1] = cFolder
	Endfunc


	Function UpdateFileTypesMRU (cFileTypes)
		Local nRow

		nRow = Ascan (This.aFileTypesMRU, cFileTypes, -1, -1, 1, 15)
		If nRow > 0
			= Adel (This.aFileTypesMRU, nRow)
		Endif
		= Ains (This.aFileTypesMRU, 1)
		This.aFileTypesMRU[1] = cFileTypes
	Endfunc

	Function ProjectMatch (cFileTypes, oMatchFileCollection, cFilename)
		Local i
		Local lMatch

		lMatch = .F.
		For i = 1 To oMatchFileCollection.Count
			If Empty (Justpath (oMatchFileCollection.Item (i))) Or Upper (Justpath (oMatchFileCollection.Item (i))) == Upper (Justpath (cFilename))
				If This.FileMatch (Justfname (cFilename), Justfname (oMatchFileCollection.Item (i)))
					lMatch = .T.
					Exit
				Endif
			Endif
		Endfor
		Return lMatch
	Endfunc

	* borrowed from Class Browser
	Function WildcardMatch (tcMatchExpList, tcExpressionSearched, tlMatchAsIs)
		Local lcMatchExpList, lcExpressionSearched, llMatchAsIs, lcMatchExpList2
		Local lnMatchLen, lnExpressionLen, lnMatchCount, lnCount, lnCount2, lnSpaceCount
		Local lcMatchExp, lcMatchType, lnMatchType, lnAtPos, lnAtPos2
		Local llMatch, llMatch2

		If Alltrim (tcMatchExpList) == '*.*'
			Return .T.
		Endif

		If Empty (tcExpressionSearched)
			If Empty (tcMatchExpList) Or Alltrim (tcMatchExpList) == '*'
				Return .T.
			Endif
			Return .F.
		Endif
		lcMatchExpList = Lower (Alltrim (Strtran (tcMatchExpList, Tab, ' ')))
		lcExpressionSearched = Lower (Alltrim (Strtran (tcExpressionSearched, Tab, ' ')))
		lnExpressionLen = Lenc (lcExpressionSearched)
		If lcExpressionSearched == lcMatchExpList
			Return .T.
		Endif
		llMatchAsIs = tlMatchAsIs
		If Leftc (lcMatchExpList, 1) == ["] And Rightc (lcMatchExpList, 1) == ["]
			llMatchAsIs = .T.
			lcMatchExpList = Alltrim (Substrc (lcMatchExpList, 2, Lenc (lcMatchExpList) - 2))
		Endif
		If Not llMatchAsIs And ' ' $ lcMatchExpList
			llMatch = .F.
			lnSpaceCount = Occurs (' ', lcMatchExpList)
			lcMatchExpList2 = lcMatchExpList
			lnCount = 0
			Do While .T.
				lnAtPos = At_c (' ', lcMatchExpList2)
				If lnAtPos = 0
					lcMatchExp = Alltrim (lcMatchExpList2)
					lcMatchExpList2 = ''
				Else
					lnAtPos2 = At_c (["], lcMatchExpList2)
					If lnAtPos2 < lnAtPos
						lnAtPos2 = At_c (["], lcMatchExpList2, 2)
						If lnAtPos2 > lnAtPos
							lnAtPos = lnAtPos2
						Endif
					Endif
					lcMatchExp = Alltrim (Leftc (lcMatchExpList2, lnAtPos))
					lcMatchExpList2 = Alltrim (Substrc (lcMatchExpList2, lnAtPos + 1))
				Endif
				If Empty (lcMatchExp)
					Exit
				Endif
				lcMatchType = Leftc (lcMatchExp, 1)
				Do Case
					Case lcMatchType == '+'
						lnMatchType = 1
					Case lcMatchType == '-'
						lnMatchType = -1
					Otherwise
						lnMatchType = 0
				Endcase
				If lnMatchType # 0
					lcMatchExp = Alltrim (Substrc (lcMatchExp, 2))
				Endif
				llMatch2 = This.WildcardMatch (lcMatchExp, lcExpressionSearched, .T.)
				If (lnMatchType = 1 And Not llMatch2) Or (lnMatchType = -1 And llMatch2)
					Return .F.
				Endif
				llMatch = (llMatch Or llMatch2)
				If lnAtPos = 0
					Exit
				Endif
			Enddo
			Return llMatch
		Else
			If Leftc (lcMatchExpList, 1) == '~'
				Return (Difference (Alltrim (Substrc (lcMatchExpList, 2)), lcExpressionSearched) >= 3)
			Endif
		Endif
		lnMatchCount = Occurs (',', lcMatchExpList) + 1
		If lnMatchCount > 1
			lcMatchExpList = ',' + Alltrim (lcMatchExpList) + ','
		Endif
		For lnCount = 1 To lnMatchCount
			If lnMatchCount = 1
				lcMatchExp = Lower (Alltrim (lcMatchExpList))
				lnMatchLen = Lenc (lcMatchExp)
			Else
				lnAtPos = At_c (',', lcMatchExpList, lnCount)
				lnMatchLen = At_c (',', lcMatchExpList, lnCount + 1) - lnAtPos - 1
				lcMatchExp = Lower (Alltrim (Substrc (lcMatchExpList, lnAtPos + 1, lnMatchLen)))
			Endif
			For lnCount2 = 1 To Occurs ('?', lcMatchExp)
				lnAtPos = At_c ('?', lcMatchExp)
				If lnAtPos > lnExpressionLen
					If (lnAtPos - 1) = lnExpressionLen
						lcExpressionSearched = lcExpressionSearched + '?'
					Endif
					Exit
				Endif
				lcMatchExp = Stuff (lcMatchExp, lnAtPos, 1, Substrc (lcExpressionSearched, lnAtPos, 1))
			Endfor
			If Empty (lcMatchExp) Or lcExpressionSearched == lcMatchExp Or		;
					lcMatchExp == '*' Or lcMatchExp == '?' Or lcMatchExp == '%%'
				Return .T.
			Endif
			If Leftc (lcMatchExp, 1) == '*'
				Return (Substrc (lcMatchExp, 2) == Rightc (lcExpressionSearched, Lenc (lcMatchExp) - 1))
			Endif
			If Leftc (lcMatchExp, 1) == '%' And Rightc (lcMatchExp, 1) == '%' And		;
					Substrc (lcMatchExp, 2, lnMatchLen - 2) $ lcExpressionSearched
				Return .T.
			Endif
			lnAtPos = At_c ('*', lcMatchExp)
			If lnAtPos > 0 And (lnAtPos - 1) <= lnExpressionLen And			;
					Leftc (lcExpressionSearched, lnAtPos - 1) == Leftc (lcMatchExp, lnAtPos - 1)
				Return .T.
			Endif
		Endfor
		Return .F.
	Endfunc

	* used for comparing filenames against a wildcard
	* For folder searches we can use ADIR(), but for project
	* searches we need to use this function
	Function FileMatch (cText As String, cPattern As String)
		Local i, j, k, cPattern, nPatternLen, nTextLen, ch

		If Pcount() < 2
			cPattern = This.cPattern
		Endif

		nPatternLen = Lenc (cPattern)
		nTextLen = Lenc (cText)

		If nPatternLen == 0
			Return .T.
		Endif
		If nTextLen == 0
			Return .F.
		Endif

		j = 1
		For i = 1 To nPatternLen
			If j > Lenc (cText)
				Return .F.
			Else
				ch = Substrc (cPattern, i, 1)
				If ch == FILEMATCH_ANY
					j = j + 1
				Else
					If ch == FILEMATCH_MORE
						For k = j To nTextLen
							If This.FileMatch (Substrc (cText, k), Substrc (cPattern, i + 1))
								Return .T.
							Endif
						Endfor
						Return .F.
					Else
						If j <= nTextLen And ch # Substrc (cText, j, 1)
							Return .F.
						Else
							j = j + 1
						Endif
					Endif
				Endif
			Endif
		Endfor

		* RETURN j > nTextLen
		Return .T.
	Endfunc


	* -- Parse out what's in Abstract field to return
	* -- either the search criteria or the file types searched
	Function ParseAbstract (cAbstract, cParseWhat)
		Local cDisplayText
		Local cSearchOptions

		cDisplayText = ''

		Do Case
			Case cParseWhat == 'CRITERIA'
				cSearchOptions = Leftc (cAbstract, At_c (';', cAbstract) - 1)

				If 'X' $ cSearchOptions
					cDisplayText = cDisplayText + IIf (Empty (cDisplayText), '', ', ') + COMMENTS_EXCLUDE_LOC
				Endif
				If 'C' $ cSearchOptions
					cDisplayText = cDisplayText + IIf (Empty (cDisplayText), '', ', ') + COMMENTS_ONLY_LOC
				Endif
				If 'M' $ cSearchOptions
					cDisplayText = cDisplayText + IIf (Empty (cDisplayText), '', ', ') + CRITERIA_MATCHCASE_LOC
				Endif
				If 'W' $ cSearchOptions
					cDisplayText = cDisplayText + IIf (Empty (cDisplayText), '', ', ') + CRITERIA_WHOLEWORDS_LOC
				Endif
				If 'P' $ cSearchOptions
					cDisplayText = cDisplayText + IIf (Empty (cDisplayText), '', ', ') + CRITERIA_FORMPROPERTIES_LOC
				Endif
				If 'H' $ cSearchOptions And This.ProjectFile # PROJECT_GLOBAL
					cDisplayText = cDisplayText + IIf (Empty (cDisplayText), '', ', ') + CRITERIA_PROJECTHOMEDIR_LOC
				Endif
				If 'S' $ cSearchOptions And This.ProjectFile == PROJECT_GLOBAL
					cDisplayText = cDisplayText + IIf (Empty (cDisplayText), '', ', ') + CRITERIA_SUBFOLDERS_LOC
				Endif
				If 'Z' $ cSearchOptions
					cDisplayText = cDisplayText + IIf (Empty (cDisplayText), '', ', ') + CRITERIA_WILDCARDS_LOC
				Endif

			Case cParseWhat == 'FILETYPES'
				cDisplayText = Alltrim (Substrc (Rtrim (cAbstract), At_c (';', cAbstract) + 1))

		Endcase

		Return cDisplayText
	Endfunc


	* -- Cleanup tables used by FoxRef -- removing
	* -- any references to files we can't find and
	* -- packing the files
	Function CleanupTables()
		Local nSelect
		Local cFilename
		Local lSuccess

		m.nSelect = Select()

		m.lSuccess = .F.

		If Used (Juststem (This.RefTable))
			Use In (Juststem (This.RefTable))
		Endif
		If Used (Juststem (This.DefTable))
			Use In (Juststem (This.DefTable))
		Endif
		If Used (Juststem (This.FileTable))
			Use In (Juststem (This.FileTable))
		Endif

		If This.OpenTables (.T., .T.)
			Select FileCursor
			Delete All For (FileAction == FILEACTION_ERROR Or FileAction == FILEACTION_INACTIVE) And UniqueID # 'WINDOW'
			Scan All For Not Empty (Filename)
				m.cFilename = Addbs (Rtrim (FileCursor.Folder)) + Rtrim (FileCursor.Filename)
				If Not File (m.cFilename)
					Delete In FileCursor

					Select FoxDefCursor
					Delete All For FileID == FileCursor.UniqueID In FoxDefCursor
				Endif
			Endscan

			Select FoxDefCursor
			Delete All For Inactive In FoxDefCursor

			Select FileCursor
			Pack In FileCursor

			Select FoxDefCursor
			Pack In FoxDefCursor

			m.lSuccess = .T.
		Else
			Messagebox (ERROR_EXCLUSIVE_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
		Endif


		* re-open the tables shared
		This.OpenTables()

		Select (m.nSelect)

		Return m.lSuccess
	Endfunc

	Function PrintReport (cReportName, lPreview)
		Local cFilename
		Local cExt
		Local oReportAddIn

		Try
			oReportAddIn = This.oReportAddIn.Item (m.cReportName)
		Catch
			oReportAddIn = .Null.
		Endtry

		If Not Isnull (oReportAddIn)
			With oReportAddIn
				Do Case
					Case Not Empty (.RptFilename)
						m.cExt = Upper (Justext (.RptFilename))
						Do Case
							Case m.cExt == 'PRG' Or m.cExt == 'FXP'
								Do (.RptFilename) With m.lPreview, This

							Case m.cExt == 'SCX'
								Do Form (.RptFilename) With m.lPreview, This

							Case m.cExt == 'FRX'
								Report Form (.RptFilename)
						Endcase

					Case Not Empty (This.RptClassName)
				Endcase
			Endwith
		Endif
	Endfunc

	* Abstract:
	*   This program will shell out to specified file,
	*	which can be a URL (e.g. http://www.microsoft.com),
	*	a filename, etc
	*
	* Parameters:
	*	<cFile>
	*	[cParameters]
	Function ShellTo (cFile, cParameters)
		Local cRun
		Local cSysDir
		Local nRetValue

		*-- GetDesktopWindow gives us a window handle to
		*-- pass to ShellExecute.
		Declare Integer GetDesktopWindow In user32.Dll
		Declare Integer GetSystemDirectory In kernel32.Dll		;
			String @cBuffer,									;
			Integer liSize

		Declare Integer ShellExecute In shell32.Dll			;
			Integer,										;
			String @cOperation,								;
			String @cFile,									;
			String @cParameters,							;
			String @cDirectory,								;
			Integer nShowCmd

		If Vartype (m.cParameters) # 'C'
			m.cParameters = ''
		Endif

		m.cOperation = 'open'
		m.nRetValue = ShellExecute (GetDesktopWindow(), @m.cOperation, @m.cFile, @m.cParameters, '', SW_SHOWNORMAL)
		If m.nRetValue = SE_ERR_NOASSOC && No association exists
			m.cSysDir = Space(260)  && MAX_PATH, the maximum path length

			*-- Get the system directory so that we know where Rundll32.exe resides.
			m.nRetValue = GetSystemDirectory (@m.cSysDir, Lenc (m.cSysDir))
			m.cSysDir = Substrc (m.cSysDir, 1, m.nRetValue)
			m.cRun = 'RUNDLL32.EXE'
			cParameters = 'shell32.dll,OpenAs_RunDLL '
			m.nRetValue = ShellExecute (GetDesktopWindow(), 'open', m.cRun, m.cParameters + m.cFile, m.cSysDir, SW_SHOWNORMAL)
		Endif

		Return m.nRetValue
	Endfunc

	Function CreateResourceFile (tcFileName)
		Create Table (tcFileName) (			;
			Type            C(12)			;
			, Id            C(12)			;
			, Name          M				;
			, ReadOnly      L				;
			, ckval         N(6)			;
			, Data          M				;
			, Updated       D				;
			)
		Use

	Endfunc


	*---------------------------------------------------------------------------
	#Define MAX_PATH 260
	Function	DiskFileName (lcFileName)

		Local lnFindFileData, lnHandle, lcXXX
		Declare Integer FindFirstFile In win32api String @, String @
		Declare Integer FindNextFile In win32api Integer, String @
		Declare Integer FindClose In win32api Integer

		Do Case
			Case ( Right (lcFileName, 1) == '\' )
				Return Addbs (This.DiskFileName (Left (lcFileName, Len (lcFileName) - 1)))

			Case Empty (lcFileName)
				Return ''

			Case ( Len (lcFileName) == 2 ) And ( Right (lcFileName, 1) == ':' )
				Return Upper (lcFileName)	&& win2k gives curdir() for C:
		Endcase

		lnFindFileData = Space(4 + 8 + 8 + 8 + 4 + 4 + 4 + 4 + MAX_PATH + 14)
		lnHandle		 = FindFirstFile (@lcFileName, @lnFindFileData)

		If ( lnHandle < 0 )
			If ( Not Empty (Justfname (lcFileName)) )
				lcXXX = Justfname (lcFileName)
			Else
				lcXXX = lcFileName
			Endif
		Else
			= FindClose (lnHandle)
			lcXXX	= Substr (lnFindFileData, 45, MAX_PATH)
			lcXXX	= Left (lcXXX, At (Chr(0), lcXXX) - 1)
		Endif


		Do Case
			Case Empty (Justpath (lcFileName))
				Return lcXXX
			Case ( Justpath (lcFileName) == '\' ) And (Left (lcFileName, 2) == '\\')	&& unc
				Return '\\' + lcXXX
			Otherwise
				Return Addbs (This.DiskFileName (Justpath (lcFileName))) + lcXXX
		Endcase

	Endfunc


	*---------------------------------------------------------------------------
	Function AddMRUFile (lcFileName, lcClassName)

		#Define DELIMITERCHAR Chr(0)
		#Define CR			  Chr(13)
		#Define MAXITEMS      24

		Local lcData, lcExt, lcLine25, lcList, lcMRU_ID, lcNewData, lcSearchString, lcSourceFile, lnPos
		Local lnSelect

		If 'ON' # Set ('Resource')
			Return
		Endif

		lcSourceFile = Set ('Resource', 1)
		lcExt		 = Upper (Justext (lcFileName))

		lcList = ',VCX=MRUI,PRG=MRUB,MPR=MRUB,QPR=MRUB,SCX=MRUH,MNX=MRUE,FRX=MRUG,DBC=???,LBX=???,PJX=MRUL'
		lnPos  = At (',' + lcExt + '=', lcList)
		If lnPos = 0
			lcMRU_ID = 'MRUC'
		Else
			lcMRU_ID = Substr (lcList, lnPos + 5, 4)
			If lcMRU_ID = '?'
				Return
			Endif
		Endif

		If lcMRU_ID = 'MRUI'
			If Empty (lcClassName)
				lcMRU_ID	   = 'MRU2'
				lcSearchString = lcFileName + DELIMITERCHAR
			Else
				lcSearchString = lcFileName + '|' + Lower (lcClassName) + DELIMITERCHAR
			Endif
		Else
			lcSearchString = lcFileName + DELIMITERCHAR
		Endif

		lnSelect = Select()
		Select 0
		Use (lcSourceFile) Again Shared Alias MRU_Resource_Add

		Locate For Id = lcMRU_ID
		If Found()
			lcData = DELIMITERCHAR + Substr (Data, 3)
			lnPos  = Atcc (DELIMITERCHAR + lcSearchString, lcData)
			Do Case
				Case lnPos = 1
					* already tops of the list
					lcNewData = Data
				Case lnPos = 0
					* must add to list
					lcNewData = Stuff (Data, 3, 0, lcSearchString)
					* note that GetWordNum won't accept CHR(0) as a delimiter
					lcLine25  = Getwordnum (Chrtran (Substr (lcNewData, 3), DELIMITERCHAR, CR), MAXITEMS + 1, CR)
					If Not Empty (lcLine25)
						lcNewData = Strtran (lcNewData, DELIMITERCHAR + lcLine25 + DELIMITERCHAR, DELIMITERCHAR, 1, 1, 1)
					Endif
				Otherwise
					lcNewData = Stuff (Data, lnPos + 1, Len (lcSearchString), '')
					lcNewData = Stuff (lcNewData, 3, 0, lcSearchString)
			Endcase
			Replace															;
				Data	 With  lcNewData									;
				ckval	 With  Val (Sys(2007, Substr (lcNewData, 3)))		;
				Updated	 With  Date()
		Else
			lcNewData = Chr(4) + DELIMITERCHAR + lcSearchString
			Insert Into MRU_Resource_Add			;
				(Type, Id, ckval, Data, Updated)	;
				values 								;
				('PREFW', lcMRU_ID, Val (Sys(2007, Substr (lcNewData, 3))), lcNewData, Date())

			****************************************************************
		Endif

		Use
		Select (lnSelect)
	Endfunc


	*---------------------------------------------------------------------------
	Function GetMRUList (lcMRU_ID)

		Local loCollection As 'Collection'
		Local laItems(1), lcData, lcList, lcSourceFile, lnI, lnPos, lnSelect

		loCollection = Createobject ('Collection')

		If 'ON' # Set ('Resource')
			Return loCollection
		Endif

		lnSelect	 = Select()
		lcSourceFile = Set ('Resource', 1)
		Select 0
		Use (lcSourceFile) Again Shared Alias MRU_Resource_Get

		If lcMRU_ID # 'MRU'
			lcList = ',VCX=MRUI,PRG=MRUB,MPR=MRUB,QPR=MRUB,SCX=MRUH,MNX=MRUE,FRX=MRUG,DBC=???,LBX=???,PJX=MRUL'
			lnPos  = At (',' + Upper (Justext ('.' + lcMRU_ID)) + '=', lcList)
			If lnPos = 0
				lcMRU_ID = 'MRUC'
			Else
				lcMRU_ID = Substr (lcList, lnPos + 5, 4)
				If lcMRU_ID = '?'
					Return
				Endif
			Endif
		Endif

		Locate For Id = lcMRU_ID
		If Found()
			lcData = Data
			Alines (laItems, Substr (lcData, 3), 0, DELIMITERCHAR)
			For lnI = 1 To Alen (laItems)
				If Not Empty (laItems (lnI))
					loCollection.Add (laItems (lnI))
				Endif
			Endfor
		Endif

		Use
		Select (lnSelect)
		Return loCollection

	Endfunc

Enddefine


Define Class ReportAddIn As Custom
	Name = 'ReportAddIn'

	ReportName      = ''
	RptFilename     = ''
	RptClassLibrary = ''
	RptClassName    = ''
	RptMethod       = ''
Enddefine

Define Class RefOption As Custom
	Name = 'RefOption'

	OptionName   = ''
	Description  = ''
	OptionValue  = .Null.
	PropertyName = ''
Enddefine




Procedure _EdGetLPos
Procedure _EdSelect
Procedure _EdSToPos
Procedure _edgetenv
Procedure _wfindtitl
Procedure _wselect

Procedure DeCodeTimeStamp

	Lparameters nTimestamp

	Local nDate, nDay, nHr, nMin, nMonth, nSec, nTime, nYear

	If nTimestamp = 0
		Return Datetime()
	Endif

	nDate = Bitrshift (nTimestamp, 16)
	nTime = Bitand (nTimestamp, 2^16 - 1)

	nYear = Bitand (Bitrshift (nDate, 9), 2^8 - 1) + 1980
	nMonth = Bitand (Bitrshift (nDate, 5), 2^4 - 1)
	nDay = Bitand (nDate, 2^5 - 1)

	nHr = Bitand (Bitrshift (nTime, 11), 2^5 - 1)
	nMin = Bitand (Bitrshift (nTime, 5), 2^6 - 1)
	nSec = Bitand (nTime, 2^5 - 1)

	Return Datetime (nYear, nMonth, nDay, nHr, nMin, nSec)

	Return
