#include "foxref.h"

* Abstract:
*   Class for add/retrieving values
*	from FoxUser resource file.
*

Define Class FoxResource As Custom
	Protected oCollection

	oCollection  = .Null.

	ResourceType = 'PREFW'
	ResourceFile = ''

	Procedure Init()
		This.oCollection = Createobject ('Collection')
		*** JRN 2010-06-02 : THIS.ResourceFile = SYS(2005)
		This.ResourceFile = RESOURCE_FILE
	Endproc

	* Clear out all options
	Function Clear()
		This.oCollection.Remove (-1)
	Endfunc

	Function Set (cOption, xValue)
		* remove if already exists
		If This.OptionExists (m.cOption)
			This.oCollection.Remove (Upper (m.cOption))
		Endif

		* Add back in
		Return This.oCollection.Add (m.xValue, Upper (m.cOption))
	Endfunc

	Function Get (cOption)
		Local xValue
		Local i

		m.xValue = .Null.
		m.cOption = Upper (m.cOption)
		For m.i = 1 To This.oCollection.Count
			If Upper (This.oCollection.GetKey (m.i)) == m.cOption
				m.xValue = This.oCollection.Item (m.i)
				Exit
			Endif
		Endfor

		Return m.xValue
	Endfunc

	Function OptionExists (cOption)
		Local i
		Local lExists

		m.lExists = .F.
		m.cOption = Upper (m.cOption)
		For m.i = 1 To This.oCollection.Count
			If Upper (This.oCollection.GetKey (m.i)) == m.cOption
				m.lExists = .T.
				Exit
			Endif
		Endfor

		Return m.lExists
	Endfunc

	Function OpenResource (lUpdating)
		Local lSuccess
		Local Array aFileInfo[1]

		*** JRN 2010-06-02 :
		*!*	IF !(SET("RESOURCE") == "ON")
		*!*		RETURN .F.
		*!*	ENDIF

		lSuccess = .F.

		If Not Used ('FoxResource')
			If Adir (aFileInfo, This.ResourceFile) > 0 And ( Not lUpdating Or Not ('R' $ aFileInfo[1, 5]))
				Try
					Use (This.ResourceFile) Alias FoxResource In 0 Shared Again
					lSuccess = Used ('FoxResource')
				Catch
				Endtry
			Endif
		Else
			lSuccess = Not lUpdating Or (Adir (aFileInfo, This.ResourceFile) > 0 And Not ('R' $ aFileInfo[1, 5]))
		Endif

		Return lSuccess
	Endfunc

	Procedure Save (cID, cName)
		Local nSelect
		Local cType
		Local i
		Local Array aOptions[1]

		If Vartype (m.cName) # 'C'
			m.cName = ''
		Endif
		If This.OpenResource (.T.)
			m.nSelect = Select()

			m.cType = Padr (This.ResourceType, Len (FoxResource.Type))
			m.cID   = Padr (m.cID, Len (FoxResource.Id))

			Select FoxResource
			Locate For Type == m.cType And Id == m.cID And Name == m.cName
			If Not Found()
				Append Blank In FoxResource
				Replace															;
									Type With m.cType,							;
									Name With m.cName,							;
									Id With m.cID,								;
									ReadOnly With .F.							;
									In FoxResource
			Endif

			If Not FoxResource.ReadOnly
				If This.oCollection.Count > 0
					Dimension aOptions[THIS.oCollection.Count, 2]
					For m.i = 1 To This.oCollection.Count
						aOptions[m.i, 1] = This.oCollection.GetKey (m.i)
						aOptions[m.i, 2] = This.oCollection.Item (m.i)
					Endfor
					Save To Memo Data All Like aOptions
				Else
					Blank Fields Data In FoxResource
				Endif

				Replace															;
									Updated With Date(),						;
									ckval With Val (Sys(2007, FoxResource.Data));
									In FoxResource
			Endif

			Select (m.nSelect)
		Endif
	Endproc

	Procedure Load (cID, cName)
		Local nSelect
		Local cType
		Local i
		Local nCnt
		Local Array aOptions[1]

		If Vartype (m.cName) # 'C'
			m.cName = ''
		Endif

		This.Clear()
		If This.OpenResource()
			m.nSelect = Select()

			m.cType = Padr (This.ResourceType, Len (FoxResource.Type))
			m.cID   = Padr (m.cID, Len (FoxResource.Id))

			Select FoxResource
			Locate For Type == m.cType And Id == m.cID And Name == m.cName
			If Found() And Not Empty (Data) And ckval == Val (Sys(2007, Data))
				Restore From Memo Data Additive
				If Vartype (aOptions[1, 1]) == 'C'
					m.nCnt = Alen (aOptions, 1)
					For m.i = 1 To m.nCnt
						This.Set (aOptions[m.i, 1], aOptions[m.i, 2])
					Endfor
				Endif
			Endif

			Select (m.nSelect)
		Endif
	Endproc

	Function GetData (cID, cName)
		Local cData
		Local nSelect
		Local cType

		If Vartype (m.cName) # 'C'
			m.cName = ''
		Endif

		m.cData = .Null.
		If This.OpenResource()
			m.nSelect = Select()

			m.cType = Padr (This.ResourceType, Len (FoxResource.Type))
			m.cID   = Padr (m.cID, Len (FoxResource.Id))

			Select FoxResource
			Locate For Type == m.cType And Id == m.cID And Name == m.cName
			If Found() And Not Empty (Data) && AND ckval == VAL(SYS(2007, Data))
				m.cData = FoxResource.Data
			Endif

			Select (m.nSelect)
		Endif

		Return m.cData
	Endfunc

	* save to a specific fieldname
	Function SaveTo (cField, cAlias)
		Local i
		Local nSelect
		Local lSuccess
		Local Array aOptions[1]

		If Vartype (m.cAlias) # 'C'
			m.cAlias = Alias()
		Endif

		If Used (m.cAlias)
			m.nSelect = Select()
			Select (m.cAlias)

			If This.oCollection.Count > 0
				Dimension aOptions[THIS.oCollection.Count, 2]
				For m.i = 1 To This.oCollection.Count
					aOptions[m.i, 1] = This.oCollection.GetKey (m.i)
					aOptions[m.i, 2] = This.oCollection.Item (m.i)
				Endfor
				Save To Memo &cField All Like aOptions
			Else
				Blank Fields &cField In FoxResource
			Endif
			Select (m.nSelect)
			m.lSuccess = .T.
		Else
			m.lSuccess = .F.
		Endif

		Return m.lSuccess
	Endfunc


	Function RestoreFrom (cField, cAlias)
		Local i
		Local nSelect
		Local lSuccess
		Local Array aOptions[1]

		If Vartype (m.cAlias) # 'C'
			m.cAlias = Alias()
		Endif

		If Used (m.cAlias)
			m.nSelect = Select()
			Select (m.cAlias)

			Restore From Memo &cField Additive
			If Vartype (aOptions[1, 1]) == 'C'
				m.nCnt = Alen (aOptions, 1)
				For m.i = 1 To m.nCnt
					This.Set (aOptions[m.i, 1], aOptions[m.i, 2])
				Endfor
			Endif

			Select (m.nSelect)
			m.lSuccess = .T.
		Else
			m.lSuccess = .F.
		Endif

		Return m.lSuccess
	Endfunc
Enddefine

