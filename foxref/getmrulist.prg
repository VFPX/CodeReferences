Function GetMRUList (lcMRU_ID)

	#Define DELIMITERCHAR Chr(0)

	Local loCollection As 'Collection'
	Local laItems[1], lcData, lcSourceFile, lnI, lnSelect

	loCollection = Createobject ('Collection')

	If 'ON' # Set ('Resource')
		Return loCollection
	Endif

	lnSelect	 = Select()
	lcSourceFile = Set ('Resource', 1)
	Use (lcSourceFile) Again Shared Alias MRU_Resource

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
