REM  *****  BASIC  *****

' Doc sur le basic openoffice / libreoffice
Rem https://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XCellRange.html
Rem https://wiki.documentfoundation.org/Macros/Calc/fr

Option Explicit


' Named range format for auto-insert-transaction ______________________________
' |   0   |    1    |    2   |     3     |    4     |    5     |    7     |    6     |    
' |  day  |  month  |  year  |   account |  desc.   | category |   memo   |  amount  |  
Const autoNamedRange = "AUTO_NAME"
Const autoPlanDayCol = 0
Const autoPlanMonthCol  = 1
Const autoPlanYearCol  = 2
Const autoAccountCol = 3
Const autoDescCol = 4
Const autoCategoryCol = 5
Const autoAmountCol = 6
Const autoMemoCol = 7




' Main account movements named range column format
' |   0    |     1     |    2    |     3      |    4     |    5    |  6  |  7
' |  date  |  account  |  descr. |  Category  |  amount  |  memo   |  X  |  V
Const mvtSheetName = "Mvt"
Const mvtNamedRange = "MVT_NAME" 'Must cover column until mvtMaxCol
Const mvtDateCol = 0
Const mvtAccountCol = 1
Const mvtDescCol = 2
Const mvtCategoryCol = 3
Const mvtAmountCol = 4
Const mvtMemoCol = 5
Const mvtXCol = 6
Const mvtVCol = 7
Const mvtMaxCol = mvtVCol

' Categories named range column format
' |  0   |       1       |    2  |     3        |
' |  ID  |  Categ. Name  |  Type |  Description | 

'Named range on column 1
Const categoryNamedArea = "CATEGORIES_NAME" 'deprecated
Const categoryNamedRange = "CATEGORIES_NAME"
Const categoryAllNamedArea = "CATEGORIES"
Const catIdCol = 0
Const catNameCol = 1
' -1 for expense / 1 for credit / 0 for bank account (and inside budget)
Const catTypeCol = 2
Const catDescCol= 3

' Check range column format
' |   0    |     1     |    2    |     3      |    4     |    5    |    6     |    7
' |  date  |  account  |  descr. |  catégory  |  amount  |  memo   |  report  |  Match

'Named range on column 1
Const checkSheetName = "Check"
Const checkDateCol = mvtDateCol
Const checkAccountCol = mvtAccountCol
Const checkDescCol = mvtDescCol
Const checkAmountCol = mvtAmountCol
Const checkMemoCol = mvtMemoCol
' Report : indicate line where to report if any
Const checkReportCol = 6
' When match contain matching line in mvtNamedRange
Const checkMatchCol = 7
Const checkMaxCol = checkMatchCol

Const CHECK_CHEQUE_TAG = "Ch."		' Mvt entry containing Ch. are allowed a variation of 365 days in matching opération
Const CHECK_VERIFY_TAG = "V"		' Ask to mark as verified check entry in mvt at indicated line
Const CHECK_REPORT_TAG = "Report"	' Ask to report check entry in mvt entry
Const CHECK_APPROX_MATCH_TAG = "?" 	' We guess this is the right match
Const CHECK_DONE_TAG = "Done"		' Entry is checked manually or not (nothing else to do)

' Test class
Sub Main

	dim LibOClasseur as object
	dim LibOFeuille as object
	dim LibOCellule as object
	dim LibORange as object
	Dim theDate, theCategory, theDescription, theAccount as String
	Dim theAmount as Double
	Dim iLine%

	if ( MsgBox("Copy template to destination before running tests",4) = 7) Then
		Exit Sub
	End If

	MsgBox("Check Range Size :" + mvtNamedRange)
	LibORange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells
	LibORange.getCellByPosition(mvtMaxCol,0).Formula
	MsgBox("Check Range Size :" + mvtNamedRange+ " ----> OK")

	MsgBox("Test insertLineMvt function on 30/05/2015")
	LibORange = insertLineMvt(CDate("30/05/2015"), "LCL", "Test1,5", "Budget Annuel", -1169.04 , "mémo test")
	
	if (not IsNull(LibORange)) Then
	    MsgBox("Test insertLineMvt : date=" + LibORange.getCellByPosition(mvtDateCol,0).FormulaLocal + " done on line=" + LibORange.RangeAddress.startRow _
	    	+ " of range named " + mvtNamedRange)
	Else
	    MsgBox("Test insertLineMvt on date=30/05/2015 FAILED in range named " + mvtNamedRange)
		
	End If

	MsgBox("Test insertLineMvt function on at Line 3")
	LibORange = insertLineMvt(CDate("30/05/2013"), "LCL", "Test1,5", "Budget Annuel", -1169.04 , "mémo test", 2)
	
	if (not IsNull(LibORange)) Then
	    MsgBox("Test insertLineMvt : date=" + LibORange.getCellByPosition(mvtDateCol,0).FormulaLocal + " done on line=" + LibORange.RangeAddress.startRow _
	    	+ " of range named " + mvtNamedRange)
	Else
	    MsgBox("Test insertLineMvt at Line 3 FAILED in range named " + mvtNamedRange)
		
	End If

	MsgBox("Test findDateCell function")
	LibORange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells
	LibOCellule = findDateCell(LibORange, CDate("30/05/2015"))
	
	if (not IsNull(LibOCellule)) Then
	    MsgBox("Test findDateCell : date=" + LibOCellule.FormulaLocal + " found on line=" + LibOCellule.RangeAddress.startRow _
	    	+ " of range named " + mvtNamedRange)
	Else
	    MsgBox("Test findDateCell : date=30/05/2015 NOT FOUND in range named " + mvtNamedRange)
		
	End If

	MsgBox("Test findDateCellBefore function")
	LibORange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells
	LibOCellule = findDateCellBefore(LibORange, CDate("30/05/2015"))
	
	if (not IsNull(LibOCellule)) Then
	    MsgBox("Test findDateCellBefore : date=" + LibOCellule.FormulaLocal + " found on line=" + LibOCellule.RangeAddress.startRow _
	    	+ " of range named " + mvtNamedRange)
	Else
	    MsgBox("Test findDateCellBefore : date=30/05/2015 NOT FOUND in range named " + mvtNamedRange)
		
	End If

	
	MsgBox("Test matchingMvt function")
	Dim oMvtLine as Object
	LibORange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells
	oMvtLine =  matchingMvt(CDate("06/05/2015"), "LCL", "Test2 – matchingMvt", "TestTest", -100.05)
	if (not IsNull(oMvtLine)) Then
		theDate = oMvtLine.getCellByPosition(mvtDateCol,0).FormulaLocal
		theAccount = oMvtLine.getCellByPosition(mvtAccountCol,0).Formula
		theCategory = oMvtLine.getCellByPosition(mvtCategoryCol,0).Formula
		theDescription = oMvtLine.getCellByPosition(mvtDescCol,0).Formula
		theAmount = oMvtLine.getCellByPosition(mvtAmountCol,0).Value
		iLine = oMvtLine.RangeAddress.startRow
		
	    MsgBox("Test mvtMatching :" + "date=" + theDate + ", account=" + theAccount + ", desc=" + _
	    		theDescription + ", category=" + theCategory + ", amount=" + theAmount + _
	    		" on line" + iLine)
	Else
	    MsgBox("Test mvtMatching NOT FOUND : 06/05/2015, LCL, monoprix, alimentation, -100,05")
	End If

	MsgBox("Test matchingMvtForCheck function with desc='matching'")
	LibORange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells
	Dim infoMatch as String
	oMvtLine =  matchingMvtForCheck(CDate("06/05/2015"), "LCL", "matChing", -100.05, infoMatch)
	if (not IsNull(oMvtLine)) Then
		theDate = oMvtLine.getCellByPosition(mvtDateCol,0).FormulaLocal
		theAccount = oMvtLine.getCellByPosition(mvtAccountCol,0).Formula
		theCategory = oMvtLine.getCellByPosition(mvtCategoryCol,0).Formula
		theDescription = oMvtLine.getCellByPosition(mvtDescCol,0).Formula
		theAmount = oMvtLine.getCellByPosition(mvtAmountCol,0).Value
		iLine = oMvtLine.RangeAddress.startRow
		
	    MsgBox("Test matchingMvtForCheck :" + "date=" + theDate + ", account=" + theAccount + ", desc=" + _
	    		theDescription + ", category=" + theCategory + ", amount=" + theAmount + _
	    		" on line" + iLine)
	Else
	    MsgBox("Test matchingMvtForCheck NOT FOUND : 06/05/2015, LCL, monoprix, alimentation, -100,05")
	End If
	
	

End Sub

Function dateNear(dateSource as Date, dateBound as Date, boundSize as Integer) as Boolean 
	
	Dim iVal%
	
	iVal = abs(dateDiff("d",dateSource,dateBound))
	
	dateNear = 0
	
	if ( abs(dateDiff("d",dateSource,dateBound)) <= boundSize ) Then
		dateNear = 1
	End If

	if (iVal < 2 ) Then
		iVal = iVal
	End If

End Function

Function findDateCell(aRange as Object, thedate as Date) as Object
	Dim nCurCol%, nCurRow%, nEndCol%, nEndRow%
	Dim oCell as Object
	
	nCurRow = aRange.RangeAddress.StartRow
	nEndRow = aRange.RangeAddress.EndRow
	nCurCol = aRange.RangeAddress.StartColumn
	nEndCol = aRange.RangeAddress.EndColumn
	
	For nCurCol = 0 To nEndCol
		For nCurRow = 0 To nEndRow
			oCell = aRange.GetCellByPosition( mvtDateCol, nCurRow )
			
			If (thedate = oCell.Value) then
				findDateCell = oCell
				nCurCol = nEndCol
				nCurRow = nEndRow
			End If
		Next
	Next
	

End function

Function findDateCellBefore(aRange as Object, thedate as Date) as Object
	Dim nCurCol%, nCurRow%, nEndCol%, nEndRow%
	Dim oCell as Object
	
	nCurRow = aRange.RangeAddress.StartRow
	nEndRow = aRange.RangeAddress.EndRow
	nCurCol = aRange.RangeAddress.StartColumn
	nEndCol = aRange.RangeAddress.EndColumn
	
	For nCurCol = 0 To nEndCol
		For nCurRow = 0 To nEndRow
			oCell = aRange.GetCellByPosition( mvtDateCol, nCurRow )
			
			If (thedate > oCell.Value) and IsDate(oCell.FormulaLocal) then
				findDateCellBefore = oCell
				nCurCol = nEndCol
				nCurRow = nEndRow
			End If
		Next
	Next
	

End function




Function matchingMvt(matchDate as Date, matchAccount as String, matchDesc as String, matchCategory as String, matchAmount as Double) as Object
    Dim mvtRange As Object
	Dim nCurRow%, nEndRow%
	Dim oCell as Object

	'MsgBox("matchingMvt " + matchDate + "," + matchAccount + "," + matchDesc + "," + matchCategory + "," + matchAmount)
	mvtRange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells
	
	nCurRow = mvtRange.RangeAddress.StartRow
	nEndRow = mvtRange.RangeAddress.EndRow
	
	For nCurRow = 0 To nEndRow
		oCell = mvtRange.GetCellByPosition( mvtDateCol, nCurRow )
		
		If (matchDate = oCell.Value) then
			Dim currentDesc, currentAccount, currentCategory as String
			Dim currentAmount as Double
			
			currentAccount = mvtRange.getCellByPosition(mvtAccountCol, nCurRow).Formula
			currentDesc = mvtRange.getCellByPosition(mvtDescCol, nCurRow).Formula
			currentCategory = mvtRange.getCellByPosition(mvtCategoryCol, nCurRow).Formula
			currentAmount = mvtRange.getCellByPosition(mvtAmountCol, nCurRow).Value
					
rem			if (currentAccount = matchAccount) and (currentDesc = matchDesc) and _
rem				(currentCategory = matchCategory) and (currentAmount = matchAmount) Then
			if (currentAccount = matchAccount) and (currentDesc = matchDesc) and _
				(currentCategory = matchCategory)  Then
				matchingMvt = mvtRange.getCellRangeByPosition(mvtDateCol, nCurRow, mvtAmountCol, nCurRow)
				nCurRow = nEndRow
			End If		
		End If
	Next
End Function


Function newLineMvt(newLineIndex as Long)

    Dim oSheet As Object
    
    Dim srcCellRange As Object
    Dim dstCellRange As Object
	
	oSheet = ThisComponent.Sheets.getByName(mvtSheetName)
	oSheet.Rows.insertByIndex(newLineIndex,1)
	srcCellRange = oSheet.getCellRangeByPosition(0,newLineIndex+1, 20, newLineIndex+1)
	dstCellRange = oSheet.getCellByPosition(0, newLineIndex)
	oSheet.CopyRange(dstCellRange.CellAddress, srcCellRange.RangeAddress)
End Function


Function setMvtCheck(iLine as Integer, xValue as Integer, vValue as Integer) 
	Dim insertCell, mvtRange as Object	
	mvtRange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells
	
	mvtRange.getCellByPosition(mvtXCol,iLine).setValue(xValue)
	mvtRange.getCellByPosition(mvtVCol,iLine).setValue(vValue)
End Function

Function insertLineMvt(insertDate as Date, account as String, description as String, _
					category as String, amount as Double, memo as String, optional insertLineArg as Integer) as Object

	Dim insertLine%
	Dim insertCell, mvtRange as Object	
	
	'MsgBox("line does not exist ")
	
	mvtRange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells
	
	If ( IsMissing(insertLineArg) ) then
		insertCell = findDateCellBefore(mvtRange, insertDate)
		
		If (not IsNull(insertCell)) then
			insertLine = insertCell.RangeAddress.StartRow
		End if
		if (insertLine < 1) then
			insertLine = 1
		End if
	Else
		insertLine = insertLineArg
	End If					
	
	newLineMvt(insertLine)

	mvtRange.getCellByPosition(mvtDateCol,insertLine).setValue(insertDate)
	mvtRange.getCellByPosition(mvtAccountCol,insertLine).setFormula(account)
	mvtRange.getCellByPosition(mvtDescCol,insertLine).setFormula(description)
	mvtRange.getCellByPosition(mvtCategoryCol,insertLine).setFormula(category)
	mvtRange.getCellByPosition(mvtAmountCol,insertLine).setValue(amount)
	mvtRange.getCellByPosition(mvtXCol,insertLine).setFormula("")
	mvtRange.getCellByPosition(mvtVCol,insertLine).setFormula("")
	mvtRange.getCellByPosition(mvtMemoCol,insertLine).setFormula(memo)


	insertLineMvt = mvtRange.getCellRangeByPosition(0, insertLine, mvtMaxCol, insertLine)
End Function

Rem si une date est supérieur à un pattern on ajoute une ligne
Sub AutoTransaction
    Dim oCellRange As Object
    Dim x as Long
	Dim refDate as Date, szRefDate as String
	
	refDate = now()
	
	szRefDate = InputBox("Date de complétude","Confirmer la date de complétude",FORMAT(refDate,"dd/mm/yyyy"))
	refDate = DateValue(szRefDate)
	rem InputBox("Ref Date","Confirm Ref date",FORMAT(refDate,"dd/mm/yyyy"))
	
	oCellRange = ThisComponent.NamedRanges.getByName(autoNamedRange).ReferredCells

	For x = 0 to oCellRange.Rows.Count - 1
		Dim insertDate as Date
		Dim planDay, planMonth, planYear As Long
		
		planDay   = oCellRange.getCellByPosition(autoPlanDayCol,x).getValue()
		planMonth = oCellRange.getCellByPosition(autoPlanMonthCol,x).getValue()
		planYear  = oCellRange.getCellByPosition(autoPlanYearCol,x).getValue()
		
		
		if planDay = 0 then 
			planDay = Day(refDate)
		end if
		if planMonth = 0 then 
			planMonth = Month(refDate)
		end if
		if planYear = 0 then 
			planYear = Year(refDate)
		end if
		

		rem oCellRange.getCellByPosition(autoAmountCol,x).setValue(0)

		insertDate = DateSerial(planYear,planMonth,planDay)
		if  insertDate <= now() then
			' Event is to be inserted
			Dim account, description, category, memo as String
			Dim amount as Double
			Dim alreadyExist as Object
			
			'MsgBox("adding line " + x)
			'autoMemoCol = 7
			account     = oCellRange.getCellByPosition(autoAccountCol,x).getFormula()
			description = oCellRange.getCellByPosition(autoDescCol,x).getFormula()
			category    = oCellRange.getCellByPosition(autoCategoryCol,x).getFormula()
			amount      = oCellRange.getCellByPosition(autoAmountCol,x).getValue()
			memo        = oCellRange.getCellByPosition(autoMemoCol,x).getFormula()
			
			alreadyExist = matchingMvt(insertDate, account, description, category, amount)
			If isNull(alreadyExist) Then
				' Transaction does not exist
				Dim newLine as Object
				
				newLine = insertLineMvt(insertDate, account, description, category, amount, memo )
				newLine.getCellByPosition(mvtDateCol,0).CellBackColor = RGB(0,255,0)
				
			End if
		End if
		
	Next x

End Sub

Sub CheckCategory
	Dim categoryRange as Object
	Dim mvtRange as Object
	Dim mvtLineIndex%, catLineIndex%
	
	
	' Sélectionner les data de la zone category
	categoryRange = ThisComponent.NamedRanges.getByName(categoryNamedRange).ReferredCells
	mvtRange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells


	For mvtLineIndex = mvtRange.RangeAddress.startRow + 1  To mvtRange.RangeAddress.endRow
		Dim mvtCategory, catCategory as String
		mvtCategory = mvtRange.getCellByPosition(mvtCategoryCol, mvtLineIndex).Formula
		Dim oo
		oo = mvtRange.getCellByPosition(mvtCategoryCol, mvtLineIndex).CellBackColor
		mvtRange.getCellByPosition(mvtCategoryCol, mvtLineIndex).CellBackColor = RGB(255,0,0)
		
		For catLineIndex  = categoryRange.RangeAddress.startRow + 1 To _
			categoryRange.RangeAddress.endRow
			catCategory = categoryRange.getCellByPosition(0, catLineIndex).Formula
			
			If (mvtCategory = catCategory) Then
				mvtRange.getCellByPosition(mvtCategoryCol, mvtLineIndex).CellBackColor = -1				
				catLineIndex = categoryRange.RangeAddress.endRow
			End If
		Next catLineIndex 
	
	
	Next mvtLineIndex 
	
	' Parcourir la sélection et mettre en rouge les zone qui ne sont pas dans la liste
    
End Sub


' Ajoute une ligne de saisie dans la feuille de mouvement
Sub NewLine

	Dim newDate As Date
    Dim oSheet As Object
    
    Dim srcCellRange As Object
    Dim dstCellRange As Object
	Dim newLineIndex As Long
	
	newLineIndex = ThisComponent.getCurrentController().getSelection().RangeAddress.StartRow

	newLineMvt(newLineIndex)

	oSheet = ThisComponent.Sheets.getByName(mvtSheetName)

	if (newLineIndex <= 1) Then
		' We insert at the beginning of the mvt sheet
		newLineIndex = 1
		newDate = now()
		newDate = DateSerial(Year(newDate),Month(newDate),Day(newDate))
	Else
		' We insert in the "middle" of the mvt sheet (and we use date just below)
		newDate = oSheet.getCellByPosition(mvtDateCol,newLineIndex+1).FormulaLocal
	End if


	oSheet.getCellByPosition(mvtDateCol,newLineIndex).FormulaLocal = newDate
'	oSheet.getCellByPosition(mvtAccountCol,newLineIndex).setFormula("")
	oSheet.getCellByPosition(mvtDescCol,newLineIndex).setFormula("")
	oSheet.getCellByPosition(mvtCategoryCol,newLineIndex).setFormula("")
	oSheet.getCellByPosition(mvtAmountCol,newLineIndex).setValue(0.0)
	oSheet.getCellByPosition(mvtMemoCol,newLineIndex).setFormula("")
	oSheet.getCellByPosition(mvtXCol,newLineIndex).setFormula("")
'	oSheet.getCellByPosition(mvtVCol,newLineIndex).setFormula("")


	
End Sub

' Crée la transation inverse
Sub ReverseLine

	Dim newDate As Date
    Dim oSheet As Object
    
    Dim srcCellRange As Object
    Dim dstCellRange As Object
	Dim newLineIndex As Long
	

	oSheet = ThisComponent.Sheets.getByName(mvtSheetName)
	newLineIndex = ThisComponent.getCurrentController().getSelection().RangeAddress.StartRow
	srcCellRange = oSheet.getCellRangeByPosition(mvtDateCol, newLineIndex,20, newLineIndex)
	
	newLineIndex = newLineIndex+1
	
	oSheet.Rows.insertByIndex(newLineIndex,1)
	
	dstCellRange = oSheet.getCellByPosition(0, newLineIndex)
	oSheet.CopyRange(dstCellRange.CellAddress, srcCellRange.RangeAddress)
	
	dstCellRange = oSheet.getCellRangeByPosition(mvtDateCol, newLineIndex,20, newLineIndex)

	dstCellRange.getCellByPosition(mvtAmountCol,0).Value = - srcCellRange.getCellByPosition(mvtAmountCol,0).Value
	dstCellRange.getCellByPosition(mvtAccountCol,0).Formula = srcCellRange.getCellByPosition(mvtCategoryCol,0).Formula
	dstCellRange.getCellByPosition(mvtCategoryCol,0).Formula = srcCellRange.getCellByPosition(mvtAccountCol,0).Formula

	
End Sub


' Parse check sheet
'   - mark entry checked when match found in mvt
'         match mvt if : v is on, 5+ letters word in check desc match with mvt desc + account + sum match, and category is set
'   - report entry in mvt is asked
Sub MatchMvtWithCheck
    Dim toCheckRange As Object
    Dim mvtRange As Object
	Dim nBgnRow%, nEndRow%, checkLineIndex%
	Dim oCellToMatch as Object

	'MsgBox("matchingMvt " + matchDate + "," + matchAccount + "," + matchDesc + "," + matchCategory + "," + matchAmount)
	mvtRange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells

	nBgnRow = ThisComponent.getCurrentController().getSelection().RangeAddress.StartRow
	nEndRow = ThisComponent.getCurrentController().getSelection().RangeAddress.EndRow
	
	if ( MsgBox("Confirmer l'analyse des lignes " + nBgnRow + " à " + nEndRow, 4) = 7) Then
		Exit Sub
	End If

	
	toCheckRange = ThisComponent.sheets.getByName(checkSheetName).getCellRangeByPosition(0, nBgnRow, checkMaxCol, nEndRow)
	
	
	For checkLineIndex = 0   To (nEndRow - nBgnRow)
		Dim report$, checkDesc$, checkAccount$, checkMemo$
		Dim checkDate as Date
		Dim checkAmount as Double
		Dim mvtRangeMatchingCheck as Object
		Dim newLine as Object
		Dim infoMatch as String
		
		
		
		report = toCheckRange.getCellByPosition(checkReportCol,checkLineIndex).getFormula()
		checkDate = CDate(toCheckRange.getCellByPosition(checkDateCol,checkLineIndex).FormulaLocal)
		checkAmount = toCheckRange.getCellByPosition(checkAmountCol,checkLineIndex).getValue()
		checkDesc = toCheckRange.getCellByPosition(checkDescCol,checkLineIndex).getFormula()
		checkAccount = toCheckRange.getCellByPosition(checkAccountCol,checkLineIndex).getFormula()
		checkMemo = toCheckRange.getCellByPosition(checkMemoCol,checkLineIndex).getFormula()

		if ( strComp(report,CHECK_DONE_TAG ) <> 0 ) then 
		
			' If line is asked to report (!!! should increment further number in ReportCol !!!)
			if ( strComp(report,CHECK_REPORT_TAG ) = 0 ) then
				newLine = insertLineMvt(checkDate, checkAccount, checkDesc, "", checkAmount, checkMemo)
				toCheckRange.getCellByPosition(checkMatchCol,checkLineIndex).setValue(newLine.RangeAddress.startRow + 1 )
				setMvtCheck(newLine.RangeAddress.startRow, 0, 1)
				toCheckRange.getCellByPosition(checkReportCol,checkLineIndex).setFormula(CHECK_DONE_TAG)
			Else
			
				' Search for a hit
				toCheckRange.getCellByPosition(checkMatchCol,checkLineIndex).SetFormula("...search")
				mvtRangeMatchingCheck = matchingMvtForCheck(checkDate, checkAccount, checkDesc, checkAmount, infoMatch)
		
				If (Not isNull(mvtRangeMatchingCheck)) Then
					Dim nMatchingLine%
					
					nMatchingLine = mvtRangeMatchingCheck.RangeAddress.StartRow
				
					toCheckRange.getCellByPosition(checkMatchCol,checkLineIndex).SetFormula(nMatchingLine+1)
					If ( strComp(infoMatch,CHECK_APPROX_MATCH_TAG) = 0) Then
						toCheckRange.getCellByPosition(checkMatchCol,checkLineIndex).CellBackColor = RGB(0,255,255)
					End If 
				Else
					toCheckRange.getCellByPosition(checkMatchCol,checkLineIndex).SetFormula("Aucune")
				End If
			End if
	
	
			if ( strComp(report,CHECK_VERIFY_TAG ) = 0 ) then
				Dim iLine as Integer 
				iLine = toCheckRange.getCellByPosition(checkMatchCol,checkLineIndex).getValue()
				setMvtCheck(iLine - 1, 0, 1)
				toCheckRange.getCellByPosition(checkReportCol,checkLineIndex).setFormula(CHECK_DONE_TAG)
			End If
			
		End If

	next checkLineIndex
	

End Sub



Function matchingMvtForCheck(matchDate as Date, matchAccount as String, matchDesc as String, matchAmount as Double, infoMatch as String) as Object
    Dim mvtRange As Object
	Dim nCurRow%, nEndRow%, i%
	Dim bestMatch as Object

	' Initialize to null
	matchingMvtForCheck = Nothing
	infoMatch = CHECK_APPROX_MATCH_TAG 

	'MsgBox("search matchingMvtForCheck : " + matchDate + " , " + matchAccount + " , " + matchDesc + " , " + matchAmount)
	mvtRange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells
	
	nCurRow = mvtRange.RangeAddress.StartRow
	nEndRow = mvtRange.RangeAddress.EndRow
	
	For nCurRow = 1 To nEndRow
		Dim mvtCheckStatus as Boolean
		
		'Juste because i know False = 0 and for lazyness around midnight
		mvtCheckStatus = mvtRange.getCellByPosition(mvtVCol, nCurRow).Value + mvtRange.getCellByPosition(mvtXCol, nCurRow).Value
		
		
		' Non checked mvt line
		If (Not mvtCheckStatus) then
			Dim currentDesc, currentAccount as String
			Dim currentAmount as Double
			Dim currentDate as Date
			Dim nBoundDate%
			
			currentDate =  CDate(mvtRange.GetCellByPosition(mvtDateCol, nCurRow).FormulaLocal)
			currentAccount = mvtRange.getCellByPosition(mvtAccountCol, nCurRow).Formula
			currentDesc = mvtRange.getCellByPosition(mvtDescCol, nCurRow).Formula
			currentAmount = mvtRange.getCellByPosition(mvtAmountCol, nCurRow).Value
			
			' If Check, deposit may occurs at a differed date
			nBoundDate = 4
			If ( InStr(currentDesc,CHECK_CHEQUE_TAG) > 0 ) Then
				nBoundDate = 365
			End If
		
					
			If (currentAccount = matchAccount) and (currentAmount = matchAmount) and _
				dateNear(currentDate, matchDate, nBoundDate) Then
				
				'Print "in deep match (account, amount, date similar) : " + nCurRow 
				'Parse word in description
				Dim words(0 to 10) as String
				Dim nStrPos%
				
				'Consider the last match on account, amount and similar date as a potential match
				matchingMvtForCheck = mvtRange.getCellRangeByPosition(0, nCurRow, mvtMaxCol, nCurRow)
				
				words=Split(matchDesc," ")
				For i=0 to Ubound( words() )
					If (Len(words(i)) > 3) Then
						nStrPos = InStr(currentDesc, words(i))
						If (nStrPos > 0) Then
							' Match for sure
							'Print "Found " + words(i) + " in " + currentDesc + " at " + nStrPos 'To write log
							infoMatch = ""
						End If
					End if   
				Next i 
			End If
		
		End If
	Next
End Function


' Move V check to C
Sub MoveMvtVToMvtX
	Dim nCurrentRow%
	Dim mvtRange as Object
	Dim mvtLineIndex%
	
	
	nCurrentRow = ThisComponent.getCurrentController().getSelection().RangeAddress.StartRow
	
	Dim szCurrentAccount
	szCurrentAccount = ThisComponent.sheets.getByName(mvtSheetName).getCellByPosition(mvtAccountCol, nCurrentRow).getFormula()
	
	if ( MsgBox("Confirmer le marquage à X de ligne en V pour le compte " + szCurrentAccount, 4) = 7) Then
		Exit Sub
	End If

	mvtRange = ThisComponent.NamedRanges.getByName(mvtNamedRange).ReferredCells


	For mvtLineIndex = mvtRange.RangeAddress.startRow + 1  To mvtRange.RangeAddress.endRow
		Dim nVVal%, nXVal%
		Dim szAccount as String
		nVVal = mvtRange.getCellByPosition(mvtVCol, mvtLineIndex).Value
		nXVal = mvtRange.getCellByPosition(mvtXCol, mvtLineIndex).Value
		szAccount = mvtRange.getCellByPosition(mvtAccountCol, mvtLineIndex).Formula
		If (nVVal = 1) and (nXVal = 0) and (StrComp(szAccount,szCurrentAccount)=0) Then
			mvtRange.getCellByPosition(mvtVCol, mvtLineIndex).Formula = ""
			mvtRange.getCellByPosition(mvtXCol, mvtLineIndex).Value = 1
			'MsgBox("Line " + mvtLineIndex
		End If
	Next mvtLineIndex
	
	MsgBox("End")

End Sub
