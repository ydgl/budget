' Documentation pointer on basic openoffice / libreoffice
Rem https://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XCellRange.html
Rem https://wiki.documentfoundation.org/Macros/Calc/f
Rem https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Spreadsheets
Rem https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Editing_Spreadsheet_Documents
Rem http://www.debugpoint.com/libreoffice-basic-macro-tutorial-index/
Rem https://openoffice-libreoffice.developpez.com/tutoriels/openoffice-libreoffice/xray/

Option Explicit



Const G_BACKLOG_SHEETNAME = "Product Backlog"
'     BackLog_UserStory...
CONST G_BL_US_ID_CARD_COL = 0
CONST G_BL_US_NAME_COL = 1
CONST G_BL_US_TYPE_COL = 2
CONST G_BL_US_ESTIMATION_COL = 3
CONST G_BL_US_HOWTO_COL = 4
CONST G_BL_US_NOTE_COL = 5
' Size of header of backlog
CONST G_BL_START_ROW_OFFSET = 0

Const G_CARDS_SHEETNAME = "Cards"		

Const G_TEMPLATE_SHEETNAME = "Template"	
Const G_TPL_ID_CARD = "D2"
Const G_TPL_NAME_CARD = "B2"
Const G_TPL_IMPORTANCE_CARD = "D5"
Const G_TPL_NOTE_CARD = "B5"
Const G_TPL_ESTIMATION_CARD = "D8"
Const G_TPL_HOWTO_CARD = "B8"
Const G_TPL_TYPE_CARD = "B10"
Const G_TPL_RNG_ADDRESS = "A1:F11"
' Number of line for one template
'Const G_TPL_HEIGHT = 11




Const G_LOGSHEET_NAME = "LOG"
Const G_ARRAY_STEP_SIZE = 100

Dim g_blSheet as Object


' Insert in G_LOGSHEET_NAME text at first line
Sub logDebug(text as String)

	if (ThisComponent.Sheets.hasByName(G_LOGSHEET_NAME)) then 
		Dim oSheet As Object
		Dim aLocale as new com.sun.star.lang.Locale, oFormats
			
		oSheet = ThisComponent.Sheets.getByName(G_LOGSHEET_NAME)
		oSheet.Rows.insertByIndex(0,1)
		' TODO : Set date time format
		oSheet.getCellByPosition(0,0).setValue(now())
		oSheet.getCellByPosition(1,0).setFormula(text)
		
		oFormats = ThisComponent.NumberFormats
   		oSheet.getCellByPosition(0,0).NumberFormat = oFormats.getStandardFormat(com.sun.star.util.NumberFormat.DATETIME, aLocale)
		
	End if

End Sub

'--------------------------------------------------------------

Private Sub initBl(ByRef blSheet as Object) 
	g_blSheet = blSheet
End Sub

Public Function blCardId(ByVal rowNb as Integer) as String
	
	blCardId = ""
	if (Not IsEmpty(g_blSheet)) then
		blCardId = g_blSheet.getCellByPosition(G_BL_US_ID_CARD_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	end if
		
End Function

Public Function blCardName(ByVal rowNb as Integer) as String
	
'	blCardName = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_NAME_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	
	blCardName = ""
	if (Not IsEmpty(g_blSheet)) then
		blCardName = g_blSheet.getCellByPosition(G_BL_US_NAME_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	end if
	
End Function


Public Function blCardType(ByVal rowNb as Integer) as String
	
	'blCardType = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_TYPE_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	
	blCardType = ""
	if (Not IsEmpty(g_blSheet)) then
		blCardType = g_blSheet.getCellByPosition(G_BL_US_TYPE_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	end if
	
End Function

Public Function blCardEstimation(ByVal rowNb as Integer) as String
	
	'blCardEstimation = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_ESTIMATION_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	
	blCardEstimation = ""
	if (Not IsEmpty(g_blSheet)) then
		blCardEstimation = g_blSheet.getCellByPosition(G_BL_US_ESTIMATION_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	end if
		
End Function

Public Function blCardHowto(ByVal rowNb as Integer) as String
	
	'blCardHowto = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_HOWTO_COL, G_BL_START_ROW_OFFSET+rowNb).Formula

	blCardHowto = ""
	if (Not IsEmpty(g_blSheet)) then
		blCardHowto = g_blSheet.getCellByPosition(G_BL_US_HOWTO_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	end if
	
End Function

Public Function blCardNote(ByVal rowNb as Integer) as String
	
	'blCardNote = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_NOTE_COL, G_BL_START_ROW_OFFSET+rowNb).Formula

	blCardNote = ""
	if (Not IsEmpty(g_blSheet)) then
		blCardNote = g_blSheet.getCellByPosition(G_BL_US_NOTE_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	end if
		
End Function




'----------------------------------------------

Private Sub InitTemplate(ByRef sheet As Worksheet)
		sheet.getCellRangeByName(G_TPL_ID_CARD).setFormula("ID")
		sheet.getCellRangeByName(G_TPL_NAME_CARD).setFormula("Name / Title")
		sheet.getCellRangeByName(G_TPL_TYPE_CARD).setFormula("Type")
		sheet.getCellRangeByName(G_TPL_IMPORTANCE_CARD).setFormula("Importance")
		sheet.getCellRangeByName(G_TPL_ESTIMATION_CARD).setFormula("Estimation")
		sheet.getCellRangeByName(G_TPL_HOWTO_CARD).setFormula("How To Test")
		sheet.getCellRangeByName(G_TPL_NOTE_CARD).setFormula("Note / Description")
End Sub



Private Sub FillTemplateWithBlRow(blRowIndex as Integer)
            
		Dim oSheet as Object
		
		oSheet = ThisComponent.Sheets.getByName(G_TEMPLATE_SHEETNAME)
		
		oSheet.getCellRangeByName(G_TPL_ID_CARD).setFormula(blCardId(blRowIndex))
		oSheet.getCellRangeByName(G_TPL_NAME_CARD).setFormula(blCardName(blRowIndex))
		oSheet.getCellRangeByName(G_TPL_TYPE_CARD).setFormula(blCardType(blRowIndex))
		oSheet.getCellRangeByName(G_TPL_IMPORTANCE_CARD).setFormula(""+blRowIndex)
		oSheet.getCellRangeByName(G_TPL_ESTIMATION_CARD).setFormula(blCardEstimation(blRowIndex))
		oSheet.getCellRangeByName(G_TPL_HOWTO_CARD).setFormula(blCardHowto(blRowIndex))
		oSheet.getCellRangeByName(G_TPL_NOTE_CARD).setFormula(blCardNote(blRowIndex))

   
End Sub

' Copy template to another sheet
' Note : this way template sheet drive cards sheet formating
Private Sub removeAndCopyTemplateSheet(sSheetName as String)

	If ThisComponent.Sheets().hasByName( sSheetName ) Then
		ThisComponent.Sheets().removeByName( sSheetName)
	End If
	
	' We expect less than 255 sheets
	ThisComponent.Sheets().CopyByName( G_TEMPLATE_SHEETNAME, sSheetName , ThisComponent.Sheets().getCount() )

End Sub

Private Sub CopyAndPasteTemplate(ByRef oSheetSrc As Object, ByRef oSheetDst As Worksheet, nItemNumber As Integer)
  
    Dim srcCellRange As Object
    Dim dstCellRange As Object

	srcCellRange = oSheetSrc.getCellRangeByName(G_TPL_RNG_ADDRESS)
	' srcCellRange.EndRow is the height of each template card
	dstCellRange = oSheetDst.getCellByPosition(0, nItemNumber*(srcCellRange.RangeAddress.EndRow+1))

	oSheetDst.CopyRange(dstCellRange.CellAddress, srcCellRange.RangeAddress)

	' Formating aspect ________________________________________________

	' Insert page break for clean printing
	oSheetDst.Rows(nItemNumber*(srcCellRange.RangeAddress.EndRow+1)).IsStartOfNewPage = true

	' I did not find a way to make paste including formating, we copy height of row too
	Dim n%
	For n = srcCellRange.RangeAddress.StartRow to srcCellRange.RangeAddress.EndRow
		oSheetDst.Rows(nItemNumber*(srcCellRange.RangeAddress.EndRow+1)+n).Height = oSheetSrc.Rows(n).Height
	Next n

End Sub



' Parse current selection and put selected row index in an array
' current selection can be multiple
Private Sub BuildRowIndexArrayFromCurrentSelection(ByRef rowIndexArray)
	Dim nNbRanges%
	Dim n%
	Dim oCurrSel as Object
	Dim vCursor as Object
	Dim oCurrSelPart as Object
	Dim rowIndexTmpArray(G_ARRAY_STEP_SIZE-1)
	Dim nRowIndexArraySize%

	logDebug "BEGIN BuildRowIndexArrayFromCurrentSelection"

	oCurrSel = ThisComponent.getCurrentSelection()


'	If oCurrSel.supportsService("com.sun.star.sheet.SheetCell") Then
'		logDebug "Support com.sun.star.sheet.SheetCell"
'	Else
'		logDebug "NO Support com.sun.star.sheet.SheetCell"
'	
'	End If
'
'	If oCurrSel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
'		logDebug "Support com.sun.star.sheet.SheetCellRanges"
'	Else 
'		logDebug "NO Support com.sun.star.sheet.SheetCellRanges"
'
'	End If
'
'	If oCurrSel.supportsService("com.sun.star.table.CellRange") Then
'		logDebug "Support com.sun.star.table.CellRange"
'	Else 
'		logDebug "NO Support com.sun.star.table.CellRange"
'	End If

	' We store in rowNbArray the selected row(s) number(s) ____________
	' Order will be order of selection
	nRowIndexArraySize = 0
	n = 0
	Do
	
		' If selection is multiple object in oCurrentSel is not the same
		If oCurrSel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
			nNbRanges = oCurrSel.getCount()
			oCurrSelPart = oCurrSel.getByIndex(n).queryIntersection(oCurrSel.getByIndex(n).getRangeAddress())
			
		Else
			' Be Carefull address for one cell is like 'B3', while address for a range is like 'B3:C1'
			' queryIntersection is used to have B3:B3 for one cell case
			' otherwise CreateEnumeration (below) will fail
			nNbRanges = 1
			oCurrSelPart = oCurrSel.queryIntersection(oCurrSel.getRangeAddress())
	
		End If	
	
		vCursor = oCurrSelPart.Cells.CreateEnumeration
		while vCursor.hasMoreElements
			Dim currCell
			currCell = vCursor.NextElement
			rowIndexTmpArray(nRowIndexArraySize) = currCell.RangeAddress.StartRow
			nRowIndexArraySize = nRowIndexArraySize + 1
			' TODO si nRowNbArraySize >  G_ARRAY_STEP_SIZE) alors on sort
		Wend
		
		n = n + 1
		If (n < nNbRanges) Then
			' We are sure here that object support "com.sun.star.sheet.SheetCellRanges"
			oCurrSelPart = oCurrSel.getByIndex(n).queryIntersection(oCurrSel.getByIndex(n).getRangeAddress())
		End If
		
	Loop While (n < nNbRanges)
	
	Redim Preserve rowIndexTmpArray(nRowIndexArraySize-1)

	rowIndexArray = rowIndexTmpArray
	
	logDebug "END BuildRowIndexArrayFromCurrentSelection"
End Sub


Public Sub BuildIndexCardUsingCurrentSelection()

'	Dim cardSheetName = "Cards"
	Dim rowIndexArray()
	Dim n%
	
	logDebug "BGN BuildIndexCardUsingCurrentSelection"
	
	' TODO : Check if current selection is in BL sheet

	initBl(ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME)) 


	BuildRowIndexArrayFromCurrentSelection(rowIndexArray)

	removeAndCopyTemplateSheet(G_CARDS_SHEETNAME)

	For n = 0 to UBound(rowIndexArray)
		logDebug "    BuildIndexCardUsingCurrentSelection on line : " + rowIndexArray(n) 
		FillTemplateWithBlRow(rowIndexArray(n))
		CopyAndPasteTemplate(ThisComponent.Sheets.getByName(G_TEMPLATE_SHEETNAME), ThisComponent.Sheets.getByName(G_CARDS_SHEETNAME),n)
	Next n

	'Leave template clean
	InitTemplate(ThisComponent.Sheets.getByName(G_TEMPLATE_SHEETNAME))


	logDebug "END BuildIndexCardUsingCurrentSelection"
	
End Sub






