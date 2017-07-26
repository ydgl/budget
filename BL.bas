' Documentation pointer on basic openoffice / libreoffice
Rem https://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XCellRange.html
Rem https://wiki.documentfoundation.org/Macros/Calc/f
Rem https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Spreadsheets
Rem https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Editing_Spreadsheet_Documents
Rem http://www.debugpoint.com/libreoffice-basic-macro-tutorial-index/
Rem https://openoffice-libreoffice.developpez.com/tutoriels/openoffice-libreoffice/xray/

Option Explicit



Const G_BACKLOG_SHEETNAME = "Product Backlog"
CONST G_BL_US_ID_CARD_COL = 0
CONST G_BL_US_NAME_COL = 1
CONST G_BL_US_TYPE_COL = 2
CONST G_BL_US_ESTIMATION_COL = 3
CONST G_BL_US_HOWTO_COL = 4
CONST G_BL_US_NOTE_COL = 5
CONST G_BL_START_ROW_OFFSET = 0

Const G_CARDS_SHEETNAME = "CARDS"		

Const G_TEMPLATE_SHEETNAME = "Template"	
Const G_TPL_ID_CARD = "C2"
Const G_TPL_NAME_CARD = "C3"
Const G_TPL_TYPE_CARD = "C9"
Const G_TPL_IMPORTANCE_CARD = "E5"
Const G_TPL_ESTIMATION_CARD = "E8"
Const G_TPL_HOWTO_CARD = "C8"
Const G_TPL_NOTE_CARD = "C5"


Const G_LOGSHEET_NAME = "LOG"
Const G_ARRAY_STEP_SIZE = 100


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



Public Function blCardId(ByVal rowNb as Integer) as String
	
	blCardId = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_ID_CARD_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	
End Function

Public Function blCardName(ByVal rowNb as Integer) as String
	
	blCardName = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_NAME_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	
End Function


Public Function blCardType(ByVal rowNb as Integer) as String
	
	blCardType = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_TYPE_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	
End Function

Public Function blCardEstimation(ByVal rowNb as Integer) as String
	
	blCardEstimation = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_ESTIMATION_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
		
End Function

Public Function blCardHowto(ByVal rowNb as Integer) as String
	
	blCardHowto = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_HOWTO_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	
End Function

Public Function blCardNote(ByVal rowNb as Integer) as String
	
	blCardNote = ThisComponent.Sheets.getByName(G_BACKLOG_SHEETNAME).getCellByPosition(G_BL_US_NOTE_COL, G_BL_START_ROW_OFFSET+rowNb).Formula
	
End Function




'----------------------------------------------



Private Sub RemoveCardSheet()

    On Error GoTo ErrHandler:

    Dim sheet As Object
    Dim sheetFound As Boolean
    sheetFound = True
    Set sheet = Sheets(CARDS_SHEETNAME)
    If (sheetFound = True) Then
        Application.DisplayAlerts = False
        sheet.Delete
        Application.DisplayAlerts = True
    End If

    Exit Sub
ErrHandler:
    ' error handling code
    sheetFound = False

    Resume Next
    
End Sub




Private Sub SetPageBreaks(ByRef targetSheet As Worksheet, ByRef numberOfCards)

    'add the page breaks
    ActiveWindow.View = xlPageBreakPreview
    Application.CutCopyMode = False
    Dim pageBreakIndex As Integer
    Dim i As Integer
    
    For i = 1 To numberOfCards - 1
        
        If (i Mod 2 = 0) Then
            'even
            pageBreakIndex = pageBreakIndex + 1
            
            Dim s As String
            s = Str(1 + (i * 11))
            s = "A" & Trim(s)
            Set targetSheet.HPageBreaks(pageBreakIndex).Location = Range(s)
        End If

    Next
    ActiveWindow.View = xlNormalView

End Sub

Private Sub CopyTemplate(ByRef sheet As Worksheet, ByRef rowIndex As Integer)

    Dim rng As Object
  
    Set rng = sheet.Rows(rowIndex)
    
    rng.PasteSpecial Paste:=8, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

        
    rng.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

End Sub

Private Sub FillTemplateWithBlEntry(blRowNb as Integer)
            
		Dim oSheet as Object
		
		oSheet = ThisComponent.Sheets.getByName(G_TEMPLATE_SHEETNAME)
		
		oSheet.getCellRangeByName(G_TPL_ID_CARD).setFormula(blCardId(blRowNb))
		oSheet.getCellRangeByName(G_TPL_NAME_CARD).setFormula(blCardName(blRowNb))
		oSheet.getCellRangeByName(G_TPL_TYPE_CARD).setFormula(blCardType(blRowNb))
		oSheet.getCellRangeByName(G_TPL_IMPORTANCE_CARD).setFormula(""+blRowNb)
		oSheet.getCellRangeByName(G_TPL_ESTIMATION_CARD).setFormula(blCardEstimation(blRowNb))
		oSheet.getCellRangeByName(G_TPL_HOWTO_CARD).setFormula(blCardHowto(blRowNb))
		oSheet.getCellRangeByName(G_TPL_NOTE_CARD).setFormula(blCardNote(blRowNb))

   
End Sub

  

'Private Function parseSelectionAndGetRowNbArray(oCurrSel as Object) As Integer()

	'parseSelectionAndGetRowsArray
'End Function

Public Sub testSub()

    FillTemplateWithBlEntry(1)

End Sub


Public Sub BuildIndexCardUsingCurrentSelection()

	Dim nNbRanges%
	Dim n%
	Dim s as String
	Dim oCurrSel as Object
	Dim vCursor as Object
	Dim oCurrSelPart as Object
	Dim rowNbArray(G_ARRAY_STEP_SIZE-1)
	Dim nRowNbArraySize%
	
	logDebug "BEGIN BuildIndexCardUsingCurrentSelection"
	
	' Get Sheetname from range
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
	nRowNbArraySize = 0
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
			rowNbArray(nRowNbArraySize) = currCell.RangeAddress.StartRow
			nRowNbArraySize = nRowNbArraySize + 1
			' TODO si nRowNbArraySize >  G_ARRAY_STEP_SIZE) alors on sort
		Wend
		
		n = n + 1
		If (n < nNbRanges) Then
			' We are sure here that object support "com.sun.star.sheet.SheetCellRanges"
			oCurrSelPart = oCurrSel.getByIndex(n).queryIntersection(oCurrSel.getByIndex(n).getRangeAddress())
		End If
		
	Loop While (n < nNbRanges)
	
	' We copy array and to the right size because shortening array ( redim(nRowNbArraySize) ) make the array empty !
	' We copy array to avoid to vehiculate size with array while UBound() exists !
	Dim rowNbArray2(nRowNbArraySize-1)
	For n = 0 to nRowNbArraySize-1
		rowNbArray2(n) = rowNbArray(n) 
	Next n
	


	For n = 0 to UBound(rowNbArray2)
		logDebug "line to print : " + rowNbArray2(n) 
	Next n

'	cards= ReadCards(aRange)
	
	' Build Collection of backlog item ?
	
	
	' Empty sheet destination
	
	' Print Card 1
	' Print PageBreak 1
	' Loop
	
	' Send to printer
	
'	nCurrentRow = ThisComponent.getCurrentController().getSelection().RangeAddress.StartRow
	
'	Dim szCurrentAccount
'	szCurrentAccount = ThisComponent.sheets.getByName(mvtSheetName).getCellByPosition(mvtAccountCol, nCurrentRow).getFormula()

	logDebug "END BuildIndexCardUsingCurrentSelection"
	
End Sub


' TODO
'   4 Decouper le code
'   CANCEL Utiliser un tableau d'ID de card  --> pas besoin
'   2 Fonction de copy du template
'   3 Fonction de vidage du template
'   DONE Fonction d'écriture dans le template (on ecrit dans le template puis on le copie puis on le vide)
'   3 Fonction de vidage de la la liste des card


