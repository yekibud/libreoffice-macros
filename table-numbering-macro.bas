REM  *****  BASIC  *****

sub TableNumbering

    dim document   as object
    dim dispatcher as object
    ' TODO: CELL_CURSOR_GAP should be determined by the relative default for the gap between rows
    ' - check table properties first to determine gap?
    ' cell cursor gap between rows
    'CELL_CURSOR_GAP = 690
    CELL_CURSOR_GAP = 780
    CELL_PAGE_GAP_START = 6200
    PAGE_SIZE = 28441
    'PAGE_SIZE_TOLERANCE = 1136
    PAGE_SIZE_TOLERANCE = 2000
    'CELL_PAGE_GAP_END = 6900
    ' 6841
    ' TODO: remove all these hard-coded gaps as they can vary depending on font-size, row padding, etc.
    CELL_PAGE_GAP_END = 6800
    document   = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    cursor = ThisComponent.currentController.getViewCursor()
    dim element(0) as new com.sun.star.beans.PropertyValue
    cellValue = 1
    cellName = cursor.cell.cellName
    cursorCorrected = false
    pageNum = cursor.page
    nextPageNum = cursor.page

    do	  
        element(0).Name = "Text"
        element(0).Value = Cstr(cellValue)
        dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, element())
        originalPosition = cursor.position.y
        dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, element())
        nextPosition = cursor.position.y
        
        positionDifference = nextPosition - originalPosition	
        if positionDifference > CELL_CURSOR_GAP then
            dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, element())
            if not cursorCorrected then
                dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, element()
                cursorCorrected = true
            end if
            
            ' don't add a paragraph during page transitions
            ' NOTE: this will create line numbers with blank lines until the bottom of the page if 
            ' there is a gap between the last row on the page and the end of the page
            ' TODO: make this calculation simply a matter of page size since the gap calculation can happen anywhere on the page 
            pagePosition = nextPosition Mod PAGE_SIZE
            if pagePosition < PAGE_SIZE_TOLERANCE and positionDifference > CELL_PAGE_GAP_START and positionDifference < CELL_PAGE_GAP_END then
                dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, element())
            else
            	dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())  
            end if
        end if
        
        nextPageNum = cursor.page
        if nextPageNum <> pageNum then
            cellValue = 1
            pageNum = nextPageNum
        else
            cellValue = cellValue + 1
        end if   
    ' TODO: numbering will end prematurely and cause numbers to be placed in wrong position if there is whitespace or
    ' non-left alignment after table end
    loop until isEmpty(cursor.cell)

end sub
