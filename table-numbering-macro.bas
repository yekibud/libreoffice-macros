REM  *****  BASIC  *****

sub TableNumbering

    dim document   as object
    dim dispatcher as object
    CELL_CURSOR_GAP = 690
    CELL_PAGE_GAP = 6234
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
            if positionDifference <> CELL_PAGE_GAP then
                dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())
            else
                dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, element())
            end if
        end if
        
        nextPageNum = cursor.page
        if nextPageNum <> pageNum then
            cellValue = 1
            pageNum = nextPageNum
        else
            cellValue = cellValue + 1
        end if   
    loop until isEmpty(cursor.cell)

end sub
