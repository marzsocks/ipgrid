//(C) 2016 Michael Roberg all rights reserved

//----- GLOBALS ------------------------------------------------------------------------------------------------------------------------------------

var ip_GridProps = {}; //new Array();
var defaultRowHeight = 25;    //Default row width
var defaultColWidth = 100;    //Default column width
var loadedRowsThreshold = { min: 40, max: 110 }; //Defines how many row may be directly loaded without loss of performance in the browser. Anything beyond this value will result in slower performance.
var loadedColsThreshold = { min: 15, max: 40 };
var thisBrowser = ip_Browser();

//----- PLUGIN ------------------------------------------------------------------------------------------------------------------------------------

(function ($) {

    //RANGE OPTION 1: /(?=[^"]*$)(([$a-z]+[$0-9]+:[$a-z]+[$0-9]+)|([$a-z]+[$0-9]+))(\#\w*)?/gi
    //RANGE OPTION 2: /(?=[^"]*(?:"[^"]*"[^"]*)*$)(\$?\w\$?\d+(:\$?\w+\$?\d+)?)(\#\w*)?/gi
    
    ip_GridProps['index'] = {
        browser: '',
        focusedGrid: '',
        regEx: {
            range: /(?=[^"]*(?:"[^"]*"[^"]*)*$)(([$a-z]+[$0-9]+:[$a-z]+[$0-9]+)|([$a-z]+[$0-9]+))(\#\w*)?/gi,
            notInBrackets: /(?![^(]*[,)])/gi,
            notInTags: /(?![^<]*[>])/gi,
            notInQuotes: /(?=[^"]*(?:"[^"]*"[^"]*)*$)/gi,
            inQuotes: /("(.*?)")/gi,
            words: /[\w\s"']+/gi,
            nonWords: /[^\w\s]/gi,
            fx: /(\w+)(?=\()/gi,            
            //operator: /([=%*/+-,]+\(\))/gi,
            operator: /\+|,|.|-|\*|\/|=|>|<|>=|<=|&|\||%|!|\^|\(|\)/,
            hashtag: /(\#\w*)/gi,
            

        }
    }
    
    $.fn.ip_Grid = function (options) {

        var options = $.extend({ 

            id: $(this).attr('id'),
            publicKey: null,
            rows: 10,
            cols: 10,
            frozenRows: 0,
            frozenCols: 0,
            showRowSelector: true,
            showColSelector: true,
            showGridResizerX: true,
            rowData: new Array(),
            colData: new Array(),
            mergeData: new Array(),
            dataTypes: new Array(),
            ip_Grid: null, //this is the call back which is executed once the grid is created
            defaultColWidth: defaultColWidth,
            defaultRowHeight: defaultRowHeight,
            loading: false,
            refresh: false,            
            scrollX: null,
            scrollY: null,
            hashTags: [],
            sheetName: '',

            //EVENTS
            onLoad: null
            

        }, options);
         


        //Code starts here ..        
        if (!options.refresh) { ip_UnbindAllEvents(options.id); }
        
        //Record any data  if refreshing
        if (ip_GridProps[options.id] != null) {

            if (options.refresh) {
                if (options.callbacks == null) { options.callbacks = ip_GridProps[options.id].callbacks; }
                if (options.scrollX == null) { options.scrollX = ip_GridProps[options.id].scrollX; }
                if (options.scrollY == null) { options.scrollY = ip_GridProps[options.id].scrollY; }
                if (options.showGridResizerX == null) { options.showGridResizerX = ip_GridProps[options.id].showGridResizerX }
                if (options.frozenRows != ip_GridProps[options.id].frozenRows) { options.scrollY = ( options.scrollY > options.frozenRows ? options.scrollY : options.frozenRows ); }
                if (options.frozenCols != ip_GridProps[options.id].frozenCols) { options.scrollX = ( options.scrollX > options.frozenCols ? options.scrollX : options.frozenCols ); }
            }
            delete ip_GridProps[options.id];

        }

        ip_GridProps['index'].browser = ip_Browser();

        //Setup grid constructer properties
        ip_GridProps[options.id] = $().ip_gridProperties({
            id: options.id,
            publicKey: options.publicKey,
            index:Object.keys(ip_GridProps).length,
            rows: options.rows,
            cols: options.cols,
            scrollX: options.scrollX,
            scrollY: options.scrollY,
            frozenRows: options.frozenRows,
            frozenCols: options.frozenCols,
            showRowSelector: options.showRowSelector,
            showColSelector: options.showColSelector,
            showGridResizerX: options.showGridResizerX,
            rowData: options.rowData,
            colData: options.colData,
            dataTypes: options.dataTypes,
            hashTags: options.hashTags,
            sheetName: options.sheetName,
            undo: { maxTransactions: options.maxUndoTransactions, maxUndoRangeSize: options.maxUndoRangeSize },
            dimensions: {
                defaultColWidth: options.defaultColWidth,
                defaultRowHeight: options.defaultRowHeight
            },
            callbacks: options.callbacks
            
        });

        //validate data
        ip_ValidateData(options.id, options.rows, options.cols, options.loading);

        //render grid
        ip_CreateGrid(options);

        //Execute callback once grid is created    
        if (options.onLoad != null) { ip_GridProps[options.id].callbacks.onLoad = options.onLoad; }
        if (typeof ip_GridProps[options.id].callbacks.onLoad == "function") { ip_GridProps[options.id].callbacks.onLoad(); }

        if (!options.refresh) {

            $('#' + options.id).ip_ReCalculate();
            ip_FocusGrid(options.id, true);

        }

        return $(this);


    };

    $.fn.ip_GridMeta = function (options) {
        //Manages all non UI meta data changes

        var options = $.extend({

            hashTags: null, //[#tag1,#tag2,#tag3]
            parentKeys: null, //[publicKey1,publicKey2,publicKey3]

        }, options);

        var GridID = $(this).attr('id');
        var TransactionID = ip_GenerateTransactionID();
        var Effected = {
            gridData: {
                hashTags: options.hashTags,
                parentKeys: options.parentKeys
            }
        }       

        if (options.hashTags != null) { ip_GridProps[GridID].hashTags = options.hashTags; }
        if (options.parentKeys != null) { ip_GridProps[GridID].parentKeys = options.parentKeys; }

        //Raise resize event
        ip_RaiseEvent(GridID, 'ip_GridMeta', TransactionID, { GridMeta: { Inputs: options, Effected: Effected } });

    };
    
    $.fn.ip_ResizeColumn = function (options) {

        var options = $.extend({

            columns: [-1],        
            size: ip_GridProps[$(this).attr('id')].dimensions.defaultColWidth //int, must be pixels

        }, options);

        var GridID = $(this).attr('id');
        var size = parseInt(options.size);
        var Effected = { colData: new Array() }
        var TransactionID = ip_GenerateTransactionID();
        
        if (jQuery.inArray(-1, options.columns) > -1) {

            //Create an undo stack
            var defaultColWidth = ip_GridProps[GridID].dimensions.defaultColWidth;
            var ResizeRange = { startRow: -1, startCol: -1, endRow: ip_GridProps[GridID].rows - 1, endCol: ip_GridProps[GridID].cols - 1 }
            var ColUndoData = ip_AddUndo(GridID, 'ip_ResizeColumn', TransactionID, 'ColData', ResizeRange, ResizeRange, { row: ResizeRange.startRow, col: ResizeRange.startCol });
            var FunctionUndoData = ip_AddUndo(GridID, 'ip_ResizeColumn', TransactionID, 'function', ResizeRange, null, null, null, function () { ip_GridProps[GridID].dimensions.defaultColWidth = defaultColWidth; });


            //set the default column width for everything and resize all columns
            if (options.size != '') {
                                
                ip_GridProps[GridID].dimensions.defaultColWidth = size;
                Effected.colData[Effected.colData.length] = { col: -1, width: size };

            }

            for (var c = 0; c < ip_GridProps[GridID].cols; c++) {
                                
                ip_AddUndoTransactionData(GridID, ColUndoData, ip_CloneCol(GridID, c));

                if (size == '' || !size || size == NaN) {
                    var calcsize = parseInt($('#' + GridID).ip_MinWidth(c)); //Calculate the minimum width
                    ip_GridProps[GridID].colData[c].width = calcsize;
                    Effected.colData[Effected.colData.length] = { col: c, width: calcsize }; //Set width of column
                }
                else { ip_GridProps[GridID].colData[c].width = size; }

            }

            //Rerender the columns
            ip_ReRenderCols(GridID, 'frozen');
            $(this).ip_Scrollable();
            ip_ReRenderCols(GridID, 'scroll');

        }
        else
        {


            var resizeColumnCount = options.columns.length;

            for (var i = 0; i < resizeColumnCount; i++) {


                var col = parseInt(options.columns[i]);
                var hidden = ip_GridProps[GridID].colData[col].hide == null ? false : ip_GridProps[GridID].colData[col].hide;
                var Quad = (col >= ip_GridProps[GridID].frozenCols ? 2 : 1);

                var TopHeaderCellID = GridID + '_q' + Quad + '_columnSelectorCell_' + col;
                var BottomHeaderCellID = GridID + '_q' + (Quad == 1 ? '3' : '4') + '_columnSelectorCell_' + col;
                var TopHeaderCell = $('#' + TopHeaderCellID);
                var BottomHeaderCell = $('#' + BottomHeaderCellID);
                
                var ResizeRange = { startRow: -1, startCol: col, endRow: ip_GridProps[GridID].rows - 1, endCol: col }
                var ColUndoData = ip_AddUndo(GridID, 'ip_ResizeColumn', TransactionID, 'ColData', ResizeRange, ResizeRange, { row: ResizeRange.startRow, col: ResizeRange.startCol }, null, ip_CloneCol(GridID, col));

                //Set the width of cells in memory
                if (options.size == '') { size = $('#' + GridID).ip_MinWidth(col); }

                var SizeDifference = (hidden ? 0 : size - ip_GridProps[GridID].colData[col].width);
                ip_GridProps[GridID].colData[col].width = size;
                Effected.colData[Effected.colData.length] = { col: col, width: size };

                if (TopHeaderCell.length > 0) {
                               
                    if (col >= ip_GridProps[GridID].frozenCols) { ip_GridProps[GridID].dimensions.accumulativeScrollWidth += SizeDifference; }

                    //Resize the loaded cells
                    $(TopHeaderCell).css('width', ip_ColWidth(GridID, col, false));
                    $(BottomHeaderCell).css('width', ip_ColWidth(GridID, col, false));

                }

            }
            
            $(this).ip_Scrollable();
            ip_OptimizeLoadedCols(GridID); 

        }
                
        ip_ShowColumnFrozenHandle(GridID, true);
        ip_RePoistionRanges(GridID, 'all', true, false);



        //Raise resize event
        ip_RaiseEvent(GridID, 'ip_ResizeColumn', TransactionID, { ResizeColumn: { Inputs: options, Effected: Effected } });

    };

    $.fn.ip_ResizeRow = function (options) {

        var options = $.extend({

            rows: [-1],
            size: ip_GridProps[$(this).attr('id')].dimensions.defaultRowHeight // int, must be pixels

        }, options);

        var GridID = $(this).attr('id');
        var Effected = { rowData: new Array() }
        var size = (options.size == '' ? ip_GridProps[GridID].dimensions.defaultRowHeight : parseInt(options.size));
        var TransactionID = ip_GeneratePublicKey();

        //Resize ALL rows
        if (jQuery.inArray(-1, options.rows) > -1) {            

            //Create an undo stack
            var defaultRowHeight = ip_GridProps[GridID].dimensions.defaultRowHeight;
            var ResizeRange = { startRow: -1, startCol: -1, endRow: ip_GridProps[GridID].rows - 1, endCol: ip_GridProps[GridID].cols - 1 }
            var RowUndoData = ip_AddUndo(GridID, 'ip_ResizeRow', TransactionID, 'RowData', ResizeRange, ResizeRange, { row: ResizeRange.startRow, col: ResizeRange.startCol });
            var FunctionUndoData = ip_AddUndo(GridID, 'ip_ResizeRow', TransactionID, 'function', ResizeRange, null, null, null, function () { ip_GridProps[GridID].dimensions.defaultRowHeight = defaultRowHeight; });

            //set default row height if resizing everything  and resize the rows 
            if (options.size != '') { ip_GridProps[GridID].dimensions.defaultRowHeight = size; }
            
            Effected.rowData[Effected.rowData.length] = { row: -1, height: size };

            for (var i = 0; i < ip_GridProps[GridID].rows; i++) {
                
                ip_AddUndoTransactionData(GridID, RowUndoData, ip_CloneRow(GridID, i));
                if (options.size == '') { ip_GridProps[GridID].rowData[i].height = size; }
                else { ip_GridProps[GridID].rowData[i].height = size; }

            }

            //Rerender the rows
            ip_ReRenderRows(GridID, 'frozen');
            $(this).ip_Scrollable();
            ip_ReRenderRows(GridID, 'scroll');
        }
        else
        {
            //Resize SPECIFIC rows
            var resizeRowCount = options.rows.length;
            for (var i = 0; i < resizeRowCount; i++) {

                
                var row = parseInt(options.rows[i]);
                var hidden = ip_GridProps[GridID].rowData[row].hide == null ? false : ip_GridProps[GridID].rowData[row].hide;
                var SizeDifference = (hidden ? 0 : size - ip_GridProps[GridID].rowData[row].height);
                var Quad = (row >= ip_GridProps[GridID].frozenRows ? 3 : 1);
                var LeftRowID = GridID + '_q' + Quad + '_gridRow_' + row;
                var RightRowID = GridID + '_q' + (Quad == 1 ? '2' : '4') + '_gridRow_' + row;
                var LeftRow = $('#' + LeftRowID);
                var RightRow = $('#' + RightRowID);

                //Add to undo stack
                var ResizeRange = { startRow: row, startCol: -1, endRow: row, endCol: ip_GridProps[GridID].cols - 1 }
                var RowUndoData = ip_AddUndo(GridID, 'ip_ResizeRow', TransactionID, 'RowData', ResizeRange, ResizeRange, { row: ResizeRange.startRow, col: ResizeRange.startCol }, null, ip_CloneRow(GridID, row));

                                
                ip_GridProps[GridID].rowData[row].height = size;
                Effected.rowData[Effected.rowData.length] = { row: row, height: size };

                if (LeftRow.length > 0 && RightRow.length > 0) {

                    if (row >= ip_GridProps[GridID].frozenRows) { ip_GridProps[GridID].dimensions.accumulativeScrollHeight += SizeDifference; }

                    $(LeftRow).css('max-height', ip_RowHeight(GridID, row, false));
                    $(LeftRow).css('height', ip_RowHeight(GridID, row, false));

                    $(RightRow).css('max-height', ip_RowHeight(GridID, row, false));
                    $(RightRow).css('height', ip_RowHeight(GridID, row, false));

                }
                
            }

            $(this).ip_Scrollable();
            ip_OptimzeLoadedRows(GridID); 

        }

        
        ip_ShowRowFrozenHandle(GridID, true);
        ip_RePoistionRanges(GridID, 'all', false, true);
        
        //Raise resize event
        ip_RaiseEvent(GridID, 'ip_ResizeRow', TransactionID, { ResizeRow: { Inputs: options, Effected: Effected } });

    };

    $.fn.ip_ColumnSelector = function (options) {

        var options = $.extend({

            show: false

        }, options);

        var GridID = $(this).attr('id');
        var Effected = { gridData: { showColSelector: ip_GridProps[GridID].showColSelector } }

        if (!options.show) {

            ip_GridProps[GridID].showColSelector = false;
            Effected.gridData.showColSelector = ip_GridProps[GridID].showColSelector;

            $('#' + GridID + '_q1_container th.ip_grid_columnSelectorCellCorner div.ip_grid_cell_outerContent').css('display', 'none');
            $('#' + GridID + '_q1_container th.ip_grid_columnSelectorCellCorner').css('border-top', 'none');

            $('#' + GridID + '_q1_container th.ip_grid_columnSelectorCell div.ip_grid_cell_outerContent').css('display', 'none');
            $('#' + GridID + '_q1_container th.ip_grid_columnSelectorCell').css('border-top', 'none');


            $('#' + GridID + '_q2_container th.ip_grid_columnSelectorCellCorner div.ip_grid_cell_outerContent').css('display', 'none');
            $('#' + GridID + '_q2_container th.ip_grid_columnSelectorCellCorner').css('border-top', 'none');

            $('#' + GridID + '_q2_container th.ip_grid_columnSelectorCell div.ip_grid_cell_outerContent').css('display', 'none');
            $('#' + GridID + '_q2_container th.ip_grid_columnSelectorCell').css('border-top', 'none'); 
            
            $('#' + GridID + '_q1_gridRow_-1').css('height', '');
            $('#' + GridID + '_q2_gridRow_-1').css('height', '');

            if (ip_GridProps[GridID].frozenRows == 0) { $('#' + GridID + '_rowFrozenHandle').hide(); }
            else {  ip_ShowRowFrozenHandle(GridID); }
            $('#' + GridID + '_columnFrozenHandle').hide();

        }
        else {

            ip_GridProps[GridID].showColSelector = true;
            Effected.gridData.showColSelector = ip_GridProps[GridID].showColSelector;

            $('#' + GridID + '_q1_container th.ip_grid_columnSelectorCellCorner div.ip_grid_cell_outerContent').css('display', '');
            $('#' + GridID + '_q1_container th.ip_grid_columnSelectorCellCorner').css('border-top', '');

            $('#' + GridID + '_q1_container th.ip_grid_columnSelectorCell div.ip_grid_cell_outerContent').css('display', '');
            $('#' + GridID + '_q1_container th.ip_grid_columnSelectorCell').css('border-top', '');

            $('#' + GridID + '_q2_container th.ip_grid_columnSelectorCellCorner div.ip_grid_cell_outerContent').css('display', '');
            $('#' + GridID + '_q2_container th.ip_grid_columnSelectorCellCorner').css('border-top', '');

            $('#' + GridID + '_q2_container th.ip_grid_columnSelectorCell div.ip_grid_cell_outerContent').css('display', '');
            $('#' + GridID + '_q2_container th.ip_grid_columnSelectorCell').css('border-top', '');

            $('#' + GridID + '_q1_gridRow_-1').css('height', ip_GridProps[GridID].dimensions.columnSelectorHeight + 'px');
            $('#' + GridID + '_q2_gridRow_-1').css('height', ip_GridProps[GridID].dimensions.columnSelectorHeight + 'px');

            ip_ShowRowFrozenHandle(GridID);
            ip_ShowColumnFrozenHandle(GridID);
        
        }
            
        $(this).ip_Scrollable();
        ip_OptimzeLoadedRows(GridID);
        ip_RePoistionRanges(GridID, 'all', false, true);

        //Raise resize event
        ip_RaiseEvent(GridID, 'ip_ColumnSelector', null, { ColumnSelector: { Inputs: options, Effected: Effected } });

            
    };

    $.fn.ip_RowSelector = function (options) {

        var options = $.extend({

            show: false

        }, options);

        var GridID = $(this).attr('id');
        var Effected = { gridData: { showRowSelector: ip_GridProps[GridID].showRowSelector } }

        if (!options.show) {

            ip_GridProps[GridID].showRowSelector = false;
            Effected.gridData.showRowSelector = ip_GridProps[GridID].showRowSelector;

            $('#' + GridID + '_q1_columnSelectorCell_-1').css('width', '0px');
            $('#' + GridID + '_q1_columnSelectorCell_-1').css('border-left', 'none');

            $('#' + GridID + '_q3_columnSelectorCell_-1').css('width', '0px');
            $('#' + GridID + '_q3_columnSelectorCell_-1').css('border-left', 'none');

            $('#' + GridID + '_rowFrozenHandle').hide();
            if (ip_GridProps[GridID].frozenCols == 0) { $('#' + GridID + '_columnFrozenHandle').hide(); }
            else { ip_ShowColumnFrozenHandle(GridID); }

        }
        else {

            ip_GridProps[GridID].showRowSelector = true;
            Effected.gridData.showRowSelector = ip_GridProps[GridID].showRowSelector;

            $('#' + GridID + '_q1_columnSelectorCell_-1').css('width', '');
            $('#' + GridID + '_q1_columnSelectorCell_-1').css('border-left', '');

            $('#' + GridID + '_q3_columnSelectorCell_-1').css('width', '');
            $('#' + GridID + '_q3_columnSelectorCell_-1').css('border-left', '');

            ip_ShowRowFrozenHandle(GridID);
            ip_ShowColumnFrozenHandle(GridID); 
            
        }

        $(this).ip_Scrollable();
        ip_RePoistionRanges(GridID, 'all', true, false);

        //Raise resize event
        ip_RaiseEvent(GridID, 'ip_RowSelector', null, { RowSelector: { Inputs: options, Effected: Effected } });
    };

    $.fn.ip_CellHtml = function (row,col, returnLastIfNull) {
        
        
        var GridID = $(this).attr('id');

        return ip_CellHtml(GridID, row, col, returnLastIfNull);

    };

    $.fn.ip_CellData = function (row, col) {

        var GridID = $(this).attr('id');

        return ip_CellData(GridID, row, col);

    };

    $.fn.ip_MinWidth = function (col) {

        var MinWidth = 0;
        var GridID = $(this).attr('id');
        var value = '';
        var style = '';
        
        for (var r = 0; r < ip_GridProps[GridID].rows; r++) {

            var cellvalue = (ip_GridProps[GridID].rowData[r].cells[col].display == null ? '' : ip_GridProps[GridID].rowData[r].cells[col].display);
            
            if (cellvalue.length > MinWidth) {

                value = cellvalue;
                MinWidth = value.length;
                style = (ip_GridProps[GridID].rowData[r].cells[col].style == null ? '' : ip_GridProps[GridID].rowData[r].cells[col].style);
            }

        }

        if (value != '') {
            $('#' + GridID + '_cellContentWidthTool').html(value);
            if (style != '') { $('#' + GridID + '_cellContentWidthTool').attr('style', 'display:none;' + style); }
            
            //STILL NEED TO DO THIS:
            //$('#' + GridID + '_cellContentWidthTool').css('font-size', '');
            //$('#' + GridID + '_cellContentWidthTool').css('font-family', '');
            MinWidth = $('#' + GridID + '_cellContentWidthTool').width();
        }
        else {
            MinWidth = ip_GridProps[GridID].dimensions.defaultColWidth;
        }

        return parseInt(MinWidth);
    }

    $.fn.ip_Scrollable = function (options) {

        var options = $.extend({

            scrollbarX: true,
            scrollbarY: true,
            scrollSpeed: 0.01,
            placementX: 'container',
            placementY: 'container',
            initX: true,
            initY: true

        }, options);

        var GridID = $(this).attr('id');
        var result = true;

        ip_SetQuadContainerWidth(GridID);

        if (result && options.initX) { result = ip_InitScrollX_Large(GridID, options.scrollSpeed, options.scrollbarX, options.scrollbarY, options.placementX, options.placementY); }
        if (result && options.initY) { result = ip_InitScrollY_Large(GridID, options.scrollSpeed, options.scrollbarY); }

        return result;
    }

    $.fn.ip_ScrollTo = function (options) {

        var options = $.extend({

            row: null,
            col: null

        }, options);

        var GridID = $(this).attr('id');

        //Scroll to a column
        if (options.col != null) {
            
            var col = options.col;

            if (col < ip_GridProps[GridID].frozenCols) { col = ip_GridProps[GridID].frozenCols; }
            else if (col > ip_GridProps[GridID].cols) { col = ip_GridProps[GridID].cols - 1; }

            ip_ScrollToX(GridID, col, true, 'all', 'none');

        }

        //Scroll to a row
        if (options.row != null) {

            var row = options.row;

            if (row < ip_GridProps[GridID].frozenRows) { row = ip_GridProps[GridID].frozenRows; }
            else if (row > ip_GridProps[GridID].rows) { row = ip_GridProps[GridID].rows - 1; }

            ip_ScrollToY(GridID, row, true, 'all', 'none');
        }
    }

    $.fn.ip_SelectCell = function (options) {
        //SELECTS A CELL, if not cell options are specified (cell, row, col) will select a row instead
        var options = $.extend({

            row: null,
            col: null,
            cell: null,
            multiselect: false,
            scrollIncrement: 0,
            scrollTo: false,
            direction: '',
            selectAsRange: true,
            raiseEvent: true,
            debug:false

        }, options);

        var GridID = $(this).attr('id');
        var cell = null;
        var TransactionID = ip_GenerateTransactionID();
        
        //Check if we should select a column instead
        if (ip_GridProps[GridID].selectedCell == null && options.cell == null && (options.row == null || options.col == null || options.row == -1 || options.col == -1)) {

            if (ip_GridProps[GridID].selectedColumn.length > 0 || options.col == -1) { $('#' + GridID).ip_SelectColumn({ col: (options.col != null ? options.col : ip_GridProps[GridID].selectedColumn[0]), count: ip_GridProps[GridID].selectedColumn.length, unselect: (options.col != null ? true : false) }); return; }
            if (ip_GridProps[GridID].selectedRow.length > 0 || options.row == -1) { $('#' + GridID).ip_SelectColumn({ row: (options.row != null ? options.row : ip_GridProps[GridID].selectedRow[0]), count: ip_GridProps[GridID].selectedRow.length, unselect: (options.row != null ? true : false) }); return; }
            return;
        }

        if (!options.multiselect) {
            $(this).ip_UnselectColumn();
            $(this).ip_UnselectRow();
            $(this).ip_RemoveRange();
        }

        if (options.row == null) { options.row = (ip_GridProps[GridID].selectedCell == null ? 0 : ip_GridProps[GridID].selectedCell.row); }
        if (options.col == null) { options.col = (ip_GridProps[GridID].selectedCell == null ? 0 : ip_GridProps[GridID].selectedCell.col); }
        
        if (options.row < 0) { options.row = 0; }
        if (options.row >= ip_GridProps[GridID].rows) { options.row = ip_GridProps[GridID].rows - 1; }
        if (options.col < 0) { options.col = 0; }
        if (options.col >= ip_GridProps[GridID].cols) { options.col = ip_GridProps[GridID].cols - 1; }
        


        var MergedCell = ip_GetNextMergedCell(GridID, options.row, options.col, options.direction);
        options.row = MergedCell.row;
        options.col = MergedCell.col;

        if (options.scrollIncrement != 0) {
            options.row = ip_NextNonHiddenRow(GridID, options.row, null, options.row, (options.scrollIncrement > 0 ? 'down' : 'up'));
            options.col = ip_NextNonHiddenCol(GridID, options.col, null, options.col, (options.scrollIncrement > 0 ? 'right' : 'left'));
        }

        //Scroll to cell options
        if (options.scrollTo == true) { 
            if (!ip_IsRowLoaded(GridID, options.row)) { ip_ScrollToY(GridID, options.row - Math.ceil(ip_GridProps[GridID].loadedRows / 2), true); }
            if (!ip_IsColLoaded(GridID, options.col)) { ip_ScrollToX(GridID, options.col - Math.ceil(ip_GridProps[GridID].loadedCols / 2), true); }
        }
        if (options.scrollIncrement > 0) { ip_ScrollCell(GridID, options.direction, options.row - 1, options.col, null, null, options.scrollIncrement, 'ip_grid_cell_rangeHighlight'); } //'ip_grid_cell_rangeHighlight'
        else if (options.scrollIncrement < 0) { ip_ScrollCell(GridID, options.direction, options.row, options.col, null, null, options.scrollIncrement, 'ip_grid_cell_rangeHighlight'); }  //ip_grid_cell_rangeHighlight
        
        if (options.cell != null) { cell = options.cell; options.row = parseInt($(cell).attr('row')); options.col = parseInt($(cell).attr('col')); } //use the specified cell
        else { cell = ip_CellHtml(GridID, options.row, options.col, true); } //find the cell based row/col values
        
        if(!options.multiselect || ip_GridProps[GridID].selectedRange.length == 0){

            //We can only have one actively selected sell
            ip_GridProps[GridID].selectedCell = cell;
            ip_GridProps[GridID].selectedCell.row = options.row;
            ip_GridProps[GridID].selectedCell.col = options.col;
            
            ip_SetFxBarValue(GridID, { cell: cell });
            $(this).ip_RangeHighlight({ startCell: cell, multiselect: options.multiselect });

        
        }

        if (options.selectAsRange) { $(this).ip_SelectRange({ startCell: cell, multiselect: options.multiselect }); }

        //Set focus to grid - importaint for events to be focused on grid
        $('#' + GridID).focus();
        
        if (options.raiseEvent) {
            var Effected = { row: options.row, col: options.col, cell: options.cell }
            ip_RaiseEvent(GridID, 'ip_SelectCell', TransactionID, { SelectCell: { Inputs: options, Effected: Effected } });
        }
    }
    
    $.fn.ip_UnSelectCell = function (options) {
        var options = $.extend({

            //No options becausewe can only have one selected cell
        }, options);

        var GridID = $(this).attr('id');

        ip_GridProps[GridID].selectedCell = null;        
        $(this).ip_RemoveRangeHighlight({ highlightType: 'ip_grid_cell_rangeHighlight_activecell' });

    }

    $.fn.ip_SelectColumn = function (options) {

        var options = $.extend({

            col: -1,
            count:1,
            unselect: true,
            multiselect: false,
            fetchRange: true,
            rangeType: null,
            raiseEvent: true

        }, options);

        var GridID = $(this).attr('id');
        var col = options.col;
        var count = options.count;
        var Effected = { col: [] }
        var TransactionID = ip_GenerateTransactionID();

        if (col >= -1 && col < ip_GridProps[GridID].cols) {

            //if (col == -1) { col = 0; count = ip_GridProps[GridID].cols; }

            //Validate merges
            var ColMerge = ip_GetColumnMerge(GridID, col);
            count += col - ColMerge.mergedWithCol;
            if (count < ColMerge.colSpan) { count = ColMerge.colSpan; }
            col = ColMerge.mergedWithCol;

            var IsAllSelected = jQuery.inArray(-1, ip_GridProps[GridID].selectedColumn);
            var IsSelected = jQuery.inArray(col, ip_GridProps[GridID].selectedColumn);

            //Clear column selection
            if (!options.multiselect || IsAllSelected != -1) {

                $(this).ip_UnselectRow();
                $(this).ip_UnselectColumn();
                $(this).ip_RemoveRange();
                $(this).ip_UnSelectCell();
                IsAllSelected = -1;
            }

            if (col != -1) {
                for (var i = 0; i < count; i++) {

                    var selCol = col + i;

                    if (IsSelected == -1 || options.unselect == false) {

                        //this loop allows us to select merged columns

                        var Quad = (selCol >= ip_GridProps[GridID].frozenCols ? 2 : 1);

                        //add column to selection
                        Effected.col[Effected.col.length] = selCol;
                        ip_GridProps[GridID].colData[selCol].selected = true;
                        ip_GridProps[GridID].selectedColumn[ip_GridProps[GridID].selectedColumn.length] = parseInt((selCol));
                        if (ip_IsColLoaded(GridID, (selCol))) { $('#' + GridID + '_q' + Quad + '_columnSelectorCell_' + (selCol)).addClass('selected'); }


                    }
                    else {
                        //remove column from selection if it is already selected
                        $('#' + GridID).ip_UnselectColumn({ col: selCol });
                    }

                }
            }
            else {

                //Select all
                Effected.col[0] = -1;
                ip_GridProps[GridID].selectedColumn = new Array();
                ip_GridProps[GridID].selectedColumn[0] = -1;
                IsAllSelected = 0;
                for (var c = 0; c < ip_GridProps[GridID].cols; c++) {

                    var Quad = (c >= ip_GridProps[GridID].frozenCols ? 2 : 1);
                    ip_GridProps[GridID].colData[c].selected = true;
                    if (ip_IsColLoaded(GridID, (c))) { $('#' + GridID + '_q' + Quad + '_columnSelectorCell_' + (c)).addClass('selected'); }
                    
                }
                
            }

            if (IsAllSelected == -1) {


                //Sort and Remove duplicates
                ip_GridProps[GridID].selectedColumn.sort(function sortNumber(a, b) { return a - b; });

                var prevC = -2;
                for (var c = ip_GridProps[GridID].selectedColumn.length - 1; c >= 0; c--) {
                    if (prevC == ip_GridProps[GridID].selectedColumn[c]) { ip_GridProps[GridID].selectedColumn.splice(c, 1); }
                    prevC = ip_GridProps[GridID].selectedColumn[c];
                }

                //Show the column as selected
                if (ip_GridProps[GridID].selectedColumn.length > 0 && options.fetchRange) {

                    //calculate how to select ranges
                    var ranges = [];
                    var prevCol = -2;
                    for (var c = 0; c < ip_GridProps[GridID].selectedColumn.length; c++) {

                        var col = ip_GridProps[GridID].selectedColumn[c];
                        var diff = col - prevCol;

                        if (diff > 1) { ranges[ranges.length] = { startRow: -1, startCol: col, endRow: ip_GridProps[GridID].rows - 1, endCol: col } }
                        else { ranges[ranges.length - 1].endCol = col; }

                        prevCol = col;
                    }

                    for (var r = 0; r < ranges.length; r++) { $(this).ip_SelectRange({ multiselect: (r > 0 ? true : false), range: ranges[r], rangeType: options.rangeType }); }

                }

            }
            else {
                //select all
                var range = { startRow: -1, startCol: 0, endRow: ip_GridProps[GridID].rows - 1, endCol: ip_GridProps[GridID].cols - 1 };
                $(this).ip_SelectRange({ multiselect: false, range: range, rangeType: options.rangeType });
            }

            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_SelectColumn', TransactionID, { SelectColumn: { Inputs: options, Effected: Effected } }); }

        }
       
    }
       
    $.fn.ip_SelectRow = function (options) {

        var options = $.extend({

            row: -1,
            count: 1,
            unselect:true,
            multiselect: false,
            fetchRange: true,
            rangeType: null,
            raiseEvent: true,

        }, options);

        var GridID = $(this).attr('id');
        var row = options.row;
        var count = options.count;
        var TransactionID = ip_GenerateTransactionID();
        var Effected = { row: [] }

        if (row >= -1 && row < ip_GridProps[GridID].rows) {

            //Validate merges
            var RowMerge = ip_GetRowMerge(GridID, row);
            count += row - RowMerge.mergedWithRow;
            if (count < RowMerge.rowSpan) { count = RowMerge.rowSpan; }
            row = RowMerge.mergedWithRow;

            var IsAllSelected = jQuery.inArray(-1, ip_GridProps[GridID].selectedRow);
            var IsSelected = jQuery.inArray(row, ip_GridProps[GridID].selectedRow);

            //Clear column selection
            if (!options.multiselect || IsAllSelected != -1) {

                $(this).ip_UnselectRow();
                $(this).ip_UnselectColumn();
                $(this).ip_RemoveRange();
                $(this).ip_UnSelectCell();
                IsAllSelected = -1;
            }   

            //if (row == -1) { row = 0; count = ip_GridProps[GridID].rows; }
            
            if (row != -1) {
                for (var i = 0; i < count; i++) {

                    var selRow = row + i;

                    if (IsSelected == -1 || options.unselect == false) {

                        //this loop allows us to select merged columns
                        var Quad = (selRow >= ip_GridProps[GridID].frozenRows ? 3 : 1);

                        //add column to selection
                        Effected.row[Effected.row.length] = selRow;
                        ip_GridProps[GridID].rowData[selRow].selected = true;
                        ip_GridProps[GridID].selectedRow[ip_GridProps[GridID].selectedRow.length] = parseInt((selRow));
                        if (ip_IsRowLoaded(GridID, (selRow))) { $('#' + GridID + '_q' + Quad + '_rowSelecterCell_' + (selRow)).addClass('selected'); }

                    }
                    else {
                        //remove column from selection if it is already selected
                        $('#' + GridID).ip_UnselectRow({ row: selRow });
                    }

                }
            }
            else {

                //Select all
                ip_GridProps[GridID].selectedRow = new Array();
                ip_GridProps[GridID].selectedRow[0] = -1;
                IsAllSelected = 0;
                Effected.row[0] = -1;

                for (var r = 0; r < ip_GridProps[GridID].rows; r++) {

                    var Quad = (r >= ip_GridProps[GridID].frozenRows ? 3 : 1);
                    ip_GridProps[GridID].rowData[r].selected = true;
                    if (ip_IsRowLoaded(GridID, (r))) { $('#' + GridID + '_q' + Quad + '_rowSelecterCell_' + (r)).addClass('selected'); }

                }

            }



            if (IsAllSelected == -1) {

                //Sort and Remove duplicates
                ip_GridProps[GridID].selectedRow.sort(function sortNumber(a, b) { return a - b; });

                var prevR = -2;
                for (var r = ip_GridProps[GridID].selectedRow.length - 1; r >= 0; r--) {
                    if (prevR == ip_GridProps[GridID].selectedRow[r]) { ip_GridProps[GridID].selectedRow.splice(r, 1); }
                    prevR = ip_GridProps[GridID].selectedRow[r];
                }

                //Show the column as selected
                if (ip_GridProps[GridID].selectedRow.length > 0 && options.fetchRange) {

                    //calculate how to select ranges
                    var ranges = [];
                    var prevRow = -2;
                    for (var r = 0; r < ip_GridProps[GridID].selectedRow.length; r++) {

                        var row = ip_GridProps[GridID].selectedRow[r];
                        var diff = row - prevRow;

                        if (diff > 1) { ranges[ranges.length] = { startRow: row, startCol: -1, endRow: row, endCol: ip_GridProps[GridID].cols - 1 } }
                        else { ranges[ranges.length - 1].endRow = row; }

                        prevRow = row;
                    }

                    for (var r = 0; r < ranges.length; r++) { $(this).ip_SelectRange({ multiselect: (r > 0 ? true : false), range: ranges[r], rangeType: options.rangeType }); }

                }
            }
            else {
                //select all
                var range = { startRow: 0, startCol: 0, endRow: ip_GridProps[GridID].rows - 1, endCol: ip_GridProps[GridID].cols - 1 };
                $(this).ip_SelectRange({ multiselect: false, range: range, rangeType: options.rangeType });
            }

            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_SelectRow', TransactionID, { SelectRow: { Inputs: options, Effected: Effected } }); }


        }

    }

    $.fn.ip_UnselectColumn = function (options) {
        var options = $.extend({

            col: -1,

        }, options);

        var GridID = $(this).attr('id');
        var col = options.col;
        var Quad = (col >= ip_GridProps[GridID].frozenCols ? 2 : 1);
        var IsSelected = ip_GridProps[GridID].selectedColumn.indexOf(col); //jQuery.inArray(col, ip_GridProps[GridID].selectedColumn);



        if (col == -1) {


            //remove visual effect on selected column
            $('#' + GridID + ' .ip_grid_columnSelectorCell').removeClass('selected');
            $('#' + GridID + ' .ip_grid_columnSelectorCellCorner').removeClass('selected');

            ////Remove ranges for column
            for (var c = 0; c < ip_GridProps[GridID].colData.length; c++) { ip_GridProps[GridID].colData[c].selected = false;  }

            //Clear all column selections
            ip_GridProps[GridID].selectedColumn = new Array();

        } else if (IsSelected > -1) {

            ip_GridProps[GridID].colData[col].selected = false;

            //Clear specific column selections            
            ip_GridProps[GridID].selectedColumn.splice(IsSelected, 1);
            $('#' + GridID + '_q' + Quad + '_columnSelectorCell_' + col).removeClass('selected');

        }
    }

    $.fn.ip_UnselectRow = function (options) {
        var options = $.extend({

            row: -1,
            

        }, options);

        var GridID = $(this).attr('id');
        var row = options.row;
        var Quad = (row >= ip_GridProps[GridID].frozenRows ? 3 : 1);
        var IsSelected = ip_GridProps[GridID].selectedRow.indexOf(row);


        if (row == -1) {

            //Clear all column selections           
            $('#' + GridID + ' .ip_grid_rowSelecterCell').removeClass('selected');
            
            //Remove ranges for row
            for (var r = 0; r < ip_GridProps[GridID].rowData.length; r++) { ip_GridProps[GridID].rowData[r].selected = false;  }

            ip_GridProps[GridID].selectedRow = new Array();

        } else if (IsSelected > -1) {

            ip_GridProps[GridID].rowData[row].selected = false;

            //Clear specific column selections
            ip_GridProps[GridID].selectedRow.splice(IsSelected, 1);
            $('#' + GridID + '_q' + Quad + '_rowSelecterCell_' + row).removeClass('selected');

        }

    }
    
    $.fn.ip_RangeHighlight = function (options) {

        var options = $.extend({

            range: { startRow: null, startCol: null, endRow: null, endCol: null },
            startCellOrdinates: [0, 0], //[row,col]
            endCellOrdinates: null, //[row,col]
            startCell: null, //td object
            endCell: null, //td object
            multiselect: false,
            borderStyle:'',
            borderColor: '',
            fillColor: null,
            opacity: null,
            highlightType: 'ip_grid_cell_rangeHighlight_activecell', //css class defining the highlight type
            fadeIn: false,
            fadeOut: false,
            expireTimeout:0 //Removes the highlight after x miliseconds

        }, options);

        var GridID = $(this).attr('id');
        var startCell = null;
        var endCell = null;

        if (ip_GridProps[GridID].rows == 0 || ip_GridProps[GridID].cols == 0) { return false; }        

        //clear any previously selected range
        if (!options.multiselect) { $(this).ip_RemoveRangeHighlight({ highlightType: options.highlightType });   }
        
        options.considerMerges = false;
        options = ip_ValidateRangeOptions(GridID, options);

        if (!options) { return false; }

        startCell = options.startCell;
        
        if (startCell != null) {
            
            endCell = options.endCell;

            var RangeHighlight = $('#' + GridID + '_rangeHighlight').clone(); //Clone the range selector tool (so that we can add multiple range selectors)
            var CellQuad = $(startCell).parent().parent().parent().parent().parent(); //choose the cells container div so we know where to add the range selector tool
            var RangeHighlightID = ip_setRangeHighlightID(GridID, RangeHighlight, options.startCellOrdinates, options.endCellOrdinates);

            $(RangeHighlight).appendTo(CellQuad);

            var NewTop = ip_CalculateRangeTop(GridID, RangeHighlight, startCell, options.range.startRow, options.range.startCol, options.range.endRow, options.range.endCol); //, (options.mergedStartRow != null ? options.mergedStartRow : options.startCellOrdinates[0, 0]), (options.mergedStartCol != null ? options.mergedStartCol : options.startCellOrdinates[0, 1])
            var NewLeft = ip_CalculateRangeLeft(GridID, RangeHighlight, startCell, options.range.startRow, options.range.startCol, options.range.endRow, options.range.endCol); //, (options.mergedStartRow != null ? options.mergedStartRow : options.startCellOrdinates[0, 0]), (options.mergedStartCol != null ? options.mergedStartCol : options.startCellOrdinates[0, 1])

            $(RangeHighlight).width($(startCell).width());
            $(RangeHighlight).height($(startCell).height());

            $(RangeHighlight).css('top', NewTop + 'px');
            $(RangeHighlight).css('left', NewLeft + 'px');
            if (options.borderStyle != '') { $(RangeHighlight).css('border-style', options.borderStyle); }
            if (options.borderColor != '') { $(RangeHighlight).css('border-color', options.borderColor); }
            if (options.fillColor != '') { $(RangeHighlight).css('background-color', options.fillColor); }
            if (options.opacity != '') { $(RangeHighlight).css('opacity', options.opacity); }
            $(RangeHighlight).addClass('ip_grid_cell_rangeHighlight_selected');
            $(RangeHighlight).addClass(options.highlightType);

            if (options.fadeIn) { $(RangeHighlight).fadeIn(); }
            else { $(RangeHighlight).show(); }

            //Select full range (if multiple cells are covered in range)
            if (endCell != startCell) {

                var NewWidth = ip_SetRangeWidth(GridID, RangeHighlight, startCell, endCell, options.startCellOrdinates, options.endCellOrdinates);
                var NewHeight = ip_SetRangeHeight(GridID, RangeHighlight, startCell, endCell, options.startCellOrdinates, options.endCellOrdinates);

                $(RangeHighlight).width(NewWidth);
                $(RangeHighlight).height(NewHeight);

            }

            var rangeIndex = ip_GridProps[GridID].highlightRange.length;
            ip_GridProps[GridID].highlightRange[rangeIndex] = new Array([options.startCellOrdinates[0, 0], options.startCellOrdinates[0, 1]], [options.endCellOrdinates[0, 0], options.endCellOrdinates[0, 1]], options.borderStyle, options.borderColor, options.highlightType, options.expireTimeout);

            if (options.expireTimeout != 0) {
                setTimeout(function () {
                    $('#' + GridID).ip_RemoveRangeHighlight({ fadeOut: options.fadeOut, rangeHighlightElement: RangeHighlight });
                }, options.expireTimeout);
            }
        }

    }

    $.fn.ip_RemoveRangeHighlight = function (options) {

        var options = $.extend({

            rangeHighlightElement: null, //Optional, physical visual range object (div)
            rangeHighlightID: null, //Optional, ID for visual range object (div), e.g. MyGrid_range_7_3_15_6
            highlightType: null, //CSS class defining the hilight to remove
            fadeOut: false
        }, options);


        var GridID = $(this).attr('id');
        var RangeHighlight = null;

        if (options.rangeHighlightElement != null) { RangeHighlight = options.rangeHighlightElement; }
        else if (options.rangeHighlightID != null) { RangeHighlight = $('#' + options.rangeHighlightID); }
        else if (options.highlightType != null) { RangeHighlight = $('#' + GridID + ' .' + options.highlightType); }

        if (RangeHighlight == null) {

            //remove all active ranges
            if (options.fadeOut) { $(RangeHighlight).fadeOut(function () { $('.ip_grid_cell_rangeHighlight_selected').remove(); }); }
            else { $('.ip_grid_cell_rangeHighlight_selected').remove(); }

            ip_GridProps[GridID].highlightRange = new Array();

        }
        else {

            if (RangeHighlight.length > 0) {

                 ////remove a specific range
                if (options.fadeOut) {
                    $(RangeHighlight).fadeOut(function () { $(RangeHighlight).remove();  });
                }
                else { $(RangeHighlight).remove(); }

                for (var i = 0; i < RangeHighlight.length; i++) {

                    var index = ip_GetRangeHighlightIndex(GridID, RangeHighlight[i]);
                    if (index != -1) { ip_GridProps[GridID].highlightRange.splice(index, 1); }
                }
            }
        }
    }

    $.fn.ip_SelectRange = function (options) {

        var options = $.extend({

            range: { startRow: null, startCol: null, endRow: null, endCol: null },
            startCellOrdinates : [0,0], //[row,col]
            endCellOrdinates: null, //[row,col]
            startCell: null, //td object
            endCell: null, //td object
            multiselect: false,
            allowColumn: true,
            allowRow: true,
            showContextButton: false,
            showKey:true,
            rangeType: null,
            allowMove: true,
            fadeIn: false,
            raiseEvent:true

        }, options);
                

        var TransactionID = ip_GeneratePublicKey();
        var GridID = $(this).attr('id');
        var startCell = null;
        var endCell = null;
        
        if (ip_GridProps[GridID].rows == 0 || ip_GridProps[GridID].cols == 0) { return false; }

        //clear any previously selected range
        if (!options.multiselect) { $(this).ip_RemoveRange(); }
        if (!options.allowColumn) { $(this).ip_UnselectColumn(); }
        if (!options.allowRow) { $(this).ip_UnselectRow(); }
            
            


        //We validate merge's inside nomalize range options
        options = ip_ValidateRangeOptions(GridID, options);
        if (!options) { return false; }

        startCell = options.startCell;

        if (startCell != null) {

            //Get range's end cell
            endCell = options.endCell;

            var Range = $('#' + GridID + '_rangeselector').clone(); //Clone the range selector tool (so that we can add multiple range selectors)
            var CellQuad = ($(startCell).parent().parent().parent().parent().parent())[0]; //choose the cells container div so we know where to add the range selector tool
            var CellSelectorKey = $(Range).children('.ip_grid_cell_rangeselector_key');
            var RangeBorder = $(Range).children('.ip_grid_cell_rangeselector_border');
            var RangeID = ip_setRangeID(GridID, Range, options.startCellOrdinates, options.endCellOrdinates);

            if (document.getElementById(RangeID) == null) {


                //Add range selector tool to grid quadrant
                $(Range).appendTo(CellQuad);

                var NewTop = ip_CalculateRangeTop(GridID, Range, startCell, options.range.startRow, options.range.startCol, options.range.endRow, options.range.endCol); //, (options.mergedStartRow != null ? options.mergedStartRow : options.startCellOrdinates[0, 0]), (options.mergedStartCol != null ? options.mergedStartCol : options.startCellOrdinates[0, 1])
                var NewLeft = ip_CalculateRangeLeft(GridID, Range, startCell, options.range.startRow, options.range.startCol, options.range.endRow, options.range.endCol); //, (options.mergedStartRow != null ? options.mergedStartRow : options.startCellOrdinates[0, 0]), (options.mergedStartCol != null ? options.mergedStartCol : options.startCellOrdinates[0, 1])

                $(Range).width($(startCell).width());
                $(Range).height($(startCell).height());

                $(Range).css('top', NewTop + 'px');
                $(Range).css('left', NewLeft + 'px');
                $(Range).addClass('ip_grid_cell_rangeselector_selected');
                $(Range).addClass(options.rangeType);
                if (!options.showKey) { $(CellSelectorKey).hide(); }
                if (!options.allowMove) { $(RangeBorder).remove(); }

                if (options.fadeIn) { $(Range).fadeIn(); }
                else { $(Range).css('display', 'block'); }


                //Select full range (if multiple cells are covered in range)
                if (endCell != startCell) {
                                        
                    var NewWidth = ip_SetRangeWidth(GridID, Range, startCell, endCell, options.startCellOrdinates, options.endCellOrdinates);
                    var NewHeight = ip_SetRangeHeight(GridID, Range, startCell, endCell, options.startCellOrdinates, options.endCellOrdinates);

                    $(Range).width(NewWidth);
                    $(Range).height(NewHeight);

                }


                //Add range to grid properties
                var rangeIndex = ip_GridProps[GridID].selectedRange.length;
                ip_GridProps[GridID].selectedRange[rangeIndex] = new Array([options.startCellOrdinates[0, 0], options.startCellOrdinates[0, 1]], [options.endCellOrdinates[0, 0], options.endCellOrdinates[0, 1]], [options.startCellOrdinates[0, 0], options.startCellOrdinates[0, 1]], [NewTop,NewLeft]);
                ip_GridProps[GridID].selectedRangeIndex = rangeIndex;
  
                var rangeBorder_MouseDown = null;
                $(RangeBorder).mousedown(rangeBorder_MouseDown = function (e) {

                    if (e.which == 1 && !ip_GridProps[GridID].resizing) {

                        ip_ShowRangeMove(GridID, Range, e, this);
                        if (!ip_GridProps[GridID].scrollAnimate) { ip_EnableScrollAnimate(GridID); } //This is now handleded in the mouseenter event for the cell                       


                    }

                });
 
                //Add handler for mouse move range selection                
                var reloadRanges_MouseDown = null;
                $(CellSelectorKey).mousedown(reloadRanges_MouseDown = function () {

                    ip_GridProps[GridID].resizing = true;

                    var MinWidth = $(startCell).width();
                    var MinHeight = $(startCell).height();

                    ip_GridProps[GridID].selectedRangeIndex = rangeIndex;

                    $(Range).addClass('ip_grid_cell_rangeselector_Resize');
                    $(Range).children('.ip_grid_cell_rangeselector_border').hide();

                    ip_EnableScrollAnimate(GridID);

                    var reloadRanges_MouseMove = null;
                    $('#' + GridID).mousemove(reloadRanges_MouseMove = function (e) {

                        ip_CalculateRangeDimensions(GridID, Range, rangeIndex, e, 'both', MinHeight, MinWidth);

                    });


                    var reloadRanges_MouseUp = null;
                    $(document).mouseup(reloadRanges_MouseUp = function (e) {

                        
                        reloadRanges_MouseMove = ip_UnBindEvent('#' + GridID, 'mousemove', reloadRanges_MouseMove);
                        reloadRanges_MouseUp = ip_UnBindEvent(document, 'mouseup', reloadRanges_MouseUp);
                        
                        ip_SetHoverCell(GridID, 'ip_grid_cell_rangeselector', e);
                        $(Range).children('.ip_grid_cell_rangeselector_border').show();

                        var startRange = { startRow: ip_GridProps[GridID].selectedRange[rangeIndex][0][0, 0], startCol: ip_GridProps[GridID].selectedRange[rangeIndex][0][0, 1], endRow: ip_GridProps[GridID].selectedRange[rangeIndex][1][0, 0], endCol: ip_GridProps[GridID].selectedRange[rangeIndex][1][0, 1] }
                        var endRange = ip_ChangeRange(GridID, ip_GridProps[GridID].selectedRange[rangeIndex], ip_GridProps[GridID].hoverCell, 0, 0 , 0, true);

                        ip_DragRange(GridID, { range: startRange, dragToRange: endRange });

                        ip_GridProps[GridID].resizing = false;
                    });

                });

                if (options.raiseEvent) {
                    var Effected = { startRow: options.startCellOrdinates[0, 0], startCol: options.startCellOrdinates[0, 1], endRow: options.endCellOrdinates[0, 0], endCol: options.endCellOrdinates[0, 1] }
                    ip_RaiseEvent(GridID, 'ip_SelectRange', TransactionID, { SelectRange: { Inputs: options, Effected: Effected } });
                }

            }
            else {
                //remove range            
                $(this).ip_RemoveRange({ rangeID: RangeID });
            }
        }

    }
       
    $.fn.ip_RemoveRange = function (options) {

        var options = $.extend({

            rangeElement: null, //Optional, physical visual range object (div)
            rangeID: null //Optional, ID for visual range object (div), e.g. MyGrid_range_7_3_15_6
    
        }, options);

        var GridID = $(this).attr('id');
        var Range = null;

        if (options.rangeElement != null) { Range = options.rangeElement; }
        else if (options.rangeID != null) { Range = $('#' + options.rangeID); }

        if (Range == null) {
            
            //remove all active ranges
            $('.ip_grid_cell_rangeselector_selected').remove();
            ip_GridProps[GridID].selectedRange = new Array();

        }
        else {

            //remove a specific range
            $(Range).remove();
            var i = ip_GetRangeIndex(GridID, Range);
            if (i != -1) { ip_GridProps[GridID].selectedRange.splice(i, 1); }

        }
    }

    $.fn.ip_ResetRange = function (options) {

        var options = $.extend({

            range: null, //[{ startRow: null, startCol: null, endRow: null, endCol: null }] array of range object
            preserveRange: null, //[{ startRow: null, startCol: null, endRow: null, endCol: null }] array of range object
            valuesOnly: false,

        }, options);

        var GridID = $(this).attr('id');
        var arrRange = new Array(); //{ startRow: null, startCol: null, endRow: null, endCol: null, cut: false }
        var fullReset = false;
        var TransactionID = ip_GenerateTransactionID();
        var Effected = { rangeData: arrRange, preserveRangeData: options.preserveRange, resetGrid: fullReset, rowData: null, valuesOnly: options.valuesOnly };

        //Validate and up range
        if (options.range) {
            arrRange = options.range;
            for (var i = 0; i < options.range.length; i++) {

                if (arrRange[i].startRow < 0) { arrRange[i].startRow = 0; }
                if (arrRange[i].startCol < 0) { arrRange[i].startCol = 0; }
                if (arrRange[i].endRow >= ip_GridProps[GridID].rows) { arrRange[i].endRow = ip_GridProps[GridID].rows - 1; }
                if (arrRange[i].endCol >= ip_GridProps[GridID].cols) { arrRange[i].endCol = ip_GridProps[GridID].cols - 1; }

            }
        }       
        else {

            for (var i = 0; i < ip_GridProps[GridID].selectedRange.length; i++) {
                arrRange[i] = {
                    startRow: ip_GridProps[GridID].selectedRange[i][0][0],
                    startCol: ip_GridProps[GridID].selectedRange[i][0][1],
                    endRow: ip_GridProps[GridID].selectedRange[i][1][0],
                    endCol: ip_GridProps[GridID].selectedRange[i][1][1]
                }

                if (arrRange[i].startRow < 0) { arrRange[i].startRow = 0; }
                if (arrRange[i].startCol < 0) { arrRange[i].startCol = 0; }
                if (arrRange[i].endRow >= ip_GridProps[GridID].rows) { arrRange[i].endRow = ip_GridProps[GridID].rows - 1; }
                if (arrRange[i].endCol >= ip_GridProps[GridID].cols) { arrRange[i].endCol = ip_GridProps[GridID].cols - 1; }
            }
        }

        //Do delete action
        for (i = 0; i < arrRange.length; i++) {
            
            if (arrRange[i].startRow == 0 && arrRange[i].endRow == ip_GridProps[GridID].rows - 1 && arrRange[i].startCol == 0 && arrRange[i].endCol == ip_GridProps[GridID].cols - 1) { fullReset = true; }

            var CellUndoData = ip_AddUndo(GridID, 'ip_ResetRange', TransactionID, 'CellData', arrRange[i], arrRange[i], (i == 0 ? { row: arrRange[i].startRow, col: arrRange[i].startCol } : null));
            var MergeUndoData = ip_AddUndo(GridID, 'ip_ResetRange', TransactionID, 'MergeData', arrRange[i]);
            var merges = ip_ValidateRangeMergedCells(GridID, arrRange[i].startRow, arrRange[i].startCol, arrRange[i].endRow, arrRange[i].endCol);


            for (r = arrRange[i].startRow; r <= arrRange[i].endRow; r++) {

                for (c = arrRange[i].startCol; c <= arrRange[i].endCol; c++) {
                    
                    var Ignore = false;


                    if (options.preserveRange != null) {
                        for (ir = 0; ir < options.preserveRange.length; ir++) {
                            if (r >= options.preserveRange[ir].startRow && r <= options.preserveRange[ir].endRow && c >= options.preserveRange[ir].startCol && c <= options.preserveRange[ir].endCol) { Ignore = true; }
                        }
                    }

                    
                    if (!Ignore) {

                        //Add to undo stack
                        ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, r, c));

                        //Reset the cell
                        if (options.valuesOnly) {
                            ip_SetValue(GridID, r, c, null);
                            ip_RemoveCellFormulaIndex(GridID, { row: r, col: c });
                            ip_GridProps[GridID].rowData[r].cells[c].formula = null;
                            ip_GridProps[GridID].rowData[r].cells[c].error = null;
                        }
                        else { ip_ResetCell(GridID, r, c, true, true); }

                    }

                }
            }

            //Remove merges that dont overlap
            if (!options.valuesOnly) {
                for (var m = 0; m < merges.merges.length; m++) {
                    if (!merges.merges[m].containsOverlap) {

                        //Add to undo stack
                        ip_AddUndoTransactionData(GridID, MergeUndoData, ip_mergeObject(merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol, merges.merges[m].rowSpan, merges.merges[m].colSpan));

                        //Reset the cell merge
                        ip_ResetCellMerge(GridID, merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol);

                    }
                }
            }

        }

        Effected.rowData = ip_ReCalculateFormulas(GridID, { range: arrRange, transactionID: TransactionID, render: true, raiseEvent: false }).Effected.rowData;

        //Remove cut object from copied cells
        for (i = ip_GridProps[GridID].copiedRange.length - 1 ; i >= 0 ; i--) {

            if (ip_GridProps[GridID].copiedRange[i][2] == 'cut') {
                ip_GridProps[GridID].copiedRange.splice(i, 1);
            }

        }

        $(this).ip_RemoveRangeHighlight({ highlightType: 'ip_grid_cell_rangeHighlight_cut' });
        ip_SetFxBarValue(GridID); 
        ip_ReRenderRanges(GridID, arrRange);

        //Raise event                            
        ip_RaiseEvent(GridID, 'ip_ResetRange', TransactionID, { ResetRange: { Inputs: options, Effected: Effected } });

    }

    $.fn.ip_RemoveRow = function (options) {
        
        var GridID = $(this).attr('id');

        ip_RemoveRow(GridID, options);

    }

    $.fn.ip_AddRow = function (options) {
               
        var GridID = $(this).attr('id');
        ip_AddRow(GridID, options);              

    }

    $.fn.ip_InsertRow = function (options) {

        var GridID = $(this).attr('id');
        ip_InsertRow(GridID, options);

    }

    $.fn.ip_RemoveCol = function (options) {


        var GridID = $(this).attr('id');

        ip_RemoveCol(GridID, options);

    }

    $.fn.ip_AddCol = function (options) {
        
        var GridID = $(this).attr('id');
        ip_AddCol(GridID, options);

    }

    $.fn.ip_InsertCol = function (options) {

        var GridID = $(this).attr('id');
        ip_InsertCol(GridID, options);
        
    }
    
    $.fn.ip_FrozenRowsCols = function (options) {

        var options = $.extend({

            rows: null, //integer, amount of frozen rows to set
            cols: null,  //integer, amount of frozen cols to set
            raiseEvent: true,
            createUndo: true            

        }, options);

        var GridID = $(this).attr('id');
        var error = '';
        var Effected = { gridData: { frozenRows: options.rows, frozenCols: options.cols } };
        var TransactionID = ip_GenerateTransactionID();

        rows = options.rows;
        cols = options.cols;

        if (rows != null) {

            //Validate rows
            if (rows < 0) { rows = 0; }
            if (rows > ip_GridProps[GridID].loadedRows) {

                rows = ip_GridProps[GridID].loadedRows - 1;
                ip_RaiseEvent(GridID, 'warning',arguments.callee.caller, 'Your frozen rows may not exceed the viewable area of rows 0 - ' + rows + '');
                return false;

            }

            var validateMerges = ip_ValidateRangeMergedCells(GridID, rows, 0, rows, ip_GridProps[GridID].cols - 1);
            for (var i = 0; i < validateMerges.merges.length; i++) {
                if (validateMerges.merges[i].mergedWithRow != rows) {
                    $(this).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, color: '#ff5a00', range: { startRow: validateMerges.merges[i].mergedWithRow, startCol: validateMerges.merges[i].mergedWithCol, endRow: validateMerges.merges[i].mergedWithRow, endCol: validateMerges.merges[i].mergedWithCol } });
                    error = 'Cant freeze rows over merged cells, first unmerge your cells'
                }
            }

            if (error == '') {

                //Add to undo stack
                if (options.createUndo) {
                    var frozenRows = ip_GridProps[GridID].frozenRows;
                    var FunctionUndoData = ip_AddUndo(GridID, 'ip_FrozenRowsCols', TransactionID, 'function', null, null, null, null, function () { $('#' + GridID).ip_FrozenRowsCols({ rows: frozenRows, raiseEvent: false, createUndo: false }); });
                }

                var scrollDiff = rows - ip_GridProps[GridID].frozenRows;

                
                ip_GridProps[GridID].frozenRows = rows;
                ip_GridProps[GridID].scrollY += scrollDiff;

                ip_ReRenderRows(GridID, 'frozen'); 
            }


        }

        if (cols != null) {


            //Validate rows
            if (cols < 0) { cols = 0; }
            if (cols > ip_GridProps[GridID].loadedCols) {

                cols = ip_GridProps[GridID].loadedCols - 1;
                ip_RaiseEvent(GridID, 'warning',arguments.callee.caller, 'Your frozen columns may not exceed the viewable area of columns 0 - ' + cols + '');
                return false;
                
            }

            //Validate merges
            var validateMerges = ip_ValidateRangeMergedCells(GridID, 0, cols, ip_GridProps[GridID].rows - 1, cols );
            for (var i = 0; i < validateMerges.merges.length; i++) {
                if (validateMerges.merges[i].mergedWithCol != cols) {
                    $(this).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, color: '#ff5a00', range: { startRow: validateMerges.merges[i].mergedWithRow, startCol: validateMerges.merges[i].mergedWithCol, endRow: validateMerges.merges[i].mergedWithRow, endCol: validateMerges.merges[i].mergedWithCol } });
                    error = 'Cant freeze columns over merged cells, first unmerge your cells'
                }
            }
                        
            if (error == '') {

                //Add to undo stack
                if (options.createUndo) {
                    var frozenCols = ip_GridProps[GridID].frozenCols;
                    var FunctionUndoData = ip_AddUndo(GridID, 'ip_FrozenRowsCols', TransactionID, 'function', null, null, null, null, function () { $('#' + GridID).ip_FrozenRowsCols({ cols: frozenCols, raiseEvent: false, createUndo: false }); });
                }

                var frozenFiff = cols - ip_GridProps[GridID].frozenCols;
                var scrollDiff = cols - ip_GridProps[GridID].frozenCols;                


                ip_GridProps[GridID].frozenCols = cols;
                ip_GridProps[GridID].scrollX += scrollDiff;
                

                ip_ReRenderCols(GridID, 'frozen');
               

            }
        }

        if (error != '') { ip_RaiseEvent(GridID, 'warning',arguments.callee.caller, error); }
        else {



            $('#' + GridID).ip_Scrollable();

            if (cols != null) { ip_ReRenderCols(GridID, 'scroll'); }
            if (rows != null) { ip_ReRenderRows(GridID, 'scroll'); }

            ip_ReloadRanges(GridID);


            //Raise event            
            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_FrozenRowsCols', TransactionID, { FrozenRowsCols: { Inputs: options, Effected: Effected } }); }
        }

        
    }

    $.fn.ip_MoveCol = function (options) {

        var GridID = $(this).attr('id');
        return ip_MoveColumn(GridID, options);

    }

    $.fn.ip_MoveRow = function (options) {

        var GridID = $(this).attr('id');
        return ip_MoveRow(GridID, options);
    }

    $.fn.ip_Copy = function (options) {

        var options = $.extend({

            range: null, //[{ startRow: null, startCol: null, endRow: null, endCol: null }] array of range object
            rangeElement: null, //[div] array of actual range elemnt
            rangeID: null, //[ID] array of actual range element ID
            cut: false,
            toClipBoard: true

        }, options);

        var GridID = $(this).attr('id');
        var arrRange = new Array(); //{ startRow: null, startCol: null, endRow: null, endCol: null }

        //Load up range
        if (options.range) {
            arrRange = options.range;
        }
        else if (options.rangeID) {

            for (var i = 0; i < options.rangeID.length; i++) {
                arrRange[i] = {
                    startRow: parseInt('#' + $(options.rangeID[i]).attr('startrow')),
                    startCol: parseInt('#' + $(options.rangeID[i]).attr('startcol')),
                    endRow: parseInt('#' + $(options.rangeID[i]).attr('endrow')),
                    endCol: parseInt('#' + $(options.rangeID[i]).attr('endcol'))
                }
            }

        }
        else {

            if (!options.rangeElement) { options.rangeElement = $('#' + GridID + ' .ip_grid_cell_rangeselector_selected'); }

            for (var i = 0; i < options.rangeElement.length; i++) {
                arrRange[i] = {
                    startRow: parseInt($(options.rangeElement[i]).attr('startrow')),
                    startCol: parseInt($(options.rangeElement[i]).attr('startcol')),
                    endRow: parseInt($(options.rangeElement[i]).attr('endrow')),
                    endCol: parseInt($(options.rangeElement[i]).attr('endcol'))
                }
            }

        }

        var ClipBoardText = '';

        if (options.cut) { ip_GridProps[GridID].cutRange = new Array(); } //[[startrow,startcol],[endrow,endcol]]
        ip_GridProps[GridID].copiedRange = new Array(); //[[startrow,startcol],[endrow,endcol]]

        $('#' + GridID).ip_RemoveRangeHighlight();

        //Copy range to clipboard
        for (var i = 0; i < arrRange.length; i++) {
            
            //Validate the range
            if (arrRange[i].startRow >= ip_GridProps[GridID].rows) { arrRange[i].startRow = ip_GridProps[GridID].rows - 1; }
            if (arrRange[i].startCol >= ip_GridProps[GridID].cols) { arrRange[i].startCol = ip_GridProps[GridID].cols - 1; }
            if (arrRange[i].endRow >= ip_GridProps[GridID].rows) { arrRange[i].endRow = ip_GridProps[GridID].rows - 1; }
            if (arrRange[i].endCol >= ip_GridProps[GridID].cols) { arrRange[i].endCol = ip_GridProps[GridID].cols - 1; }


            //if (options.toClipBoard) { ip_SelectTextTool(GridID, arrRange[i].startRow, arrRange[i].startCol, arrRange[i].endRow, arrRange[i].endCol, (i == 0 ? false : true)); }

            var RangeID = GridID + '_range_' + arrRange[i].startRow + '_' + arrRange[i].startCol + '_' + arrRange[i].endRow + '_' + arrRange[i].endCol;
            
            //Copy the range

            ip_GridProps[GridID].copiedRange[i] = new Array([arrRange[i].startRow, arrRange[i].startCol], [arrRange[i].endRow, arrRange[i].endCol], (options.cut ? 'cut' : 'copy'));
            

            $('#' + GridID).ip_RangeHighlight({ range: arrRange[i], multiselect: true, fadeIn: true, highlightType: (options.cut ? 'ip_grid_cell_rangeHighlight_cut' : 'ip_grid_cell_rangeHighlight_copy') });

            //if (options.toClipBoard) { ip_CopyToClipboard(GridID); }

            ip_GridProps[GridID].cut = options.cut;
        }

        



    }

    $.fn.ip_Paste = function (options) {

        fxid_ip_Paste: 

        var options = $.extend({

            row: null, //int, Starting row to paste into
            col: null, //int, Starting col to paste into
            fromRange: null, //[{ startRow: null, startCol: null, endRow: null, endCol: null }] array of range object
            fromClipboard: false,
            cut: ip_GridProps[$(this).attr('id')].cut,
            changeFormula: true

        }, options);

        var GridID = $(this).attr('id');      
        var TransactionID = ip_GenerateTransactionID();
        var Error = '';

        //VALIDATE OPTIONS
        if (options.row == null && ip_GridProps[GridID].selectedRow.length > 0) { ip_GridProps[GridID].selectedRow.sort(); options.row = ip_GridProps[GridID].selectedRow[0]; options.col = 0; }
        if (options.col == null && ip_GridProps[GridID].selectedColumn.length > 0) { ip_GridProps[GridID].selectedColumn.sort(); options.row = 0; options.col = ip_GridProps[GridID].selectedColumn[0]; }
        if (options.row == null && ip_GridProps[GridID].selectedCell != null) { options.row = ip_GridProps[GridID].selectedCell.row; } else if (options.row == null) { Error = 'Row to paste to is not specified' }
        if (options.col == null && ip_GridProps[GridID].selectedCell != null) { options.col = ip_GridProps[GridID].selectedCell.col; } else if (options.col == null) { Error = 'Column to paste to is not specified' }
        if (options.row < 0) { options.row = 0; }
        if (options.col < 0) { options.col = 0; }
        if (options.row >= ip_GridProps[GridID].rows) { options.row = ip_GridProps[GridID].rows - 1; }
        if (options.col >= ip_GridProps[GridID].cols) { options.col = ip_GridProps[GridID].cols - 1; }

        if (options.fromClipboard) { }
        else if (options.fromRange == null) {

            options.fromRange = new Array();
            for (var i = 0; i < ip_GridProps[GridID].copiedRange.length; i++) { options.fromRange[i] = ip_rangeObject(ip_GridProps[GridID].copiedRange[i][0][0], ip_GridProps[GridID].copiedRange[i][0][1], ip_GridProps[GridID].copiedRange[i][1][0], ip_GridProps[GridID].copiedRange[i][1][1]);  }

        }
        
        //FROM RANGE VALIDATION
        if (options.fromRange != null)
        {

            var arrFromRange = new Array();
            var arrToRange = new Array();
            var rowDiff = 0; //options.row - options.fromRange[0].startRow;
            var colDiff = 0; //options.col - options.fromRange[0].startCol;

            for (var i = 0; i < options.fromRange.length; i++) {

                arrFromRange[i] = ip_rangeObject(options.fromRange[i].startRow, options.fromRange[i].startCol, options.fromRange[i].endRow, options.fromRange[i].endCol);

                //VALIDATE FROM RANGE AREA
                if (arrFromRange[i].startRow < 0) { arrFromRange[i].startRow = 0; }
                if (arrFromRange[i].startCol < 0) { arrFromRange[i].startCol = 0; }
                if (arrFromRange[i].endRow >= ip_GridProps[GridID].rows) { arrFromRange[i].endRow = ip_GridProps[GridID].rows - 1; }
                if (arrFromRange[i].endCol >= ip_GridProps[GridID].cols) { arrFromRange[i].endCol = ip_GridProps[GridID].cols - 1; }
                

                rowDiff = options.row - arrFromRange[0].startRow;
                colDiff = options.col - arrFromRange[0].startCol;


                arrToRange[i] = ip_rangeObject(arrFromRange[i].startRow + rowDiff, arrFromRange[i].startCol + colDiff, arrFromRange[i].endRow + rowDiff, arrFromRange[i].endCol + colDiff);


                //VALIDATE TO RANGE AREA
                if (arrToRange[i].startRow < 0) { arrToRange[i].startRow = 0; }
                if (arrToRange[i].startCol < 0) { arrToRange[i].startCol = 0; }
                if (arrToRange[i].endRow >= ip_GridProps[GridID].rows) { Error = 'You are trying to paste cells beyond the sheet area'; $(this).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, color: '#ff5a00', range: { startRow: arrFromRange[i].startRow, startCol: arrFromRange[i].startCol, endRow: arrFromRange[i].endRow, endCol: arrFromRange[i].endCol } }); }
                if (arrToRange[i].endCol >= ip_GridProps[GridID].cols) { Error = 'You are trying to paste cells beyond the sheet area'; $(this).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, color: '#ff5a00', range: { startRow: arrFromRange[i].startRow, startCol: arrFromRange[i].startCol, endRow: arrFromRange[i].endRow, endCol: arrFromRange[i].endCol } }); }
                
                if (Error == '') {

                    arrToRange[i].existingMerges = ip_ValidateRangeMergedCells(GridID, arrToRange[i].startRow, arrToRange[i].startCol, arrToRange[i].endRow, arrToRange[i].endCol);


                    //VALIDATE THE PASTE DOES NOT OVERLAP MERGES IN THE TO RANGE AREA
                    if (arrToRange[i].existingMerges.containsOverlap) {
                        for (var m = 0; m < arrToRange[i].existingMerges.merges.length; m++) {
                                                        
                            if (!options.cut || ( arrToRange[i].existingMerges.merges[m].containsOverlap && !ip_IsMergeInRange(GridID, arrToRange[i].existingMerges.merges[m], arrFromRange))) { 
                                $(this).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, color: '#ff5a00', range: { startRow: arrToRange[i].existingMerges.merges[m].mergedWithRow, startCol: arrToRange[i].existingMerges.merges[m].mergedWithCol } }); 
                                Error = 'When placing cells, make sure they dont partcially overlap an existing merge';
                            }

                        }
                    }
                }



            }

        }

        //ACTUAL PASTE PROCESSING STARTS HERE
        if (Error == '') {


            if (options.fromClipboard) { var Effected = { rowData: [], mergeData: new Array(), mergeDataForRemove: new Array(), cut: options.cut }; }
            else {

                //Grab the from values and store them
                var Effected = { rowData:[],  row: options.row, col: options.col, fromRange: arrFromRange, cut: options.cut };
                var fromRangeData = new Array();
                for (var i = 0; i < arrFromRange.length; i++) { fromRangeData[i] = ip_GetRangeData(GridID, false, arrFromRange[i].startRow, arrFromRange[i].startCol, arrFromRange[i].endRow, arrFromRange[i].endCol); }


                //Clear existing cell data
                if (options.cut) {

   
                    for (var i = 0; i < fromRangeData.length; i++) {

                        //Create an undo stack
                        var CellUndoData = ip_AddUndo(GridID, 'ip_Paste', TransactionID, 'CellData', fromRangeData[i]);
                        var MergeUndoData = ip_AddUndo(GridID, 'ip_Paste', TransactionID, 'MergeData', fromRangeData[i]);
                        
                        //Reset cells
                        for (var r = 0; r < fromRangeData[i].rowData.length; r++) {
                            for (var c = 0; c < fromRangeData[i].rowData[r].cells.length; c++) {

                                //Add data to undo stack
                                ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, fromRangeData[i].rowData[r].cells[c].row, fromRangeData[i].rowData[r].cells[c].col));                                
                                ip_ResetCell(GridID, fromRangeData[i].rowData[r].cells[c].row, fromRangeData[i].rowData[r].cells[c].col, true, true);

                            }
                        }

                        //Remove merges if they dont overlap
                        for (var m = 0; m < fromRangeData[i].mergeData.length; m++) {
                            if (!fromRangeData[i].mergeData[m].containsOverlap) {

                                //Record undo data for merge
                                ip_AddUndoTransactionData(GridID, MergeUndoData, ip_mergeObject(fromRangeData[i].mergeData[m].mergedWithRow, fromRangeData[i].mergeData[m].mergedWithCol, fromRangeData[i].mergeData[m].rowSpan, fromRangeData[i].mergeData[m].colSpan));                                
                                ip_ResetCellMerge(GridID, fromRangeData[i].mergeData[m].mergedWithRow, fromRangeData[i].mergeData[m].mergedWithCol);

                            }
                        } 
                    }
                }

                //Paste values in new position
                for (var i = 0; i < arrToRange.length; i++) {

                    //Create an undo stack
                    var CellUndoData = ip_AddUndo(GridID, 'ip_Paste', TransactionID, 'CellData', arrToRange[i], arrToRange[i], (i == 0 ? { row: arrToRange[i].startRow, col: arrToRange[i].startCol } : null));
                    var MergeUndoData = ip_AddUndo(GridID, 'ip_Paste', TransactionID, 'MergeData', arrToRange[i]); 

                    //Remove merges in the to-range area
                    for (var m = 0; m < arrToRange[i].existingMerges.merges.length; m++) {

                        //Add data to undo stack
                        ip_AddUndoTransactionData(GridID, MergeUndoData, ip_mergeObject(arrToRange[i].existingMerges.merges[m].mergedWithRow, arrToRange[i].existingMerges.merges[m].mergedWithCol, arrToRange[i].existingMerges.merges[m].rowSpan, arrToRange[i].existingMerges.merges[m].colSpan));                        
                        ip_ResetCellMerge(GridID, arrToRange[i].existingMerges.merges[m].mergedWithRow, arrToRange[i].existingMerges.merges[m].mergedWithCol);

                    }

                    var rIndex = 0;

                    for (var r = arrToRange[i].startRow; r <= arrToRange[i].endRow; r++) {

                        var cIndex = 0;

                        for (var c = arrToRange[i].startCol; c <= arrToRange[i].endCol; c++) {

                            //Add data to undo stack
                            ip_AddUndoTransactionData(GridID,CellUndoData,ip_CloneCell(GridID, r, c));
                            
                            //Paste the origonal values,  removing merges (which will get added later)
                            var fromCell = fromRangeData[i].rowData[rIndex].cells[cIndex];
                            var fxIndex = ip_GridProps[GridID].rowData[r].cells[c].fxIndex;

                            ip_GridProps[GridID].rowData[r].cells[c] = fromCell;
                            ip_GridProps[GridID].rowData[r].cells[c].fxIndex = fxIndex; //set the old index so it can be properly removed in SetCellFormula method                            

                            //Update this cells formula
                            ip_SetCellFormula(GridID, { row: r, col: c, formula: (options.changeFormula ? ip_MoveFormulaOrigon(GridID, ip_GridProps[GridID].rowData[r].cells[c].formula, fromCell.row, fromCell.col, r, c, null, (options.cut ? arrFromRange[i] : null)) : ip_GridProps[GridID].rowData[r].cells[c].formula) }); 

                            //Update formulas linked to this cell
                            if (options.changeFormula && options.cut) { Effected.rowData = Effected.rowData.concat(ip_ChangeFormulasForCellMove(GridID, CellUndoData, fromCell.row, fromCell.col, r, c, fromRangeData[i]).rowData); }

                            ip_SetValue(GridID, r, c);
                            
                            delete ip_GridProps[GridID].rowData[r].cells[c].merge; //Remove any merge data as we need to set it manually at the end of processing

                            cIndex++;

                        }

                        rIndex++;
                    }
                }

                //Set new merges
                for (var i = 0; i < fromRangeData.length; i++) {
                    for (var m = 0; m < fromRangeData[i].mergeDataContained.length; m++) {

                        var StartRow = fromRangeData[i].mergeDataContained[m].mergedWithRow + rowDiff;
                        var StartCol = fromRangeData[i].mergeDataContained[m].mergedWithCol + colDiff;
                        var EndRow = StartRow + fromRangeData[i].mergeDataContained[m].rowSpan - 1;
                        var EndCol = StartCol + fromRangeData[i].mergeDataContained[m].colSpan - 1;

                        ip_SetCellMerge(GridID, StartRow, StartCol, EndRow, EndCol);

                    }
                }

            }
                        
            //Effected.rowData = Effected.rowData.concat(ip_ReCalculateFormulas(GridID, { createUndo: true, range: (options.cut ? arrToRange.concat(arrFromRange) : arrToRange), transactionID: TransactionID, render: true, raiseEvent: false }).Effected.rowData);
            if (options.cut) { Effected.rowData = Effected.rowData.concat(ip_ReCalculateFormulas(GridID, { createUndo: true, range: arrFromRange, transactionID: TransactionID, render: true, raiseEvent: false }).Effected.rowData); }
            Effected.rowData = Effected.rowData.concat(ip_ReCalculateFormulas(GridID, { recalcSource:true, createUndo: true, range: arrToRange, transactionID: TransactionID, render: true, raiseEvent: false }).Effected.rowData);
            Effected  = ip_PurgeEffectedRowData(GridID, Effected);

            //Re-render the grid
            ip_ReRenderRanges(GridID, (options.cut ? arrToRange.concat(arrFromRange) : arrToRange));
            $(this).ip_SelectCell({ row: options.row, col: options.col });
            for (i = 0; i < arrToRange.length; i++) { $(this).ip_SelectRange({ range: arrToRange[i], multiselect: (i == 0 ? false : true) }); }            

            //Remove cut object from copied cells
            for (i = ip_GridProps[GridID].copiedRange.length - 1 ; i >= 0 ; i--) {  if (ip_GridProps[GridID].copiedRange[i][2] == 'cut') { ip_GridProps[GridID].copiedRange.splice(i, 1); }   }
            $(this).ip_RemoveRangeHighlight({ highlightType: 'ip_grid_cell_rangeHighlight_cut' });

            //Raise event            
            ip_RaiseEvent(GridID, 'ip_Paste', TransactionID, { Paste: { Inputs: options, Effected: Effected } });

        }

        if(Error != '') {

            //Handle any errors            
            $(this).ip_SelectCell({ cell: ip_GridProps[GridID].selectedCell });
            for (i = 0; i < arrFromRange.length; i++) { $(this).ip_SelectRange({ range: arrFromRange[i], multiselect: (i == 0 ? false : true) }); }
            ip_RaiseEvent(GridID, 'warning', arguments.callee.caller, Error);

        }

    }

    $.fn.ip_ClearCopy = function (options) {

        var GridID = $(this).attr('id');
        $(this).ip_RemoveRangeHighlight({ highlightType: 'ip_grid_cell_rangeHighlight_cut' });
        $(this).ip_RemoveRangeHighlight({ highlightType: 'ip_grid_cell_rangeHighlight_copy' });
        ip_GridProps[GridID].copiedRange = new Array();

    }

    $.fn.ip_MergeRange = function (options) {

        var options = $.extend({

            range: null//, //[{ startRow: null, startCol: null, endRow: null, endCol: null }] array of range object

        }, options);

        var GridID = $(this).attr('id');
        var arrRange = new Array(); //{ startRow: null, startCol: null, endRow: null, endCol: null }
        var error = '';
        var Effected = { mergeData: new Array(), rowData:null };
        var TransactionID = ip_GenerateTransactionID();

        //Validate range
        if (options.range) {
            arrRange = options.range;

            for (var i = 0; i < options.range.length; i++) {

                if (arrRange[i].startRow < 0) { arrRange[i].startRow = 0; }
                if (arrRange[i].startCol < 0) { arrRange[i].startCol = 0; }
                if (arrRange[i].endRow >= ip_GridProps[GridID].rows) { arrRange[i].endRow = ip_GridProps[GridID].rows - 1; }
                if (arrRange[i].endCol >= ip_GridProps[GridID].cols) { arrRange[i].endCol = ip_GridProps[GridID].cols - 1; }

            }

        }    
        else {

            if (!options.rangeElement) { options.rangeElement = $('#' + GridID + ' .ip_grid_cell_rangeselector_selected'); }

            for (var i = 0; i < options.rangeElement.length; i++) {
                arrRange[i] = {
                    startRow: parseInt($(options.rangeElement[i]).attr('startrow')),
                    startCol: parseInt($(options.rangeElement[i]).attr('startcol')),
                    endRow: parseInt($(options.rangeElement[i]).attr('endrow')),
                    endCol: parseInt($(options.rangeElement[i]).attr('endcol'))
                }

                if (arrRange[i].startRow < 0) { arrRange[i].startRow = 0; }
                if (arrRange[i].startCol < 0) { arrRange[i].startCol = 0; }
                if (arrRange[i].endRow >= ip_GridProps[GridID].rows) { arrRange[i].endRow = ip_GridProps[GridID].rows - 1; }
                if (arrRange[i].endCol >= ip_GridProps[GridID].cols) { arrRange[i].endCol = ip_GridProps[GridID].cols - 1; }
            }

        }


        //Validate merges
        var PassValidation = true;
        for (var i = 0; i < arrRange.length; i++) {            

            //Checks that we are not merging ontop of merges
            var validateMerges = ip_ValidateRangeMergedCells(GridID, arrRange[i].startRow, arrRange[i].startCol, arrRange[i].endRow, arrRange[i].endCol);
            if (validateMerges.merges.length > 0) {

                PassValidation = false;
                $(this).ip_RangeHighlight({ fadeOut:true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, color: '#ff5a00', range: { startRow: arrRange[i].startRow, startCol: arrRange[i].startCol, endRow: arrRange[i].endRow, endCol: arrRange[i].endCol } });
                error = 'Please adjust your range[s] so that they dont overlap an existing merge';
            }

        }

        if (PassValidation) {
            var validateMerges = ip_ValidateRangeOverlap(GridID, arrRange);
            if (validateMerges.length > 0) {
                PassValidation = false;
                for (var i = 0; i < validateMerges.length; i++) {
                    $(this).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, color: '#ff5a00', range: { startRow: validateMerges[i].startRow, startCol: validateMerges[i].startCol, endRow: validateMerges[i].endRow, endCol: validateMerges[i].endCol } });
                }
                error = 'Please adjust your range[s] so that they dont overlap an each other';
            }
        }

        if (PassValidation) {

            
            for (var i = 0; i < arrRange.length; i++) {

                //Validate that the merge exceedes one cell
                if ((arrRange[i].startRow != arrRange[i].endRow) || (arrRange[i].startCol != arrRange[i].endCol)) {


                    var CellUndoData = ip_AddUndo(GridID, 'ip_MergeRange', TransactionID, 'CellData', arrRange[i], arrRange[i], { row: arrRange[i].startRow, col: arrRange[i].startCol }, null, null, arrRange[i]);
                    var MergeUndoData = ip_AddUndo(GridID, 'ip_MergeRange', TransactionID, 'MergeData', arrRange[i]);


                    //Do actual merge foreach range
                    ip_SetCellMerge(GridID, arrRange[i].startRow, arrRange[i].startCol, arrRange[i].endRow, arrRange[i].endCol, CellUndoData, MergeUndoData);
                    Effected.mergeData[Effected.mergeData.length] = {
                        mergedWithRow: arrRange[i].startRow,
                        mergedWithCol: arrRange[i].startCol,
                        rowSpan: arrRange[i].endRow - arrRange[i].startRow + 1,
                        colSpan: arrRange[i].endCol - arrRange[i].startCol + 1,
                    }
                }
            }

            Effected.rowData = ip_ReCalculateFormulas(GridID, { range: arrRange, transactionID: TransactionID, render: false, raiseEvent: false, createUndo: true }).Effected.rowData;

            ip_ReRenderRows(GridID);

            //Reselect the merge ranges            
            $(this).ip_RemoveRangeHighlight();
            for (var i = 0; i < arrRange.length; i++) {

                var MergeRow = arrRange[i].startRow;
                var MergeCol = arrRange[i].startCol;
                $(this).ip_SelectRange({ range: { startRow: MergeRow, startCol: MergeCol }, multiselect: (i == 0 ? false : true) });
                if (i == 0) { $('#' + GridID).ip_SelectCell({ row: arrRange[0].startRow, col: arrRange[0].startCol }); }
            }


            

            //Raise event            
            ip_RaiseEvent(GridID, 'ip_MergeRange', TransactionID, { MergeRange: { Inputs: { rangeData: arrRange }, Effected: Effected } });

            if (error == '') { return true; }
        }
        else {

            ip_RaiseEvent(GridID, 'warning',arguments.callee.caller, error);
        }

        return false;
    }

    $.fn.ip_UnMergeRange = function (options) {

        var options = $.extend({

            range: null//, //[{ startRow: null, startCol: null, endRow: null, endCol: null }] array of range object

        }, options);

        var GridID = $(this).attr('id');
        var arrRange = new Array(); //{ startRow: null, startCol: null, endRow: null, endCol: null }
        var error = '';
        var Effected = { rangeData: new Array() };
        var TransactionID = ip_GenerateTransactionID();

        //Validate range
        if (options.range) {
            arrRange = options.range;

            for (var i = 0; i < options.range.length; i++) {

                if (arrRange[i].startRow < 0) { arrRange[i].startRow = 0; }
                if (arrRange[i].startCol < 0) { arrRange[i].startCol = 0; }
                if (arrRange[i].endRow >= ip_GridProps[GridID].rows) { arrRange[i].endRow = ip_GridProps[GridID].rows - 1; }
                if (arrRange[i].endCol >= ip_GridProps[GridID].cols) { arrRange[i].endCol = ip_GridProps[GridID].cols - 1; }

            }

        }
        else {

            if (!options.rangeElement) { options.rangeElement = $('#' + GridID + ' .ip_grid_cell_rangeselector_selected'); }

            for (var i = 0; i < options.rangeElement.length; i++) {
                arrRange[i] = {
                    startRow: parseInt($(options.rangeElement[i]).attr('startrow')),
                    startCol: parseInt($(options.rangeElement[i]).attr('startcol')),
                    endRow: parseInt($(options.rangeElement[i]).attr('endrow')),
                    endCol: parseInt($(options.rangeElement[i]).attr('endcol'))
                }

                if (arrRange[i].startRow < 0) { arrRange[i].startRow = 0; }
                if (arrRange[i].startCol < 0) { arrRange[i].startCol = 0; }
                if (arrRange[i].endRow >= ip_GridProps[GridID].rows) { arrRange[i].endRow = ip_GridProps[GridID].rows - 1; }
                if (arrRange[i].endCol >= ip_GridProps[GridID].cols) { arrRange[i].endCol = ip_GridProps[GridID].cols - 1; }
            }

        }


        //Validate merges
        for (var i = 0; i < arrRange.length; i++) {

            //Checks that we are not merging ontop of merges
            var ValidateMerges = ip_ValidateRangeMergedCells(GridID, arrRange[i].startRow, arrRange[i].startCol, arrRange[i].endRow, arrRange[i].endCol);
            arrRange[i].merges = ValidateMerges.merges;

            for (var m = 0; m < ValidateMerges.merges.length; m++) {
                if (ValidateMerges.merges[m].containsOverlap) {
                    $(this).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, color: '#ff5a00', range: { startRow: arrRange[i].startRow, startCol: arrRange[i].startCol, endRow: arrRange[i].endRow, endCol: arrRange[i].endCol } });
                    error = 'Please adjust your range[s] so that they full cover the merge you would like to remove';
                }
            }            
        }

        if (error == '') {

            Effected.rangeData = arrRange;
            
            //Do the unmerge
            for (var i = 0; i < arrRange.length; i++) {

                //Validate that the merge exceedes one cell
                //var merges = ip_ValidateRangeMergedCells(GridID, arrRange[i].startRow, arrRange[i].startCol, arrRange[i].endRow, arrRange[i].endCol);
                var merges = arrRange[i].merges;
                var MergeUndoData = ip_AddUndo(GridID, 'ip_UnMergeRange', TransactionID, 'MergeData', arrRange[i], arrRange[i], { row: arrRange[i].startRow, col: arrRange[i].startCol });

                for (var m = 0; m < merges.length; m++) {

                    if (!merges[m].containsOverlap) {
                        ip_AddUndoTransactionData(GridID, MergeUndoData, ip_mergeObject(merges[m].mergedWithRow, merges[m].mergedWithCol, merges[m].rowSpan, merges[m].colSpan));
                        ip_ResetCellMerge(GridID, merges[m].mergedWithRow, merges[m].mergedWithCol);
                    }

                }



            }

            ip_ReRenderRows(GridID);

            ////Reselect the merge ranges            
            //$(this).ip_RemoveRangeHighlight();
            //for (var i = 0; i < arrRange.length; i++) {

            //    var MergeRow = arrRange[i].startRow;
            //    var MergeCol = arrRange[i].startCol;
            //    $(this).ip_SelectRange({ range: { startRow: MergeRow, startCol: MergeCol }, multiselect: (i == 0 ? false : true) });
            //    if (i == 0) { $('#' + GridID).ip_SelectCell({ row: arrRange[0].startRow, col: arrRange[0].startCol }); }
            //}



            //Raise event            
            ip_RaiseEvent(GridID, 'ip_UnMergeRange', TransactionID, { UnMergeRange: { Inputs: { rangeData: arrRange }, Effected: Effected } });
        }

        if (error == '') { return true; }
        else { ip_RaiseEvent(GridID, 'warning', arguments.callee.caller, error); }

        return false;
    }

    $.fn.ip_Sort = function (options) {

        var options = $.extend({

            col: null, //Column to sort by
            range: null, //[{ startRow: 0, startCol: 0, endRow: 0, endCol: 0, col:0 }], //Col should be column you want to sort by E.G. Column 3
            order: 'az'         

        }, options);

        var GridID = $(this).attr('id');
        var Error = '';
        var Effected = { range: new Array(), order: options.order, col:options.col, rowDataLoading: false, rowData:null  };
        var TransactionID = ip_GenerateTransactionID();

        if (options.col != null && ip_GridProps[GridID].selectedRange.length == 0) { options.range = [ip_rangeObject(0, 0, ip_GridProps[GridID].rows - 1, ip_GridProps[GridID].cols - 1, null, options.col)];  }       
        else if (options.range == null) {

            //Copied ranges
            options.range = new Array();
            for (var i = 0; i < ip_GridProps[GridID].selectedRange.length; i++) { options.range[i] = ip_rangeObject(ip_GridProps[GridID].selectedRange[i][0][0], ip_GridProps[GridID].selectedRange[i][0][1], ip_GridProps[GridID].selectedRange[i][1][0], ip_GridProps[GridID].selectedRange[i][1][1], null, options.col); }
            
        }
        
        //Validate ranges
        for (var i = 0; i < options.range.length; i++) {
            if (options.range[i].startRow < 0) { options.range[i].startRow = 0; }
            if (options.range[i].startCol < 0) { options.range[i].startCol = 0; }
            if (options.range[i].endRow >= ip_GridProps[GridID].rows) { options.range[i].endRow = ip_GridProps[GridID].rows - 1; }
            if (options.range[i].endCol >= ip_GridProps[GridID].cols) { options.range[i].endCol = ip_GridProps[GridID].cols - 1; }

            //Valiudate sort by column 
            if (options.range[i].col == null) { options.range[i].col = (options.col == null ? options.range[i].startCol : options.col); }
            if (options.range[i].col < options.range[i].startCol) { options.range[i].col = options.range[i].startCol; }
            if (options.range[i].col > options.range[i].endCol) { options.range[i].col = options.range[i].endCol; }
        }

        if (options.range != null) {

            //Validate merges
            for (var i = 0; i < options.range.length; i++) {

                var validateMerges = ip_ValidateRangeMergedCells(GridID, options.range[i].startRow, options.range[i].startCol, options.range[i].endRow, options.range[i].endCol);
                var PassValidation = true;

                if (validateMerges.merges.length > 0) {
                    if (validateMerges.containsOverlap) { PassValidation = false; Error = 'Cannot sort a range with overlapping merges'; }
                    else if (validateMerges.containsUnmergedCells) { PassValidation = false; Error = 'Sorting requires all cells in range to be identical (all cells within sort range must have the same merge row/column span or no merge at all)'; }
                    else if (!validateMerges.mergesIdentical) { PassValidation = false; Error = 'Sorting requires all merges within range to be identical'; }
                }

                if (!PassValidation) { $(this).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, color: '#ff5a00', range: { startRow: options.range[i].startRow, startCol: options.range[i].startCol, endRow: options.range[i].endRow, endCol: options.range[i].endCol } }); }

            }

            if (Error == '') {

                //Validation pass, do the sort here        
                for (var i = 0; i < options.range.length; i++) {
                           

                                        
                    var rangeData = ip_GetRangeData(GridID, false, options.range[i].startRow, options.range[i].startCol, options.range[i].endRow, options.range[i].endCol, true, false, true, false, true);
                    var rowSort = options.range[i].startCol == 0 && options.range[i].endCol == ip_GridProps[GridID].cols - 1  ? true : false;

                    //Only do the sort if the range we are trying to sort is full loaded into the grid - else sort on the server side
                    if (rangeData.rowData.length > 0) {

                        if (!rangeData.rowDataLoading) {

                            //Create undo stack
                            var CellUndoData = ip_AddUndo(GridID, 'ip_Sort', TransactionID, 'CellData', options.range[i], options.range[i], { row: options.range[i].startRow, col: options.range[i].startCol });
                            var RowUndoData = ip_AddUndo(GridID, 'ip_Sort', TransactionID, 'RowData', options.range[i]);

                            //Identify which is the correct column index within the rangeData to sort by
                            var col = options.range[i].col - rangeData.startCol;
                            

                            //Sort actual happens here
                            rangeData.rowData.sort(ip_CompareRow(GridID, col, options.order, rowSort));

                            //Update grid with sorted data
                            var rIndex = 0;
                            for (var r = rangeData.startRow; r <= rangeData.endRow; r++) {

                                var cIndex = 0;

                                //Add sorted cells range
                                for (var c = rangeData.startCol; c <= rangeData.endCol; c++) {

                                    //Add to undostack
                                    ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, r, c));

                                    var SortedRow = rIndex;
                                    var SortedCol = cIndex;
                                    var sortedCell = rangeData.rowData[SortedRow].cells[SortedCol];
                                    var merge = ip_GridProps[GridID].rowData[r].cells[c].merge;
                                    
                                    ip_GridProps[GridID].rowData[r].cells[c] = sortedCell;
                                    if (merge != null) { ip_GridProps[GridID].rowData[r].cells[c].merge = merge; }
                                    if (sortedCell.fxIndex != null) { ip_ChangeFormulaOrigin(GridID, { fxIndex: sortedCell.fxIndex, toRow: r, toCol: c }); }

                                    cIndex++
                                }


                                //Add sorted row range
                                if (rowSort) {

                                    //Add to undostack
                                    ip_AddUndoTransactionData(GridID, RowUndoData, ip_CloneRow(GridID, r));

                                    var cells = ip_GridProps[GridID].rowData[r].cells;
                                    var containsMerges = ip_GridProps[GridID].rowData[r].containsMerges;

                                    ip_GridProps[GridID].rowData[r] = rangeData.rowData[rIndex];
                                    ip_GridProps[GridID].rowData[r].containsMerges = containsMerges;
                                    ip_GridProps[GridID].rowData[r].cells = cells;

                                }

                                rIndex++;
                            }

                        }
                        else {

                            //Range not full loaded so we need to sort on the server                            
                            var FunctionUndoData = ip_AddUndo(GridID, 'ip_Sort', TransactionID, 'undoServer', options.range[i], options.range[i], { row: options.range[i].startRow, col: options.range[i].startCol });                      
                            Effected.rowDataLoading = true;
                            
                        }

                        Effected.col = options.col;
                        Effected.range[Effected.range.length] = ip_rangeObject(options.range[i].startRow, options.range[i].startCol, options.range[i].endRow, options.range[i].endCol, null, options.range[i].col);
                        
                    }
                    
                }

                if (!Effected.rowDataLoading) {

                    Effected.rowData = ip_ReCalculateFormulas(GridID, { range: Effected.range, transactionID: TransactionID, render: false, raiseEvent: false, createUndo: true }).Effected.rowData;
                    ip_ReRenderRows(GridID);

                }

                //Raise event  
                if (Effected.range.length > 0) { ip_RaiseEvent(GridID, 'ip_Sort', TransactionID, { Sort: { Inputs: options, Effected: Effected } }); }

            }
            else { ip_RaiseEvent(GridID, 'warning', TransactionID, Error); }

        }
    }

    $.fn.ip_TextEditable = function (options) {

        var options = $.extend({
            
            row: null,
            col: null,
            element: null, //Element to edit
            ToolBar: null, //Works with the 'ip_toolbar_font' by highlighting toolbar buttons, bold, italic, font etc
            contentType: 'ip_grid_CellEdit',
            defaultValue: null,
            clear: false,
            cursor:null,
            dropDown: { allowEmpty: true, validate: false, autoComplete: true, keyField: null, displayField: null, data: [], noMatchData: null }

        }, options);
        
        if (options.element != null) {
            
            var GridID = $(this).attr('id');
            var top = 0;
            var left = 0;
            var width = 0;
            var height = 0;
            var appendTo = options.element;
            var defaultValue = options.defaultValue;
            var elementToEdit = (options.element.length > 0 ? options.element[0] : options.element);
            var style = '';
            

            //if (options.element != ip_GridProps[GridID].editing.element) {

                ip_EnableSelection(GridID);

                if (ip_GridProps[GridID].editing.editTool == null) { ip_GridProps[GridID].editing.editTool = $('#' + GridID + '_editTool')[0]; }

                var editTool = ip_GridProps[GridID].editing.editTool; //Cache the edit tool
                var editToolInput = $(editTool).children('.ip_grid_EditTool_Input')[0];
                var fxBarText = $(ip_GridProps[GridID].fxBar).children('.ip_grid_fbar_text')[0];
                         
                editTool.text = function (value, key, cursorX, cursorLength) {

                    if (key != null) { $(editToolInput).attr('key', key); }
                    if (value != null) {

                        if (typeof value == 'string') { value = value.replace(/\s$/, "&nbsp;"); }

                        $(editToolInput).html(value);
                        $(fxBarText).html(value);
                        $(elementToEdit).find('.ip_grid_cell_innerContent').html(value);

                    }                    
                    if (cursorX != null) { ip_GridProps[GridID].editing.cursor = ip_SetCursorPos(GridID, editToolInput, cursorX, cursorLength); }
                    //if (value == null) { value = ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].value; }
                    

                   return $(editToolInput).text();
                    // return value;
                }
                editTool.isFX = function (returnFX) {

                    var text = editTool.text().trim();

                    if (text != '') {
                        if (!returnFX && text[0] == '=') { return true; }
                        else { return ip_fxObject(GridID, text); }
                    }

                    return false;
                }
                editTool.row = parseInt($(elementToEdit).attr('row'));
                editTool.col = parseInt($(elementToEdit).attr('col'));

                $(editToolInput).attr('key', '');
                                
                if (options.contentType == 'ip_grid_CellEdit' || options.contentType == 'ip_grid_ColumnType') {

                    elementToStyle = elementToEdit;

                    appendTo = ($(elementToEdit).parent().parent().parent().parent().parent())[0];

                    top = ip_CalculateRangeTop(GridID, editTool, elementToEdit, editTool.row, editTool.col, editTool.row, editTool.col, true);
                    left = ip_CalculateRangeLeft(GridID, editTool, elementToEdit, editTool.row, editTool.col, editTool.row, editTool.col, true);

                    width = $(elementToEdit).width();
                    height = $(elementToEdit).height();

                    if (defaultValue == null && !options.clear && editTool.row >= 0) { defaultValue = (ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].formula != null && ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].formula != '' ? ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].formula : ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].value) } //$(elementToEdit).find('.ip_grid_cell_innerContent').html();
                    else if(defaultValue == null) { defaultValue = ''; }

                    if (options.contentType == 'ip_grid_ColumnType') { style = 'background-color:inherit;box-shadow:none;'; }
                    else {  style = $(elementToEdit).attr('style');   }

                    //Set dropdown data to a range - if we are validating from a range
                    if (options.dropDown.data.length == 0 && ip_GetEnabledControlType(GridID, null, editTool.row, editTool.col, true) == 'dropdown') { options.dropDown.data = (ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].validation.validationCriteria ? ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].validation.validationCriteria : ip_GridProps[GridID].colData[editTool.col].validation.validationCriteria); }
       

                }
                else if (options.contentType = 'ip_grid_fxBar') {

                    elementToStyle = ip_GridProps[GridID].fxBar;

                    appendTo = ip_GridProps[GridID].fxBar;

                    top = 0;
                    left = 60;

                    width = $(fxBarText).width() - left; // - parseInt($(editToolInput).css('padding-right').replace('px', ''));
                    height = $(fxBarText).height();// - ($(editTool).outerHeight(true) - $(editTool).height());

                    if (defaultValue == null && !options.clear && editTool.row >= 0) { defaultValue = (ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].formula != null && ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].formula != '' ? ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].formula : ip_GridProps[GridID].rowData[editTool.row].cells[editTool.col].value) } //$(elementToEdit).find('.ip_grid_cell_innerContent').html();
                    else if (defaultValue == null) { defaultValue = ''; }
                    
                    style = 'box-shadow:none;background-color:' + $(appendTo).css('background-color') + ';';
                    

                }
                
                if (defaultValue == null) { defaultValue = ''; }
                if (style == null) { style = ''; }

                //Record the element that is being edited
                ip_GridProps[GridID].editing.editing = true;
                ip_GridProps[GridID].editing.row = editTool.row;
                ip_GridProps[GridID].editing.col = editTool.col;
                ip_GridProps[GridID].editing.element = elementToEdit;
                ip_GridProps[GridID].editing.contentType = options.contentType;
                ip_GridProps[GridID].editing.carret = options.cursor;

                //Set position of edit tool
                $(appendTo).append(editTool);                
                ip_GridProps[GridID].editing.editTool.text(defaultValue);                

                $(editTool).width(width);//.css('min-width', width + 'px');
                $(editTool).height(height);//.css('min-height', height + 'px');
                $(editTool).css('line-height', height + 'px');
                $(editTool).css('top', top + 'px');
                $(editTool).css('left', left + 'px');
                $(editTool).css('text-align', $(elementToStyle).css('text-align'));
                $(editTool).css('vertical-align', $(elementToStyle).css('vertical-align'));
                $(editTool).css('font-family', $(elementToStyle).css('font-family'));
                $(editTool).css('font-size', $(elementToStyle).css('font-size'));
                $(editTool).css('color', $(elementToStyle).css('color'));

                $(editToolInput).attr('style', style);
                
                ip_GridProps[GridID].editing.carret

                //Deal setting the carret position for the editing text as well as showing tool tip
                ip_UnBindEvent(editToolInput, 'mouseup', ip_GridProps[GridID].events.textEditTool_MouseUp);
                $(editToolInput).mouseup(ip_GridProps[GridID].events.textEditTool_MouseUp = function (e) {


                    var carret = ip_GetCursorPos(GridID, editToolInput);
                    var text = ip_GridProps[GridID].editing.editTool.text();
                    ip_GridProps[GridID].editing.carret = carret;
                    ip_EditToolHelp(GridID, text, carret, 0, true, true, true);                     
    

                });

                //Deal with the enter key, moving cursor, and selecting items in the autocomplete dropdown
                ip_UnBindEvent(editToolInput, 'keyup', ip_GridProps[GridID].events.textEditTool_KeyUp);
                $(editToolInput).keyup(ip_GridProps[GridID].events.textEditTool_KeyUp = function (e) {

                    if (e.keyCode == 8 || e.keyCode == 37 || e.keyCode == 39) {

                        var carret = ip_GetCursorPos(GridID, editToolInput);
                        var text = ip_GridProps[GridID].editing.editTool.text();

                        ip_GridProps[GridID].editing.carret = carret;

                        ip_EditToolHelp(GridID, text, carret, 0, true, true, true);
                        e.preventDefault();
                    }

                });

                //Deal with the enter key, moving cursor, and selecting items in the autocomplete dropdown
                ip_UnBindEvent(editToolInput, 'keydown', ip_GridProps[GridID].events.textEditTool_KeyDown);
                $(editToolInput).keydown(ip_GridProps[GridID].events.textEditTool_KeyDown = function (e) {

                    clearTimeout(ip_GridProps[GridID].timeouts.textEditToolBlurTimeout);//Prevents hiding of autocomplete dropdown
                    
                    
                    if (e.keyCode == 13) {
                                      
                        //Enter key
                        return ip_TextEditbleBlur(GridID, defaultValue, editTool, editToolInput, ip_GridProps[GridID].editing.element, ip_GridProps[GridID].editing.contentType, editTool.row, editTool.col, true);

                    }
                    else if (e.keyCode == 27) {

                        //Escape
                        return ip_TextEditbleCancel(GridID, editTool, editToolInput, ip_GridProps[GridID].editing.element, ip_GridProps[GridID].editing.contentType, editTool.row, editTool.col, true);
                        
                    }
                    else if (e.keyCode == 8) {

                        //Backspace key
                        var carret = ip_GetCursorPos(GridID, editToolInput);
                        var text = ip_GridProps[GridID].editing.editTool.text();

                        text = text.slice(0, carret.x - (carret.length == 0 ? 1 : carret.length)) + text.substring(carret.x + carret.length, text.length);
                        ip_EditToolDropDownFilter(GridID, text);
                        //ip_SetFxBarValue(GridID, { value: text, cell: null });

                    }
                    else if (e.keyCode == 37 || e.keyCode == 39) {

                        //Move cursor left or right     
                        var carret = ip_GetCursorPos(GridID, editToolInput);
                        var cursorIncr = (e.keyCode == 37 ? -1 : 1);
                        var cursorI = carret.x + cursorIncr;

                        if (cursorI < 0 || cursorI > $(this).text().length) { return ip_TextEditbleBlur(GridID, defaultValue, editTool, editToolInput, ip_GridProps[GridID].editing.element, ip_GridProps[GridID].editing.contentType, editTool.row, editTool.col); }

                    }
                    else if (e.keyCode == 38 || e.keyCode == 40) {

                        //Up down keys
                        if (options.dropDown.data != null && options.dropDown.data.length > 0 && options.dropDown.autoComplete) {
                            //Filter dropdown                            
                            ip_EditToolDropDownHighlightNext(GridID, (e.keyCode == 38 ? 'up' : 'down'));
                        }
                        else {

                            return ip_TextEditbleBlur(GridID, defaultValue, editTool, editToolInput, ip_GridProps[GridID].editing.element, ip_GridProps[GridID].editing.contentType, editTool.row, editTool.col);

                        }

                    }
                    


                });

                //Deals with regular keypress, filtering drop down. NOTE we cant use the keydown event because it wont capture the character when shift is pressed
                ip_UnBindEvent(editToolInput, 'keypress', ip_GridProps[GridID].events.textEditTool_KeyPress);
                $(editToolInput).keypress(ip_GridProps[GridID].events.textEditTool_KeyPress = function (e) {

                    var key = e.key || String.fromCharCode(e.keyCode);

                    if (key == 'Spacebar') { key = ' '; };
                    if (key.length == 1) {

                        var carret = ip_GetCursorPos(GridID, editToolInput);
                        var text = ip_GridProps[GridID].editing.editTool.text();

                       
                        ip_GridProps[GridID].editing.carret = $.extend({},carret); //Create a clone
                        ip_GridProps[GridID].editing.carret.x++;

                        text = text.slice(0, carret.x) + key + text.substring(carret.x + carret.length, text.length);
                        
                        if (options.dropDown.data != null && options.dropDown.data.length > 0 && options.dropDown.autoComplete) {

                            //Filter dropdown                            
                            ip_EditToolDropDownFilter(GridID, text);

                        }
                        else {

                            //Show help
                            carret.length = 0;                            
                            var FX = ip_EditToolHelp(GridID, text, carret, 1, true, true, true);
                            e.preventDefault(); 

                        }

                        
                    }
                });

                //Setup event to save changes after editing
                ip_UnBindEvent(editToolInput, 'focusin', ip_GridProps[GridID].events.textEditTool_FocusIn);
                $(editToolInput).focusin(ip_GridProps[GridID].events.textEditTool_FocusIn = function (e) {

                    clearTimeout(ip_GridProps[GridID].timeouts.textEditToolBlurTimeout);

                });
                    
                //Setup event to save changes after editing
                ip_UnBindEvent(editToolInput, 'focusout', ip_GridProps[GridID].events.textEditTool_FocusOut);
                $(editToolInput).focusout(ip_GridProps[GridID].events.textEditTool_FocusOut = function (e) {
                    
                    ip_GridProps[GridID].timeouts.textEditToolBlurTimeout = setTimeout(function () {
                        return ip_TextEditbleBlur(GridID, defaultValue, editTool, editToolInput, ip_GridProps[GridID].editing.element, ip_GridProps[GridID].editing.contentType, editTool.row, editTool.col, true);

                    }, 10);

                });                       

                $(editTool).show();                
                ip_EditToolShowDropDown(GridID, editTool, options.dropDown);
                ip_EditToolHelp(GridID, defaultValue, options.cursor, 0, false, true, true);                
                editToolInput.focus();
            }
        
        //}
    }

    $.fn.ip_Undo = function (options) {

        var options = $.extend({

            transactionID:''

        }, options);

        var GridID = $(this).attr('id');
        var Effected = { transactionID: '', rowDataLoading:false };
        var Error = '';

        if (Object.keys(ip_GridProps[GridID].undo.undoStack).length > 0) {

            var sortedStack = [];

            $('#' + GridID).ip_RemoveRangeHighlight();

            //Buld an array of the undo stack so we can sort it
            for (var key in ip_GridProps[GridID].undo.undoStack) {
                if (ip_GridProps[GridID].undo.undoStack.hasOwnProperty(key)) {
                    sortedStack.push(ip_GridProps[GridID].undo.undoStack[key]);
                }
            }

            //Sort stack descending with the latest transaction frist
            sortedStack.sort(function (a, b) { return b.transactionSeq - a.transactionSeq;  });
            

            //Get latest transactions
            if (options.transactionID == '') { options.transactionID = sortedStack[0].transactionID }
        
            //Loop through undo transactions and undo them until we reach the transaction we've specified
            for (var i = 0; i < sortedStack.length; i++) {

                var IsTransaction = (sortedStack[i].transactionID == options.transactionID ? true : false);
                var UndoResult = ip_UndoTransaction(GridID, sortedStack[i].transactionID, IsTransaction, IsTransaction);
                if (UndoResult.success) { Effected.transactionID = options.transactionID; ip_RaiseEvent(GridID, 'message', null, 'Undo: ' + UndoResult.method.replace('ip_','')); } else { i = sortedStack.length; }
                if (UndoResult.rowDataLoading) { Effected.rowDataLoading = true; }
                if (IsTransaction) { i = sortedStack.length; }
                
            }
              
            if (Effected.transactionID != '') { ip_RaiseEvent(GridID, 'ip_Undo', null, { Undo: { Inputs: options, Effected: Effected } }); }
            else { Error = 'Unable to undo'; }
        }
        else { Error = 'There is nothing to undo'; }

        if (Error != '') { ip_RaiseEvent(GridID, 'warning', null, Error); }
    }

    $.fn.ip_HideShowRows = function (options) {

        var options = $.extend({

            action: 'hide',
            range: null, //{fromRow:0, toRow: 0}
            rows: null, //Array containing the rows to hide
            selectRows: true,
            render: true,
            raiseEvent: true,
            createUndo: true

        }, options);

        var GridID = $(this).attr('id');        
        var selectRows = [];
        var TransactionID = ip_GenerateTransactionID();
        var Effected = { action: options.action, rowData: [] };
        var Error = '';

        //Validate inputs
        if (options.range != null) {
            var i = 0;
            if (options.rows == null) { options.rows = new Array(); }
            for (var r = options.range.startRow; r <= options.range.endRow; r++) {  options.rows[i] = r; i++;   }
        }
        else if (options.rows == null) { options.rows = ip_GridProps[GridID].selectedRow.slice(0); }

        if (jQuery.inArray(-1, options.rows) != -1) { Error = 'Cannot ' + options.action + ' all rows, please choose specific rows to hide' }
        if (options.rows.length == 0) { Error = 'Nothing specified to hide'; }        

        if (Error == '') {

            if (options.action == 'hide') {

                for (var i = 0; i < options.rows.length; i++) {

                    var row = options.rows[i];

                    if (row >= 0 && row < ip_GridProps[GridID].rows) {
                        if (!ip_GridProps[GridID].rowData[row].hide) {

                            //Add to undo stack
                            if (options.createUndo) {
                                var ResizeRange = { startRow: row, startCol: -1, endRow: row, endCol: ip_GridProps[GridID].cols - 1 }
                                var RowUndoData = ip_AddUndo(GridID, 'ip_HideShowRows', TransactionID, 'RowData', ResizeRange, ResizeRange, { row: ResizeRange.startRow, col: ResizeRange.startCol }, null, ip_CloneRow(GridID, row));
                            }

                            ip_GridProps[GridID].rowData[row].hide = true;
                            Effected.rowData[Effected.rowData.length] = { row: row, hide: true };

                        }
                    }

                }


                $('#' + GridID).ip_RemoveRange();
                $('#' + GridID).ip_RemoveRangeHighlight();
            }
            else {


                for (var i = 0; i < options.rows.length; i++) {

                    var row = options.rows[i];



                    if (row >= 0 && row < ip_GridProps[GridID].rows) {
                        if (ip_GridProps[GridID].rowData[row].hide) {


                            //Add to undo stack
                            if (options.createUndo) {
                                var ResizeRange = { startRow: row, startCol: -1, endRow: row, endCol: ip_GridProps[GridID].cols - 1 }
                                var RowUndoData = ip_AddUndo(GridID, 'ip_HideShowRows', TransactionID, 'RowData', ResizeRange, ResizeRange, { row: ResizeRange.startRow, col: ResizeRange.startCol }, null, ip_CloneRow(GridID, row));
                            }

                            ip_GridProps[GridID].rowData[row].hide = false;
                            Effected.rowData[Effected.rowData.length] = { row: row, hide: false };

                            //Select rows afterwards
                            if (selectRows.length > 0 && selectRows[selectRows.length - 1].lastRow == row - 1) { selectRows[selectRows.length - 1].count++; selectRows[selectRows.length - 1].lastRow = row; }
                            else { selectRows[selectRows.length] = { row: row, count: 1, lastRow: row } }

                        }
                    }

                }

            }



            if (options.render) {

                ip_ReRenderRows(GridID, 'frozen');
                $(this).ip_Scrollable();
                ip_ReRenderRows(GridID, 'scroll');

                //Select rows
                if (options.selectRows) { for (i = 0; i < selectRows.length; i++) { $('#' + GridID).ip_SelectRow({ row: selectRows[i].row, count: selectRows[i].count, unselect: false, multiselect: (i == 0 ? false : true) }); } }
                else {
                    $('#' + GridID).ip_UnselectRow();
                    $('#' + GridID).ip_RemoveRange();
                    $('#' + GridID).ip_RemoveRangeHighlight();
                }

            }

        }

        if (Error == '' && options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_HideShowRows', TransactionID, { HideShowRows: { Inputs: options, Effected: Effected } }); }

        if (Error != '') { ip_RaiseEvent(GridID, 'warning', null, Error); }
    }

    $.fn.ip_HideShowColumns = function (options) {

        var options = $.extend({

            action: 'hide',
            range: null, //{startCol:0, endCol:0 },
            cols: null, //Array containing the rows to hide
            render: true,
            raiseEvent: true,
            createUndo: true

        }, options);

        var GridID = $(this).attr('id');
        var selectCols = [];
        var TransactionID = ip_GenerateTransactionID();
        var Effected = { action: options.action, colData: [] };
        var Error = '';

        //Validate inputs
        if (options.range != null) {
            var i = 0;
            if (options.cols == null) { options.cols = new Array(); }
            for (var c = options.range.startCol; c <= options.range.endCol; c++) { options.cols[i] = c; i++; }
        }
        if (options.cols == null) { options.cols = ip_GridProps[GridID].selectedColumn.slice(0); }

        if (jQuery.inArray(-1, options.cols) != -1) { Error = 'Cannot ' + options.action + ' all columns, please choose specific column to hide' }
        if (options.cols.length == 0) { Error = 'Nothing specified to hide'; }

        if (Error == '') {

            if (options.action == 'hide') {

                for (var i = 0; i < options.cols.length; i++) {

                    var col = options.cols[i];

                    if (col >= 0 && col < ip_GridProps[GridID].cols) {
                        if (!ip_GridProps[GridID].colData[col].hide) {

                            //Add to undo stack
                            if (options.createUndo) {
                                var ResizeRange = { startRow: -1, startCol: col, endRow: ip_GridProps[GridID].rows - 1, endCol: col }
                                var ColUndoData = ip_AddUndo(GridID, 'ip_HideShowColumns', TransactionID, 'ColData', ResizeRange, ResizeRange, { row: ResizeRange.startRow, col: ResizeRange.startCol }, null, ip_CloneCol(GridID, col));
                            }

                            ip_GridProps[GridID].colData[col].hide = true;
                            Effected.colData[Effected.colData.length] = { col: col, hide: true };

                        }
                    }

                }

                $('#' + GridID).ip_RemoveRange();
                $('#' + GridID).ip_RemoveRangeHighlight();

            }
            else {

                for (var i = 0; i < options.cols.length; i++) {

                    var col = options.cols[i];

                    if (col >= 0 && col < ip_GridProps[GridID].cols) {
                        if (ip_GridProps[GridID].colData[col].hide) {

                            //Add to undo stack
                            if (options.createUndo) {
                                var ResizeRange = { startRow: -1, startCol: col, endRow: ip_GridProps[GridID].rows - 1, endCol: col }
                                var ColUndoData = ip_AddUndo(GridID, 'ip_HideShowColumns', TransactionID, 'ColData', ResizeRange, ResizeRange, { row: ResizeRange.startRow, col: ResizeRange.startCol }, null, ip_CloneCol(GridID, col));
                            }

                            ip_GridProps[GridID].colData[col].hide = false;
                            Effected.colData[Effected.colData.length] = { col: col, hide: false };

                            //Select rows afterwards
                            if (selectCols.length > 0 && selectCols[selectCols.length - 1].lastCol == col - 1) { selectCols[selectCols.length - 1].count++; selectCols[selectCols.length - 1].lastCol = col; }
                            else { selectCols[selectCols.length] = { col: col, count: 1, lastCol: col } }

                        }
                    }

                }

            }

            if (options.render) {

                $(this).ip_Scrollable();
                ip_ReRenderCols(GridID);
                ip_RePoistionRanges(GridID, 'all', true, false);

                //Select rows
                for (i = 0; i < selectCols.length; i++) { $('#' + GridID).ip_SelectColumn({ col: selectCols[i].col, count: selectCols[i].count, unselect: false, multiselect: (i == 0 ? false : true) }); }
                
            }


        }


        if (Error == '' && options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_HideShowColumns', TransactionID, { HideShowColumns: { Inputs: options, Effected: Effected } }); }

        if (Error != '') { ip_RaiseEvent(GridID, 'warning', null, Error); }
    }

    $.fn.ip_GroupRows = function (options) {

        var options = $.extend({

            groupColumn: 0,
            action: 'hide',
            range: null, //{fromRow:0, toRow: 0}  rows are grouped under first row
            rows: null, //Array containing the rows to group, rows are grouped under first row
            render: true,
            raiseEvent: true,
            createUndo: true

        }, options);

        var GridID = $(this).attr('id');
        var selectRows = [];
        var TransactionID = ip_GenerateTransactionID();
        var Effected = { rowData: [] };
        var Error = '';

        //Validate inputs
        if (options.range != null) {
            var i = 0;
            if (options.rows == null) { options.rows = new Array(); }
            for (var r = options.range.startRow; r <= options.range.endRow; r++) { options.rows[i] = r; i++; }
        }
        else if (options.rows == null) { options.rows = ip_GridProps[GridID].selectedRow.slice(0); }

        if (jQuery.inArray(-1, options.rows) != -1) { Error = 'Cannot group all rows, please choose specific rows to hide' }
        if (options.rows.length == 0) { Error = 'Nothing specified to group'; }

        if (Error == '') {

            options.rows.sort(function (a, b) { return a - b });
            
            //Get rows to group
            var CurrentGroupedRow = -1;
            var CurrentEffectedIndex = 0;
            var PrevRow = -2;
            for (var r = 0; r < options.rows.length; r++) {

                //Add to undo stack
                if (options.createUndo) {
                    var GroupRange = { startRow: options.rows[r], startCol: -1, endRow: options.rows[r], endCol: ip_GridProps[GridID].cols - 1 }
                    var RowUndoData = ip_AddUndo(GridID, 'ip_GroupRows', TransactionID, 'RowData', GroupRange, GroupRange, { row: GroupRange.startRow, col: GroupRange.startCol }, null, ip_CloneRow(GridID, options.rows[r]));
                }

                if (PrevRow != options.rows[r] - 1) {

                    CurrentGroupedRow = options.rows[r];
                    CurrentEffectedIndex = Effected.rowData.length;

                    Effected.rowData[CurrentEffectedIndex] = { row: CurrentGroupedRow, groupCount: 0, groupColumn: options.groupColumn }
                    ip_GridProps[GridID].rowData[CurrentGroupedRow].groupCount = 0;
                    ip_GridProps[GridID].rowData[CurrentGroupedRow].groupColumn = options.groupColumn;
                    selectRows[selectRows.length] = { row: CurrentGroupedRow, count: 1 }

                }
                else
                {

                    Effected.rowData[CurrentEffectedIndex].groupCount++;
                    ip_GridProps[GridID].rowData[CurrentGroupedRow].groupCount++;

                    if (options.action == 'hide') {
                        
                        Effected.rowData[Effected.rowData.length] = { row:options.rows[r], hide: true };
                        ip_GridProps[GridID].rowData[options.rows[r]].hide = true;
                        
                    }

                }

                //groupCount
                PrevRow = options.rows[r];
            }


            if (options.render) {

                ip_ReRenderRows(GridID, 'frozen');
                $(this).ip_Scrollable();
                ip_ReRenderRows(GridID, 'scroll');

                //Select rows
                for (i = 0; i < selectRows.length; i++) { $('#' + GridID).ip_SelectRow({ row: selectRows[i].row, count: selectRows[i].count, unselect: false, multiselect: (i == 0 ? false : true) }); }

            }

        }

        if (Error == '' && options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_GroupRows', TransactionID, { GroupRows: { Inputs: options, Effected: Effected } }); }

        if (Error != '') { ip_RaiseEvent(GridID, 'warning', null, Error); }
    }

    $.fn.ip_UngroupRows = function (options) {

        var options = $.extend({

            range: null, //{fromRow:0, toRow: 0}  rows are grouped under first row
            rows: null, //Array containing the rows to group, rows are grouped under first row
            render: true,
            raiseEvent: true,
            createUndo: true

        }, options);

        var GridID = $(this).attr('id');
        var selectRows = [];
        var TransactionID = ip_GenerateTransactionID();
        var Effected = { rowData: [] };
        var Error = '';

        //Validate inputs
        if (options.range != null) {
            var i = 0;
            if (options.rows == null) { options.rows = new Array(); }
            for (var r = options.range.startRow; r <= options.range.endRow; r++) { options.rows[i] = r; i++; }
        }
        else if (options.rows == null) { options.rows = ip_GridProps[GridID].selectedRow.slice(0); }

        if (jQuery.inArray(-1, options.rows) != -1) { Error = 'Cannot group all rows, please choose specific rows to hide' }
        if (options.rows.length == 0) { Error = 'Nothing specified to group'; }

        if (Error == '') {

            var GroupedRows = ip_GetGroupedRows(GridID);
            var UngroupStrategy = {};

            options.rows.sort(function (a, b) { return a - b });

            //Create a strategy to ungroup. Do this by finding the first row in the group that we want to ungroup from and how much to ungroup by          
            for (var r = 0; r < options.rows.length; r++) {

                var row = options.rows[r];


                if (GroupedRows.rowData[row] != null) {

                    var groupedWithRow = GroupedRows.rowData[row].groupedWithRow;

                    if (UngroupStrategy[groupedWithRow] == null) {
                        UngroupStrategy[groupedWithRow] = {
                            groupedWithRow: GroupedRows.rowData[row].groupedWithRow,
                            removeCount: GroupedRows.rowData[row].groupCountDown,
                            showFromRow: row,
                            showToRow: row + GroupedRows.rowData[row].groupCountDown - 1
                        }
                    }
                    
              
                }

            }

            //Decrease the group count
            for (key in UngroupStrategy) {
                
                var groupedWithRow = UngroupStrategy[key].groupedWithRow;

                //Add to undo stack
                if (options.createUndo) {
                    var GroupRange = { startRow: groupedWithRow, startCol: -1, endRow: groupedWithRow, endCol: ip_GridProps[GridID].cols - 1 }
                    var RowUndoData = ip_AddUndo(GridID, 'ip_UngroupRows', TransactionID, 'RowData', GroupRange, GroupRange, { row: GroupRange.startRow, col: GroupRange.startCol }, null, ip_CloneRow(GridID, groupedWithRow));
                }
                
                if (ip_GridProps[GridID].rowData[groupedWithRow].groupCount > 0) {

                    ip_GridProps[GridID].rowData[groupedWithRow].groupCount -= UngroupStrategy[key].removeCount;
                    if (ip_GridProps[GridID].rowData[groupedWithRow].groupCount < 0) { ip_GridProps[GridID].rowData[groupedWithRow].groupCount = 0; }

                    Effected.rowData[Effected.rowData.length] = {
                        row: groupedWithRow,
                        groupCount: ip_GridProps[GridID].rowData[groupedWithRow].groupCount,
                    }
                }



            }


            if (options.render) {

                ip_ReRenderRows(GridID, 'frozen');
                $(this).ip_Scrollable();
                ip_ReRenderRows(GridID, 'scroll');

                //Select rows
                for (i = 0; i < selectRows.length; i++) { $('#' + GridID).ip_SelectRow({ row: selectRows[i].row, count: selectRows[i].count, unselect: false, multiselect: (i == 0 ? false : true) }); }

            }

        }

        if (Error == '' && options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_UngroupRows', TransactionID, { UngroupRows: { Inputs: options, Effected: Effected } }); }

        if (Error != '') { ip_RaiseEvent(GridID, 'warning', null, Error); }
    }

    $.fn.ip_FormatCell = function (options) {

        var GridID = $(this).attr('id');

        return ip_SetCellFormat(GridID, options);

    }

    $.fn.ip_FormatCellModal = function (options) {
        //Shows the format cell modal

        var options = $.extend({

            toolType: 'cellFormatting',
            toolDefault: 'celltype',

        }, options);

        
        var GridID = $(this).attr('id');

        //Validate
        if (!ip_GridProps[GridID].selectedRange && ip_GridProps[GridID].selectedRange || ip_GridProps[GridID].selectedRange && ip_GridProps[GridID].selectedRange.length == 0) { return; }
        
        var ranges = ip_ValidateRangeObjects(GridID, ip_GridProps[GridID].selectedRange);
       
        options.id = GridID;
        options.enabledFormats = ip_EnabledFormats(GridID, { maxCells:1, getValidation: true, getControlType: true, getDataType: true, getHashTags: true, getMask: true, getStyle: false, adviseDefault: true });

        if (options.enabledFormats.controlType == null) { options.enabledFormats.controlType = ''; }
        if (options.enabledFormats.hashTags == null) { options.enabledFormats.hashTags = ''; }
        if (options.enabledFormats.mask == null) { options.enabledFormats.mask = ''; }
        
        ip_EnableSelection(GridID);
                
        $().ip_Modal({
            greyOut: screen,            
            speech: false,
            title: 'Cell Formatting',
            message: ip_CreateGridTools(options),
            buttons: {

                ok: {
                    text: 'SAVE', style: "background: #4d9c3b; background: -moz-linear-gradient(top,  #4d9c3b 0%, #1b6d00 100%);background: -webkit-gradient(linear, left top, left bottom, color-stop(0%,#4d9c3b), color-stop(100%,#1b6d00)); background: -webkit-linear-gradient(top,  #4d9c3b 0%,#1b6d00 100%); background: -o-linear-gradient(top,  #4d9c3b 0%,#1b6d00 100%); background: -ms-linear-gradient(top,  #4d9c3b 0%,#1b6d00 100%); background: linear-gradient(to bottom,  #4d9c3b 0%,#1b6d00 100%); filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#4d9c3b', endColorstr='#1b6d00',GradientType=0 ); ", onClick: function () {

                        //This code warrents an explination:
                        //It checks to see what was formatting options were changed, and only if an option changed does it save the change. 
                        //It looks at the ip-dirty class to indecate if a filed has changed.
                        //By only saving changed stuff, allows us to show the default value (column value) but not comit it to the cell object. This is a very effcient way of managing 100000's of rows.
                        //Setting an option to "default" resets it so the cell does not contain an actual value and allows it to use the column setting -> this is more memory efficient

                        ip_GridProps[GridID].formatting.formatting = false;

                        var validationValue = $('#' + GridID + '_ValidationValue')[0];
                        var inputs = $('#' + GridID + '_ValidationValue').val().split(",");
                        var fxName = $('#' + GridID + '_ControlType').attr('key');                                               
                        
                        ranges = ip_fxRangeObject(GridID, null, null, $('#' + GridID + '_Range_F').val());

                        for (var i = 0; i < inputs.length; i++) {
                            
                            if (ip_fxRangeObject(GridID, null, null, inputs[i])) {
                                inputs[i] = inputs[i].trim();
                            }
                            else if (inputs[i] && typeof (ip_parseAny(GridID, inputs[i])) == 'string') {
                                inputs[i] = '"' + inputs[i].trim() + '"';
                            }
                        }

                        
                        var hashTags = $('#' + GridID + '_HashTags')[0].val(); 
                        var formula = '=' + fxName + '(' + ($(validationValue).hasClass('tooltip') ? '' : inputs) + ')';
                        var controlType = ip_GetCellControlType(GridID, formula, null, null, null);
                        var dataTypes = $('#' + GridID + '_DataType').attr('key').replace(/default/gi, '').split('~');
                        var mask = $('#' + GridID + '_Mask').attr('key');
                        var recalculate = false;

                        if (dataTypes[0] != '' && (!dataTypes[1] || dataTypes[1] == '')) { dataTypes[1] = dataTypes[0]; }
                                       
                        controlType.validation.validationAction = $('#' + GridID + '_Action').attr('key');                        
                        controlType.dataType = ip_dataTypeObject(dataTypes[0], dataTypes[1]);
                        
                        //Only include if value has changed - NB 
                        if (!$('#' + GridID + '_Mask').hasClass('ip-dirty') && !$('#' + GridID + '_DataType').hasClass('ip-dirty')) { mask = null; }
                        if (!$('#' + GridID + '_HashTags').hasClass('ip-dirty')) { hashTags = null; }
                        if (!$('#' + GridID + '_DataType').hasClass('ip-dirty')) { controlType.dataType = null; }
                        if (!$('#' + GridID + '_ControlType').hasClass('ip-dirty') && !$('#' + GridID + '_ValidationValue').hasClass('ip-dirty')) { controlType.controlType = null; }
                        if (!$('#' + GridID + '_ValidationValue').hasClass('ip-dirty')) { controlType.validation.validationCriteria = null; }
                        if (!$('#' + GridID + '_Action').hasClass('ip-dirty')) { controlType.validation.validationAction = null; }
                        if (!$('#' + GridID + '_ValidationValue').hasClass('ip-dirty') && !$('#' + GridID + '_Action').hasClass('ip-dirty')) { controlType.validation = null; }
                        //if (!$('#' + GridID + '_Action').hasClass('ip-dirty') && !$('#' + GridID + '_ControlType').hasClass('ip-dirty') && !$('#' + GridID + '_ValidationValue').hasClass('ip-dirty')) { controlType.controlType = null; controlType.validation = null; }
                        
                        if (controlType.validation != null || controlType.controlType != null || controlType.dataType != null || hashTags != null || mask != null) {
                            recalculate = (hashTags != null || controlType.dataType != null || mask != null || controlType.validation != null);
                            ip_SetCellFormat(GridID, { mask:mask, validation: controlType.validation, controlType: controlType.controlType, dataType: controlType.dataType, hashTags: hashTags, recalculate: recalculate });
                        }
                        ip_DisableSelection(GridID);

                        $().ip_Modal({ show: false });
                        

                    }
                },
                cancel: {
                    text: 'CANCEL', style: "background: #9c3b3b;background: -moz-linear-gradient(top,  #9c3b3b 0%, #752c2c 100%); background: -webkit-gradient(linear, left top, left bottom, color-stop(0%,#9c3b3b), color-stop(100%,#752c2c)); background: -webkit-linear-gradient(top,  #9c3b3b 0%,#752c2c 100%); background: -o-linear-gradient(top,  #9c3b3b 0%,#752c2c 100%); background: -ms-linear-gradient(top,  #9c3b3b 0%,#752c2c 100%); background: linear-gradient(to bottom,  #9c3b3b 0%,#752c2c 100%); filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#9c3b3b', endColorstr='#752c2c',GradientType=0 );", onClick: function () {
                        ip_GridProps[GridID].formatting.formatting = false;
                        ip_DisableSelection(GridID); $().ip_Modal({ show: false });
                    }
                }

            }
        });

        $('.formatting-tool').on('change', function () {
            $(this).addClass('ip-dirty');
        });
        
        $('#' + GridID + '_ControlType').ip_DropDown({
            init: true,
            defaultKey: options.enabledFormats.controlType,
            events: {
                onChange: function (e) {

                    var validationValue = $('#' + GridID + '_ValidationValue')[0];
                           
                    for (var i = 0; i < ip_GridProps[GridID].controlTypes.length; i++) {
                        if (e.key == ip_GridProps[GridID].controlTypes[i].key) {
                            $(validationValue).attr('placeholder', ip_GridProps[GridID].controlTypes[i].example);
                            $(validationValue).val('');
                            break;
                        }
                    }
                    
                    if (options.enabledFormats.validation.validationCriteria && options.enabledFormats.validation.validationCriteria.indexOf(e.key) != -1 && options.enabledFormats.validation.inputs) {
                        $(validationValue).val(options.enabledFormats.validation.inputs.join(', ').replace(/["]/g, ''));
                        $(validationValue).removeClass('tooltip');
                    }

                    if (e.key == "") { $(validationValue).hide(); } else { $(validationValue).show(); }
                    
                    
                }
            }, dropDownItems: ip_GridProps[GridID].controlTypes
        });
           
        var initMaskDropdown = function (dataType, defaultIndex) {
            
            var defaultKey = options.enabledFormats.mask;
            if (defaultIndex != null) { defaultKey = ""; }

            $('#' + GridID + '_Mask').html('');
            $('#' + GridID + '_Mask').ip_DropDown({
                init: true,
                defaultKey: defaultKey,
                defaultIndex: defaultKey == "" ? 0 : null,
                events: {
                    onChange: function (e) {



                    }
                },
                dropDownItems: function () {

                    var masks = ip_GetMasksForDataType(GridID, dataType);
                    
                    if (masks != null) {

                        var options = [];
                        for (var i = 0 ; i < masks.length; i++) {
                            var item = '<table style="width:100%;" cellpadding="0" cellspacing="0"><tr><td style="text-align:left;color:#46b3ff;">' + (masks[i].title == null ? '' : masks[i].title) + '</td><td style="text-align:right;padding:0px 4px 0px 10px;font-size:90%;">' + masks[i].mask + '</td><tr/></table>';
                            options.push({ key: masks[i].mask, value: item });
                        }
                        $('#' + GridID + '_Mask').show();
                        return options;
                        
                    }
                    
                    //No format options
                    $('#' + GridID + '_Mask').hide();
                    return [];
                }
                
            });
        }
        
        $('#' + GridID + '_HashTags').ip_TagInput({ tags: options.enabledFormats.hashTags });
        
        var initDataType = function () {

            //Get data types        
            var dropdownDataTypeNames = [{ value: 'default', key: '~default' }];
            var defaultKey = options.enabledFormats.dataType.dataType + '~' + (options.enabledFormats.dataType.dataTypeName == '' ? 'default' : options.enabledFormats.dataType.dataTypeName);

            initMaskDropdown(options.enabledFormats.dataType);

            for (var i = 0; i < ip_GridProps[GridID].dataTypes.length; i++) {

                var value = ip_GridProps[GridID].dataTypes[i].dataTypeName + (ip_GridProps[GridID].dataTypes[i].dataType != ip_GridProps[GridID].dataTypes[i].dataTypeName && ip_GridProps[GridID].dataTypes[i].dataType != null && ip_GridProps[GridID].dataTypes[i].dataType != '' ? ' (' + ip_GridProps[GridID].dataTypes[i].dataType + ')' : '');
                var key = (ip_GridProps[GridID].dataTypes[i].dataType == null ? '' : ip_GridProps[GridID].dataTypes[i].dataType) + '~' + ip_GridProps[GridID].dataTypes[i].dataTypeName;
                dropdownDataTypeNames[dropdownDataTypeNames.length] = { value: value, key: key }

            }

            $('#' + GridID + '_DataType').ip_DropDown({
                init: true,
                events: {
                    onChange: function (d) {
                        var arrDataType = d.key.split('~');
                        initMaskDropdown({ dataType: arrDataType[0], dataTypeName: arrDataType[1] }, 0);
                    }
                },
                defaultKey: defaultKey,
                dropDownItems: dropdownDataTypeNames
            });

        }
        initDataType();
               
        $('#' + GridID + '_Action').ip_DropDown({ init: true, defaultKey: options.enabledFormats.validation.validationAction, dropDownItems: [{ key: '', value: 'default' }, { key: 'allow', value: 'allow incorrect input' }, { key: 'warn', value: 'warn' }, { key: 'prevent', value: 'reject incorrect input' }] });

        if (options.enabledFormats.validation.inputs) {
            $('#' + GridID + '_ValidationValue').val(options.enabledFormats.validation.inputs.join(', ').replace(/["]/g, ''));
            $('#' + GridID + '_ValidationValue').removeClass('tooltip');
        }
        $('#' + GridID + '_ValidationValue').on('keydown', function () {
            if ($(this).hasClass('tooltip')) {                
                $(this).removeClass('tooltip');
            }
        });
        
        $('#' + GridID + '_Range_F').on('focus', function () { ip_GridProps[GridID].formatting.focusedControl = this; });
        $('#' + GridID + '_Range_F').val(ip_fxRangeToString(GridID, ranges));
        $('#' + GridID + '_Range_F').focus();
        
        ip_GridProps[GridID].formatting.formatting = true;
    }

    $.fn.ip_EnabledFormats = function (options) {

        var GridID = $(this).attr('id');

        return ip_EnabledFormats(GridID, options);

    }

    $.fn.ip_ResizeGrid = function (options) {

        var GridID = $(this).attr('id');

        ip_ResizeGrid(GridID, options.width);

    }
    
    $.fn.ip_CellInput = function (options) {

        var GridID = $(this).attr('id');

        ip_CellInput(GridID, options);

    }

    $.fn.ip_Render = function (options) {

        var GridID = $(this).attr('id');

        ip_ReRenderRows(GridID)

    }

    $.fn.ip_ReCalculate = function (options) {

        var GridID = $(this).attr('id');

        options = $.extend({
            range: [{ startRow: 0, startCol: 0, endRow: ip_GridProps[GridID].rows - 1, endCol: ip_GridProps[GridID].cols - 1 }]
        }, options);
        
        ip_ReCalculateFormulas(GridID, options)

    }

    $.fn.ip_Border = function (options) {
        //Sets a border in relation to the range
        var options = $.extend({

            range: null,//, //[{ startRow: null, startCol: null, endRow: null, endCol: null }] array of range object
            row: null,
            col: null,
            borderStyle: 'solid',
            borderSize: 1,
            borderColor: 'black',
            borderPlacement: '', //all, 'top', 'right', bottom, left, inner, outer, horizontal, vertical, none
            
            raiseEvent: true,
            render: true,

        }, options);

        if (options.borderPlacement == 'none') { options.borderStyle = 'none'; }

        var GridID = $(this).attr('id');
        var TransactionID = ip_GenerateTransactionID();
        var Effected = {}
        var borderPlacement = options.borderPlacement;
        var borderStyle = options.borderStyle;
        var borderSize = options.borderSize;
        var borderColor = options.borderColor;
        var borderTop = '';
        var borderRight = '';
        var borderBottom = '';
        var borderLeft = '';
        var Error = '';

        if (options.range == null && options.row != null && options.col != null) { options.range = [ip_rangeObject(options.row, options.col, options.row, options.col)] }

        //Validate range
        if (options.range == null && ip_GridProps[GridID].selectedRange.length == 0) { Error = 'Please specify a range'; }
        else if (options.range == null) {

            //Copied ranges
            options.range = new Array();
            for (var i = 0; i < ip_GridProps[GridID].selectedRange.length; i++) { options.range[i] = ip_rangeObject(ip_GridProps[GridID].selectedRange[i][0][0], ip_GridProps[GridID].selectedRange[i][0][1], ip_GridProps[GridID].selectedRange[i][1][0], ip_GridProps[GridID].selectedRange[i][1][1], null, options.col); }

        }

        options.range = ip_ValidateRangeObjects(GridID, options.range);

        Effected = { range: options.range, border: { borderStyle: borderStyle, borderPlacement: borderPlacement, borderSize: borderSize, borderColor: borderColor } }
        
        for (var i = 0; i < options.range.length; i++) {

            var range = options.range[i];
            var startRow = range.startRow;
            var endRow = range.endRow;
            var startCol = range.startCol;
            var endCol = range.endCol;

            var CellUndoData = ip_AddUndo(GridID, 'ip_Border', TransactionID, 'CellData', range, range, (i == 0 ? { row: range.startRow, col: range.startCol } : null));

            if (borderPlacement == 'top') { endRow = startRow; }
            else if (borderPlacement == 'right') { startCol = endCol; }
            else if (borderPlacement == 'bottom') { startRow = endRow; }
            else if (borderPlacement == 'left') { endCol = startCol; }
                        
            for (var r = startRow; r <= endRow; r++) {

                for (var c = startCol; c <= endCol; c++) {

                    ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, r, c));

                    if (borderPlacement == 'top') { ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, ['top']); }
                    else if (borderPlacement == 'right') { ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, ['right']); }
                    else if (borderPlacement == 'bottom') { ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, ['bottom']); }
                    else if (borderPlacement == 'left') { ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, ['left']); }
                    else if (borderPlacement == 'all') { ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, ['top', 'right', 'left', 'bottom']); }
                    else if (borderPlacement == 'none') { ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, ['top', 'right', 'left', 'bottom']); }
                    else if (borderPlacement == 'horizontal') {
                        if (r < endRow) { ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, ['bottom']); }
                        if (r > startRow) { ip_cellBorder(GridID, r, c, 'none', borderSize, borderColor, ['top']); }
                    }
                    else if (borderPlacement == 'vertical') {
                        if (c < endCol) { ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, ['right']); }
                        if (c > startCol) { ip_cellBorder(GridID, r, c, 'none', borderSize, borderColor, ['left']); }
                    }
                    
                    else if (borderPlacement == 'inner') {

                        var placements = [];

                        if (r != endRow) { placements.push('bottom'); }
                        if (c != endCol) { placements.push('right'); }

                        ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, placements);

                        placements = [];

                        if (r != startRow) { placements.push('top'); }
                        if (c != startCol) { placements.push('left'); }

                        ip_cellBorder(GridID, r, c, 'none', borderSize, borderColor, placements);
                        
                    }
                    else if (borderPlacement == 'outer') {

                        var placements = [];

                        if (r == startRow) { placements.push('top');  }
                        if (c == startCol) { placements.push('left'); }
                        if (r == endRow) { placements.push('bottom'); }
                        if (c == endCol) { placements.push('right');  }

                        ip_cellBorder(GridID, r, c, borderStyle, borderSize, borderColor, placements);
       

                    }
                    
                }

            }

        }
        


        if (Error == '') {

            if (options.render) { ip_ReRenderRanges(GridID, options.range); }
            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_Border', TransactionID, { Border: { Inputs: options, Effected: Effected } }); }

        }
        else { ip_RaiseEvent(GridID, 'warning', TransactionID, Error); return null; }

    }

    $.fn.ip_AddFormula = function (options) {
        //Registers a custom formula with ip sheets
        var options = $.extend({

            formulaName: null,
            functionName: null,
            tip: null,
            inputs: null,
            example: null,

        }, options);


        var GridID = $(this).attr('id');
        var functionName = options.functionName;
        var formulaName = options.formulaName;
        var tip = options.tip;
        var inputs = options.inputs;
        var example = options.example;

        if (formulaName == null || functionName == null) { return 'error - formulaName and functionName options must be specified'; }
        if (ip_GridProps[GridID].fxList[formulaName] == undefined) { ip_GridProps[GridID].fxList[formulaName] = { fxName: functionName }; }
        else { ip_GridProps[GridID].fxList[formulaName].fxName = functionName; }

        if (ip_GridProps[GridID].fxList[formulaName].fxInfo == null) { ip_GridProps[GridID].fxList[formulaName].fxInfo = { name: formulaName, tip: '', inputs: '', example: '' } }
        if (tip != null) { ip_GridProps[GridID].fxList[formulaName].fxInfo.tip = tip; }
        if (inputs != null) { ip_GridProps[GridID].fxList[formulaName].fxInfo.inputs = inputs; }
        if (example != null) { ip_GridProps[GridID].fxList[formulaName].fxInfo.example = example; }

        return true;
    }

    $.fn.ip_AddMask = function (options) {
        //Registers a custom masks with ip sheets
        var options = $.extend({

            dataType: null,
            input: function (value) { return value; },
            output: function (value) { return value; },
            mask: null,
            title: ''

        }, options);
        

        var GridID = $(this).attr('id');
        var input = options.input;
        var output = options.output;
        var dataType = options.dataType.toLowerCase();
        var mask = options.mask;
        var title = options.title;
        var newMask = null;

        if (dataType == null || mask == null) { return 'error - mask and dataType options must be specified'; }
        if (ip_GridProps[GridID].mask[dataType] == null) { ip_GridProps[GridID].mask[dataType] = []; }

        for (var i = 0; i < ip_GridProps[GridID].mask[dataType].length; i++) {

            if (ip_GridProps[GridID].mask[dataType][i].mask == mask) { newMask = ip_GridProps[GridID].mask[dataType][i]; break; }

        }

        if (newMask != null) { newMask.title = title }
        else { 
            newMask = { mask: mask, title: title }
            ip_GridProps[GridID].mask[dataType].push(newMask);
        }

        if (ip_GridProps[GridID].mask['input'] == null) { ip_GridProps[GridID].mask['input'] = {} }
        if (ip_GridProps[GridID].mask['output'] == null) { ip_GridProps[GridID].mask['output'] = {} }
        ip_GridProps[GridID].mask['input'][mask] = input;
        ip_GridProps[GridID].mask['output'][mask] = output;
        
        return true;
    }
    

    //----- CORE SUPPORTING UI PLUGINS --------------------------------------------------------------------------------------------------------------------------------------------------
    
    $.fn.ip_PositionElement = function (options) {

        var options = $.extend({

            relativeTo: document.body,
            position: null,// Note there the difference between [right,top] and [top,right] -> it is what it means unless axisContainment is off
            positionNAN: [],
            offset: { y: 0, x: 0 },
            axisContainment: true,
            windowContainment: true,
            animate: 600

        }, options);

        var ControlID = $(this).attr('id');
        
        if ($(document.body).children(this) == 0) { $(this).appendTo(document.body); }

        //Automatically decide the  best position
        if (options.relativeTo != null) {
            if (options.position == null) {

                var Q = 1;
                var rQ = { top: 0, left: 0, bottom: 0, right: 0 }
                var qWidth = Math.ceil($(window).width() / 3);
                var qHeight = Math.ceil($(window).height() / 3);
                var rWidth = $(options.relativeTo).outerWidth(true);
                var rHeight = $(options.relativeTo).outerWidth(true);
                var rTop = $(options.relativeTo).offset().top;
                var rLeft = $(options.relativeTo).offset().left;
                var rRight = rWidth + rLeft;
                var rBottom = rHeight + rTop;
                //var objWidth = $(this).width();
                //var objHeight = $(this).height();

                var q1 = { top: qHeight * 0, left: qWidth * 0, bottom: (qHeight * 0) + qHeight, right: (qWidth * 0) + qWidth }
                var q2 = { top: qHeight * 0, left: qWidth * 1, bottom: (qHeight * 0) + qHeight, right: (qWidth * 1) + qWidth }
                var q3 = { top: qHeight * 0, left: qWidth * 2, bottom: (qHeight * 0) + qHeight, right: (qWidth * 2) + qWidth }

                var q4 = { top: qHeight * 1, left: qWidth * 0, bottom: (qHeight * 1) + qHeight, right: (qWidth * 0) + qWidth }
                var q5 = { top: qHeight * 1, left: qWidth * 1, bottom: (qHeight * 1) + qHeight, right: (qWidth * 1) + qWidth }
                var q6 = { top: qHeight * 1, left: qWidth * 2, bottom: (qHeight * 1) + qHeight, right: (qWidth * 2) + qWidth }

                var q7 = { top: qHeight * 2, left: qWidth * 0, bottom: (qHeight * 2) + qHeight, right: (qWidth * 0) + qWidth }
                var q8 = { top: qHeight * 2, left: qWidth * 1, bottom: (qHeight * 2) + qHeight, right: (qWidth * 1) + qWidth }
                var q9 = { top: qHeight * 2, left: qWidth * 2, bottom: (qHeight * 2) + qHeight, right: (qWidth * 2) + qWidth }

                if (rTop <= q1.bottom && rLeft <= q1.right) { options.position = ['right', 'top']; Q = 1; }
                else if (rTop <= q2.bottom && rLeft <= q2.right) { options.position = ['left', 'top']; Q = 2; }
                else if (rTop <= q3.bottom && rLeft <= q3.right) { options.position = ['left', 'top']; Q = 3; }
                else if (rTop <= q4.bottom && rLeft <= q4.right) { options.position = ['right', 'center']; Q = 4; }
                else if (rTop <= q5.bottom && rLeft <= q5.right) { options.position = ['left', 'center']; Q = 5; }
                else if (rTop <= q6.bottom && rLeft <= q6.right) { options.position = ['left', 'center']; Q = 6; }
                else if (rTop <= q7.bottom && rLeft <= q7.right) { options.position = ['right', 'bottom']; Q = 7; }
                else if (rTop <= q8.bottom && rLeft <= q8.right) { options.position = ['left', 'bottom']; Q = 8; }
                else if (rTop <= q9.bottom && rLeft <= q9.right) { options.position = ['left', 'bottom']; Q = 9; }

            }
        }

        if (options.positionNAN) {
            for (var i = 0; i < options.positionNAN.length; i++) {

                if (options.positionNAN[i] == options.position[0] && options.position[0] == 'left') { options.position[0] = (Q >= 7 ? 'top' : 'bottom') }
                else if (options.positionNAN[i] == options.position[0] && options.position[0] == 'right') { options.position[0] = (Q >= 7 ? 'top' : 'bottom') }
                else if (options.positionNAN[i] == options.position[0] && options.position[0] == 'top') { options.position[0] = (Q == 2 || Q == 3 || Q == 5 || Q == 6 || Q == 8 || Q == 9 ? 'left' : 'right') }
                else if (options.positionNAN[i] == options.position[0] && options.position[0] == 'bottom') { options.position[0] = (Q == 2 || Q == 3 || Q == 5 || Q == 6 || Q == 8 || Q == 9 || options.positionNAN.indexOf('right') != -1 ? 'left' : 'right') }
                else if (options.positionNAN[i] == options.position[1] && options.position[1] == 'center') { options.position[1] = (Q == 2 || Q == 3 || Q == 5 || Q == 6 || Q == 8 || Q == 9 || options.positionNAN.indexOf('right') != -1 ? 'left' : 'right') }
                 
                if ((options.position[0] == 'bottom' || options.position[0] == 'top') && (options.position[1] == 'bottom' || options.position[1] == 'top')) { options.position[1] = (Q == 1 || Q == 2 || Q == 4 || Q == 5 || Q == 7 || Q == 8 ? 'left' : 'right'); }
                else if ((options.position[0] == 'left' || options.position[0] == 'left') && (options.position[1] == 'left' || options.position[1] == 'right')) { options.position[1] = 'bottom'; }
            }
        }
        
        var topR = (options.relativeTo == screen || options.relativeTo == window ? $(document).scrollTop() : $(options.relativeTo).offset().top);
        var leftR = (options.relativeTo == screen || options.relativeTo == window ? $(document).scrollLeft() : $(options.relativeTo).offset().left);
        var widthR = (options.relativeTo == screen || options.relativeTo == window ? window.innerWidth : $(options.relativeTo).outerWidth());
        var heightR = (options.relativeTo == screen || options.relativeTo == window ? window.innerHeight : $(options.relativeTo).outerHeight());
        var placeTop = 0;
        var placeLeft = 0;
        var axisContainment = { x: false, y: false };
        var returnObject = { callout: '', position: null };

        options.size = {
            width: $(this).outerWidth(true),
            height: $(this).outerHeight(true)
        }

        //Workout axis containment
        for (var i = 0; i < options.position.length; i++) {

            options.position[i] = options.position[i].toLowerCase().trim();

            if (options.axisContainment) {
                if (!axisContainment.x && (options.position[i] == 'left' || options.position[i] == 'right')) { axisContainment.y = true; }
                if (!axisContainment.y && (options.position[i] == 'top' || options.position[i] == 'bottom')) { axisContainment.x = true; }
            }

        }



        if (options.position.indexOf('inside') != -1) {
            placeTop = topR;
            placeLeft = leftR;
        }

        if (options.position.indexOf('top') != -1) {
            placeTop = topR - (options.position.indexOf('inside') == -1 ? options.size.height : 0);
            if (options.position.indexOf('left') == -1 && options.position.indexOf('right') == -1) { placeLeft = leftR; }
        }

        if (options.position.indexOf('bottom') != -1) {
            placeTop = topR + (options.position.indexOf('inside') == -1 ? heightR : Math.max(0, heightR - options.size.height));
            if (options.position.indexOf('left') == -1 && options.position.indexOf('right') == -1) { placeLeft = leftR; }
        }

        if (options.position.indexOf('left') != -1) {
            if (options.position.indexOf('top') == -1 && options.position.indexOf('bottom') == -1) { placeTop = heightR; }
            placeLeft = leftR - (options.position.indexOf('inside') == -1 ? options.size.width : 0);
        }

        if (options.position.indexOf('right') != -1) {
            if (options.position.indexOf('top') == -1 && options.position.indexOf('bottom') == -1) { placeTop = heightR; }
            placeLeft = leftR + (options.position.indexOf('inside') == -1 ? widthR : Math.max(0, widthR - options.size.width));
        }

        if (options.position.indexOf('center') != -1) {

            if (options.position.indexOf('top') == -1 && options.position.indexOf('bottom') == -1) {
                placeTop = topR;
                placeTop += Math.max(0, ((heightR - options.size.height) / 2));
            }

            if (options.position.indexOf('right') == -1 && options.position.indexOf('left') == -1) {

                placeLeft = leftR;
                placeLeft += Math.max(0, ((widthR - options.size.width) / 2));
            }

        }

        if (axisContainment.y) {
            if (options.position.indexOf('bottom') == -1 && placeTop < topR) { placeTop = topR; }
            if (options.position.indexOf('center') == -1 && options.position.indexOf('top') == -1 && (placeTop + options.size.height > topR + heightR)) { placeTop = topR + (heightR - options.size.height); } //topR + Math.max(0, heightR - options.size.height);
            if (options.position.indexOf('center') != -1 && options.position.indexOf('top') == -1 && (placeTop + options.size.height > topR + heightR)) { placeTop = topR + ((heightR - options.size.height) / 2); } //topR + Math.max(0, heightR - options.size.height);
        }
        if (axisContainment.x) {
            if (options.position.indexOf('right') == -1 && placeLeft < leftR) { placeLeft = leftR; }
            if (options.position.indexOf('center') == -1 && options.position.indexOf('left') == -1 && (placeLeft + options.size.width > leftR + widthR)) { placeLeft = leftR + (widthR - options.size.width); } //leftR + Math.max(0, widthR - options.size.width);
            if (options.position.indexOf('center') != -1 && (placeLeft + options.size.width > leftR + widthR)) { placeLeft = leftR + ((widthR - options.size.width) / 2); }
        }

        //placeTop += options.offset.top;
        if (options.position[0] == 'top') { placeTop -= options.offset.y; } //else { placeTop += options.offset.top; }
        else if (options.position[0] == 'bottom') { placeTop += options.offset.y; }
        else if (options.position[0] == 'inside') { placeTop += options.offset.y; }

        if (options.position[0] == 'left') { placeLeft -= options.offset.x; } //else { placeLeft += options.offset.left; }
        else if (options.position[0] == 'right') { placeLeft += options.offset.x; }
        else if (options.position[0] == 'inside') { placeLeft += options.offset.x; }

        //placeLeft += options.offset.left;

        if (placeLeft > $(window).width()) { placeLeft = $(window).width() - $(this).outerWidth(true); }
        if (placeTop > $(window).height()) { placeTop = $(window).height() - $(this).outerHeight(true); }

        $(this).animate({ top: placeTop, left: placeLeft }, options.animate, 'easeInOutQuint'); 


        if (axisContainment.y && options.position.indexOf('left') != -1 && options.position.indexOf('top') != -1) { returnObject.callout = 'rightTop' }
        else if (axisContainment.y && options.position.indexOf('left') != -1 && options.position.indexOf('bottom') != -1) { returnObject.callout = 'rightBottom' }
        else if (axisContainment.y && options.position.indexOf('right') != -1 && options.position.indexOf('top') != -1) { returnObject.callout = 'leftTop' }
        else if (axisContainment.y && options.position.indexOf('right') != -1 && options.position.indexOf('bottom') != -1) { returnObject.callout = 'leftBottom' }
        else if (axisContainment.y && options.position.indexOf('right') != -1 && options.position.indexOf('top') != -1) { returnObject.callout = 'leftTop' }
        else if (axisContainment.x && options.position.indexOf('top') != -1 && options.position.indexOf('left') != -1) { returnObject.callout = 'bottomLeft' }
        else if (axisContainment.x && options.position.indexOf('top') != -1 && options.position.indexOf('right') != -1) { returnObject.callout = 'bottomRight' }
        else if (axisContainment.x && options.position.indexOf('bottom') != -1 && options.position.indexOf('left') != -1) { returnObject.callout = 'topLeft' }
        else if (axisContainment.x && options.position.indexOf('bottom') != -1 && options.position.indexOf('right') != -1) { returnObject.callout = 'topRight' }
        else if (options.position.indexOf('top') != -1) { returnObject.callout = 'bottom'; }
        else if (options.position.indexOf('left') != -1) { returnObject.callout = 'right'; }
        else if (options.position.indexOf('right') != -1) { returnObject.callout = 'left'; }
        else if (options.position.indexOf('bottom') != -1) { returnObject.callout = 'top'; }




        returnObject.position = options.position;

        return returnObject;

    }
   
    $.fn.ip_Modal = function (options) {


        var options = $.extend({

            modalSessionID: null,
            modalType:null,
            buttons: null, //of type buttons (see below)
            title:'',
            message: '',
            relativeTo: screen,
            cssClass: 'ip_Modal',
            position: ['inside', 'center'], // Note there the difference between [right,top] and [top,right] -> it is what it means unless axisContainment is off
            positionNAN: [],
            offset: { x: 0, y: 0 },
            size: null, //{ width: 150, height: 250 },
            axisContainment: true,
            windowContainment: true,
            animate: 0,
            fade: 0,
            zIndex:101,
            speech: true,
            show: true,
            greyOut: screen,
            focus: 'input:text'

        }, options);

        var buttons = {

            ok: { text: 'OK', style: "background: #4d9c3b; background: -moz-linear-gradient(top,  #4d9c3b 0%, #1b6d00 100%);background: -webkit-gradient(linear, left top, left bottom, color-stop(0%,#4d9c3b), color-stop(100%,#1b6d00)); background: -webkit-linear-gradient(top,  #4d9c3b 0%,#1b6d00 100%); background: -o-linear-gradient(top,  #4d9c3b 0%,#1b6d00 100%); background: -ms-linear-gradient(top,  #4d9c3b 0%,#1b6d00 100%); background: linear-gradient(to bottom,  #4d9c3b 0%,#1b6d00 100%); filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#4d9c3b', endColorstr='#1b6d00',GradientType=0 ); ", onClick: null },
            cancel: { text: 'CANCEL', style: "background: #9c3b3b;background: -moz-linear-gradient(top,  #9c3b3b 0%, #752c2c 100%); background: -webkit-gradient(linear, left top, left bottom, color-stop(0%,#9c3b3b), color-stop(100%,#752c2c)); background: -webkit-linear-gradient(top,  #9c3b3b 0%,#752c2c 100%); background: -o-linear-gradient(top,  #9c3b3b 0%,#752c2c 100%); background: -ms-linear-gradient(top,  #9c3b3b 0%,#752c2c 100%); background: linear-gradient(to bottom,  #9c3b3b 0%,#752c2c 100%); filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#9c3b3b', endColorstr='#752c2c',GradientType=0 );", onClick: function () { $().ip_Modal({ show: false }) } }

        };


                       
        var control = $('#ip_modal');
        var greyOutObject = (options.greyOut == window || options.greyOut == screen ? document.body : options.greyOut);
        var inOpenSession = control.length > 0 && (control[0].modalSessionID == options.modalSessionID && control[0].modalSessionID != null);
               
        if (!options.show) {
            if (options.fade == 0) {

                $(control).removeClass('show');
                $(control).addClass('hide');

            }
            else { $(control).fadeOut();  }

            $(greyOutObject).ip_GreyOut({ show: false });
            if (control.length > 0) { control[0].modalSessionID = options.modalSessionID; }
            return;
        }

        if (!inOpenSession) {

            if (control.length == 0) { $(document.body).append('<div id="ip_modal" class="' + options.cssClass + ' ' + (options.anitmate > 0 ? 'hide' : '') + '"><div  class="ip_ModalTitle"></div><div  class="ip_ModalBody"></div><div  class="ip_ModalButtons"></div>'); control = $('#ip_modal')[0] }
            else { $(control).attr('class',options.cssClass + ' ' + (options.anitmate > 0 ? 'hide' : '')); control = $('#ip_modal')[0]; }
                        
        }

        control.modalSessionID = options.modalSessionID;
       
        $(control).css('z-index', options.zIndex);
        $(control).find('.ip_ModalTitle').html(options.title);
        $(control).find('.ip_ModalBody').html(options.message);                
        $(control).find('.ip_ModalButtons').html('');

        
        //Validate and insert Buttons
        for (var key in options.buttons) {
            if (buttons[key]) {
                if (options.buttons[key].text != undefined || options.buttons[key].onClick == null) { buttons[key].text = options.buttons[key].text; }
                if (options.buttons[key].style != undefined || options.buttons[key].onClick == null) { buttons[key].style = options.buttons[key].style; }
                if (options.buttons[key].onClick != undefined || options.buttons[key].onClick == null) { buttons[key].onClick = options.buttons[key].onClick; }
            }
            else {
                buttons[key] = options.buttons[key];
            }
        }
        for (var key in buttons) {
            if (typeof(buttons[key].onClick) == 'function') {
                $(control).find('.ip_ModalButtons').append('<input type="button" key="' + key + '" style="' + buttons[key].style + '" tabindex="1" class="button ip_ModalButton" value="' + buttons[key].text + '" />');
            }
        }
        
        //Button Events
        $('.ip_ModalButton').unbind('click');
        $('.ip_ModalButton').on('click', function (args) {

            $(this).focus();
            var key = $(this).attr('key');
            if (typeof (buttons[key].onClick) == 'function') { buttons[key].onClick() }
            
        });
        
        //Set size
        if (options.size != null) {

            $(this).outerWidth(options.size.width);
            $(this).outerHeight(options.size.height);

        }


        if (!inOpenSession) {

            //Position speech bubble
            var Position = { callout: '' }
            if (options.relativeTo != null) { Position = $(control).ip_PositionElement(options); }
            if (Position.callout == '') { Position.callout = 'bottomRight'; }

            $('#ip_ModalSpeech').attr('class', '');
            if (options.speech && Position.callout != '') {

                if ($('#ip_ModalSpeech').length == 0) { $(control).append('<b id="ip_ModalSpeech"></b>'); }
                $('#ip_ModalSpeech').addClass('ip_Speech ' + Position.callout);

            }

            if (options.fade == 0) {
                
                $(control).css('display', '');
                $(control).removeClass('hide');
                $(control).addClass('show');

            }
            else { $(control).fadeIn(); }

            $(control).draggable({
                cancel: ".no-drag",
                start:
                    function () {
                        $('#ip_ModalSpeech').attr('class', '');
                    }
            });
        }

        if (options.greyOut) { $(greyOutObject).ip_GreyOut(); }

        $(control).find(options.focus).focus();
        

        
       
    }

    $.fn.ip_GreyOut = function (options) {

        //? Shows a greyed out layer on top of a defined element

        var options = $.extend({

            titleClass: 'GreyOutTitle',
            title: '',
            show: true,
            color: 'white',
            fade: true,
            onclick: null,
            opacity: 0.5,
            zIndex: 100,
            showProgress: true
        }, options);

        //code starts here            
        return this.each(function () {

            var ControlID = $(this).attr('id');
            var dvGreyOut = $(this).children('.ip_GreyOut');

            //Create a greyout div if it does not exist
            if (dvGreyOut.length == 0) {
                dvGreyOut = $('<div class="ip_GreyOut" style="display:none;position:absolute;top:0px;left:0px;background-color:' + options.color + ';opacity:' + options.opacity + ';filter: alpha(opacity=' + (100 * options.opacity) + ');z-index:' + options.zIndex + ';width:100%;height:100%;"></div>');
            }

            $(dvGreyOut).appendTo(this);
            dvGreyOut = $(this).children('.ip_GreyOut');

            var spTitle = $(this).children('#ec_GreyOutTitle');
            if (spTitle.length == 0) {

                $('<div id="ec_GreyOutTitle" class="GreyOutTitle" style="position: absolute;"></div>').appendTo(this);
                spTitle = $(this).children('#ec_GreyOutTitle');

            }


            if (options.show) {

                if (options.title != '') {

                    var top = ($(dvGreyOut).outerHeight() - $(spTitle).outerHeight()) / 2;

                    if (this == document.body) {
                        top = ($(window).outerHeight() - $(spTitle).outerHeight()) / 2;
                        $(spTitle).attr('style', 'position: fixed;');
                    }
                    else { $(spTitle).attr('style', 'position: absolute;'); }

                    $(spTitle).removeClass();
                    $(spTitle).addClass(options.titleClass);
                    $(spTitle).show();
                    $(spTitle).html(options.title);
                    $(spTitle).css('top', top + 'px');
                    $(spTitle).css('left', (($(dvGreyOut).outerWidth() - $(spTitle).outerWidth()) / 2) + 'px');
                    $(spTitle).css('z-index', parseInt($(dvGreyOut).css('z-index')) + 1);

                }
                else { $(spTitle).hide(); }


                if (options.fade) { $(dvGreyOut).fadeIn(); }
                else { $(dvGreyOut).show(); }

            }
            else {

                $(spTitle).hide();
                if (options.fade) { $(dvGreyOut).fadeOut(); }
                else { $(dvGreyOut).hide(); }

            }

            $(dvGreyOut).unbind('click');
            if (typeof options.onclick == "function") {
                $(dvGreyOut).click(function () { options.onclick(); });
            }

        });
    }

    $.fn.ip_TagInput = function (options) {

        var options = $.extend({
            tags: null, //string or []
            includeHashTag: true,            
        }, options);

        var self = this;
        var ControlID = $(this).attr('id');
        var Controls = this;

        for (i = 0; i < Controls.length; i++) {

            var Control = Controls[i];
            var TagInput;
            var loadTags = null;
            var placeHolder = $(Control).attr('placeholder');

            if (typeof (options.tags) == 'string') { options.tags = options.tags.split(','); }
            if (options.tags != null) { loadTags = options.tags; options.tags = []; }
            else { options.tags = []; }

            if (placeHolder == undefined) { placeHolder = ''; }

            $(Control).html('');

            $(Control).append('<div class="ip_TagInput"><div contenteditable="true" tabindex="0" class="ip_TagInputTextBox" /><div class="ip_TagInputPlaceholder">' + placeHolder + '</div></div>');

            var TagInput = $(Control).children('.ip_TagInput')[0];
            var TagInputTextBox = $(TagInput).children('.ip_TagInputTextBox')[0];

            Control.addTag = function (el, init) {

                if (el == undefined) { el = TagInputTextBox; }

                var tagtext = (typeof (el) == 'object' ? $(el).text() : el);

                if (tagtext == '' || tagtext == undefined || tagtext == null) { return }

                if (tagtext[0] != '#' && options.includeHashTag) { tagtext = '#' + tagtext; }
                                
                if (options.tags.indexOf(tagtext) == -1) {
                    $('<div class="ip_TagInputTag">' + tagtext + '<div class="ip_TagInputTagClose">x</div></div>').insertBefore(TagInputTextBox);
                    options.tags.push(tagtext);
                }

                if (typeof (el) == 'object') { $(el).text(''); }

                if (!init) { $(Control).trigger('AddTag', { tag: tagtext, tags: options.tags }); }
                if (!init) { $(Control).trigger('change', { tag: tagtext, tags: options.tags, action: 'AddTag' }); }
            }

            Control.removeTag = function (index, tagObj) {

                var tags = $(TagInput).children('.ip_TagInputTag');

                if (index == undefined && tagObj == undefined) { index = tags.length - 1; }
                else if (tagObj != undefined && tagObj != null) { index = $(tags).index(tagObj); }

                var tagtext = $(tags[index]).text();

                $(tags[index]).remove();
                options.tags.splice(index, 1);

                $(Control).trigger('RemoveTag', { tag: tagtext, tags: options.tags });
                $(Control).trigger('change', { tag: tagtext, tags: options.tags, action:'RemoveTag' });
                
            }

            Control.val = function () {

                var tags = '';
                for (var i = 0; i < options.tags.length; i++) {

                    tags += options.tags[i];
                    if (i != options.tags.length - 1) { tags += ','; }

                }
                return tags;
            }

            if (loadTags != null) {
                for (var t = 0; t < loadTags.length; t++) { Control.addTag(loadTags[t].trim(), true); }
            }

            $(Control).unbind('keydown');
            $(Control).unbind('click');
            $(Control).unbind('focus');
            $(Control).unbind('focusout');

            $(Control).on('click', function () {

                ip_SetCursorPos(null, TagInputTextBox);

            });            
            $(Control).on('keydown', '.ip_TagInputTextBox', function (e) {
                switch (e.keyCode) {
                    case 13:
                        Control.addTag(this);
                        return false;
                    case 8:
                        var tag = $(this).text();
                        if (tag == '') { Control.removeTag(); }
                        break;
                    case 188:
                        Control.addTag(this);
                        return false;
                }
            });
            $(Control).on('click', '.ip_TagInputTagClose', function (e) { Control.removeTag(null, $(this).parent()[0]); });            
            $(Control).on('focus', '.ip_TagInputTextBox', function () {
                $(TagInput).find('.ip_TagInputPlaceholder').hide();
            });            
            $(Control).on('focusout', '.ip_TagInputTextBox', function () {
                Control.addTag();
                if (options.tags.length == 0 && $(TagInputTextBox).text() == '') { $(TagInput).find('.ip_TagInputPlaceholder').show(); }
            });

            if (options.tags.length != 0) { $(TagInput).find('.ip_TagInputPlaceholder').hide(); }

            Control.options = options;
        }

    }

    $.fn.ip_DropDown = function (options) {

        var options = $.extend({

            select: null,
            editable: true,
            defaultIndex: null, //will set the control to this index
            defaultValue: null, //will set the control to this values index
            defaultKey: null, //will set the control to this keys index
            text: null, //defaults the control with this text without setting index
            key: null,    //defaults the control with this key without setting index                      
            dropDownItems: [], // [{ key:'', value:'' }]
            init: false,
            events: { onChange: null },


        }, options);

        var Control;
        var ControlID;
        var DropDownText;
        var DropDownItemContainer;
        var DropDownIcon;
        var DropDownItems;
        var Error = '';
        var elControl = this[0];

        elControl.initObjects = function () {

            Control = this;
            ControlID = $(this).attr('id');
            DropDownText = $(this).find('.ip_DropDownText');
            DropDownItemContainer = $(this).find('.ip_DropDownItems');
            DropDownIcon = $(this).find('.ip_DropDownIcon');
            DropDownItems = $(DropDownItemContainer).find('.ip_DropDownItem');
            $(Control).attr('key', '');
            Error = '';
        }

        this[0].initObjects();

        this.val = function () { return $(this).attr('key'); }

        if (options.select != null) {
            if (typeof (options.select) == 'number') { $(DropDownItems)[options.select].click(); }
            $(this).find('.ip_DropDownItem').each(function () { if ($(this).attr('key') == options.select) { $(this).click(); } });
            return;
        }
        if (typeof (options.dropDownItems) == 'function') { options.dropDownItems = options.dropDownItems(); }
        if (options.dropDownItems == null) { options.dropDownItems = []; }
        if (DropDownItemContainer.length == 0) { $(this).append('<div class="ip_DropDownItems"></div>'); DropDownItemContainer = $(this).find('.ip_DropDownItems'); }
        if (DropDownIcon.length == 0) { $(this).append('<div class="ip_DropDownIcon down"></div>'); DropDownIcon = $(this).find('.ip_DropDownIcon'); }
        if (DropDownText.length == 0) { $(this).append('<div class="ip_DropDownText" contenteditable="true"></div>'); DropDownText = $(this).find('.ip_DropDownText'); }

        DropDownItemContainer = DropDownItemContainer[0];
        DropDownText = DropDownText[0];
        DropDownIcon = DropDownIcon[0];

        $(DropDownText).css('width', $(Control).outerWidth() - parseInt($(DropDownText).css('margin-left').replace(/[a-z]/gi, '')) - ($(DropDownIcon).outerWidth() * 2) + 'px')
        $(DropDownText).attr('placeholder', $(Control).attr('placeholder'));

        if (options.editable) { $(DropDownText).css('content-editable', 'true'); }

        if (options.dropDownItems.length > 0) {

            for (var i = 0; i < options.dropDownItems.length; i++) {

                if (options.defaultKey != null && options.defaultKey == options.dropDownItems[i].key) { options.defaultIndex = i; }
                else if (options.defaultValue != null && options.defaultValue == options.dropDownItems[i].value) { options.defaultIndex = i; }
                $(DropDownItemContainer).append('<div key="' + options.dropDownItems[i].key + '" class="ip_DropDownItem ' + (options.dropDownItems[i].title ? 'title' : '') + ' ' + ((options.text != null && options.dropDownItems[i].value.trim().toLowerCase() == options.text.trim().toLowerCase()) || i == options.defaultIndex ? 'hover' : '') + '" title="' + (options.dropDownItems[i].tooltip ? options.dropDownItems[i].tooltip : '') + '">' + options.dropDownItems[i].value + '</div>');

            }

        }
        DropDownItems = $(DropDownItemContainer).find('.ip_DropDownItem');

        $(DropDownIcon).css('top', (($(this).outerHeight() - parseInt($(DropDownIcon).css('border-top-width').replace('px', ''))) / 2) + 'px');
        $(DropDownItemContainer).css('min-width', $(this).width());

        $(this).attr('defaultText', options.text);

        if (options.init) {

            DropDownItemContainer.timeout = null;
            DropDownItemContainer.show = function (args) {

                clearTimeout(DropDownItemContainer.timeout);
                if (!$(DropDownItemContainer).is(":visible")) {

                    var args = $.extend({ delay: 150, scrollTop: null }, args);

                    if (args.scrollTop != null) { $(DropDownItemContainer).scrollTop(args.scrollTop); }

                    //Calculate height
                    $(DropDownItemContainer).css('max-height', '');
                    $(DropDownItemContainer).show();
                    var top = $(DropDownItemContainer).offset().top;
                    var height = $(DropDownItemContainer).outerHeight();
                    $(DropDownItemContainer).hide();

                    var windowHeight = $(window).height();
                    var containerBottom = top + height;
                    var maxHeight = '';

                    if (containerBottom > windowHeight) { maxHeight = windowHeight - top - 20; }
                    if (maxHeight != '') { $(DropDownItemContainer).css('max-height', maxHeight + 'px'); }

                    $(DropDownItems).each(function () {
                        $(this).removeClass('nohover');
                        $(this).removeClass('hover');
                        $(this).show();
                    });

                    $(DropDownItemContainer).slideDown(args.delay);

                }
            }

            $(DropDownText).on('keypress', function (e) {
                if (e.keyCode == 13) { return false; }
            });
            $(DropDownText).on('keyup', function (e) {

                if (e.keyCode != 40 && e.keyCode != 38 && e.keyCode != 13) {
                    var value = $(this).text().toLowerCase().trim();

                    $(Control).closest('.ip_DropDown').attr('key', value);

                    $(DropDownItems).each(function (e) {

                        if ($(this).text().toLowerCase().indexOf(value) == -1) { $(this).hide(); }
                        else { $(this).show(); }

                    })
                }
                if (e.keyCode == 13) {

                    $(DropDownItems).each(function () {
                        var clicked = false;
                        if ($(this).hasClass('hover')) {
                            $(this).click();
                            return;
                        }
                    })
                    return false;

                }


            });

            $(DropDownText).on('keydown', function (e) {



                if (e.keyCode == 40) { //down

                    if (DropDownItems.length == 0) { return; }

                    DropDownItemContainer.show({ delay: 0 });

                    var found = null;
                    var set = null;

                    for (var i = 0; i < DropDownItems.length; i++) {

                        if (found != null && $(DropDownItems[i]).is(":visible")) {
                            $(DropDownItems[i]).addClass('hover');
                            $(DropDownItems[i]).removeClass('nohover');
                            set = i;
                            found = null;
                        }
                        else if ($(DropDownItems[i]).hasClass('hover')) {
                            found = i;
                            $(DropDownItems[i]).removeClass('hover');
                            $(DropDownItems[i]).addClass('nohover');
                        }

                    }

                    if (set == null) {
                        $(DropDownItems[0]).addClass('hover');
                        $(DropDownItems[0]).removeClass('nohover');
                        found = 0;
                        set = 0;
                    }

                    //Make sure the element in focus maintains scroll focus
                    var scrollTop = $(DropDownItemContainer).scrollTop();
                    var containerHeight = $(DropDownItemContainer).height();

                    if ($(DropDownItems[set]).position().top > containerHeight || $(DropDownItems[set]).position().top < 0) {
                        var newScrollTop = scrollTop + $(DropDownItems[set]).position().top;
                        $(DropDownItemContainer).animate({ scrollTop: newScrollTop }, 100);
                    }

                    $(DropDownText).text($(DropDownItems[set]).text());
                    $(DropDownText).select();

                }
                else if (e.keyCode == 38) { //down

                    if (DropDownItems.length == 0) { return; }

                    DropDownItemContainer.show({ delay: 0 });

                    var found = null;
                    var set = null;

                    for (var i = DropDownItems.length - 1; i >= 0; i--) {

                        if (found != null && $(DropDownItems[i]).is(":visible")) {
                            $(DropDownItems[i]).addClass('hover');
                            $(DropDownItems[i]).removeClass('nohover');
                            set = i;
                            found = null;
                        }
                        else if ($(DropDownItems[i]).hasClass('hover')) {
                            found = i;
                            $(DropDownItems[i]).removeClass('hover');
                            $(DropDownItems[i]).addClass('nohover');
                        }

                    };

                    if (set == null) {
                        $(DropDownItems[DropDownItems.length - 1]).addClass('hover');
                        $(DropDownItems[DropDownItems.length - 1]).removeClass('nohover');
                        found = DropDownItems.length - 1;
                        set = DropDownItems.length - 1;
                    }

                    //Make sure the element in focus maintains scroll focus
                    var scrollTop = $(DropDownItemContainer).scrollTop();
                    var containerHeight = $(DropDownItemContainer).height();

                    if ($(DropDownItems[set]).position().top > containerHeight || $(DropDownItems[set]).position().top < 0) {
                        var newScrollTop = scrollTop + $(DropDownItems[set]).position().top - $(DropDownItemContainer).height() + $(DropDownItems[set]).height();
                        $(DropDownItemContainer).animate({ scrollTop: newScrollTop }, 100);
                    }


                    $(DropDownText).text($(DropDownItems[set]).text());
                    $(DropDownText).select();


                }


            });

            $(this).unbind('click');
            $(this).on('click', function () {

                //solves the problem of cloning the dropdown control
                this.initObjects();
                DropDownItemContainer[0].show({ scrollTop: 0 });
                $(DropDownText).focus();

            });

            $(this).unbind('mouseenter');
            $(this).on('mouseenter', function () {

                clearTimeout(DropDownItemContainer.timeout);

            });

            $(this).unbind('mouseleave');
            $(this).on('mouseleave', function () {

                DropDownItemContainer.timeout = setTimeout(function () {
                    $(DropDownItemContainer).slideUp(300);
                    //$(DropDownItemContainer).attr('closeTimeout', setTimeout(function () { $(DropDownItemContainer).slideUp(300); }, 100));
                }, 300);

            });

            $(this).unbind('mouseenter.ip_controls');
            $(this).on('mouseenter.ip_controls', '.ip_DropDownItem', function () {

                if (!$(this).hasClass('title')) {
                    $(this).addClass('hover');
                    $(this).removeClass('nohover');
                }

            });

            $(this).unbind('mouseleave.ip_controls');
            $(this).on('mouseleave.ip_controls', '.ip_DropDownItem', function () {

                clearTimeout(DropDownItemContainer.timeout);
                //clearTimeout($(this).parent().attr('closeTimeout'));

                $(this).addClass('nohover');
                $(this).removeClass('hover');

            });

            $(this).unbind('click.ip_controls');
            $(this).on('click.ip_controls', '.ip_DropDownItem', { init: false }, function (e) {

                var el = this;

                if ($(el).hasClass('title')) {
                    var index = $.inArray(el, DropDownItems);
                    if (index < DropDownItems.length) {
                        el = DropDownItems[index + 1];
                    }
                }

                $(el).addClass('selected');

                var key = $(el).attr('key');
                var value = $(el).html();

                $(el).closest('.ip_DropDown').attr('key', key);
                $(el).closest('.ip_DropDown').find('.ip_DropDownText').html(value);
                $(el).closest('.ip_DropDownItems').slideUp(100);

                if (!initControl) {

                    $(Control).addClass('ip-dirty');
                    if (options.events.onChange != null) { options.events.onChange({ key: key, value: value }); }

                }

                e.stopPropagation();

            });

            //Simulate the onchange event for contenteditable
            var DropDownTextVal = '';
            $(this).unbind('focus.ip_controls');
            $(this).unbind('blur.ip_controls');
            $(this).on('focus', '.ip_DropDownText', function () { DropDownTextVal = $(this).text(); });
            $(this).on('blur', '.ip_DropDownText', function () { if (DropDownTextVal != $(this).text()) { $(Control).addClass('ip-dirty'); } });
        }

        var initControl = true;

        if (options.text != null) { $(DropDownText).html(options.text); }
        if (options.key != null) { $(this).attr('key', options.key); }
        if (options.defaultKey != null) { $(this).attr('key', options.defaultKey); }
        if (options.defaultIndex != null && DropDownItems.length > options.defaultIndex) { $(DropDownItems)[options.defaultIndex].click();  }
        initControl = false;

        if (Error != '') { ip_RaiseEvent(GridID, 'warning', null, Error); }
    }

    $.fn.ip_FloatingToolbar = function (options) {

        var options = $.extend({

            menus: [],
            relativeTo: document.body,
            position: null, // ['right', 'top'],// Note there the difference between [right,top] and [top,right] -> it is what it means unless axisContainment is off
            positionNAN: null,
            offset: { x: 15, y: 15 },
            size: null, //{ width: 150, height: 250 },
            axisContainment: true,
            windowContainment: true,
            animate: 500,
            events: { handleDblClick: null, onLoad: null, click: null },
            mode: 'LeftClick',
            setAnchor: null,
            resetPosition: false,
            speech: true,
            toolbarHandle: true,
            theme: 'theme-dark'

        }, options);

        var ControlID = $(this).attr('id');
        var ToolbarHandle = $(this).find('.ip_FloatingToolbarHandle');
        var FloatingToolbarContainer = $(this).find('.ip_FloatingToolbarContainer');
        var Error = '';
        var Position = { callout: '' }
        var isVisible = $(this).is(":visible");

        if (options.resetPosition && this[0].anchor) {

            options.relativeTo = this[0].anchor.relativeTo;
            options.position = this[0].anchor.position;
            options.offset = this[0].anchor.offset;
            options.speech = this[0].anchor.speech;
            if (this[0].anchorX && this[0].anchorY) {
                options.relativeTo = null;
                $(this).animate({ top: this[0].anchorY, left: this[0].anchorX }, options.animate, 'easeInOutQuint');
            }
            else { Position = $(this).ip_PositionElement(options); }

            if (options.menus.length == 0) { return; }
        }

        $(this).appendTo(document.body);
        $(this).attr('mode', options.mode);
        $(this).addClass('theme-dark');

        if (ToolbarHandle.length == 0 && options.toolbarHandle) { $(this).prepend('<div class="ip_FloatingToolbarHandle" title="Double click to reset position"></div>'); ToolbarHandle = $(this).find('.ip_FloatingToolbarHandle'); }
        if (FloatingToolbarContainer.length == 0) { $(this).append('<div class="ip_FloatingToolbarContainer"></div>'); FloatingToolbarContainer = $(this).find('.ip_FloatingToolbarContainer')[0]; }

        //Load up menus
        if (options.menus.length > 0) {

            $(FloatingToolbarContainer).find('.ip_FloatingToolbarMenu').hide();
            for (var i = 0; i < options.menus.length; i++) {
                if ($(FloatingToolbarContainer).find(options.menus[i]).length == 0) {

                    $(options.menus[i]).clone(true, true).appendTo(FloatingToolbarContainer);

                }
                $(FloatingToolbarContainer).find(options.menus[i]).show();
            }

        }

        if (options.size != null) {

            $(this).outerWidth(options.size.width);
            $(this).outerHeight(options.size.height);

        }

        if (!isVisible) { options.animate = 0; }
        if (options.relativeTo != null) { Position = $(this).ip_PositionElement(options); }
        else if (options.position != null && options.position.top && options.position.left) { $(this).animate({ top: options.position.top, left: options.position.left }, options.animate, 'easeInOutQuint'); }


        $('#ip_FloatingToolbarSpeech').attr('class', '');
        if (options.speech && Position.callout != '') {

            var SpeechClass = Position.callout;

            if ($('#ip_FloatingToolbarSpeech').length == 0) { $(this).append('<b id="ip_FloatingToolbarSpeech"></b>'); }
            $('#ip_FloatingToolbarSpeech').addClass('ip_Speech ' + SpeechClass);
            $('#ip_FloatingToolbarSpeech').addClass(options.theme);

        }

        $(this).draggable({

            start: function () {

                $('#ip_FloatingToolbarSpeech').removeClass("top");
                $('#ip_FloatingToolbarSpeech').removeClass("topLeft");
                $('#ip_FloatingToolbarSpeech').removeClass("topRight");
                $('#ip_FloatingToolbarSpeech').removeClass("Bottom");
                $('#ip_FloatingToolbarSpeech').removeClass("BotomLeft");
                $('#ip_FloatingToolbarSpeech').removeClass("BottomRight");
                $('#ip_FloatingToolbarSpeech').removeClass("right");
                $('#ip_FloatingToolbarSpeech').removeClass("rightTop");
                $('#ip_FloatingToolbarSpeech').removeClass("rightBottom");
                $('#ip_FloatingToolbarSpeech').removeClass("left");
                $('#ip_FloatingToolbarSpeech').removeClass("leftTop");
                $('#ip_FloatingToolbarSpeech').removeClass("leftBottom");
                $('#ip_FloatingToolbarSpeech').removeClass("ip_Speech");

            },
            stop: function (args) {

                this.anchorY = $(this).position().top;
                this.anchorX = $(this).position().left;
            },
            handle: '.ip_FloatingToolbarHandle',
            containment:
            [$(document).scrollLeft(), $(document).scrollTop(),
            $(document).scrollLeft() + $(window).width() - $(this).outerWidth(),
            $(document).scrollTop() + $(window).height() - $(this).outerHeight()]
        });

        //SETUP EVENTS
        if (typeof options.events.dblclick == "function") {

            $(this).unbind('dblclick');
            $(this).on('dblclick', function () { options.events.dblclick() });

        }

        $(this).unbind('click');
        if (typeof options.events.click == "function") {

            $(this).on('click', function () { options.events.click() });

        }

        $(ToolbarHandle).unbind('click');
        $(ToolbarHandle).on('click', function (e) { e.stopPropagation(); });

        if (options.setAnchor) {

            this[0].anchor = {
                relativeTo: options.relativeTo,
                position: options.position,
                offset: options.offset,
                speech: options.speech
            }
            this[0].anchorX = null,
            this[0].anchorY = null
        }

        if (typeof options.events.onLoad == "function") { options.events.onLoad(this) }

        if (!isVisible) { $(this).fadeIn(100); }

        if (Error != '') { ip_RaiseEvent(GridID, 'warning', null, Error); }
    }

    $.fn.ip_FooterDialog = function (options) {

        //? Briefly shows a dialog message at the bottom of the window/ element

        var options = $.extend({

            Message: '',
            BackColor: '#2fad75',
            cssClass: '',
            Duration: 5000,

        }, options);

        //code starts here            
        return this.each(function () {


            clearTimeout(this.ip_ShowFooterAlert_Timeout_JQ);

            var FooterObject = $(this).children('.ip_FooterDialog');
            if (FooterObject.length == 0) {

                Footer = '<div class="ip_FooterDialog"><div class="ip_FooterDialog_Text"></div></div>';
                $(this).append(Footer);
                FooterObject = $('.ip_FooterDialog')[0];

            }
            else { FooterObject = FooterObject[0]; }

            if (options.cssClass) {
                $(FooterObject).attr('class', 'ip_FooterDialog');
                $(FooterObject).addClass(options.cssClass);
            }
            $(FooterObject).children('.ip_FooterDialog_Text').html(options.Message);
            $(FooterObject).css('background-color', options.BackColor);

            $(FooterObject).slideDown(200);

            this.ip_ShowFooterAlert_Timeout_JQ = setTimeout(function () { $(FooterObject).slideUp(200); }, options.Duration);

        });
    }


    //----- DATA OBJECTS --------------------------------------------------------------------------------------------------------------------------------------------------


    $.fn.ip_gridProperties = function (options) {

        var dimensions = $.extend({

            defaultRowHeight: 30, //Default height for a row
            defaultColWidth: 130, //Default width for a column
            gridHeight: 0, //The overall height of the entire grid
            gridWidth: 0, //The overall with of the entire grid
            scrollHeight: 0,
            scrollWidth: 0,
            columnSelectorHeight: 30, //Column selector height
            rowSelectorWidth: 50, //Row selector width
            defaultBorderHeight: 1,
            accumulativeScrollHeight: 0,
            accumulativeScrollWidth: 0,
            fBarHeight:40

        }, options.dimensions);


        var undo = $.extend({

            maxTransactions: 10,
            maxUndoRangeSize: 10000,
            undoStack: {}, //This contains the undo data for the grid
            

        }, options.undo);

        var callbacks = $.extend({

            onLoad: null

        }, options.callbacks);


        var options = $.extend({
            
            id: null, //grid id   
            publicKey: null,
            index: 0, //Index within grid object  
            starred: false,
            showColSelector: true,
            showRowSelector: true,
            showGridResizerX: true,
            hoverCell: null, //td cell object
            hoverColumnIndex: null, //td cell header object
            hoverRowIndex: null,
            fxBar: null, //formula bar element to use - may be a custom fbar element
            fxList: {},
            selectedCell: null, //td cell object
            selectedColumn: new Array(), //Array of column indexes
            selectedRow: new Array(), //Array of row indexes
            selectedRange: new Array(), //[[startrow,startcol],[endrow,endcol], [pvivitstartrow,pivitendrow], [pivitTop,pivitLeft]]
            selectedRangeIndex: -1, //The last range to be clicked or edited
            copiedRange: new Array(),  //[[startrow,startcol],[endrow,endcol], 'copy or cut']            
            highlightRange: new Array(), //[[startrow,startcol],[endrow,endcol]]
            cut: false, //Are we in cut mode
            rows: 0,
            cols: 0,
            frozenRows:0,
            frozenCols: 0,
            loadedRows: 80, //20
            loadedCols: 20, //13
            scrollX: null,
            scrollY: null,
            scrollXInterval: 1,
            scrollYInterval: 1,
            scrolling: false,
            controlTypes: [{ key: '', value: 'default', example: 'Must contain value, e.g. red', defaultInputs: null }, { key: 'text', value: 'cell', example: 'Must contain value, e.g. red [or] red,green', defaultInputs: null }, { key: 'dropdown', value: 'dropdown - CSV or Range', example: 'Must contain value, e.g.: a1:a10, red, green', defaultInputs: null }], //supported controls
            rowData: new Array(), // array of ip_rowObject
            colData: new Array(), // array of ip_colObject
            dataTypes: new Array(), //array of the data type object
            mergeData: {}, // key value pair of ip_mergeObject
            indexedData: {
                colSymbols: ip_ColumnSymbols(26), //column symboles 0 = a, 1 = b, 2 = c
                cellIndex: {}, //Static cell indexes
                formulaData: {} //stores formula objects, cells link to this
            },
            //colSymbols: ip_ColumnSymbols(200),
            resizing: false,
            mouseButton: 0,
            shiftKey: false,
            ctrlKey: false,
            scrollAnimate: false,
            currentMousePos: { pageX: -1, pageY: -1 },
            mask: { },
            undo: undo,            
            editing: {
                row: -1,
                col: -1,
                editing: false,
                element: null,
                contentType: '',
                carret: null,
                editTool: null,
                selectionState: false,
                cancelFocusOut: false

            },
            formatting: {
                formatting: false,
                range: null,
                focusedControl: null,

            },
            dimensions: dimensions,
            timeouts: {
                scrollAnimateInterval: null,
                keyDownTimeout: null,
                scrollYCompleteTimeout: null,
                scrollXCompleteTimeout: null,
                textEditToolBlurTimeout: null,
                disableSelectionTimout: null,
            },
            callbacks: callbacks,
            events: {
                ip_InitScrollX_Large_MouseDown: null,
                ip_InitScrollX_Large_MouseMove: null,
                ip_InitScrollY_Large_MouseDown: null,
                ip_InitScrollY_Large_MouseMove: null,
                ip_InitScrollY_Large_scrollTimeout: null,
               
                FrozenHandle_MouseDown: null,
                FrozenHandle_MouseUp: null,
                FrozenHandle_MouseLeave: null,
                FrozenHandle_DragStart: null,
                FrozenHandle_Drag: null,
                FrozenHandle_DragEnd: null,

                rowFrozenHandle_MouseDown: null,
                rowFrozenHandle_MouseUp: null,
                rowFrozenHandle_MouseLeave: null,

                colResizer_MouseDown: null,
                colResizer_MouseUp: null,
                colResizer_MouseLeave: null,
                colResizer_dblClick: null,
                colResizer_Start: null,
                colResizer_Drag: null,
                colResizer_End: null,

                rowResizer_MouseDown: null,
                rowResizer_MouseUp: null,
                rowResizer_MouseLeave: null,
                rowResizer_dblClick: null,
                rowResizer_Start: null,
                rowResizer_Drag: null,
                rowResizer_End: null,

                moveRange_mouseMove: null,
                moveRange_mouseUp: null,
                
                showRangeMove_highlightMovePosition: null,

                scrollAnimate_MouseMove: null,
                scrollAnimate_MouseUp: null,

                scrollToXY_RangeTimeout: null,

                textEditToolDropdown_FocusIn: null,
                textEditToolDropdown_Click: null,
                textEditToolDropdown_Mouseenter:null,

                textEditTool_FocusIn: null,
                textEditTool_FocusOut: null,
                textEditTool_KeyDown: null,
                textEditTool_KeyPress: null,
                textEditTool_KeyUp: null,
                textEditTool_MouseUp: null,

                SetupEvents_document_mouseup: null,
                SetupEvents_document_keydown: null,
                SetupEvents_document_keyup: null
            }

        }, options);

        options.dimensions = dimensions;
        options.undo = undo;

        return options;
    }

    $.fn.ip_rowObject = function (options) {

        var options = $.extend(ip_rowObject(), options);
        
        //Setup a the columns
        if (options.cols != 0) {
            for (var col = 0; col < options.cols; col++) {
                options.cells[col] = ip_cellObject(null, null, null, options.row, col); //options.row + ' x ' + col
            }
        }

        return options;
    }

    $.fn.ip_colObject = function (options) {


        var options = $.extend(ip_colObject(), options);

        return options;
    }

    $.fn.ip_cellObject = function (options) {

        var options = $.extend(ip_cellObject(), options);

        return options;
    }

    $.fn.ip_mergeObject = function (options) {

        var options = $.extend({

            mergedWithRow: null,
            mergedWithCol: null,
            rowSpan: 0,
            colSpan: 0

        }, options);

        return options;
    }


})(jQuery);


//----- OBJECTS ------------------------------------------------------------------------------------------------------------------------------------

function ip_colObject(width, style, hide, dataType, formula, validation, controlType, hashTags, border, mask, decimals) {

    if (width == null) { width = defaultColWidth; }
    if (style == null) { style = null; }
    if (formula == null) { formula = null; }
    if (hide == null) { hide = false; }    
    if (dataType == null) { dataType = ip_dataTypeObject(); }
    if (validation == null) { validation = ip_validationObject(); }
    if (controlType == null) { controlType = null; }
    if (hashTags == null) { hashTags = null; }
    if (mask == null) { mask = null; }
    if (border == null) { border = null; }
    if (decimals == null) { decimals = null; }

    var colObject = {

        width: width, // a width
        style: style,
        formula: formula,
        hide: false,
        dataType: dataType,
        validation: validation,
        controlType: controlType,
        hashTags: hashTags,
        mask: mask,
        border: border,
        decimals: decimals

    }

    return colObject;
}

function ip_rowObject(height, cells, row, cols, loading, hide, groupCount, groupColumn) {

    if (height == null) { height = defaultRowHeight; }
    if (cells == null) { cells = (cols == null || cols <= 0 ? new Array() : Array.apply(null, Array(cols)).map(function () { return new ip_cellObject(); })); }
    if (loading == null) { loading = false; }
    if (hide == null) { hide = false; }
    if (groupCount == null) { groupCount = 0; }
    //if (groupColumn == null) { groupColumn = null; }
    
    //Setup a the columns
    //if (cols != null) { for (var col = 0; col < cols; col++) { cells[col] = ip_cellObject(null, null, null, row, col); } }

    rowObject = {
        height: height,
        cells: cells, // array of ip_cellObject
        loading: loading,
        hide: false,
        groupCount: groupCount,
        groupColumn: groupColumn
    }
    return rowObject;
}

function ip_cellObject(value, style, editing, row, col, dataType, formula, validation, controlType, error, hashTags, border, mask, decimals) {

    if (dataType == null) { dataType = ip_dataTypeObject(); }
    if (validation == null) { validation = ip_validationObject(); }
    
    cellObject = {

        display: '',
        value: value,
        style: style,
        editing: editing,
        dataType: dataType,
        formula: formula,
        validation: validation,
        controlType: controlType,
        hashTags: hashTags,
        mask: mask,
        decimals: decimals,
        error: error,
        border: border
        //REM: COMPARE AND CLONE
    }
    return cellObject;
}

function ip_mergeObject(mergedWithRow, mergedWithCol, rowSpan, colSpan) {

    if (mergedWithRow == null) { mergedWithRow = null; }
    if (mergedWithCol == null) { mergedWithCol = null; }
    if (rowSpan == null) { rowSpan = 0; }
    if (colSpan == null) { colSpan = 0; }

    var mergeObject = {

        mergedWithRow: mergedWithRow,
        mergedWithCol: mergedWithCol,
        rowSpan: rowSpan,
        colSpan: colSpan
    }

    return mergeObject;
}

function ip_rangeObject(startRow, startCol, endRow, endCol, row, col, hashtags) {

    return { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol, row: row, col: col, hashtags: hashtags }

}

function ip_formatObject() {
    var formatObject = {

        fontSize: '',
        fontFamily: '',
        bold: false,
        italic: false,
        linethrough: false,
        overline: false,
        underline: false,
        color: '',
        fill: '',
        merge: false,
        hasUnmergedCells: false,
        aligntop: false,
        alignmiddle: false, //default
        alignbottom: false,
        alignleft: false,
        alignright: false,
        aligncenter: false,   //default
        validation: ip_validationObject(),
        dataType: ip_dataTypeObject(),
        controlType: null,
        hashTags: null,
        mask: null,
        decimals: null
        
    }

    return formatObject;
}

function ip_eventObject(gridID, methodName, transactionID, options) {
        
    return { gridID: gridID, publicKey: ip_GridProps[gridID].publicKey, cMethod: methodName, transactionID: transactionID, options: options }

}

function ip_dataTypeObject(dataType, dataTypeName) {

    return { dataType: dataType, dataTypeName: dataTypeName}

}

function ip_validationObject(validationCriteria, validationAction) {

    return { validationCriteria: validationCriteria, validationAction: validationAction }

}

function ip_errorObject(errorCode, errorDescription) {
    return { errorCode: errorCode, errorDescription: errorDescription }
}



//----- OBJECT CLONES ------------------------------------------------------------------------------------------------------------------------------------

function ip_CloneCell(GridID, row, col, toRow, toCol) {

    var Clone = ip_cellObject();

    if (ip_GridProps[GridID].rowData[row].cells[col].merge != null) {  Clone.merge = ip_mergeObject(ip_GridProps[GridID].rowData[row].cells[col].merge.mergedWithRow, ip_GridProps[GridID].rowData[row].cells[col].merge.mergedWithCol, ip_GridProps[GridID].rowData[row].cells[col].merge.rowSpan, ip_GridProps[GridID].rowData[row].cells[col].merge.colSpan);  }
    

    Clone.display = ip_GridProps[GridID].rowData[row].cells[col].display;
    Clone.value = ip_GridProps[GridID].rowData[row].cells[col].value;
    Clone.style = ip_GridProps[GridID].rowData[row].cells[col].style;
    Clone.row = (toRow == null ? row : toRow);
    Clone.col = (toCol == null ? col : toCol);
    Clone.dataType = ip_GridProps[GridID].rowData[row].cells[col].dataType;
    Clone.formula = ip_GridProps[GridID].rowData[row].cells[col].formula;
    Clone.fxIndex = ip_GridProps[GridID].rowData[row].cells[col].fxIndex;
    Clone.validation = ip_GridProps[GridID].rowData[row].cells[col].validation;
    Clone.controlType = ip_GridProps[GridID].rowData[row].cells[col].controlType;
    Clone.error = ip_GridProps[GridID].rowData[row].cells[col].error;
    Clone.hashTags = ip_GridProps[GridID].rowData[row].cells[col].hashTags;
    Clone.border = ip_GridProps[GridID].rowData[row].cells[col].border;
    Clone.mask = ip_GridProps[GridID].rowData[row].cells[col].mask;
    Clone.decimals = ip_GridProps[GridID].rowData[row].cells[col].decimals;

    return Clone;
}

function ip_CloneCol(GridID, col) {

    var Clone = ip_colObject();

    Clone.width = ip_GridProps[GridID].colData[col].width;
    Clone.style = ip_GridProps[GridID].colData[col].style;
    Clone.hide = ip_GridProps[GridID].colData[col].hide;
    Clone.col = col;
    Clone.dataType.dataType = ip_GridProps[GridID].colData[col].dataType.dataType;
    Clone.dataType.dataTypeName = ip_GridProps[GridID].colData[col].dataType.dataTypeName;
    Clone.validation.validationAction = ip_GridProps[GridID].colData[col].validation.validationAction;
    Clone.validation.validationCriteria = ip_GridProps[GridID].colData[col].validation.validationCriteria;
    Clone.controlType = ip_GridProps[GridID].colData[col].controlType;
    Clone.hashTags = ip_GridProps[GridID].colData[col].hashTags;
    Clone.border = ip_GridProps[GridID].colData[col].border;
    Clone.mask = ip_GridProps[GridID].colData[col].mask;

    return Clone;
}

function ip_CloneRow(GridID, row, includeCells) {

    var Clone = ip_rowObject();

    Clone.height = ip_GridProps[GridID].rowData[row].height;
    Clone.hide = ip_GridProps[GridID].rowData[row].hide;
    Clone.groupCount = ip_GridProps[GridID].rowData[row].groupCount;
    Clone.groupColumn = ip_GridProps[GridID].rowData[row].groupColumn;
    Clone.cells = [];
    Clone.row = row;
    
    

    //For now cloning row cells is not needed anywhere so we will add this feature when it is required
    //Create cell undo stack
    if (includeCells) { for (var c = 0; c < ip_GridProps[GridID].rowData[row].cells.length; c++) { Clone.cells[Clone.cells.length] = ip_CloneCell(GridID, row, c); } }


    return Clone;
}


//----- OBJECT COMPARES ------------------------------------------------------------------------------------------------------------------------------------

function ip_CompareCell(GridID, cell1, cell2) {

    if (ip_parseAny(GridID, cell1.value) != ip_parseAny(GridID, cell2.value)) { return false; }
    if (cell1.formula != cell2.formula) { return false; }
    if (cell1.border != cell2.border) { return false; }
    if (cell1.hashtags != cell2.hashtags) { return false; }
    if (cell1.style != cell2.style) { return false; }
    if (cell1.controlType != cell2.controlType) { return false; }
    if (cell1.error != null || cell2.error != null)
    {
        var errorCode1 = (cell1.error ? cell1.error.errorCode : '');
        var errorDescription1 = (cell1.error ? cell1.error.errorDescription : '');
        var errorCode2 = (cell2.error ? cell2.error.errorCode : '');
        var errorDescription2 = (cell2.error ? cell2.error.errorDescription : '');

        if (errorCode1 == null) { errorCode1 = ''; }
        if (errorDescription1 == null) { errorDescription1 = ''; }
        if (errorCode2 == null) { errorCode2 = ''; }
        if (errorDescription2 == null) { errorDescription2 = ''; }

        if (errorCode1 != errorCode2) { return false; }
        if (errorDescription1 != errorDescription2) { return false; }

    }
    if (cell1.validation != null || cell2.validation != null) {
        if (cell1.validation == null) { return false; }
        if (cell2.validation == null) { return false; }
        if (cell1.validation.validationCriteria != cell2.validation.validationCriteria) { return false; }
        if (cell1.validation.validationAction != cell2.validation.validationAction) { return false; }
    }
    if (cell1.dataType != null || cell2.dataType != null) {
        if (cell1.dataType == null) { return false; }
        if (cell2.dataType == null) { return false; }
        if (cell1.dataType.dataType != cell2.dataType.dataType) { return false; }
        if (cell1.dataType.dataTypeName != cell2.dataType.dataTypeName) { return false; }
    }

    return true;
}

function ip_CompareError(GridID, error1, error2) {

    if (error1 == null && error2 == null) { return true; }
    if (error1 == null) { error1 = ip_errorObject("",""); }
    if (error2 == null) { error2 = ip_errorObject("", ""); }

    if (error1.errorCode != error2.errorCode) { return false; }
    if (error1.errorDescription != error2.errorDescription) { return false; }

    return true;
}


//----- GRID ARCHITECHTURE --------------------------------------------------------------------------------------------------------------------------------------------------

function ip_DisposeGrid(GridID) {

    if (ip_GridProps[GridID] != null) {
        
        ip_RaiseEvent(GridID, 'ip_DisposeGrid', ip_GenerateTransactionID(), { ResizeGrid: { Inputs: null, Effected: null } });

        ip_UnbindAllEvents(GridID);
        
        $('#' + GridID).html('');
        $('#' + GridID + '_gridResizer').remove();

        delete ip_GridProps[GridID];
    }

}

function ip_CreateGrid(options) {
    
    
    ip_GridProps[options.id].dimensions.gridHeight = $('#' + options.id).height() - ip_GridProps[options.id].dimensions.fBarHeight;
    ip_GridProps[options.id].dimensions.gridWidth = $('#' + options.id).width();
    ip_GridProps[options.id].scrollY = (ip_GridProps[options.id].scrollY == null ? ip_GridProps[options.id].frozenRows : ip_GridProps[options.id].scrollY);
    ip_GridProps[options.id].scrollX = (ip_GridProps[options.id].scrollX == null ? ip_GridProps[options.id].frozenCols : ip_GridProps[options.id].scrollX);
    

    ip_RecalculateLoadedRowsCols(options.id, false, true, true);
    ip_SetupMerges(options.id, options.mergeData);


    var optionsQ1 = jQuery.extend({ startCol: 0, startRow: 0 }, options);
    optionsQ1.showColSelector = (options.showColSelector == false ? false : true); //default: true
    optionsQ1.showRowSelector = (options.showRowSelector == false ? false : true); //default: true
    optionsQ1.id = options.id + '_q1';
    optionsQ1.rows = options.frozenRows;
    optionsQ1.cols = options.frozenCols;
    optionsQ1.Quad = 1;
    optionsQ1.GridID = options.id;



    var optionsQ2 = jQuery.extend({ startCol: ip_GridProps[options.id].scrollX, startRow: 0 }, options);
    optionsQ2.showColSelector = (options.showColSelector == false ? false : true); //default: true
    optionsQ2.showRowSelector = false;
    optionsQ2.id = options.id + '_q2';
    optionsQ2.rows = options.frozenRows;
    optionsQ2.cols = options.cols - options.frozenCols;
    optionsQ2.Quad = 2;
    optionsQ2.GridID = options.id;

 

    var optionsQ3 = jQuery.extend({ startCol: 0, startRow: ip_GridProps[options.id].scrollY }, options);
    optionsQ3.showColSelector = false;
    optionsQ3.showRowSelector = (options.showRowSelector == false ? false : true) //default:true
    optionsQ3.id = options.id + '_q3';
    optionsQ3.rows = options.rows - options.frozenRows;
    optionsQ3.cols = options.frozenCols;
    optionsQ3.Quad = 3;
    optionsQ3.GridID = options.id;

    var optionsQ4 = jQuery.extend({ startCol: ip_GridProps[options.id].scrollX, startRow: ip_GridProps[options.id].scrollY }, options);
    optionsQ4.showColSelector = false;
    optionsQ4.showRowSelector = false;
    optionsQ4.id = options.id + '_q4';
    optionsQ4.rows = options.rows - options.frozenRows;
    optionsQ4.cols = options.cols - options.frozenCols;
    optionsQ4.Quad = 4;
    optionsQ4.GridID = options.id;

    var Quads = '';

    if (!$('#' + options.id).hasClass('ip_grid_sheet')) { $('#' + options.id).addClass('ip_grid_sheet'); };
    
    options.toolType = 'gridControls';

    Quads += ip_CreateGridTools(options);
    Quads += '<table id="' + options.id + '_table" border="0" cellpadding="0" cellspacing="0" >';
    Quads += '<tr>';
    Quads += '<td id="' + optionsQ1.id + '" class="ip_grid_quadrant" valign="top" align="left"><div id="' + optionsQ1.id + '_div_container" style="position:relative;">' + ip_CreateGridQuadrant(optionsQ1) + '<div></td>';
    Quads += '<td id="' + optionsQ2.id + '" class="ip_grid_quadrant" valign="top" align="left"><div id="' + optionsQ2.id + '_div_container" style="position:relative;">' + ip_CreateGridQuadrant(optionsQ2) + '<div></td>';
    Quads += '</tr>';
    Quads += '<tr>';
    Quads += '<td id="' + optionsQ3.id + '" class="ip_grid_quadrant" valign="top" align="left"><div id="' + optionsQ3.id + '_div_container" style="position:relative;">' + ip_CreateGridQuadrant(optionsQ3) + '<div></td>';
    Quads += '<td id="' + optionsQ4.id + '" class="ip_grid_quadrant" valign="top" align="left"><div id="' + optionsQ4.id + '_div_container" style="position:relative;">' + ip_CreateGridQuadrant(optionsQ4) + ip_CreateScrollBar('x', optionsQ4) + ip_CreateScrollBar('y', optionsQ4) + '<div></td>';
    Quads += '</tr>';
    Quads += '</table>';

    var $Quads = $(Quads);    
    $('#' + options.id).html(Quads);

    //Setup fxBar
    if (ip_GridProps[options.id].fxBar == null) { ip_GridProps[options.id].fxBar = $('#' + options.id + '_fBar')[0]; }
    
    //setup container
    $('#' + options.id).attr('tabindex', ip_GridProps[options.id].index); //needed for keypress events on grid

    //Setup grid objects and events    
    if (!options.refresh) { ip_SetupEvents(options.id);  }// ip_InitFxBar(options.id); }

    if (!$('#' + options.id).ip_Scrollable()) { return false; }

    ip_ShowColumnFrozenHandle(options.id);
    ip_ShowRowFrozenHandle(options.id);
    ip_ShowGridResizerHandle(options.id);

    //remove text selection    
    ip_DisableSelection(options.id, true);

    ip_SetupFx(options.id);
    ip_SetupMask(options.id);

    return $Quads;
}

function ip_CreateGridTools(options) {

    var GridTools = '';

    if (options.toolType == "gridControls") {

        //Ranger selector
        GridTools += '<div id="' + options.id + '_rangeselector" class="ip_grid_cell_rangeselector"  style="pointer-events:none;" >' +

                        '<svg class="ip_grid_cell_rangeselector_innerContent" xmlns="http://www.w3.org/2000/svg"  style="pointer-events:none;"  pointer-events="none">  ' +
                            '<rect x="0" y="0" width="100%" height="100%" pointer-events="none" style="pointer-events:none;" />' +
                        '</svg>' +

                        '<div class="ip_grid_cell_rangeselector_border" borderPosition="top" style="width:100%;height:3px;top:-2px;"></div>' +
                        '<div class="ip_grid_cell_rangeselector_border" borderPosition="left" style="width:3px;height:100%;left:-2px;"></div>' +
                        '<div class="ip_grid_cell_rangeselector_border" borderPosition="right" style="width:3px;height:100%;right:-2px;"></div>' +
                        '<div class="ip_grid_cell_rangeselector_border" borderPosition="bottom" style="width:100%;height:3px;bottom:-2px;"></div>' +

                        '<div class="ip_grid_cell_rangeselector_key"></div>' +

                     '</div>';

        GridTools += '<div id="' + options.id + '_editTool" class="ip_grid_EditTool"><div contenteditable="true" tabindex="0" class="ip_grid_EditTool_Input" ></div><div tabindex="0"  class="ip_grid_EditTool_DropDown"></div></div>';

        GridTools += '<div id="' + options.id + '_rangeHighlight" class="ip_grid_cell_rangeHighlight"></div>';

        GridTools += '<div id="' + options.id + '_columnFrozenHandle"  class="ip_grid_columnFrozenHandle"><div id="' + options.id + '_columnFrozenHandleLine" class="ip_grid_columnFrozenHandleLine"></div></div>';
        GridTools += '<div id="' + options.id + '_rowFrozenHandle"  class="ip_grid_rowFrozenHandle"><div id="' + options.id + '_rowFrozenHandleLine" class="ip_grid_rowFrozenHandleLine"></div></div>';

        GridTools += '<div id="' + options.id + '_columnResizer"  class="ip_grid_columnSelectorResizeTool"><div id="' + options.id + '_columnLine" class="ip_grid_columnSelectorResizeLine"></div></div>';
        GridTools += '<div id="' + options.id + '_rowResizer"  class="ip_grid_rowSelectorResizeTool"><div id="' + options.id + '_rowLine" class="ip_grid_rowSelectorResizeLine"></div></div>';

        GridTools += '<span id="' + options.id + '_cellContentWidthTool"  class="ip_grid_cell" style="display:none;position:absolute;"></span>';

        GridTools += '<textarea id="' + options.id + '_selectTool" cols="50" rows="10" style="position:absolute;z-index:0;"></textarea>';

        GridTools += '<div id="' + options.id + '_fBar" class="ip_grid_fbar" style="line-height:' + ip_GridProps[options.id].dimensions.fBarHeight + 'px;height:' + ip_GridProps[options.id].dimensions.fBarHeight + 'px;"><div class="ip_grid_fbar_title"><span class="ip_grid_fbar_f">f</span><span class="ip_grid_fbar_x">x</span></div><div class="ip_grid_fbar_text"></div></div>';

        //This tools must lay outside of the grid objects containment realm    
        $('#' + options.id + '_gridResizer').remove();
        $('#' + options.id).parent().append('<div id="' + options.id + '_gridResizer" title="Resize Grid (double click to split or maximize)"  class="ip_grid_gridResizeTool"><div id="' + options.id + '_gridResizerLine" class="ip_grid_gridResizeToolLine"></div></div>');

    }
    else if (options.toolType == 'cellFormatting') {

        var enabledFormats = options.enabledFormats;
        
        GridTools += '<div  style="min-width:670px">';
        GridTools += '<table border="0" cellpadding="0" cellspacing="10">';''
        GridTools += '<tr><td>Range:</td><td><input id="' + options.id + '_Range_F" placeholder="Effected range, example: A1:A10" class="ip_ControlBase ip_TextBox no-drag formatting-tool"  style="width:164px;" type="text" /></td></tr>';
        GridTools += '<tr><td>Cell Type:</td><td><div id="' + options.id + '_ControlType" placeholder="Please choose .." style="width:170px;" class="ip_ControlBase ip_DropDown no-drag formatting-tool"  ></div><input id="' + options.id + '_ValidationValue" class="ip_ControlBase ip_TextBox tooltip no-drag formatting-tool" style="width:280px;margin-left:10px;' + (enabledFormats.controlType == '' ? 'display:none;' : '') + '" type="text" /></td></tr>'; 
        GridTools += '<tr><td>Incorrect Input Action:</td><td><div id="' + options.id + '_Action" style="width:170px;" placeholder="Please choose .." class="ip_ControlBase ip_DropDown no-drag formatting-tool"  ></div></td></tr>';
        GridTools += '<tr><td>Data Type:</td><td><div id="' + options.id + '_DataType" style="width:170px;" placeholder="Please choose .." class="ip_ControlBase ip_DropDown no-drag formatting-tool"  ></div><div id="' + options.id + '_Mask" style="width:170px;margin-left:10px;display:none;" placeholder="Formatting .." class="ip_ControlBase ip_DropDown no-drag formatting-tool"  ></div></td></tr>';
        GridTools += '<tr><td>#Tags:</td><td colspan="2"><div id="' + options.id + '_HashTags" placeholder="Example: #Red, #Green, #Blue" class="ip_ControlBase ip_TextBox no-drag formatting-tool"  style="width:470px;" /></td></tr>';
        GridTools += '</table>';
        GridTools += '</div>';

    }

    return GridTools;
}

function ip_CreateGridQuadrant(options)
{
    // create table
    var TableContent = new Array();
 
    TableContent += '<div id="' + options.id + '_container" class="ip_grid_container">'
    TableContent += '<table id="' + options.id + '_table" Quad="' + options.Quad + '" class="ip_grid" border="0" cellpadding="0" cellspacing="0" >';
    TableContent += ip_CreateGridQuadColumnHeaders(options);
    TableContent += ip_CreateGridQuadRows(options);
    TableContent += '</table>';
    TableContent += '</div>';

    return TableContent;
}

function ip_CreateGridQuadColumnHeaders(options) {
    
    //var VisibleCols = ip_GridProps[options.GridID].loadedCols > options.cols ? options.cols : options.startCol + ip_GridProps[options.GridID].loadedCols - ip_GridProps[options.GridID].frozenCols;
    var VisibleCols = options.startCol + (ip_GridProps[options.GridID].loadedCols > options.cols ? options.cols : ip_GridProps[options.GridID].loadedCols - ip_GridProps[options.GridID].frozenCols);
    var TableHeader = ['<thead style="" >',
                       '<tr  class="ip_grid_columnSelectorRow"  id="' + options.id + '_gridRow_-1" row="-1" col="-1" style="' + (!options.showColSelector ? '' : 'height:' + ip_GridProps[options.GridID].dimensions.columnSelectorHeight + 'px') + '">',
                       '<th class="ip_grid_columnSelectorCellCorner q' + options.Quad + '" id="' + options.id + '_columnSelectorCell_-1" quadrant="' + options.Quad + '" row="-1" col="-1"  style="' + (!options.showRowSelector ? 'width:0px;' : 'width:' + ip_GridProps[options.GridID].dimensions.rowSelectorWidth + 'px;') + '" >',
                       '<div class="ip_grid_cell_outerContent" style="' + (!options.showColSelector ? 'display:none;' : '') + (!options.showRowSelector ? 'display:none;' : '') + '" >',
                       '</div>',
                       '</th>'];
    
    if (VisibleCols < 0) { VisibleCols = ip_GridProps[options.GridID].loadedCols; }
    if (options.Quad == 2) { ip_GridProps[options.GridID].dimensions.accumulativeScrollWidth = 0; }

    for (var col = options.startCol; col < VisibleCols; col++) {

        options.cellType = 'ColSelector';
        options.col = col;
        options.row = -1;
        TableHeader.push(ip_CreateGridQuadCell(options));
        if (options.Quad == 2) { ip_GridProps[options.GridID].dimensions.accumulativeScrollWidth += ip_ColWidth(options.GridID, col, true); }

    }

    TableHeader.push('</tr>');
    TableHeader.push('</thead>');


    return TableHeader.join('');
}

function ip_CreateGridQuadRows(options) {

    var options = jQuery.extend({ row:0 }, options);
    var TableColumns = ['<tbody id="' + options.id + '_table_tbody">']
    var VisibleRows = options.startRow + (ip_GridProps[options.GridID].loadedRows > options.rows ? options.rows : ip_GridProps[options.GridID].loadedRows - ip_GridProps[options.GridID].frozenRows);
        
    if (VisibleRows < 0) { VisibleRows = ip_GridProps[options.GridID].loadedRows; }
    if (options.Quad == 4) { ip_GridProps[options.GridID].dimensions.accumulativeScrollHeight = 0;  }

    for (var row = options.startRow; row < VisibleRows; row++) {

        options.row = row;
        TableColumns.push(ip_CreateGridQuadRow(options, TableColumns));

        //Record the current scroll height
        if (options.Quad == 4) { ip_GridProps[options.GridID].dimensions.accumulativeScrollHeight += ip_RowHeight(options.GridID, row, true); }
    }

    TableColumns.push('</tbody>');


    return TableColumns.join('');
}

function ip_CreateGridQuadRow(options, tableColumnsArray) {

    var loadedScrollable = (options.loadedScrollable == null ? ip_LoadedRowsCols(options.GridID) : options.loadedScrollable);    
    var VisibleCols = (options.Quad == 1 || options.Quad == 3 ? loadedScrollable.colCount_frozen : loadedScrollable.colTo_scroll + 1);
    var options = jQuery.extend({ col: 0, cellType: '' }, options);
    var row = parseInt(options.row);
    var rowData = ip_GridProps[options.GridID].rowData[row];
    var height = ip_RowHeight(options.GridID, row, false);
    
    if (rowData.hide) { return ''; }


    tableColumnsArray = (tableColumnsArray != null ? tableColumnsArray : new Array());        
    tableColumnsArray.push('<tr class="ip_grid_row" id="' + options.id + '_gridRow_' + row + '" row="' + row + '" style="height:' + height + 'px;max-height:' + height + 'px;">');

    options.cellType = 'RowSelector';
    options.col = -1;
    ip_CreateGridQuadCell(options, tableColumnsArray);
    
    for (var col = options.startCol; col < VisibleCols; col++) {  //for (var col = options.startCol; col < options.cols + options.startCol; col++) {

        options.cellType = '';
        options.col = col;
        ip_CreateGridQuadCell(options, tableColumnsArray);

    }

    tableColumnsArray.push('</tr>');
}

function ip_CreateGridQuadCell(options, tableColumnsArray) {

    var col = options.col;
    var row = options.row;
    var TableColumns = '';
    var Quad = options.Quad;

    if (row >= 0 && ip_GridProps[options.GridID].rowData[row].hide) { return ''; }
    if (col >= 0 && ip_GridProps[options.GridID].colData[col].hide) { return ''; }

    if (options.cellType == 'RowSelector') {

        //Row Selector Column
        var selected = (ip_GridProps[options.GridID].rowData[row].selected ? " selected" : "");
        var hide = (ip_GridProps[options.GridID].rowData[row].hide ? " ip_grid_cell_hideRow"  : "");
        var groupCount = (ip_GridProps[options.GridID].rowData[row].groupCount > 0 ? ' ip_grid_rowGroup' : '')
        var hideIcon = '';
        
        //if (row < ip_GridProps[options.GridID].rows - 1 && !ip_GridProps[options.GridID].rowData[row].hide && ip_GridProps[options.GridID].rowData[row + 1].hide) { hideIcon = '<div class="ip_grid_cell_hideIcon bottom"></div>'; }
        if (row > 0 && !ip_GridProps[options.GridID].rowData[row].hide && ip_GridProps[options.GridID].rowData[row - 1].hide) { hideIcon += '<div class="ip_grid_cell_hideIcon top"></div>'; }

        TableColumns = '<th cellType="RowSelector" class="ip_grid_rowSelecterCell q' + Quad + selected + " " + hide + groupCount + '" id="' + options.id + '_rowSelecterCell_' + row + '"  quadrant="' + options.Quad + '" col="-1" row="' + row + '" style="' +
                                                                                                                                                        (!options.showRowSelector ? 'width:0px;' : '') +
                                                                                                                                                        '">' +
                            hideIcon +
                            '<div class="ip_grid_cell_outerContent">' +
                                (Quad == 1 || Quad == 3 ? row : '') + 
                            '</div>' +
                        '</th>';
    }
    else if (options.cellType == 'ColSelector') {

        var colData = ip_GridProps[options.GridID].colData[col];
        var colWidth = ip_ColWidth(options.GridID, col, false);
        var selected = (colData.selected ? " selected" : "");
        var hide = (colData.hide ? " ip_grid_cell_hideCol" : "");
        var hideIcon = '';
        var dataTypeName = colData.dataType.dataTypeName != null ? colData.dataType.dataTypeName : ''

        if (colData.hide && colWidth == 0) { colWidth = 0.01; } //this is a chrome fix, without this the border wont show

        //if (col < ip_GridProps[options.GridID].cols - 1 && !hide && ip_GridProps[options.GridID].colData[col + 1].hide) { hideIcon = '<div class="ip_grid_cell_hideIcon right"></div>'; }
        if (options.showColSelector && col > 0 && !hide && ip_GridProps[options.GridID].colData[col - 1].hide) { hideIcon += '<div class="ip_grid_cell_hideIcon left"></div>'; }



        //Col Selector Column
        TableColumns = '<th cellType="ColSelector" class="ip_grid_columnSelectorCell q' + Quad + selected + hide + '" id="' + options.id + '_columnSelectorCell_' + col + '"  quadrant="' + options.Quad + '" row="-1" col="' + col + '"  style="' +
                                                                                                                                                           (!options.showColSelector ? 'border-top:none;' : '') +
                                                                                                                                                           'width:' + colWidth + 'px;' +
                                                                                                                                                           '">' +
                               hideIcon +
                               '<div class="ip_grid_cell_outerContent" style="position:relative;' +
                                                                                (!options.showColSelector ? 'display:none;' : '') +
                                                                                '" >' +

                                    '(' + ip_GridProps[options.GridID].indexedData.colSymbols.colSymbols[col] + ') ' + //ip_GridProps[options.GridID].colSymbols[col]
                                    '<div  class="ip_grid_cell_innerContent">' +
                                    dataTypeName +
                                    '</div>' +
                               '</div>' +
                       '</th>';
    }
    else if(row >= 0 && col >= 0) {

        //Standard Cell         
        var CellData = ip_CellData(options.GridID, row, col, true);      
        TableColumns = '<td cellType="Cell" class="ip_grid_cell' + CellData.css + '" id="' + options.GridID + '_cell_' + row + '_' + col + '"  quadrant="' + options.Quad + '" col="' + col + '" row="' + row + '" style="' + CellData.style + '" ' + CellData.rowSpan + ' ' + CellData.colSpan + ' >' +
                            ip_CellBorder(options.GridID, row, col, CellData) +
                            ip_CellOuter(options.GridID, row, col, CellData) +

                        '</td>';
    }

    if (tableColumnsArray != null) { tableColumnsArray.push(TableColumns); }

    return TableColumns;
}

function ip_CellBorder(GridID, row, col, CellData) {

    if (!CellData.border) { return ''; }

    var border = '';

    border = '<div class="ip_grid_cell_border" style="' + CellData.border + '"></div>';

    return border;

}

function ip_CreateScrollBar(orientation, options) {

    var ScrollBar = '<div id="' + options.id + '_scrollbar_container_' + orientation + '" class="ip_grid_scrollBarContainer ' + (orientation == 'x' ? 'horizontal' : 'vertical' ) +'" >' +
                    '<div class="ip_grid_scrollBar_innerContent">' +
                        '<div id="' + options.id + '_scrollbar_gadget_' + orientation + '" class="ip_grid_scrollBar_gadget ' + (orientation == 'x' ? 'horizontal' : 'vertical') + '" >' +
                        '<div id="' + options.id + '_scrollbar_gadget_scrollIndecator_' + orientation + '" class="ip_grid_scrollBar_indecator ' + (orientation == 'x' ? 'horizontal' : 'vertical') + '">0</div>' +
                        '</div>' +
                    '</div>' +
                    '</div>';

    return ScrollBar;

}

function ip_CellHtml(GridID, row, col, returnLastIfNull) {
        
    
    if (row >= 0 && col >= 0 && row < ip_GridProps[GridID].rows && col < ip_GridProps[GridID].cols) {

        var merge = ip_GridProps[GridID].rowData[row].cells[col].merge;



        if (merge != null) {

            //row = merge.mergedWithRow;
            //col = merge.mergedWithCol;

            row = ip_NextNonHiddenRow(GridID, merge.mergedWithRow, merge.mergedWithRow + merge.rowSpan - 1, row, 'down');
            //--this
            col = ip_NextNonHiddenCol(GridID, merge.mergedWithCol, merge.mergedWithCol + merge.colSpan - 1, col, 'right');

        }
    }

    //Choose the last row if requested row exceeds grid rows
    var visibleScrollRows = ip_GridProps[GridID].scrollY + ip_GridProps[GridID].loadedRows - ip_GridProps[GridID].frozenRows - 1; // 
    if (visibleScrollRows >= ip_GridProps[GridID].rows) { visibleScrollRows = ip_GridProps[GridID].rows - 1; }

    if (row >= visibleScrollRows && returnLastIfNull) {

        row = visibleScrollRows;
    }
    else if (row < ip_GridProps[GridID].scrollY && row >= ip_GridProps[GridID].frozenRows && returnLastIfNull) {

        row = ip_GridProps[GridID].scrollY;

    }

    //Choose the last col if requested col exceeds grid cols
    var visibleScrollCols = ip_GridProps[GridID].scrollX + ip_GridProps[GridID].loadedCols - ip_GridProps[GridID].frozenCols - 1;
    if (visibleScrollCols >= ip_GridProps[GridID].cols) { visibleScrollCols = ip_GridProps[GridID].cols - 1; }


    if (col >= visibleScrollCols && returnLastIfNull) {

        col = visibleScrollCols;
    }
    else if (col < ip_GridProps[GridID].scrollX && col >= ip_GridProps[GridID].frozenCols && returnLastIfNull) {

        col = ip_GridProps[GridID].scrollX;

    }


    //Retreive cell object
    var CellID = ''
    if (row == -1) {
        var Quad = ip_GetQuad(GridID, row, col);
        CellID = GridID + '_q' + Quad + '_columnSelectorCell_' + col;
    }
    else if (col == -1) {
        var Quad = ip_GetQuad(GridID, row, col);
        CellID = GridID + '_q' + Quad + '_rowSelecterCell_' + row;
    }
    else {

        CellID = GridID + '_cell_' + row + '_' + col;
    }
    
    var Cell = $('#' + CellID);

    if (Cell.length > 0) { return Cell[0]; }

    return null;

}

function ip_CellInner(GridID, row, col) {

    //This should be the inner VISUAL markup of the cell
    var cellInner = '';
    var cell = ip_GridProps[GridID].rowData[row].cells[col];

    if (cell.controlType == 'gantt') {
        cellInner = (ip_parseBool(cell.value) ? '<div class="ip_grid_cell_gantt"></div>' : '');
    }
    else {
        cellInner = (cell.value == null ? '' : cell.display);
    }

    return cellInner;

}

function ip_CellOuter(GridID, row, col, cellData) {

    if (cellData == null) { cellData = ip_CellData(GridID, row, col, true); }

    var border = '';
   
    var cellOuter = '' +
    
    cellData.iconGP +
    '<div  class="ip_grid_cell_outerContent">' +
        
        '<div  class="ip_grid_cell_innerContent">' +
        ip_CellInner(GridID, row, col) +
        '</div>' +

    '</div>' +
    
    cellData.iconDD;

    return cellOuter;
}

function ip_IsCellEmpty(GridID, row, col) {

    var merge = ip_GridProps[GridID].rowData[row].cells[col].merge;

    if (merge != null) {

        row = merge.mergedWithRow;
        col = merge.mergedWithCol;

    }

    if (ip_GridProps[GridID].rowData[row].cells[col].value != null && ip_GridProps[GridID].rowData[row].cells[col].value != '') { return false }
    if (ip_GridProps[GridID].rowData[row].cells[col].formula != null && ip_GridProps[GridID].rowData[row].cells[col].formula != '') { return false }
    if (ip_GridProps[GridID].rowData[row].cells[col].style != null && ip_GridProps[GridID].rowData[row].cells[col].style != '') { return false }
    if (ip_GridProps[GridID].rowData[row].cells[col].validation != null && ip_GridProps[GridID].rowData[row].cells[col].validation.validationCriteria != null && ip_GridProps[GridID].rowData[row].cells[col].validation.validationCriteria != '') { return false }
    if (ip_GridProps[GridID].rowData[row].cells[col].validation != null && ip_GridProps[GridID].rowData[row].cells[col].validation.validationAction != null && ip_GridProps[GridID].rowData[row].cells[col].validation.validationAction != '') { return false }
    if (ip_GridProps[GridID].rowData[row].cells[col].dataType != null && ip_GridProps[GridID].rowData[row].cells[col].dataType.dataType != null && ip_GridProps[GridID].rowData[row].cells[col].dataType.dataType != '') { return false }
    if (ip_GridProps[GridID].rowData[row].cells[col].dataType != null && ip_GridProps[GridID].rowData[row].cells[col].dataType.dataTypeName != null && ip_GridProps[GridID].rowData[row].cells[col].dataType.dataTypeName != '') { return false }
    if (ip_GridProps[GridID].rowData[row].cells[col].controlType != null && ip_GridProps[GridID].rowData[row].cells[col].controlType != '') { return false }

    return true;

}

function ip_ColSymbol(GridID, col) {

    
    if (col == 0) { return 'A'; }
    
    return col;

}

function ip_ColumnSymbols(columns) {

    var keyCounter = 0,
		intSymbols = [],
		chrSymbols = [],
		loops = columns / 26,
		maxCols = 1024; // (AMJ)
    
    for (var x = -1; x < loops - 1; x++) {
        for (var i = 0; i < maxCols; i++) {
            var code = ip_ColumnSymboldCharCode(i);
            intSymbols[keyCounter] = code
            chrSymbols[code] = keyCounter;
            keyCounter++;
        }
    }

    return { colSymbols: intSymbols, symbolCols: chrSymbols };
}

function ip_ColumnSymboldCharCode(number) {
    // function to generate unlimited predefined cols
    var l = '';
    if(number > 701) {
        l += String.fromCharCode(64 + parseInt(number / 676));
        l += String.fromCharCode(64 + parseInt((number % 676) / 26));
    } else if(number > 25) {
        l += String.fromCharCode(64 + parseInt(number / 26));
    }
    l += String.fromCharCode(65 + (number % 26));

    return l;
}

//----- INITIALIZATION ------------------------------------------------------------------------------------------------------------------------------------

function ip_SetupEvents(GridID) {
      
    //Deals managing which grid is focused
    $('#' + GridID).on('focus', function () {
        ip_FocusGrid(GridID);
    });

    //Deals with setting the mouse button currently pressed
    $('#' + GridID).on('mousedown', function (e) {
        ip_GridProps[GridID].mouseButton = e.which;
        ip_GridProps[GridID].shiftKey = e.shiftKey;
        ip_GridProps[GridID].ctrlKey = e.ctrlKey;
    });

    //Deals with resetting the mouse button currently pressed
    $(document).on('mouseup', ip_GridProps[GridID].events.SetupEvents_document_mouseup = function (e) {

        ip_GridProps[GridID].mouseButton = 0;
        ip_GridProps[GridID].shiftKey = e.shiftKey;
        ip_GridProps[GridID].ctrlKey = e.ctrlKey;

    });

    //Deals with selecing a cell & or changing a range ordinates
    $('#' + GridID ).on('mousedown','.ip_grid_cell', function (e) {
        
        var shitClick = e.shiftKey;
        var ctrlClick = e.ctrlKey;
        if (!ip_GridProps[GridID].resizing && !ip_GridProps[GridID].editing.editing) {
            
            //Check if we have clicked on a range, an make it the last selected range
            if (e.which != 1) {

                var row = parseInt($(this).attr('row'));
                var col = parseInt($(this).attr('col'));

                for (var r = 0; r < ip_GridProps[GridID].selectedRange.length; r++) {

                    var range1 = {
                        startRow: ip_GridProps[GridID].selectedRange[r][0][0,0],
                        startCol: ip_GridProps[GridID].selectedRange[r][0][0,1],
                        endRow: ip_GridProps[GridID].selectedRange[r][1][1,0],
                        endCol: ip_GridProps[GridID].selectedRange[r][1][1,1]
                    }

                    var range2 = { startRow: row, startCol: col, endRow: row, endCol: col }

                    if (ip_IsRangeOverLap(GridID, range1, range2)) {

                        //Range has been clicked, fire event
                        var rangeID = GridID + '_range_' + range1.startRow + '_' + range1.startCol + '_' + range1.endRow + '_' + range1.endCol;
                        ip_GridProps[GridID].selectedRangeIndex = r;                        
                        Effected = { range: range1, rangeEl: $('#' + rangeID)[0] }
                        ip_RaiseEvent(GridID, 'ip_RangeMouseDown', ip_GenerateTransactionID(), { RangeMouseDown: { Inputs: e, Effected: Effected } });
                        return false;

                    }

                }


            }

            if (shitClick) {

                $('#' + GridID).ip_UnselectColumn();
                $('#' + GridID).ip_UnselectRow();
                ip_ChangeRange(GridID, ip_GridProps[GridID].selectedRange[ip_GridProps[GridID].selectedRangeIndex], this);
                
            }
            else
            {      
                $('#' + GridID).ip_SelectCell({ cell: this, multiselect: ctrlClick });    
            }
      

        }
        else if (ip_GridProps[GridID].editing.editing) {

            //If editing allow user to select a range 
            if (ip_GridProps[GridID].editing.editTool.isFX()) {
                
                var row = parseInt($(this).attr('row'));
                var col = parseInt($(this).attr('col'));
                
                if (row != ip_GridProps[GridID].editing.row || col != ip_GridProps[GridID].editing.col) {
                    e.preventDefault();//NB: used by text editor to prevent BLUR from triggering when selecting cell AND prevents chrome from selecting the range
                    $('#' + GridID).ip_SelectRange({ startCell: this, multiselect: false, rangeType: 'ip_grid_cell_rangeselector_transparent', allowColumn: false, allowRow: false, allowMove: false, showKey: false });
                }
            }
            else
            {                
                $('#' + GridID).ip_SelectCell({ cell: this });
            }
        }
    });
    
    //Deals with: setting the current hovered over cell & adjusting the range
    $('#' + GridID).on('mouseenter', '.ip_grid_cell', function (e) {

        if (ip_GridProps[GridID].hoverCell != this) { ip_GridProps[GridID].hoverCell = this; }

        //This is outside the  above IF statement to offer improved support for IE
        if (ip_GridProps[GridID].mouseButton == 1 && !ip_GridProps[GridID].resizing) {

            ip_EnableScrollAnimate(GridID);
            ip_ChangeRange(GridID, ip_GridProps[GridID].selectedRange[ip_GridProps[GridID].selectedRangeIndex], this);
            


        }
        

        

    });
       
    //Deals with clearing copy
    $('#' + GridID).on('dblclick', '.ip_grid_cell', function (e) {

        if (!ip_GridProps[GridID].resizing) {

            $('#' + GridID).ip_ClearCopy();
            $('#' + GridID).ip_TextEditable({ contentType: 'ip_grid_CellEdit', element: this });
        }

    });
    
    //This is an IE fix because pointer-events does not work in IE
    if (ip_GridProps['index'].browser.name == 'ie') {

        $('#' + GridID).on('mousedown', '.ip_grid_cell_rangeselector', function (e) {

            ip_GridProps[GridID].shiftKey = e.shiftKey;
            ip_GridProps[GridID].ctrlKey = e.ctrlKey;

            var hoverCell = ip_SetHoverCell(GridID, this, e, false, null);
            $(hoverCell).mousedown();

            return false;

        });

    }    
    
    //Deals with: setting the current hovered over column
    $('#' + GridID).on('mouseenter', '.ip_grid_columnSelectorCell', function (e) {

        if (!ip_GridProps[GridID].resizing) { ip_HideFloatingHandles(GridID); }
        if (ip_GridProps[GridID].hoverCell != this) {            
            ip_GridProps[GridID].hoverCell = this;
        }
    });

    //Deals with: setting the current hovered over row
    $('#' + GridID).on('mouseenter', '.ip_grid_rowSelecterCell', function (e) {

        if (!ip_GridProps[GridID].resizing) { ip_HideFloatingHandles(GridID); }
        if (ip_GridProps[GridID].hoverCell != this) {
            ip_GridProps[GridID].hoverCell = this;
            //ip_GridProps[GridID].hoverRowIndex = parseInt($(this).attr('row'));
        }
    });
    
    //Deals with: moving selected columns
    $('#' + GridID).on('mousedown', '.ip_grid_columnSelectorCell', function (e, e2) {

        if (ip_GridProps[GridID].editing.editing && ip_GridProps[GridID].editing.editTool.isFX()) { e.preventDefault(); return; }

        var shitClick = e.shiftKey;
        var ctrlClick = e.ctrlKey;
        var col = parseInt($(this).attr('col'));

        if (ip_GridProps['index'].browser.name == 'ie' && shitClick == null) {

            shitClick = ip_GridProps[GridID].shiftKey;
            ctrlClick = ip_GridProps[GridID].ctrlKey;
        }

        if (!ip_GridProps[GridID].resizing) {

            var allowColSelect = true;
            var isSelected = jQuery.inArray(col, ip_GridProps[GridID].selectedColumn) != -1;

            if (e.which == 1) {

                if (isSelected && !shitClick && !ctrlClick) {
                    ip_EnableScrollAnimate(GridID);
                    ip_ShowColumnMove(GridID, col);
                    allowColSelect = false;
                }
                     
            }
            else if (isSelected) { allowColSelect = false; }

            //Do the column selection
            if (allowColSelect) {

                var rangeType = null;

                if (ip_GridProps[GridID].editing.editing && ip_GridProps[GridID].editing.editTool.isFX()) { rangeType = 'ip_grid_cell_rangeselector_transparent' }

                if (shitClick) {

                    if (ip_GridProps[GridID].selectedColumn.length > 0) {

                        //Select a bunch of outer columns
                        ip_GridProps[GridID].selectedColumn.sort(function (a, b) { return a - b });

                        var minCol = ip_GridProps[GridID].selectedColumn[0];
                        var maxCol = ip_GridProps[GridID].selectedColumn[ip_GridProps[GridID].selectedColumn.length - 1];
                        var count = 1;

                        if (col < minCol) { count = minCol - col; }
                        else if (col > maxCol) { count = col - maxCol; col = maxCol + 1; }

                        $('#' + GridID).ip_SelectColumn({ multiselect: true, col: col, count: count, rangeType: rangeType });

                    }
                    else {
                        $('#' + GridID).ip_SelectColumn({ multiselect: false, col: col, rangeType: rangeType });
                    }
                }
                else {
                    $('#' + GridID).ip_SelectColumn({ multiselect: (ip_GridProps[GridID].selectedColumn.length > 0 ? ctrlClick : false), col: col, rangeType: rangeType });
                }
            }

        }

    });

    //Deals with: moving selected rows
    $('#' + GridID).on('mousedown', '.ip_grid_rowSelecterCell', function (e) {

        if (ip_GridProps[GridID].editing.editing && ip_GridProps[GridID].editing.editTool.isFX()) { e.preventDefault(); return; }

        var shitClick = e.shiftKey;
        var ctrlClick = e.ctrlKey;
        var row = parseInt($(this).attr('row'));
        
        if (ip_GridProps['index'].browser.name == 'ie' && shitClick == null) {

            shitClick = ip_GridProps[GridID].shiftKey;
            ctrlClick = ip_GridProps[GridID].ctrlKey;
        }

        if (!ip_GridProps[GridID].resizing) {

            var allowRowSelect = true;
            var isSelected = jQuery.inArray(row, ip_GridProps[GridID].selectedRow) != -1;

            if (e.which == 1) {

                if (isSelected && !shitClick && !ctrlClick) {

                    allowRowSelect = false;
                    ip_EnableScrollAnimate(GridID);
                    ip_ShowRowMove(GridID, row);

                }

            }
            else if (isSelected) { allowRowSelect = false; }
     
            if (allowRowSelect) {

                var rangeType = null;

                if (ip_GridProps[GridID].editing.editing && ip_GridProps[GridID].editing.editTool.isFX()) { rangeType = 'ip_grid_cell_rangeselector_transparent' }

                if (shitClick) {

                    if (ip_GridProps[GridID].selectedRow.length > 0) {

                        //Select a bunch of outer rows
                        ip_GridProps[GridID].selectedRow.sort(function (a, b) { return a - b });

                        var minRow = ip_GridProps[GridID].selectedRow[0];
                        var maxRow = ip_GridProps[GridID].selectedRow[ip_GridProps[GridID].selectedRow.length - 1];
                        var count = 1;

                        if (row < minRow) { count = minRow - row; }
                        else if (row > maxRow) { count = row - maxRow; row = maxRow + 1; }

                        $('#' + GridID).ip_SelectRow({ multiselect: true, row: row, count: count, rangeType: rangeType });

                    }
                    else {
                        $('#' + GridID).ip_SelectRow({ multiselect: false, row: row, rangeType: rangeType });
                    }
                }
                else {
                    $('#' + GridID).ip_SelectRow({ multiselect: (ip_GridProps[GridID].selectedRow.length > 0 ? ctrlClick : false), row: row, rangeType: rangeType });
                }
            }
            
        }

    });
    
    //Deals with: selecting entire grid as range
    $('#' + GridID).on('click', '.ip_grid_columnSelectorCellCorner', function (e) {

        var shitClick = e.shiftKey;
        var col = -1;
        var row = -1;

        $('#' + GridID).ip_SelectColumn({ multiselect: false, col: col });
        $('#' + GridID).ip_SelectRow({ multiselect: true, row: row, fetchRange: false });

    });

    //Deals with: resize column, frozen columns
    $('#' + GridID).on('mousemove', '.ip_grid_columnSelectorCell', function (e) {
        
        var resizeActiveCell = null;
        var frozenActiveCell = null;
        
        if (!ip_GridProps[GridID].resizing) {

            var cellWidth = $(this).width();
            var offsetLeft = $(this).offset().left + parseInt($(this).css('border-left-width').replace('px',''));
            var pageX = e.pageX;
            var mouseX = pageX - offsetLeft;
            var DistanceToCellEnd = cellWidth - mouseX;

            //--This deals with the columns resizer tool
            var resizerActiveZone = (DistanceToCellEnd <= 5 && DistanceToCellEnd >= 0 ? true : false);
            if (resizerActiveZone && resizeActiveCell != this) {

                resizeActiveCell = this;
                ip_ShowColumnResizer(GridID, parseInt($(this).attr('col')));

            }

        }


    });
      
    //Deals with: resize row
    $('#' + GridID).on('mousemove', '.ip_grid_rowSelecterCell', function (e) {
        
        var activeCell = null;
        var frozenActiveCell = null;

        if (!ip_GridProps[GridID].resizing) {

            var cellHeight = $(this).height();
            var offsetTop = $(this).offset().top;
            var pageY = e.pageY;
            var mouseY = pageY - offsetTop;
            var resizerActiveZone = (cellHeight - mouseY <= 5 && cellHeight - mouseY >= 0 ? true : false);
            var DistanceToCellEnd = cellHeight - mouseY + 3;


            //this deals with the row resizer tool
            if (resizerActiveZone && activeCell != this) {

                activeCell = this;

                var row = parseInt($(this).parent().attr('row'));
                
                ip_ShowRowResizer(GridID, row);
                
                
            }

        }

    });
    
    //Deals with range movement, copy, cut, paste, edit cell, column data types        
    $('#' + GridID).on('keydown', ip_GridProps[GridID].events.SetupEvents_document_keydown = function (e) {

        

        if (!ip_GridProps[GridID].editing.editing) {

            if ((e.keyCode == 37 || e.keyCode == 38 || e.keyCode == 39 || e.keyCode == 40 || e.keyCode == 34 || e.keyCode == 33)) {

                //Arrow keys, page up & page down
                //This offers a major performance improvement on the re-iterating key down event
                
                if (ip_GridProps[GridID].timeouts.keyDownTimeout == null) { ip_KeyDownLoop(GridID, e, 0); }
                return false;

            }
            else if (e.ctrlKey && (e.keyCode == 67 || e.keyCode == 88 || e.keyCode == 86 || e.keyCode == 90)) {
                
                //Copy, cut, paste, undo
                if (e.keyCode == 67) { $("#" + GridID).ip_Copy(); }
                else if (e.keyCode == 86) { $("#" + GridID).ip_Paste({ changeFormula: true }); } //!ip_GridProps[GridID].cut
                else if (e.keyCode == 88) { $("#" + GridID).ip_Copy({ cut: true }); }
                else if (e.keyCode == 90) { $("#" + GridID).ip_Undo(); }
                

            }
            else if (e.keyCode == 46 ) {

                $('#' + GridID).ip_ResetRange({ valuesOnly: (e.ctrlKey || e.shiftKey ? false : true) });

            }
            else if (e.keyCode == 36 || e.keyCode == 35) {

                var row = ip_GridProps[GridID].selectedCell.row;
                var col = ip_GridProps[GridID].selectedCell.col;

                //Home, end, pagedown, pageup
                if (e.keyCode == 36) { $('#' + GridID).ip_SelectCell({ col: 0, row: row, scrollTo:true }); } //Home
                else if (e.keyCode == 35) { $('#' + GridID).ip_SelectCell({ col: ip_GridProps[GridID].cols - 1, row: row, scrollTo: true }); } //End
                //else if (e.keyCode == 34) { $('#' + GridID).ip_SelectCell({ row: row + ip_GridProps[GridID].loadedRows, col: col, scrollTo: true }); } //PageDown
                //else if (e.keyCode == 33) { $('#' + GridID).ip_SelectCell({ row: row - ip_GridProps[GridID].loadedRows, col: col, scrollTo: true }); } //PageDown

            }
            else {



                var Key = String.fromCharCode(e.keyCode);

                if (/[a-z0-9=]/i.test(Key) || e.keyCode == 187 || e.keyCode == 8) {

                    if (e.ctrlKey) {

                        //SHORTCUTS
                        var formatObject = ip_EnabledFormats(GridID);
                        if (Key == 'B') { $('#' + GridID).ip_FormatCell({ style: (formatObject.bold ? 'font-weight:;' : 'font-weight:bold;') }); }
                        else if (Key == 'I') { $('#' + GridID).ip_FormatCell({ style: (formatObject.italic ? 'font-style:;' : 'font-style:italic;') }); }
                        return false;

                    }
                    else {

                        if (ip_GridProps[GridID].selectedCell != null) {

                            //EDIT A CELL
                            var clear = (e.keyCode == 8 ? false : true);

                            $('#' + GridID).ip_TextEditable({
                                clear: clear,
                                contentType: 'ip_grid_CellEdit',
                                element: ip_GridProps[GridID].selectedCell,
                                //dropDown: {

                                //    allowEmpty: true,
                                //    validate: false,
                                //    autoComplete: true,                                
                                //    displayField: 'displayField' ,
                                //    data: ip_GetColumnValues(GridID, 0)                                

                                //}

                            });

                        }
                        else if (ip_GridProps[GridID].selectedColumn.length == 1) {

                            //GIVE A COLUMN A DATATYPE
                            var row = -1;
                            var col = ip_GridProps[GridID].selectedColumn[0];
                            var Quad = ip_GetQuad(GridID, row, col);

                            $('#' + GridID).ip_TextEditable({
                                contentType: 'ip_grid_ColumnType',
                                element: $('#' + GridID + '_q' + Quad + '_columnSelectorCell_' + col),
                                dropDown: {
                                    addNonMatched: true,
                                    allowEmpty: true,
                                    validate: true,
                                    autoComplete: true,
                                    keyField: 'dataType',
                                    displayField: [{ displayField: 'dataTypeName', style: 'font-weight:bold;' }, { displayField: 'dataType', style: 'font-style:italic;' }],
                                    data: ip_GridProps[GridID].dataTypes,
                                    noMatchData: { dataTypeName: null, dataType: '' }
                                }
                            });

                        }
                    }

                }
            }

        }

      
    });
    
    // Deals with range movement timeouts
    $(document).on('keyup', ip_GridProps[GridID].events.SetupEvents_document_keyup = function (e) {

        if (ip_GridProps[GridID].timeouts.keyDownTimeout != null) {
            clearTimeout(ip_GridProps[GridID].timeouts.keyDownTimeout);
            ip_GridProps[GridID].timeouts.keyDownTimeout = null;
        }

    });
    
    //Deals with: grid mouse wheel
    $("#" + GridID).mousewheel(function (event, delta, deltaX, deltaY) {

        var ScrollY = ip_GridProps[GridID].scrollY;

        var o = '';
        if (delta > 0) {
            ScrollY = ScrollY - 4;
        }
        else if (delta < 0) {
            ScrollY = ScrollY + 4;
        }

        //Make sure we dont scroll beyond our grid range
        if (ScrollY >= ip_GridProps[GridID].rows) { ScrollY = ip_GridProps[GridID].rows - 1; }
        else if (ScrollY < ip_GridProps[GridID].frozenRows) { ScrollY = ip_GridProps[GridID].frozenRows; }

        ip_ScrollToY(GridID, ScrollY, true, 'all', 'none');

        return false; // prevent default


    });

    //Deals with: selecting entire grid as range
    $('#' + GridID).on('click', '.ip_grid_cell_hideIcon', function (e) {

        //ip_NextNonHiddenCol
        var ParentCell = $(this).parent();
        var increment = $(this).hasClass('top') || $(this).hasClass('left') ? -1 : 1;
        var row = parseInt($(ParentCell).attr('row'));
        var col = parseInt($(ParentCell).attr('col'));
        var startVal = (col == -1 ? row : col);
        var endVal = (col == -1 ? ip_NextNonHiddenRow(GridID, row + increment, null, row, (increment > 0 ? 'down' : 'up'), false) : ip_NextNonHiddenCol(GridID, col + increment, null, col, (increment > 0 ? 'right' : 'left'), false))

        if (startVal > endVal) { startVal = endVal; endVal = (col == -1 ? row : col )}

        if (row != -1) { $('#' + GridID).ip_HideShowRows({ action: 'show', range: { startRow: startVal, endRow: endVal } }); }
        if (col != -1) { $('#' + GridID).ip_HideShowColumns({ action: 'show', range: { startCol: startVal, endCol: endVal } }); }

        return false;

    });
        
    //Deals with: selecting entire grid as range
    $('#' + GridID).on('mousedown', '.ip_grid_cell_groupIcon', function (e) {

        e.stopPropagation();

        var ParentCell = $(this).parent().parent();
        var row = parseInt($(ParentCell).attr('row'));        
        var startVal = row + 1;
        var endVal = startVal + ip_GridProps[GridID].rowData[row].groupCount - 1;
        var HideShow = (ip_GridProps[GridID].rowData[startVal].hide == true ? 'show' : 'hide')
        var rows = [];

        var GroupCounter = 0;
        for (var r = startVal; r <= endVal; r++) {

            if (ip_GridProps[GridID].rowData[r].groupCount != null && ip_GridProps[GridID].rowData[r].groupCount > 0) {
                if (endVal < r + ip_GridProps[GridID].rowData[r].groupCount) { endVal = r + ip_GridProps[GridID].rowData[r].groupCount; }
            }

            if (HideShow == 'show') {
                
                if (GroupCounter == 0) { rows[rows.length] = r; }

                if (ip_GridProps[GridID].rowData[r].groupCount > GroupCounter) { GroupCounter = ip_GridProps[GridID].rowData[r].groupCount;  }
                else if (GroupCounter > 0) { GroupCounter--; }

                
            }
            else {
                rows[rows.length] = r;
            }
        }
        

        if (row != -1) { $('#' + GridID).ip_HideShowRows({ action: HideShow, rows: rows, selectRows:false }); }
        

        return false;

    });

    //Deals with focus on the formula bar
    $('#' + GridID).on('mousedown', '.ip_grid_fbar_text', function (e) {

        //This is here to solve a chrome bug
        ip_EnableSelection(GridID);
        if (ip_GridProps[GridID].editing.editing) { e.preventDefault(); } //Stops text editable focusOut from firing

        
    });

    //Deals with focus on the formula bar
    $('#' + GridID).on('click', '.ip_grid_fbar_text', function (e) {


        if (ip_GridProps[GridID].selectedCell != null) {

            var CI = ip_GetCursorPos(GridID, this);
            var defaultValue = (ip_GridProps[GridID].editing.editing ? ip_GridProps[GridID].editing.editTool.text() : null);

            $('#' + GridID).ip_TextEditable({
                defaultValue: defaultValue,
                clear: false,
                contentType: 'ip_grid_fxBar',
                element: ip_GridProps[GridID].selectedCell,
                cursor: CI
            });

        }

    });

    //Deals with showing a cell dropdown
    $('#' + GridID).on('click', '.ip_grid_cell_dropdownIconContainer', function (e) {

        if (ip_GridProps[GridID].selectedCell != null) {

            var CI = { x: 0, length: $(ip_GridProps[GridID].selectedCell).text().length }
            $('#' + GridID).ip_TextEditable({ contentType: 'ip_grid_CellEdit', element: ip_GridProps[GridID].selectedCell, cursor: CI });

        }

    });
    
    //Deals with TextEditable range selection
    $('#' + GridID).on('ip_SelectRange', function (e, args) {

        //If editing, update formula with new range
        if (ip_GridProps[GridID].editing.editing) {

            var text = ip_GridProps[GridID].editing.editTool.text().trim();            
            var newRange = args.options.SelectRange.Effected;            
            var carret = (ip_GridProps[GridID].editing.carret == null ? { x: 0, length: 0 } : ip_GridProps[GridID].editing.carret );
            var fx = ip_fxObject(GridID, text, null, null, carret.x - 1);                             

            if (fx) {

                //Work out where in the formula we want to edit text
                var isRange = fx.parts[fx.focused.partIndex].match(ip_GridProps['index'].regEx.range);
 
                if (!isRange) {
                    if (fx.parts.length > (fx.focused.partIndex + 1) && fx.parts[fx.focused.partIndex + 1].match(ip_GridProps['index'].regEx.range)) { fx.focused.cursorI++; fx.focused.partIndex++; }
                    else { fx.focused.cursorI++; fx.focused.partIndex++; fx.parts.splice(fx.focused.partIndex, 0, ''); }
                }

                var prePart = fx.parts[fx.focused.partIndex];
                fx.parts[fx.focused.partIndex] = ip_fxRangeToString(GridID, newRange);

                carret.x = fx.focused.cursorI + fx.parts[fx.focused.partIndex].length; // { x: fx.focused.cursorI + fx.parts[fx.focused.partIndex].length, length: 0 }
                carret.length = 0;
                
                ip_EditToolHelp(GridID, fx.parts.join(""), carret, 0, true, true, true);
                

            }
            else{ ip_EnableSelection(GridID,false); }

        }
        else if (ip_GridProps[GridID].formatting.formatting) {

            var ranges = ip_ValidateRangeObjects(GridID, ip_GridProps[GridID].selectedRange);
            var rangeText = ip_fxRangeToString(GridID, ranges);
            //var newRange = args.options.SelectRange.Effected;
            //var rangeText = ip_fxRangeToString(GridID, newRange);
            $(ip_GridProps[GridID].formatting.focusedControl).val(rangeText);

        }

        

    });
}

function ip_ValidateData(GridID, rows, cols, loading) {

    if (loading == null) { loading = false; }

    var colCount = ip_GridProps[GridID].colData.length;
    var rowCount = ip_GridProps[GridID].rowData.length;
     

    //Validate columns
    if (colCount < cols) {
        for (var col = colCount; col < cols; col++) { ip_GridProps[GridID].colData[col] = ip_colObject(); }
    }

    //make sure we validate all rows
    if (rows == -1) {
        rows = ip_GridProps[GridID].rows;
        rowCount = 0;
    }

    if (rowCount < rows) {
        for (var row = rowCount; row < rows; row++) {

            if (ip_GridProps[GridID].rowData[row] == null) { ip_GridProps[GridID].rowData[row] = ip_rowObject(ip_GridProps[GridID].dimensions.defaultRowHeight, null, row, cols, loading); }
            else if (ip_GridProps[GridID].rowData[row].cells == null) { ip_GridProps[GridID].rowData[row] = ip_rowObject(ip_GridProps[GridID].dimensions.defaultRowHeight, null, row, cols, loading); }

            ip_ValidateDataCols(GridID, ip_GridProps[GridID].rowData[row], cols, row);

        }
    }

}

function ip_ValidateDataCols(GridID, rowObject, cols, rowIndex) {

    var colCount = rowObject.cells.length;

    if (cols == -1) {
        cols = ip_GridProps[GridID].cols;
        colCount = 0;
    }

    if(ip_GridProps[GridID].colData == null) { ip_GridProps[GridID].colData = new Array(cols); }

    if (colCount < cols) {
        for (var col = colCount; col < cols; col++) {

            var tmpVal = rowIndex + ' x ' + col;
            
            if (rowObject.cells[col] == null) { rowObject.cells[col] = ip_cellObject(tmpVal, null, null, rowIndex, col); }
            if (ip_GridProps[GridID].colData[col] == null) { ip_GridProps[GridID].colData[col] = ip_colObject(); }

        }
    }
    
}

function ip_SetupMerges(GridID, mergeData) {

    if (mergeData != null && mergeData.length > 0) {

        for (var i = 0; i < mergeData.length; i++) {

            var startRow = mergeData[i].mergedWithRow;
            var startCol = mergeData[i].mergedWithCol;
            var endRow = startRow + mergeData[i].rowSpan - 1;
            var endCol = startCol + mergeData[i].colSpan - 1;

            ip_SetCellMerge(GridID, startRow, startCol, endRow, endCol);

        }

    }

}

function ip_SetupFx(GridID) {

    $('#' + GridID).ip_AddFormula({ formulaName: 'range', functionName: 'ip_fxRange', tip: 'Specifies a range of cells, by default the first cell in range is returned.', inputs: '[row][col]:[row][col]', example: 'A1 or A1:B5' });
    $('#' + GridID).ip_AddFormula({ formulaName: 'count', functionName: 'ip_fxCount', tip: 'Counts the number of cells which have numberic values. Ignores cells that are empty or text.', inputs: '(range1, range2, ... )', example: 'count( a1, b1:b5, a3 )' });
    $('#' + GridID).ip_AddFormula({ formulaName: 'sum', functionName: 'ip_fxSum', tip: 'Adds the numbers in a range. Ignores cells that are empty or text.', inputs: '( range1, range2, ... )', example: 'sum( a1, b1:b5, a3 )' });
    $('#' + GridID).ip_AddFormula({ formulaName: 'concat', functionName: 'ip_fxConcat', tip: 'Joins the values of cells into one text string.', inputs: '( range1, range2, ... )', example: 'concat( a1, b1:b5, a3 )' });
    $('#' + GridID).ip_AddFormula({ formulaName: 'dropdown', functionName: 'ip_fxDropDown', tip: 'Fetches a range and returns the values as a dropdown object', inputs: '( range1, range2, ... )', example: 'dropdown( a1, b1:b5, a3 )' });
    $('#' + GridID).ip_AddFormula({ formulaName: 'gantt', functionName: 'ip_fxGantt', tip: 'Returns true or false if the base date falls within start and end date ranges', inputs: '( BaseDate, StartDate, EndDate, TaskName, ProjectName (optional) )', example: '<br/>gantt( 2014-07-15, 2014-07-01, 2014-07-31, Cost of sales report, General Management )<br/>gantt( today(0), a1, a2, Cost of sales report, General Management )' });
    $('#' + GridID).ip_AddFormula({ formulaName: 'max', functionName: 'ip_fxMax', tip: 'Returns the largest number in a range. Ignores cells that are empty or text.', inputs: '( range1, range2, ... )', example: 'sum( a1, b1:b5, a3 )' });
    $('#' + GridID).ip_AddFormula({ formulaName: 'today', functionName: 'ip_fxToday', tip: 'Returns todays date', inputs: '(increment in days)', example: '<br/>today( 0 )<br/>today( -1 )<br/>today( 1 )' });
    $('#' + GridID).ip_AddFormula({ formulaName: 'date', functionName: 'ip_fxDate', tip: 'Returns the current date', inputs: '(increment in days)', example: '<br/>date( 0 )<br/>date( -1 )<br/>date( 1 )' });
    $('#' + GridID).ip_AddFormula({ formulaName: 'day', functionName: 'ip_fxDay', tip: 'Returns the calendar day in month', inputs: '(increment in days)', example: '<br/>day( 0 )<br/>day( -1 )<br/>day( 1 )' });

}

function ip_SetupMask(GridID) {

    $('#' + GridID).ip_AddMask({
        title: 'Short Date', dataType: 'date', mask: 'yyyy-mm-dd',
        input: function (value) {
            var date = ip_parseDate(value, 'yyyy-mm-dd');
            if (!date) { date = ip_parseDate(value); }
            return date;
        },
        output: function (value, oldMask) {
            
            var date = ip_formatDate(GridID, value, oldMask, 'yyyy-mm-dd');
            if (date == false && typeof (date) == "boolean") { return value; }
            return date;

        }
    });
    $('#' + GridID).ip_AddMask({
        title: 'Short Date/Time', dataType: 'date', mask: 'yyyy-mm-dd hh:mm',
        input: function (value) {
            var date = ip_parseDate(value, 'yyyy-mm-dd hh:mm');
            if (!date) { date = ip_parseDate(value); }
            return date;
        },
        output: function (value, oldMask) {
            
            var date = ip_formatDate(GridID, value, oldMask, 'yyyy-mm-dd hh:mm');
            if (date == false && typeof (date) == "boolean") { return value; }
            return date;

        }
    });
    $('#' + GridID).ip_AddMask({
        title: 'Short Date/Time/Seconds', dataType: 'date', mask: 'yyyy-mm-dd hh:mm:ss',
        input: function (value) {
            var date = ip_parseDate(value, 'yyyy-mm-dd hh:mm:ss');
            if (!date) { date = ip_parseDate(value); }
            return date;
        },
        output: function (value, oldMask) {
            
            var date = ip_formatDate(GridID, value, oldMask, 'yyyy-mm-dd hh:mm:ss');
            if (date == false && typeof (date) == "boolean") { return value; }
            return date;

        }
    });
    $('#' + GridID).ip_AddMask({
        title: 'Long Date', dataType: 'date', mask: 'dd-mon-yyyy',
        input: function (value) {
            var date = ip_parseDate(value, 'dd-mon-yyyy');
            if (!date) { date = ip_parseDate(value); }
            return date;
        },
        output: function (value, oldMask) {
            
            var date = ip_formatDate(GridID, value, oldMask, 'dd-mon-yyyy');
            if (date == false && typeof (date) == "boolean") { return value; }
            return date;

        }
    });
    $('#' + GridID).ip_AddMask({
        title: 'Full Date', dataType: 'date', mask: 'Wed Jul 01 2015 00:00:00 GMT+0200 (US Standard Time)',
        input: function (value) {
            var date = ip_parseDate(value, 'full');
            if (!date) { date = ip_parseDate(value); }
            return date;
        },
        output: function (value, oldMask) {
            var date = ip_formatDate(GridID, value, oldMask, 'full');
            if (date == false && typeof (date) == "boolean") { return value; }
            return date;
        }
    });

    $('#' + GridID).ip_AddMask({
        title: 'Standard', dataType: 'number', mask: '123',
        input: function (value) {
            if (value == null) return null;
            var number = value.toString().replace(/[\s,]/gi, '');
            return ip_parseNumber(number);
        },
        output: function (value, oldMask, decimals) {
            var number = ip_formatNumber(GridID, value, oldMask, '123', decimals);
            if (number == false && typeof (number) == "boolean") { return value; }
            return number;
        }
    });
    $('#' + GridID).ip_AddMask({
        title: 'International', dataType: 'number', mask: '1 000 000.00',
        input: function (value) {
            if (value == null) return null;
            var number = value.toString().replace(/[\s,]/gi, '');
            return ip_parseNumber(number);
        },
        output: function (value, oldMask, decimals) {
            var number = ip_formatNumber(GridID, value, oldMask, '1 000 000.00', decimals);
            if (number == false && typeof (number) == "boolean") { return value; }
            return number;
        }
    });
    $('#' + GridID).ip_AddMask({
        title: 'US Number', dataType: 'number', mask: '1,000,000.00',
        input: function (value) {
            if (value == null) return null;
            var number = value.toString().replace(/[^0-9.]/gi, '');
            return ip_parseNumber(number);
        },
        output: function (value, oldMask, decimals) {

            var number = ip_formatNumber(GridID, value, oldMask, '1,000,000.00', decimals);
            if (number == false && typeof (number) == "boolean") { return value; }
            return number;
            
        }
    });
    $('#' + GridID).ip_AddMask({
        title: 'Test number', dataType: 'Test Number', mask: '1#000#000.00',
        input: function (value) {
            if (value == null) return null;
            var number = value.toString().replace(/[^0-9.]/gi, '');
            return ip_parseNumber(number);
        },
        output: function (value, oldMask, decimals) {

            if (value == null) return null;
            var number = value.toString().replace(/[\s,]/gi, '');
            number = ip_formatNumber(GridID, number, oldMask, '1,000,000.00', decimals);
            
            if (number == false && typeof (number) == "boolean") { return value; }

            number = number.replace(/[,]/gi, '#');

            return number
        }
    });

    $('#' + GridID).ip_AddMask({
        title: 'US Dollar', dataType: 'currency', mask: '$1 000.00',
        input: function (value) {
            if (value == null) return null;
            var number = value.toString().replace(/[^0-9.]/gi, '');
            return ip_parseNumber(number);
        },
        output: function (value, oldMask, decimals) {
            if (value == null || value == '') { return value }
            var currency = ip_formatCurrency(GridID, value, oldMask, '$1,000,000.00', decimals);
            if (currency == false && typeof (currency) == "boolean") { return value; }
            return currency
        }
    });
    $('#' + GridID).ip_AddMask({
        title: 'Rand', dataType: 'currency', mask: 'R1 000.00',
        input: function (value) {
            if (value == null) return null;
            var number = value.toString().replace(/[^0-9.]/gi, '');
            return ip_parseNumber(number);
        },
        output: function (value, oldMask, decimals) {
            if (value == null || value == '') { return value }
            var currency = ip_formatCurrency(GridID, value, oldMask, 'R1,000,000.00', decimals);
            if (currency == false && typeof (currency) == "boolean") { return value; }
            return currency
        }
    });

}


//----- RANGE ------------------------------------------------------------------------------------------------------------------------------------

function ip_ReloadRanges(GridID, fadeIn) {

    if (!fadeIn) { fadeIn = false; }

    //These are rages ...
    var ActiveRanges = ip_GridProps[GridID].selectedRange;

    $('#' + GridID).ip_RemoveRange();

    for (var i = 0; i < ActiveRanges.length; i++) {


        $('#' + GridID).ip_SelectRange({ multiselect: true, startCellOrdinates: ActiveRanges[i][0], endCellOrdinates: ActiveRanges[i][1], fadeIn: fadeIn });

    }

    ip_GridProps[GridID].selectedRange = ActiveRanges;


    //These are cell highlights...
    var ActiveRangeHighlights = ip_GridProps[GridID].highlightRange;

    $('#' + GridID).ip_RemoveRangeHighlight();

    for (var i = 0; i < ActiveRangeHighlights.length; i++) {


        $('#' + GridID).ip_RangeHighlight({ multiselect: true, startCellOrdinates: ActiveRangeHighlights[i][0], endCellOrdinates: ActiveRangeHighlights[i][1], borderStyle: ActiveRangeHighlights[i][2], borderColor: ActiveRangeHighlights[i][3], highlightType: ActiveRangeHighlights[i][4], fadeIn: fadeIn, expireTimeout: ActiveRangeHighlights[i][5] });

    }

    ip_GridProps[GridID].highlightRange = ActiveRangeHighlights;
}

function ip_RePoistionRanges(GridID, rangeType, xAxis, yAxis) {

    rangeType = (rangeType == null ? '' : rangeType);

    //These are ranges    
    if (rangeType == 'ip_grid_cell_rangeselector' || rangeType == 'all') {
        for (var i = 0; i < ip_GridProps[GridID].selectedRange.length; i++) {
            ip_RePoistionRange(GridID, false, ip_GridProps[GridID].selectedRange[i], xAxis, yAxis);
        }
    }

    //These are cell highlights...
    if (rangeType == 'ip_grid_cell_rangeHighlight' || rangeType == 'all') {
        for (var i = 0; i < ip_GridProps[GridID].highlightRange.length; i++) {
            ip_RePoistionRange(GridID, true, ip_GridProps[GridID].highlightRange[i], xAxis, yAxis);
        }
    }

}

function ip_RePoistionRange(GridID, IsRangeHighlight, rangeOrdinates, xAxis, yAxis) {

    
    var startCellOrdinates = rangeOrdinates[0];
    var endCellOrdinates = rangeOrdinates[1];
    var pivitCellOrdinates = rangeOrdinates[2];
    var pivitPositionOrdinates = rangeOrdinates[3];
    var startRow = startCellOrdinates[0, 0];
    var startCol = startCellOrdinates[0, 1];
    var endRow = endCellOrdinates[0, 0];
    var endCol = endCellOrdinates[0, 1];
    var startCell = ip_CellHtml(GridID, startRow, startCol, true);
    var endCell = ip_CellHtml(GridID, endRow, endCol, true);
    var Range = $('#' + GridID + (IsRangeHighlight ? '_RangeHighlight_' : '_range_') + startRow + '_' + startCol + '_' + endRow + '_' + endCol + ':not(.ip_grid_cell_rangeselector_move)');


    if (Range.length > 0) {

        if (yAxis) {

            var NewTop = ip_CalculateRangeTop(GridID, Range, startCell, startRow, startCol, endRow, endCol);
            pivitPositionOrdinates[0, 0] = NewTop;
            $(Range).css('top', NewTop + 'px');

            

            if (!$(Range).hasClass('ip_grid_cell_rangeselector_Resize')) {

                //Reposition the range as per normal       

                var NewHeight = ip_SetRangeHeight(GridID, Range, startCell, endCell, startCellOrdinates, endCellOrdinates)
                
                $(Range).height(NewHeight);


            }
            else {

                //This is used when scrolling and dragging a range //
                ip_CalculateRangeDimensions(GridID, Range, null, null, '', $(startCell).height(), $(startCell).width(), NewTop, pivitPositionOrdinates[0,1]);

            }

        }

        if (xAxis) {

            var NewLeft = ip_CalculateRangeLeft(GridID, Range, startCell, startRow, startCol, endRow, endCol);
            pivitPositionOrdinates[0, 1] = NewLeft;
            $(Range).css('left', NewLeft + 'px');


            if (!$(Range).hasClass('ip_grid_cell_rangeselector_Resize')) {

                //Reposition the range as per normal 
                var NewWidth = ip_SetRangeWidth(GridID, Range, startCell, endCell, startCellOrdinates, endCellOrdinates, null, NewLeft)
                $(Range).width(NewWidth);
            }
            else {

                //This is used when scrolling and dragging a range //
                ip_CalculateRangeDimensions(GridID, Range, null, null, '', $(startCell).height(), $(startCell).width(), pivitPositionOrdinates[0, 0], NewLeft);

            }

        }
    }
}

function ip_CalculateRangeDimensions(GridID, Range, rangeIndex, mouseArgs, axis, MinHeight, MinWidth, AnchorTop, AnchorLeft) {

    if (mouseArgs != null) {

        ip_GridProps[GridID].currentMousePos.pageX = mouseArgs.pageX;
        ip_GridProps[GridID].currentMousePos.pageY = mouseArgs.pageY;
    }

    if (AnchorTop == null || AnchorLeft == null) {

        var rangeIndex = (rangeIndex == null ? ip_GetRangeIndex(GridID, Range) : rangeIndex);
        AnchorTop = (AnchorTop == null ? ip_GridProps[GridID].selectedRange[rangeIndex][3][0, 0] : AnchorTop);
        AnchorLeft = (AnchorLeft == null ? ip_GridProps[GridID].selectedRange[rangeIndex][3][0, 1] : AnchorLeft);

    }

    if (axis != 'x') {


        var anchorTop = Math.round(AnchorTop); //parseInt($(Range).attr('anchorTop'));
        var PositionTop = Math.round($(Range).position().top);
        var OffsetTop = Math.round($(Range).offset().top);
        var NewHeight = ip_GridProps[GridID].currentMousePos.pageY - OffsetTop;
        var NewTop = anchorTop;


        

        if (NewHeight <= 0 || PositionTop <= anchorTop - 1) { // || PositionTop < anchorTop 

            //Range resizing happens to the top
            NewTop = PositionTop + NewHeight + 10;
            NewHeight = (anchorTop - NewTop + MinHeight); // + RangeHeight + (NewHeight * -1);
                       

            if (NewHeight < MinHeight) { NewHeight = MinHeight; }
            if (NewTop > anchorTop) { NewTop = anchorTop; }

            $(Range).css('top', NewTop + 'px');
            $(Range).height(NewHeight);
            
            
        }
        else {


            //Range resizing happens to the bottom
            NewHeight -= 10;
            if (NewHeight < MinHeight) { NewHeight = MinHeight; }
            if (NewTop < anchorTop) { NewTop = anchorTop; }

            $(Range).css('top', NewTop + 'px');
            $(Range).height(NewHeight);

          

        }
        

    }

    if (axis != 'Y') {

        var anchorLeft = AnchorLeft; //parseInt($(Range).attr('anchorLeft'));
        var PositionLeft = Math.round($(Range).position().left);        
        var OffsetLeft = Math.round($(Range).offset().left);
        var NewWidth = ip_GridProps[GridID].currentMousePos.pageX - OffsetLeft;
        var NewLeft = anchorLeft;
        
        if (NewWidth < 0 || PositionLeft < anchorLeft) {  // || PositionLeft < anchorLeft

            //Range resizing happens to the left
            NewLeft = PositionLeft + NewWidth + 10;
            NewWidth = (anchorLeft - NewLeft + MinWidth);

            if (NewWidth < MinWidth) { NewWidth = MinWidth; }
            if (NewLeft > anchorLeft) { NewLeft = anchorLeft; }

            $(Range).css('left', NewLeft + 'px');
            $(Range).width(NewWidth);

        }
        else {

            NewWidth -= 10;

            //Range resizing happens to the right
            if (NewWidth < MinWidth) { NewWidth = MinWidth; }
            if (NewLeft < anchorLeft) { NewLeft = anchorLeft; }

            $(Range).css('left', NewLeft + 'px');
            $(Range).width(NewWidth);

        }
    }
    
}

function ip_SetRangeWidth(GridID, Range, startCell, endCell, startCellOrdinates, startEndOrdinates) {

    //Note startCell, endCell is no longer used except for border

    var width = 0;

    if (startCell == null && endCell == null) { return width; }

    var startRow = startCellOrdinates[0, 0];
    var startCol = startCellOrdinates[0, 1];
    var endRow = startEndOrdinates[1, 0];
    var endCol = startEndOrdinates[1, 1];
    var loadedRowsCols = ip_LoadedRowsCols(GridID, false, false, true);

    var sCol = startCol;
    var eCol = endCol;

    // Calculate frozen zone
    if (startCol < ip_GridProps[GridID].frozenCols) {

        if (endCol < ip_GridProps[GridID].frozenCols) { width -= parseInt($(startCell).css("border-left-width").replace('px', '')) + parseInt($(startCell).css("border-right-width").replace('px', '')); }

        if (startCol < loadedRowsCols.colFrom_frozen) { sCol = loadedRowsCols.colFrom_frozen; }
        if (endCol > loadedRowsCols.colTo_frozen) { eCol = loadedRowsCols.colTo_frozen; }

        for (var c = sCol; c <= eCol; c++) { width += ip_ColWidth(GridID, c, true); }

    }

    // Calculate scroll zone
    if (endCol >= ip_GridProps[GridID].frozenCols && endCol <= loadedRowsCols.colTo_scroll) {

        if (width == 0) { width -= parseInt($(startCell).css("border-left-width").replace('px', '')) + parseInt($(startCell).css("border-right-width").replace('px', '')); }

        sCol = startCol;
        eCol = endCol;

        if (startCol < loadedRowsCols.colFrom_scroll) { sCol = loadedRowsCols.colFrom_scroll; }
        if (endCol > loadedRowsCols.colTo_scroll) { eCol = loadedRowsCols.colTo_scroll; }

        if (startCol > loadedRowsCols.colTo_scroll) { return width; }
        if (endCol < loadedRowsCols.colFrom_scroll) { return width; }

        for (var c = sCol; c <= eCol; c++) { width += ip_ColWidth(GridID, c, true); }

    }
    else if (endCol >= loadedRowsCols.colTo_scroll) {

        width += ip_GridProps[GridID].dimensions.accumulativeScrollWidth;

    }


    return width;

}

function ip_SetRangeHeight(GridID, Range, startCell, endCell, startCellOrdinates, startEndOrdinates) {

    //Note startCell, endCell is no longer used except for border

    var height = 0;

    if (startCell == null && endCell == null) { return 0; }

    var startRow = startCellOrdinates[0, 0];
    var startCol = startCellOrdinates[0, 1];
    var endRow = startEndOrdinates[1, 0];
    var endCol = startEndOrdinates[1, 1];
    var loadedRowsCols = ip_LoadedRowsCols(GridID, false, true, false);

    var sRow = startRow;
    var eRow = endRow;

    // Calculate frozen zone
    if (startRow < ip_GridProps[GridID].frozenRows) {

        //if (endRow < ip_GridProps[GridID].frozenRows) { height -= parseInt($(startCell).css("border-top-width").replace('px', '')) + parseInt($(startCell).css("border-bottom-width").replace('px', '')); }

        if (startRow < loadedRowsCols.rowFrom_frozen) { sRow = loadedRowsCols.rowFrom_frozen; }
        if (endRow > loadedRowsCols.rowTo_frozen) { eRow = loadedRowsCols.rowTo_frozen; }

        for (var r = sRow; r <= eRow; r++) { height += ip_RowHeight(GridID, r, true); }

    }

    // Calculate scroll zone
    if (endRow >= ip_GridProps[GridID].frozenRows && endRow <= loadedRowsCols.rowTo_scroll) {

        //if (height == 0) { height -= parseInt($(startCell).css("border-top-width").replace('px', '')) + parseInt($(startCell).css("border-bottom-width").replace('px', '')); }
        //height -= parseInt($(startCell).css("border-top-width").replace('px', '')) + parseInt($(startCell).css("border-bottom-width").replace('px', '')); 

        sRow = startRow;
        eRow = endRow;

        if (startRow < loadedRowsCols.rowFrom_scroll) { sRow = loadedRowsCols.rowFrom_scroll; }
        if (endRow > loadedRowsCols.rowTo_scroll) { eRow = loadedRowsCols.rowTo_scroll; }

        if (startRow > loadedRowsCols.rowTo_scroll) { return height; }
        if (endRow < loadedRowsCols.rowFrom_scroll) { return height; }

        for (var r = sRow; r <= eRow; r++) { height += ip_RowHeight(GridID, r, true); }

    }
    else if (endRow >= loadedRowsCols.rowTo_scroll) {

        height += ip_GridProps[GridID].dimensions.accumulativeScrollHeight; // $('#' + GridID + '_q3').outerHeight()

    }

    if (endRow < ip_GridProps[GridID].frozenRows || startRow >= ip_GridProps[GridID].frozenRows) { height -= parseInt($(startCell).css("border-top-width").replace('px', '')) + parseInt($(startCell).css("border-bottom-width").replace('px', '')); }

    return height;

}

function ip_ChangeRange(GridID, RangeOridantes, endCell, animateSpeed, rowIncrement, colIncrement, resetPivit, scrollIncrement) {
    //NOTE, Required: GridID, RangeOridantes

    if (ip_GridProps[GridID].selectedRange.length > 0) {

        ip_DisableSelection(GridID, true);

        //Initialize & validate variables
        resetPivit = (resetPivit == null ? false : resetPivit);
        rowIncrement = (rowIncrement == null ? 0 : rowIncrement);
        colIncrement = (colIncrement == null ? 0 : colIncrement);                
        animateSpeed = (animateSpeed == null ? 80 : animateSpeed);
        scrollIncrement = (scrollIncrement == null ? 0 : scrollIncrement);

        //Calcuations for new start/end cells
        var startRowCurrent = RangeOridantes[0][0, 0];
        var startColCurrent = RangeOridantes[0][0, 1];
        var endRowCurrent = RangeOridantes[1][0, 0];
        var endColCurrent = RangeOridantes[1][0, 1];
        var CurrentQuad = ip_GetQuad(GridID, startRowCurrent, startColCurrent);
        var MoveToQuad = CurrentQuad;
        var pivitRow = RangeOridantes[2][0, 0];
        var pivitCol = RangeOridantes[2][0, 1];
        var startRow = startRowCurrent;
        var startCol = startColCurrent;
        var endRow = (endCell == null ? (startRowCurrent < pivitRow ? startRowCurrent : endRowCurrent) : parseInt($(endCell).attr('row'))) + rowIncrement;
        var endCol = (endCell == null ? (startColCurrent < pivitCol ? startColCurrent : endColCurrent) : parseInt($(endCell).attr('col'))) + colIncrement;
        var NewTop = -2;
        var NewLeft = -2;
        var ReGetEndCell = true;
        var ReGetStartCell = true;
        var Range = $('#' + GridID + '_range_' + startRowCurrent + '_' + startColCurrent + '_' + endRowCurrent + '_' + endColCurrent); //MyGrid_range_10_4_10_4
        var rangeIndex = ip_GetRangeIndex(GridID, Range);
        var startCell = null;
        var direction = '';
        var scroll = null;
        var TransactionID = ip_GenerateTransactionID();
        

        //Calculate new start/end rows
        if (endRow < pivitRow) { startRow = endRow; endRow = pivitRow; }
        else { startRow = pivitRow; }


        if (endCol < pivitCol) { startCol = endCol; endCol = pivitCol; }
        else { startCol = pivitCol; }

        

        //work out direction
        if (startRow < startRowCurrent || endRow < endRowCurrent) { direction = 'up'; }
        else if (startRow > startRowCurrent || endRow > endRowCurrent) { direction = 'down'; }
        else if (startCol < startColCurrent || endCol < endColCurrent) { direction = 'left'; }
        else if (startCol > startColCurrent || endCol > endColCurrent) { direction = 'right'; }

        //Validate & calculate how to manage the scroll taking into account frozen rows / cols
        if (scrollIncrement != 0) {

            var frozenRows = ip_GridProps[GridID].frozenRows;
            var frozenCols = ip_GridProps[GridID].frozenCols;

            //if (direction == 'down' && endRow == frozenRows) { endRow = ip_GridProps[GridID].scrollY; }
            if (direction == 'down' && endRow >= frozenRows && endRow <= ip_GridProps[GridID].scrollY) { endRow = ip_GridProps[GridID].scrollY; }
            else if (direction == 'right' && endCol >= frozenCols && endCol <= ip_GridProps[GridID].scrollX) { endCol = ip_GridProps[GridID].scrollX; }

            if (direction == 'down' && startRow <= frozenRows && endRow == endRowCurrent) { scrollIncrement = 0; }
            else if (direction == 'right' && startCol <= frozenCols && endCol == endColCurrent) { scrollIncrement = 0; }

            if (rowIncrement != 0 && endRow <= frozenRows + (direction == 'up' ? -1 : 0)) { scrollIncrement = 0; }
            else if (colIncrement != 0 && endCol <= frozenCols + (direction == 'left' ? -1 : 0)) { scrollIncrement = 0; }
            
        }
        
        

        //Validate new positions
        if (startRow < 0) { startRow = 0; }
        if (startCol < 0) { startCol = 0; }
        if (endRow < 0) { endRow = 0; }
        if (endCol < 0) { endCol = 0; }

        if (startRow >= ip_GridProps[GridID].rows) { startRow = ip_GridProps[GridID].rows - 1; }
        if (startCol >= ip_GridProps[GridID].cols) { startCol = ip_GridProps[GridID].cols - 1; }
        if (endRow >= ip_GridProps[GridID].rows) { endRow = ip_GridProps[GridID].rows - 1; }
        if (endCol >= ip_GridProps[GridID].cols) { endCol = ip_GridProps[GridID].cols - 1; }


        //Validate merged cells
        var MergedCells = ip_GetRangeMergedCells(GridID, startRow, startCol, endRow, endCol);
        startRow = MergedCells.startRow;
        startCol = MergedCells.startCol;
        endRow = MergedCells.endRow;
        endCol = MergedCells.endCol;


        var startCellOrdinates = [startRow, startCol];
        var endCellOrdinates = [endRow, endCol];

        //Scroll to new cell
        if (scrollIncrement != 0) {

            scroll = ip_ScrollCell(GridID, direction, startRow, startCol, endRow, endCol, scrollIncrement, 'ip_grid_cell_rangeHighlight');

        }


        //Get new start/end cells
        startCell = ip_CellHtml(GridID, startRow, startCol, true);
        endCell = ip_CellHtml(GridID, endRow, endCol, true);




        //Calcuate new top
        NewTop = ip_CalculateRangeTop(GridID, Range, startCell, startRow, startCol, endRow, endCol);
        NewLeft = ip_CalculateRangeLeft(GridID, Range, startCell, startRow, startCol, endRow, endCol);
        MoveToQuad = ip_GetQuad(GridID, startRow, startCol);
        
                
        //Caclulate ranges new x/y co-ordinates
        if (CurrentQuad != MoveToQuad) { $(Range).appendTo("#" + GridID + '_q' + MoveToQuad + '_div_container'); }
        if (NewTop >= -1) { $(Range).css('top', NewTop + 'px'); }
        if (NewLeft >= -1) { $(Range).css('left', NewLeft + 'px'); }
        
        //Calculate ranges new height
        var NewWidth = ip_SetRangeWidth(GridID, Range, startCell, endCell, startCellOrdinates, endCellOrdinates);
        var NewHeight = ip_SetRangeHeight(GridID, Range, startCell, endCell, startCellOrdinates, endCellOrdinates);

        

        //Update range properties
        RangeID = ip_setRangeID(GridID, Range, startCellOrdinates, endCellOrdinates); //adjust the range id
        ip_GridProps[GridID].selectedRange[rangeIndex][0] = startCellOrdinates;
        ip_GridProps[GridID].selectedRange[rangeIndex][1] = endCellOrdinates;
        if (resetPivit) {
            ip_GridProps[GridID].selectedRange[rangeIndex][2] = startCellOrdinates;
            ip_GridProps[GridID].selectedRange[rangeIndex][3] = [NewTop,NewLeft];
        }

        if (ip_GridProps[GridID].selectedColumn.indexOf(startColCurrent) > -1) { $('#' + GridID).ip_UnselectColumn({ col: startColCurrent }); }
        if (ip_GridProps[GridID].selectedRow.indexOf(startRowCurrent) > -1) { $('#' + GridID).ip_UnselectRow({ col: startRowCurrent }); }

        $(Range).removeClass('ip_grid_cell_rangeselector_Resize');

        $(Range).width(NewWidth);
        $(Range).height(NewHeight);


        if (!resetPivit && ip_GridProps[GridID].mouseButton == 1) {
            ip_GridProps[GridID].events.moveRange_mouseUp = ip_UnBindEvent(document, 'mouseup', ip_GridProps[GridID].events.moveRange_mouseUp);
            $(document).mouseup(ip_GridProps[GridID].events.moveRange_mouseUp = function () {

                ip_GridProps[GridID].events.moveRange_mouseUp = ip_UnBindEvent(document, 'mouseup', ip_GridProps[GridID].events.moveRange_mouseUp);
                ip_GridProps[GridID].selectedRange[rangeIndex][2] = startCellOrdinates;
                ip_GridProps[GridID].selectedRange[rangeIndex][3] = [NewTop, NewLeft];

            });
        }

        var newRange = { startRow: startCellOrdinates[0], startCol: startCellOrdinates[1], endRow: endCellOrdinates[0], endCol: endCellOrdinates[1] };

        ip_GridProps[GridID].resizing = false;
                
        var Effected = newRange;
        ip_RaiseEvent(GridID, 'ip_SelectRange', ip_GeneratePublicKey(), { SelectRange: { Inputs: null, Effected: Effected } });

        return newRange;
    }

}

function ip_SelectTextTool(GridID, startRow, startCol, endRow, endCol, append) {
    
    append = (append == null ? false : append);
    startRow = (startRow == null ? 0 : startRow);
    startCol = (startRow == null ? 0 : startCol);
    endRow = (startRow == null ? -1 : endRow);
    endCol = (startRow == null ? -1 : endCol);

    if (startRow < 0) { startRow = 0; }
    if (startCol < 0) { startCol = 0; }

    var ClipBoardText = '';

    for (var r = startRow; r <= endRow; r++) {
        for (var c = startCol; c <= endCol; c++) {

            ClipBoardText += ip_GridProps[GridID].rowData[r].cells[c].value + (c < endCol ? '\t' : '');

        }

        ClipBoardText += (r < endRow ? '\r\n' : '')
    }

    if (append) { ClipBoardText = $('#' + GridID + '_selectTool').val() + '\r\n' + ClipBoardText; }
    $('#' + GridID + '_selectTool').val(ClipBoardText);
    

    $('#' + GridID).enableSelection();
    $('#' + GridID + '_selectTool').select();
    $(document).disableSelection();

}

function ip_GetRangeIndex(GridID, Range) {

    var startRow = parseInt($(Range).attr('startRow'));
    var startCol = parseInt($(Range).attr('startCol'));
    var endRow = parseInt($(Range).attr('endRow'));
    var endCol = parseInt($(Range).attr('endCol'));

    for (var i = 0; i < ip_GridProps[GridID].selectedRange.length; i++) {

        if (ip_GridProps[GridID].selectedRange[i][0][0] == startRow
            && ip_GridProps[GridID].selectedRange[i][0][1] == startCol
            && ip_GridProps[GridID].selectedRange[i][1][0] == endRow
            && ip_GridProps[GridID].selectedRange[i][1][1] == endCol) {

            return i;
        }
    }

    return -1;
}

function ip_CalculateRangeTop(GridID, Range, startCell, startRow, startCol, endRow, endCol, allowHeader) {

    var row = parseInt($(startCell).attr('row'));
    var Quad = ip_GetQuad(GridID, row);

    //Validate if range falls outside of scroll area
    if (startRow >= ip_GridProps[GridID].frozenRows) {
        if (startRow > ip_GridProps[GridID].scrollY + ip_GridProps[GridID].loadedRows) { return ip_GridProps[GridID].dimensions.scrollHeight; }
        else if (endRow < ip_GridProps[GridID].scrollY) { return 0; }
    }


    //Row in scroll area    
    var col = parseInt($(startCell).attr('col'));
    var pos = ip_CellRelativePosition(GridID, row, col);
    var yCellBorder = 0; //parseInt($(startCell).css("border-top-width").replace('px', '')); //+ parseInt($(startCell).css("border-bottom-width").replace('px', ''));
    var yControllerBorder = parseInt($(Range).css("border-top-width").replace('px', ''));

        
    var NewTop = pos.localTop - yControllerBorder + yCellBorder; // + yCellBorder; //$(startCell).position().top //pos.localTop
        
    if (allowHeader != true && row == -1) {
        NewTop += ip_GridProps[GridID].dimensions.columnSelectorHeight + yControllerBorder;
    }
    else if (allowHeader == true && row == -1) { NewTop += yControllerBorder; }

    //make sure range dones not vanish into hidden scroll area
    if (Quad == 3 || Quad == 4) {
        if (NewTop > $('#' + GridID + '_q4_scrollbar_container_x').position().top) {
            NewTop = $('#' + GridID + '_q4_scrollbar_container_x').position().top - yControllerBorder;
        }
    }

    if (NewTop < -1) { NewTop = -1; }




    return NewTop;
}

function ip_CalculateRangeLeft(GridID, Range, startCell, startRow, startCol, endRow, endCol, allowHeader) {


    //Validate if range falls outside of scroll area
    if (startCol >= ip_GridProps[GridID].frozenCols) {
        if (startCol > ip_GridProps[GridID].scrollX + ip_GridProps[GridID].loadedCols) { return ip_GridProps[GridID].dimensions.scrollWidth; }
        else if (endCol < ip_GridProps[GridID].scrollX) { return 0; }
    }

    var row = parseInt($(startCell).attr('row'));
    var col = parseInt($(startCell).attr('col'));
    var pos = ip_CellRelativePosition(GridID, row, col);
    var xCellBorder = parseInt($(startCell).css("border-left-width").replace('px', '')) + parseInt($(startCell).css("border-right-width").replace('px', ''));
    var xControllerBorder = parseInt($(Range).css("border-left-width").replace('px', '')) + parseInt($(startCell).css("border-right-width").replace('px', ''));
    var NewLeft = pos.localLeft - xControllerBorder + xCellBorder;

    //This is done purly because pointer-events doesnt work in IE, remove it for IE 10
    if (allowHeader != true && col == -1) {
        NewLeft += ip_GridProps[GridID].dimensions.rowSelectorWidth + (xCellBorder / 2);
    }
    else if (allowHeader == true && col == -1) {
        NewLeft += (xCellBorder / 2);
    }
    

    if (NewLeft < -1) { NewLeft = -(xControllerBorder - xCellBorder); }

    return NewLeft;
}

function ip_setRangeID(GridID, Range, startCellOrdinates, endCellOrdinates) {

    var startRow = startCellOrdinates[0, 0];
    var startCol = startCellOrdinates[0, 1];
    var endRow = endCellOrdinates[0, 0];
    var endCol = endCellOrdinates[0, 1];
    var RangeID = GridID + '_range_' + startRow + '_' + startCol + '_' + endRow + '_' + endCol;

    $(Range).attr('id', RangeID)
    $(Range).attr('startRow', startRow)
    $(Range).attr('startCol', startCol)
    $(Range).attr('endRow', endRow)
    $(Range).attr('endCol', endCol)


    return RangeID;
}

function ip_ValidateRangeOptions(GridID, options) {

    var options = $.extend({

        range: { startRow: null, startCol: null, endRow: null, endCol: null },
        startCellOrdinates: [0, 0], //[row,col]
        endCellOrdinates: null, //[row,col]
        startCell: null, //td object
        endCell: null, //td object
        mergedstartRow: null,
        mergedEndRow: null,
        mergedEndCol: null,
        mergedEndEnd: null,
        returnLastIfNull: true,
        considerMerges: true

    }, options);
    

    //Start cell
    if (options.range.startRow != null  && options.range.startCol != null) {

        options.startCellOrdinates = [options.range.startRow, options.range.startCol];
        options.startCell = ip_CellHtml(GridID, options.startCellOrdinates[0, 0], options.startCellOrdinates[0, 1], true);

    }
    else if (options.startCell) {

        options.startCellOrdinates = [parseInt($(options.startCell).attr('row')), parseInt($(options.startCell).attr('col'))]; //set the co-ordinates of the startCell
        options.range.startRow = options.startCellOrdinates[0, 0];
        options.range.startCol = options.startCellOrdinates[0, 1];

    }
    else {

        options.range.startRow = options.startCellOrdinates[0, 0];
        options.range.startCol = options.startCellOrdinates[0, 1];
        options.startCell = ip_CellHtml(GridID, options.startCellOrdinates[0, 0], options.startCellOrdinates[0, 1], true);

    }

    //End cell
    if (options.range.endRow != null && options.range.endCol != null) {

        options.endCellOrdinates = [options.range.endRow, options.range.endCol];
        options.endCell = ip_CellHtml(GridID, options.endCellOrdinates[0, 0], options.endCellOrdinates[0, 1], true);

    }
    else if (options.endCell) {

        options.endCellOrdinates = [parseInt($(options.endCell).attr('row')), parseInt($(options.endCell).attr('col'))]; //set the co-ordinates of the startCell
        options.range.endRow = options.startCellOrdinates[0, 0];
        options.range.endCol = options.startCellOrdinates[0, 1];

    }
    else if (options.endCellOrdinates) {

        options.range.endRow = options.endCellOrdinates[0, 0];
        options.range.endCol = options.endCellOrdinates[0, 1];
        options.endCell = ip_CellHtml(GridID, options.endCellOrdinates[0, 0], options.endCellOrdinates[0, 1], true);

    }
    else {

        options.endCellOrdinates = [options.startCellOrdinates[0,0],options.startCellOrdinates[0,1]];
        options.range.endRow = options.endCellOrdinates[0, 0];
        options.range.endCol = options.endCellOrdinates[0, 1];
        options.endCell = options.startCell; 

    }

    //Validate that the range is inside the grid area
    if (options.range.startRow < (ip_GridProps[GridID].selectedColumn.length > 0 ? -1 : 0)) { return false; }
    if (options.range.startCol < (ip_GridProps[GridID].selectedRow.length > 0 ? -1 : 0)) { return false; }
    if (options.range.endRow < 0) { return false; }
    if (options.range.endCol < 0) { return false;  }
    if (options.range.startRow >= ip_GridProps[GridID].rows) { return false; }
    if (options.range.startCol >= ip_GridProps[GridID].cols) { return false; }
    if (options.range.endRow >= ip_GridProps[GridID].rows) { return false; }
    if (options.range.endCol >= ip_GridProps[GridID].cols) { return false; }

    //Validate merged cells

    if (options.considerMerges) {
        var MergedCells = ip_GetRangeMergedCells(GridID, options.startCellOrdinates[0, 0], options.startCellOrdinates[0, 1], options.endCellOrdinates[0, 0], options.endCellOrdinates[0, 1]);

        if (MergedCells.endMerge) {

            endCell = ip_CellHtml(GridID, MergedCells.endRow, MergedCells.endCol, true);
            options.endCell = endCell;
            options.endCellOrdinates[0, 0] = MergedCells.endRow;
            options.endCellOrdinates[0, 1] = MergedCells.endCol;
            options.range.endRow = MergedCells.endRow;
            options.range.endCol = MergedCells.endCol;
            options.mergedEndRow = MergedCells.endRow;
            options.mergedEndCol = MergedCells.endCol;




            startCell = ip_CellHtml(GridID, MergedCells.startRow, MergedCells.startCol, true);           

            options.startCell = startCell;
            options.startCellOrdinates[0, 0] = MergedCells.startRow;
            options.startCellOrdinates[0, 1] = MergedCells.startCol;
            options.range.startRow = MergedCells.startRow;
            options.range.startCol = MergedCells.startCol;
            options.mergedStartRow = MergedCells.startRow;
            options.mergedStartCol = MergedCells.startCol;

        }
    }


    

    return options;
}

function ip_ShowRangeMove(GridID, Range, mouseEventArgs, whichBorder) {

    ip_GridProps[GridID].resizing = true;
    
    $('#' + GridID).ip_RemoveRangeHighlight();

    //This is done to improve the accuracy of the move
    var borderPosition = $(whichBorder).attr('borderPosition');
    var mouseArgs = { clientX: mouseEventArgs.clientX, clientY: mouseEventArgs.clientY }

    if (borderPosition == 'left') { mouseArgs.clientX += 3; }
    if (borderPosition == 'right') { mouseArgs.clientX -= 3; }
    if (borderPosition == 'top') { mouseArgs.clientY += 3; }
    if (borderPosition == 'bottom') { mouseArgs.clientY -= 3; }

    var ClickStartCell = ip_SetHoverCell(GridID, Range, mouseArgs, true);
    var ClickStartRow = parseInt($(ClickStartCell).attr('row'));
    var ClickStartCol = parseInt($(ClickStartCell).attr('col'));    
    var Ranges = new Array();

    $('#' + GridID).ip_Copy({ cut: true, toClipBoard:false });

    var SelectedRanges = $('#' + GridID + ' .ip_grid_cell_rangeselector_selected');
    $(SelectedRanges).addClass('ip_grid_cell_rangeselector_move');
    $(SelectedRanges).children('.ip_grid_cell_rangeselector_key').hide();
    $(SelectedRanges).children('.ip_grid_cell_rangeselector_border').hide();

    //Record the co-ordinates of range
    for (var r = 0; r < ip_GridProps[GridID].selectedRange.length; r++) {

        var RangeID = GridID + '_range_' + ip_GridProps[GridID].selectedRange[r][0][0, 0] + '_' + +ip_GridProps[GridID].selectedRange[r][0][0, 1] + '_' + +ip_GridProps[GridID].selectedRange[r][1][1, 0] + '_' + +ip_GridProps[GridID].selectedRange[r][1][1, 1];
        var elRange = $('#' + RangeID);
        
        Ranges[r] = {
            StartX: $(elRange).position().left,
            StartY: $(elRange).position().top,
            StartCellOrdinates: ip_GridProps[GridID].selectedRange[r][0],
            EndCellOrdinates: ip_GridProps[GridID].selectedRange[r][1],
            RangeEl: elRange
        };

    }

    //Move the range along with the mouse
    var RangeStartRow = Ranges[0].StartCellOrdinates[0, 0];
    var RangeStartCol = Ranges[0].StartCellOrdinates[0, 1];
    var MouseCursorStartX = -1;
    var MouseCursorStartY = -1;
    var HoverCell = null;

    ip_GridProps[GridID].events.moveRange_mouseMove = ip_UnBindEvent('#' + GridID, 'mousemove', ip_GridProps[GridID].events.moveRange_mouseMove);
    $('#' + GridID).mousemove(ip_GridProps[GridID].events.moveRange_mouseMove = function (e) {

        

        if (MouseCursorStartX == -1) { MouseCursorStartX = e.pageX; }
        if (MouseCursorStartY == -1) { MouseCursorStartY = e.pageY; }

        var MouseMoveX = e.pageX - MouseCursorStartX;
        var MouseMoveY = e.pageY - MouseCursorStartY;

        for (var r = 0; r < Ranges.length; r++) {

            var RangeLeft = Ranges[r].StartX + MouseMoveX;
            var RangeTop = Ranges[r].StartY + MouseMoveY;

            if (borderPosition == 'left') { RangeLeft += 15; }
            if (borderPosition == 'right') { RangeLeft -= 15; }
            if (borderPosition == 'top') { RangeTop += 15; }
            if (borderPosition == 'bottom') { RangeTop -= 15; }
  
            $(Ranges[r].RangeEl).css('left', RangeLeft + 'px');
            $(Ranges[r].RangeEl).css('top', RangeTop + 'px');

        
        }

        //This code gives the user a preview for where the range will be placed
        if (HoverCell != ip_GridProps[GridID].hoverCell && !ip_GridProps[GridID].scrollAnimate) {

            clearTimeout(ip_GridProps[GridID].events.showRangeMove_highlightMovePosition);
            HoverCell = ip_SetHoverCell(GridID, '.ip_grid_cell_rangeselector_move', e, true); 

            ip_GridProps[GridID].events.showRangeMove_highlightMovePosition = setTimeout(function () {
                                
                var row = parseInt($(HoverCell).attr('row'));
                var col = parseInt($(HoverCell).attr('col'));

                var pasteRow = row - (ClickStartRow - RangeStartRow); 
                var pasteCol = col - (ClickStartCol - RangeStartCol);

                ip_HighlightMovePosition(GridID, pasteRow, pasteCol);
                

            }, 100);

        }
        else if (ip_GridProps[GridID].scrollAnimate && ip_GridProps[GridID].events.showRangeMove_highlightMovePosition != null) {

            //This code is active when we are scrolling AND moving the range at the same time - clear out the preview 
            clearTimeout(ip_GridProps[GridID].events.showRangeMove_highlightMovePosition);
            ip_GridProps[GridID].events.showRangeMove_highlightMovePosition = null;            
            $('#' + GridID).ip_RemoveRangeHighlight({ highlightType: 'ip_grid_cell_rangeHighlight_hoverCell' });            

        }


    });

    //Do range move
    ip_GridProps[GridID].events.moveRange_mouseUp = ip_UnBindEvent(document, 'mouseup', ip_GridProps[GridID].events.moveRange_mouseUp);
    $(document).mouseup(ip_GridProps[GridID].events.moveRange_mouseUp = function (e) {
        
        clearTimeout(ip_GridProps[GridID].events.showRangeMove_highlightMovePosition);

        ip_GridProps[GridID].events.moveRange_mouseMove = ip_UnBindEvent('#' + GridID, 'mousemove', ip_GridProps[GridID].events.moveRange_mouseMove);
        ip_GridProps[GridID].events.moveRange_mouseUp = ip_UnBindEvent(document, 'mouseup', ip_GridProps[GridID].events.moveRange_mouseUp);

        var ClickEndCell = ip_SetHoverCell(GridID, '.ip_grid_cell_rangeselector_move', e, true);
        var ClickEndRow = parseInt($(ClickEndCell).attr('row'));
        var ClickEndCol = parseInt($(ClickEndCell).attr('col'));

        

        $('#' + GridID).ip_RemoveRangeHighlight({ highlightType:'ip_grid_cell_rangeHighlight_hoverCell' });

        if (ClickStartCell != ClickEndCell) {

            //$('#' + GridID).ip_SelectCell({ cell: ClickEndCell });

            //This line of code validates selected rows/columns
            if (RangeStartRow < 0) { RangeStartRow = 0; }
            if (RangeStartCol < 0) { RangeStartCol = 0; }
            
            if ((RangeStartRow <= ip_GridProps[GridID].scrollY && borderPosition == 'top') || ClickStartRow < 0) { ClickStartRow = RangeStartRow; } //ip_GridProps[GridID].scrollY; }
            if ((RangeStartCol <= ip_GridProps[GridID].scrollX && borderPosition == 'left') || ClickStartCol < 0) { ClickStartCol = RangeStartCol; } //ip_GridProps[GridID].scrollY; }

            var pasteRow = ClickEndRow - (ClickStartRow - RangeStartRow);
            var pasteCol = ClickEndCol - (ClickStartCol - RangeStartCol);            

            $('#' + GridID).ip_Paste({ changeFormula: true, row: pasteRow, col: pasteCol });

        }
        else {

            //Reset the move
            $('#' + GridID).ip_ClearCopy();

            for (var r = 0; r < Ranges.length; r++) {

                //NBNB Relies on the range sequence to be in the array sequence of ip_GridProps[GridID].selectedRange
                $('#' + GridID).ip_SelectRange({ startCellOrdinates: Ranges[r].StartCellOrdinates, endCellOrdinates: Ranges[r].EndCellOrdinates, multiselect: (r == 0 ? false : true) })

            }

        }

        ip_GridProps[GridID].resizing = false;

    });

    

    //Code to reset selected ranges
}

function ip_ValidateRangeOverlap(GridID, arrRange) {

    var arrResult = new Array();
    var tested = new Array();

    for (var i = 0; i < arrRange.length; i++) {

        for (j = 0; j < arrRange.length; j++) {

            if (i != j) {
                var startRow1 = arrRange[i].startRow;
                var endRow1 = arrRange[i].endRow;
                var startCol1 = arrRange[i].startCol;
                var endCol1 = arrRange[i].endCol;

                var startRow2 = arrRange[j].startRow;
                var endRow2 = arrRange[j].endRow;
                var startCol2 = arrRange[j].startCol;
                var endCol2 = arrRange[j].endCol;

                var testIndex = startRow2 + '' + startCol2;

                if (tested[testIndex] == null) {
                    if (startRow1 <= startRow2 && endRow1 >= startRow2 && startCol1 <= startCol2 && endCol1 >= startCol2) { tested[testIndex] = true; arrResult[arrResult.length] = arrRange[j]; }
                    else if (startRow1 <= endRow2 && endRow1 >= endRow2 && startCol1 <= startCol2 && endCol1 >= startCol2) { tested[testIndex] = true; arrResult[arrResult.length] = arrRange[j]; }
                    else if (startRow1 <= startRow2 && endRow1 >= startRow2 && startCol1 <= endCol2 && endCol1 >= endCol2) { tested[testIndex] = true; arrResult[arrResult.length] = arrRange[j]; }
                    else if (startRow1 <= endRow2 && endRow1 >= endRow2 && startCol1 <= endCol2 && endCol1 >= endCol2) { tested[testIndex] = true; arrResult[arrResult.length] = arrRange[j]; }

                    else if (startRow1 <= endRow2 && endRow1 >= startRow2 && startCol1 <= endCol2 && endCol1 >= endCol2) { tested[testIndex] = true; arrResult[arrResult.length] = arrRange[j]; }
                }
            }

        }

    }

    return arrResult;
}

function ip_HighlightMovePosition(GridID, row, col) {

    if (ip_GridProps[GridID].selectedRange.length > 0) {

        var rowDiff = row - ip_GridProps[GridID].selectedRange[0][0][0];
        var colDiff = col - ip_GridProps[GridID].selectedRange[0][0][1];

        for (var r = 0; r < ip_GridProps[GridID].selectedRange.length; r++) {

            var range = ip_GridProps[GridID].selectedRange[r];
            var startRow = range[0][0] + rowDiff;
            var startCol = range[0][1] + colDiff;
            var endRow = range[1][0] + rowDiff;
            var endCol = range[1][1] + colDiff;
                        
            $('#' + GridID).ip_RangeHighlight({ fadeIn:true, multiselect: (r == 0 ? false : true), highlightType: 'ip_grid_cell_rangeHighlight_hoverCell', range: { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol } });
        }

    }
}

function ip_GetRangeData(GridID, byRef, startRow, startCol, endRow, endCol, limitResultsToEditedRowCol, processMerges, makeMergedCellValueSame, includeRowData, includeGroupValue) {
    //This method returns just cell data found within a specific rage
    //If byRef - the data is actually returned by reference so that it can be directly manipulated from the return object
    //If limitResultsToEditedRowCol - the results returned back are limited to the maximum edited row & column,  NOTE - still need to remove merge data from results if this is the case
    //If processMerges - returns a summary of merge information for the specified range
    //If makeMergedCellValueSame - gives a cell the same value as the cell it was merged into (insead of null), use full for some algorythms like sorting


    var MaxRow = -1;
    var MaxCol = -1;
    var ProcessedMerges = {};
    var RangeData = { rowData: new Array(), mergeData: new Array(), mergeData: new Array(), mergeDataContained: new Array(), startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol, rowDataLoading: false };
    var rIndex = 0;
    var GroupRow = null;
    var GroupCounter = 0;

    if (processMerges == null) { processMerges = true; }
    if (limitResultsToEditedRowCol == null) { limitResultsToEditedRowCol = false; }
    if (makeMergedCellValueSame == null) { makeMergedCellValueSame = false; } 

    for (var r = startRow; r <= endRow; r++) {

        var cIndex = 0;

        RangeData.rowData[rIndex] = ip_CloneRow(GridID, r); // ip_rowObject(null, null, null, null);

        if (ip_GridProps[GridID].rowData[r].loading) { RangeData.rowDataLoading = true; } //Is range fully loaded from server

        if (includeGroupValue) {
            
            if(GroupCounter == 0) { GroupRow = r; }

            if (ip_GridProps[GridID].rowData[r].groupCount > GroupCounter) { GroupCounter = ip_GridProps[GridID].rowData[r].groupCount; }
            else if (GroupCounter > 0) { GroupCounter--; }          
            
        }

        for (var c = startCol; c <= endCol; c++) {
            
            var merge = ip_GridProps[GridID].rowData[r].cells[c].merge;
            var row = r;
            var col = c;

            if (merge != null && makeMergedCellValueSame) { row = merge.mergedWithRow; col = merge.mergedWithCol; }

            //Calculate the maximum edited row within range
            if (!ip_IsCellEmpty(GridID, r, c)) { if (MaxRow < r) { MaxRow = r }  if (MaxCol < c) { MaxCol = c } }
            
            //Add data to results
            if (byRef) { RangeData.rowData[rIndex].cells[cIndex] = ip_GridProps[GridID].rowData[row].cells[col]; }
            else { RangeData.rowData[rIndex].cells[cIndex] = ip_CloneCell(GridID, row, col, r, c); }

            
            //Set the group value for the cell            
            if (GroupRow != null) { RangeData.rowData[rIndex].cells[cIndex].groupValue = ip_GridProps[GridID].rowData[GroupRow].cells[c].value;  }

            //Calculate merge information            
            if (merge != null && processMerges) {
                
                var MergeKey = merge.mergedWithRow + '-' + merge.mergedWithCol;
                if (ProcessedMerges[MergeKey] == null) {

                    ProcessedMerges[MergeKey] = true;

                    //Fetch the origonal instance of the merge
                    merge = ip_GridProps[GridID].rowData[merge.mergedWithRow].cells[merge.mergedWithCol].merge;

                    //Fetch merge
                    RangeData.mergeData[RangeData.mergeData.length] = ip_mergeObject();
                    RangeData.mergeData[RangeData.mergeData.length - 1].mergedWithRow = merge.mergedWithRow;
                    RangeData.mergeData[RangeData.mergeData.length - 1].mergedWithCol = merge.mergedWithCol;
                    RangeData.mergeData[RangeData.mergeData.length - 1].rowSpan = merge.rowSpan;
                    RangeData.mergeData[RangeData.mergeData.length - 1].colSpan = merge.colSpan;
                    RangeData.mergeData[RangeData.mergeData.length - 1].containsOverlap = ip_DoesRangeOverlapMerge(GridID, merge, startRow, startCol, endRow, endCol);
                    

                    //Calculate the merges confined range - this data is usefull if we want to set a range within the scope of a range
                    var mergeStartRow = merge.mergedWithRow;
                    var mergeStartCol = merge.mergedWithCol;
                    var mergeEndRow = mergeStartRow + merge.rowSpan - 1;
                    var mergeEndCol = mergeStartCol + merge.colSpan - 1;

                    if (mergeStartRow < startRow) { mergeStartRow = startRow; }
                    if (mergeStartCol < startCol) { mergeStartCol = startCol; }
                    if (mergeEndRow > endRow) { mergeEndRow = endRow; }
                    if (mergeEndCol > endCol) { mergeEndCol = endCol; }
                    
                    RangeData.mergeDataContained[RangeData.mergeDataContained.length] = ip_mergeObject();
                    RangeData.mergeDataContained[RangeData.mergeDataContained.length - 1].mergedWithRow = mergeStartRow;
                    RangeData.mergeDataContained[RangeData.mergeDataContained.length - 1].mergedWithCol = mergeStartCol;
                    RangeData.mergeDataContained[RangeData.mergeDataContained.length - 1].rowSpan = mergeEndRow - mergeStartRow + 1;
                    RangeData.mergeDataContained[RangeData.mergeDataContained.length - 1].colSpan = mergeEndCol - mergeStartCol + 1;
                    
                }

            }

            cIndex++;
         
        }

        rIndex++;
    }

    if (limitResultsToEditedRowCol) {
        
        if (MaxRow > -1) {

            RangeData.rowData.splice(MaxRow + 1, RangeData.rowData.length);

            for (var r = 0; r < RangeData.rowData.length; r++) { RangeData.rowData[r].cells.splice(MaxCol + 1, RangeData.rowData[r].cells.length); }
            for (var m = 0; m < RangeData.mergeData.length; m++) { if (RangeData.mergeData[m].mergedWithRow > MaxRow || RangeData.mergeData[m].mergedWithCol > MaxCol) { RangeData.mergeData.splice(m, 1); } }
            for (var m = 0; m < RangeData.mergeDataContained.length; m++) { if (RangeData.mergeDataContained[m].mergedWithRow > MaxRow || RangeData.mergeDataContained[m].mergedWithCol > MaxCol) { RangeData.mergeData.splice(m, 1); } }

            //Make sure that if the end cell is a merge we include it fully in the range
            if (ip_GridProps[GridID].rowData[MaxRow].cells[MaxCol].merge != null) {

                var merge = ip_GridProps[GridID].rowData[MaxRow].cells[MaxCol].merge;
                var MaxRow = merge.mergedWithRow + ip_GridProps[GridID].rowData[merge.mergedWithRow].cells[merge.mergedWithCol].merge.rowSpan - 1;
                var MaxCol = merge.mergedWithCol + ip_GridProps[GridID].rowData[merge.mergedWithRow].cells[merge.mergedWithCol].merge.colSpan - 1;
            }

        }
        else { 

            RangeData.rowData = new Array(), 
            RangeData.mergeData = new Array(), 
            RangeData.mergeData = new Array(), 
            RangeData.mergeDataContained = new Array()
            
        }

        RangeData.endRow = MaxRow;
        RangeData.endCol = MaxCol;


    }

    return RangeData;

}

function ip_DragRange(GridID, options) {

    var options = $.extend({

        range: null, // { startRow:null, startCol: null, endRow: null, endCol: null },
        dragToRange: null, // { startRow:null, startCol: null, endRow: null, endCol: null },        

    }, options);

    var TransactionID = ip_GenerateTransactionID();
    var Error = '';
    var Effected = { rowData: [], dragFromRange:ip_rangeObject(),  dragToRange: ip_rangeObject() }
    

    if (options.range != null && options.dragToRange != null) {


        //Calulcate the pattern to use when dragging. This is different for different datatypes
        var patternCols = {};

        for (var r = options.range.endRow; r >= options.range.startRow; r--) {

            for (var c = options.dragToRange.startCol; c <= options.dragToRange.endCol; c++) {

                var dataType = ip_CellDataType(GridID, r, c, true);
                var dataTypePrev = ip_CellDataType(GridID, r - 1, c, true);

                //Initialize the pattern object, each column being dragged has a patern object
                if (patternCols[c] == null) {

                    patternCols[c] = {

                        startPatternRow: options.range.startRow,
                        endPatternRow: options.dragToRange.endRow,
                        repeatPattern: Math.ceil((options.dragToRange.endRow - options.range.endRow + 1) / (options.range.endRow - options.range.startRow + 1)),
                        rangeIncrement: (options.range.endRow - options.range.startRow + 1),
                        dataTypeIncrement: {},
                        values: new Array(options.range.endRow - options.range.startRow + 1),
                        index: options.range.endRow - options.range.startRow,

                    }
                }

                if (patternCols[c].dataTypeIncrement[dataType.dataType.dataType] == null) { patternCols[c].dataTypeIncrement[dataType.dataType.dataType] = 0; }
                patternCols[c].dataTypeIncrement[dataType.dataType.dataType]++;


                //Record the actual pattern datatypes
                if (dataType.dataType.dataType == 'number') {


                    patternCols[c].values[patternCols[c].index] = {
                        //patternCols[c].rangeIncrement > 1 && 
                        increment: (r > options.range.startRow && dataTypePrev.dataType.dataType == 'number' ? (dataType.value - dataTypePrev.value) : (patternCols[c].index + 1 < patternCols[c].values.length ? patternCols[c].values[patternCols[c].index + 1].increment : 0)),
                        value: dataType.value,
                        dataType: dataType.dataType,
                        dataTypeO: ip_GridProps[GridID].rowData[r].cells[c].dataType

                    }

                }
                else if (dataType.dataType.dataType == 'text') {

                    patternCols[c].values[patternCols[c].index] = {

                        increment: 0,
                        value: dataType.value,
                        dataType: dataType.dataType,
                        dataTypeO: ip_GridProps[GridID].rowData[r].cells[c].dataType

                    }

                }
                else {

                    patternCols[c].values[patternCols[c].index] = {

                        increment: 0,
                        value: dataType.value,
                        dataType: dataType.dataType,
                        dataTypeO: ip_GridProps[GridID].rowData[r].cells[c].dataType

                    }

                }

                //clean up last
                patternCols[c].index--;
            }
        }



        var dragRange = { startRow: options.range.endRow + 1, startCol: options.dragToRange.startCol, endRow: options.dragToRange.endRow, endCol: options.dragToRange.endCol }
        var CellUndoData = ip_AddUndo(GridID, 'ip_DragRange', TransactionID, 'CellData', dragRange);


        for (var c = options.dragToRange.startCol; c <= options.dragToRange.endCol; c++) {

            var rIndex = 0;

            for (var p = 1; p <= patternCols[c].repeatPattern; p++) {

                for (var pr = 0; pr < patternCols[c].values.length; pr++) {

                    var r = patternCols[c].startPatternRow + (patternCols[c].rangeIncrement * p) + pr;

                    if (r <= patternCols[c].endPatternRow) {

                        ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, r, c));

                        var dataType = patternCols[c].values[pr].dataType.dataType;
                        var baseIncrement = patternCols[c].dataTypeIncrement[dataType];
                        var increment = patternCols[c].values[pr].increment;
                        var newValue = patternCols[c].values[pr].value;

                        if (dataType == 'number') { newValue = (newValue + ((baseIncrement * increment) * p)); }
                                                
                        ip_GridProps[GridID].rowData[r].cells[c].dataType = patternCols[c].values[pr].dataType;
                        ip_SetValue(GridID, r, c, (newValue == null ? null : newValue));

                        if (Effected.rowData[rIndex] == null) { Effected.rowData[rIndex] = { cells: [], row: r } }
                        Effected.rowData[rIndex].cells[Effected.rowData[rIndex].cells.length] = ip_CloneCell(GridID, r, c);
                        

                    }
                    rIndex++;
                }

            }




        }

        //Raise cell change event
        Effected.dragFromRange = options.range;
        Effected.dragToRange = options.dragToRange;
        Effected.dragToRange.startRow = options.range.endRow + 1;
        
        if (Effected.rowData.length > 0) { ip_RaiseEvent(GridID, 'ip_DragRange', TransactionID, { DragRange: { Inputs: options, Effected: Effected } }); }

        ip_ReRenderRows(GridID);

    }
    else if (Error == '') { Error = 'Please specify a range'; }

    if (Error != '') { ip_RaiseEvent(GridID, 'warning', TransactionID, Error); }

}

function ip_ValidateRangeObjects(GridID, ranges, hRow, hCol) {

    var rangesList = [];

    if (ranges.length > 0) {

        for (var i = 0; i < ranges.length; i++) { rangesList[rangesList.length] = ip_ValidateRangeObject(GridID, ranges[i], hRow, hCol); }

    }

    return rangesList;

}

function ip_ValidateRangeObject(GridID, range, hRow, hCol) {
    //Universal range validation obejct

    if (range != null) {

        var minRow = (hRow ? -1 : 0);
        var minCol = (hCol ? -1 : 0);

        if (range.length > 0) {
                        
            //Convert array range to regular range
            range = {
                startRow: range[0][0],
                startCol: range[0][1],
                endRow: range[1][0],
                endCol: range[1][1],
            }
        }
        
        if (range.startRow == null || range.startRow < minRow) { range.startRow = minRow; }
        if (range.startCol == null || range.startCol < minCol) { range.startCol = minCol; }
        if (range.endRow == null || range.endRow >= ip_GridProps[GridID].rows) { range.endRow = ip_GridProps[GridID].rows - 1; }
        if (range.endCol == null || range.endCol >= ip_GridProps[GridID].cols) { range.endCol = ip_GridProps[GridID].cols - 1; }
        
    }

    return range;
}

function ip_IsRangeOverLap(GridID, range1, range2) {
    //innerlap range1 falls inside range2
    //overlap range1 overlaps range2
    if (range1 != null && range2 != null) {
        var r = range2.startRow;
        var c = range2.startCol;

        //Check if exact
        if (range1.startRow == range2.startRow &&
            range1.endRow == range2.endRow &&
            range1.startCol == range2.startCol &&
            range1.endCol == range2.endCol) { return 'exact'; }

        //Check for inside / outside
        if (range1.startRow >= range2.startRow &&
            range1.startRow <= range2.endRow && 
            range1.endRow >= range2.startRow &&
            range1.endRow <= range2.endRow &&

            range1.startCol >= range2.startCol &&
            range1.startCol <= range2.endCol && 
            range1.endCol >= range2.startCol &&
            range1.endCol <= range2.endCol) { return 'inside'; }

        if (range2.startRow >= range1.startRow &&
            range2.startRow <= range1.endRow &&
            range2.endRow >= range1.startRow &&
            range2.endRow <= range1.endRow &&

            range2.startCol >= range1.startCol &&
            range2.startCol <= range1.endCol &&
            range2.endCol >= range1.startCol &&
            range2.endCol <= range1.endCol) { return 'outside'; }

        //Check for overlap/underlap
        r = range2.startRow;
        c = range2.startCol;
        if (r >= range1.startRow && r <= range1.endRow && c >= range1.startCol && c <= range1.endCol) { return 'overlap'; }

        r = range2.startRow;
        c = range2.endCol;
        if (r >= range1.startRow && r <= range1.endRow && c >= range1.startCol && c <= range1.endCol) { return 'overlap'; }

        r = range2.endRow;
        c = range2.startCol;
        if (r >= range1.startRow && r <= range1.endRow && c >= range1.startCol && c <= range1.endCol) { return 'overlap'; }

        r = range2.endRow;
        c = range2.endCol;
        if (r >= range1.startRow && r <= range1.endRow && c >= range1.startCol && c <= range1.endCol) { return 'overlap'; }

        //entire row selected
        if (r >= range1.startRow && r <= range1.endRow && range2.startCol <= range1.startCol && range2.endCol >= range1.endCol) { return 'overlap'; }

        //check inverse
        r = range1.startRow;
        c = range1.startCol;
        if (r >= range2.startRow && r <= range2.endRow && c >= range2.startCol && c <= range2.endCol) { return 'innerlap'; }

        r = range1.startRow;
        c = range1.endCol;
        if (r >= range2.startRow && r <= range2.endRow && c >= range2.startCol && c <= range2.endCol) { return 'innerlap'; }

        r = range1.endRow;
        c = range1.startCol;
        if (r >= range2.startRow && r <= range2.endRow && c >= range2.startCol && c <= range2.endCol) { return 'innerlap'; }

        r = range1.endRow;
        c = range1.endCol;
        if (r >= range2.startRow && r <= range2.endRow && c >= range2.startCol && c <= range2.endCol) { return 'innerlap'; }




    }
    return false;

}

function ip_TrimRange(GridID, range1, range2) {

    var result = ip_rangeObject();

    //this method trims range2 fit into to range1
    if (range2.startRow > range1.endRow) { return null; }
    if (range2.startCol > range1.endCol) { return null; }
    if (range2.endRow < range1.startRow) { return null; }
    if (range2.endCol < range1.startCol) { return null; }

    if (range2.startRow < range1.startRow) { result.startRow = range1.startRow; } else { result.startRow = range2.startRow }
    if (range2.startCol < range1.startCol) { result.startCol = range1.startCol; } else { result.startCol = range2.startCol }
    if (range2.endRow > range1.endRow) { result.endRow = range1.endRow; } else { result.endRow = range2.endRow }
    if (range2.endCol > range1.endCol) { result.endCol = range1.endCol; } else { result.endCol = range2.endCol }

    return result;
}


//----- RANGE HiGHLIGHT ------------------------------------------------------------------------------------------------------------------------------------

function ip_GetRangeHighlightIndex(GridID, RangeHighlight) {

    var startRow = parseInt($(RangeHighlight).attr('startRow'));
    var startCol = parseInt($(RangeHighlight).attr('startCol'));
    var endRow = parseInt($(RangeHighlight).attr('endRow'));
    var endCol = parseInt($(RangeHighlight).attr('endCol'));

    for (var i = 0; i < ip_GridProps[GridID].highlightRange.length; i++) {

        if (ip_GridProps[GridID].highlightRange[i][0][0] == startRow
            && ip_GridProps[GridID].highlightRange[i][0][1] == startCol
            && ip_GridProps[GridID].highlightRange[i][1][0] == endRow
            && ip_GridProps[GridID].highlightRange[i][1][1] == endCol) {

            return i;
        }
    }

    return -1;
}

function ip_setRangeHighlightID(GridID, RangeHighlight, startCellOrdinates, endCellOrdinates) {

    var startRow = startCellOrdinates[0, 0];
    var startCol = startCellOrdinates[0, 1];
    var endRow = endCellOrdinates[0, 0];
    var endCol = endCellOrdinates[0, 1];
    var RangeID = GridID + '_RangeHighlight_' + startRow + '_' + startCol + '_' + endRow + '_' + endCol;

    $(RangeHighlight).attr('id', RangeID)
    $(RangeHighlight).attr('startRow', startRow)
    $(RangeHighlight).attr('startCol', startCol)
    $(RangeHighlight).attr('endRow', endRow)
    $(RangeHighlight).attr('endCol', endCol)

    return RangeID;
}


//----- SCROLL ------------------------------------------------------------------------------------------------------------------------------------

function ip_InitScrollX(GridID, ScrollInterval, ShowX, ShowY, PlacementX, PlacementY) {

    var dragPos = 0;
    var ScrollX = 0;
    var ScrollCount = 0;

    //MyGrid_q2_container
    var ScrollBarY_width = (!ShowY ? 0 : parseInt($('#' + GridID + '_q4_scrollbar_container_y').width()));
    var ScrollBarX_Height = (!ShowX ? 0 : $('#' + GridID + '_q4_scrollbar_container_x').height());
    var GridTableHeight = parseInt($('#' + GridID + '_table').height());
    var FourthQuadTableWidth = ip_TableWidth(GridID + '_q4_table');
    var FourthQuadTableHeight = parseInt($('#' + GridID + '_q4_table').outerHeight());
    var FirstQuadTableWidth = ip_TableWidth(GridID + '_q1_table');
    var FirstQuadTableHeight = parseInt($('#' + GridID + '_q1_table').outerHeight());
    var GridWidth = parseInt(ip_GridProps[GridID].dimensions.gridWidth);
    var GridHeight = parseInt(ip_GridProps[GridID].dimensions.gridHeight);
    var ScrollContainerWidth = GridWidth - FirstQuadTableWidth - ScrollBarY_width;
    var ScrollBarTop = (GridHeight - FirstQuadTableHeight - ScrollBarX_Height);

    //Set the scrollbar width to the size of the grid
    if (PlacementY != 'container' && ScrollContainerWidth > FourthQuadTableWidth) { ScrollContainerWidth = FourthQuadTableWidth; }

    //Set containers to correct size
    $('#' + GridID + '_q2_container').css('width', ScrollContainerWidth + 'px');
    $('#' + GridID + '_q4_container').css('width', ScrollContainerWidth + 'px');

    //Set quadraht widths, needed bacause the absolute positioning of veritcal scroll bar 
    $('#' + GridID + '_q2').css('width', ScrollContainerWidth + ScrollBarY_width + 'px');
    $('#' + GridID + '_q4').css('width', ScrollContainerWidth + ScrollBarY_width + 'px');

    //Position scrollbar
    if (FourthQuadTableHeight >= ScrollBarTop || PlacementX == 'container') {
        $('#' + GridID + '_q3_container').css('height', ScrollBarTop + 'px');
        $('#' + GridID + '_q4_container').css('height', ScrollBarTop + 'px');
    }

    //Initialize scrollbars
    if (ShowX) {

        $('#' + GridID + '_q4_scrollbar_container_x').css('display', '');
        $('#' + GridID + '_q4_scrollbar_container_x').css('width', ScrollContainerWidth + 'px');

        var ScrollFit = 0;
        var ScrollGadgetSize = 0;

        while (ScrollGadgetSize < 10) {
            ScrollFit = ((FourthQuadTableWidth - ScrollContainerWidth) / ScrollInterval);
            ScrollGadgetSize = parseInt(ScrollContainerWidth - ScrollFit);
            if (ScrollGadgetSize < 10) { ScrollInterval++; }
        }

        //if no scrolling possible, set scrollbar_gadget to 100%
        if (ScrollGadgetSize > FourthQuadTableWidth) { ScrollGadgetSize = ScrollContainerWidth; }

        $("#" + GridID + "_q4_scrollbar_gadget_x").css('width', ScrollGadgetSize + 'px');


        //actual scroll operation
        $("#" + GridID + "_q4_scrollbar_gadget_x").draggable({
            axis: "x",
            containment: "parent",
            drag: function (event, ui) {

                if (ui.position.left > dragPos || ui.position.left < dragPos) { ScrollX = ui.position.left * ScrollInterval; }

                ip_GridProps[GridID].scrollbarX = ScrollX;
                $('#' + GridID + '_q4_container').scrollLeft(ScrollX);
                $('#' + GridID + '_q2_container').scrollLeft(ScrollX);

                dragPos = ui.position.left;

                $('#ScrollCount').html(dragPos);
            }

        }, 50);

    }
    else {
        //hide scroll bar
        $('#' + GridID + '_q4_scrollbar_container_x').css('display', 'none');
    }

}

function ip_InitScrollX_Large(GridID, ScrollInterval, ShowX, ShowY, PlacementX, PlacementY) {

    //MyGrid_q2_container
    var MinScrollGadgetSize = 50;
    var ScrollBarY_width = (!ShowY ? 0 : parseInt($('#' + GridID + '_q4_scrollbar_container_y').width()));
    var ScrollBarX_Height = (!ShowX ? 0 : $('#' + GridID + '_q4_scrollbar_container_x').height());
    var GridTableHeight = parseInt($('#' + GridID + '_table').height());
    var FourthQuadTableWidth = ip_TableWidth(GridID + '_q4_table');
    var FourthQuadTableHeight = parseInt($('#' + GridID + '_q4_table').outerHeight());
    var FirstQuadTableWidth = ip_TableWidth(GridID + '_q1_table');
    var FirstQuadTableHeight = parseInt($('#' + GridID + '_q1_table').outerHeight());
    var GridWidth = parseInt(ip_GridProps[GridID].dimensions.gridWidth);
    var GridHeight = parseInt(ip_GridProps[GridID].dimensions.gridHeight);
    var ScrollContainerWidth = GridWidth - FirstQuadTableWidth - ScrollBarY_width;
    var ScrollBarTop = (GridHeight - FirstQuadTableHeight - ScrollBarX_Height);

    if (ScrollContainerWidth < 0) { ScrollContainerWidth = GridWidth; ip_RaiseEvent(GridID, 'warning', arguments.callee.caller, 'The grid may not be able to show properly because your frozen column[s] are wider than the browser allows for!'); }
    if (ip_GridProps[GridID].loadedCols < ip_GridProps[GridID].frozenCols) { ScrollContainerWidth = ScrollBarY_width; ip_RaiseEvent(GridID, 'warning', arguments.callee.caller, 'The grid may not be able to show properly because your frozen columns[s] dont have enough space to fit!'); }

    ip_GridProps[GridID].dimensions.scrollWidth = ScrollContainerWidth;
    
    //Set the scrollbar width to the size of the grid
    if (PlacementY != 'container' && ScrollContainerWidth > FourthQuadTableWidth) { ScrollContainerWidth = FourthQuadTableWidth; }

    //Set containers to correct size
    $('#' + GridID + '_q2_container').css('width', ScrollContainerWidth + 'px');
    $('#' + GridID + '_q4_container').css('width', ScrollContainerWidth + 'px');

    //Set quadraht widths, needed bacause the absolute positioning of veritcal scroll bar 
    $('#' + GridID + '_q2').css('width', ScrollContainerWidth + ScrollBarY_width + 'px');
    $('#' + GridID + '_q4').css('width', ScrollContainerWidth + ScrollBarY_width + 'px');

    //Position scrollbar
    if (FourthQuadTableHeight >= ScrollBarTop || PlacementX == 'container') {
        $('#' + GridID + '_q3_container').css('height', ScrollBarTop + 'px');
        $('#' + GridID + '_q4_container').css('height', ScrollBarTop + 'px');
    }

    //Initialize scrollbars
    if (ShowX) {

        //if (ScrollContainerWidth > FourthQuadTableWidth) { ScrollContainerWidth = FourthQuadTableWidth; }
        if (ScrollContainerWidth < MinScrollGadgetSize) { MinScrollGadgetSize = ScrollContainerWidth - 1; }

        $('#' + GridID + '_q4_scrollbar_container_x').css('display', '');
        $('#' + GridID + '_q4_scrollbar_container_x').css('width', ScrollContainerWidth + 'px');
     
        var ScrollFit = 0;
        var ScrollGadgetSize = 0;

        while (ScrollGadgetSize < MinScrollGadgetSize) {
            //ScrollFit = ((ip_GridProps[GridID].cols - (ip_GridProps[GridID].loadedCols - (ip_GridProps[GridID].loadedCols / 4))) / ScrollInterval); 
            ScrollFit = ((ip_GridProps[GridID].cols - (ip_GridProps[GridID].loadedCols - (ip_GridProps[GridID].loadedCols - ip_GridProps[GridID].frozenCols - 1))) / ScrollInterval); 
            ScrollGadgetSize = parseInt(ScrollContainerWidth - ScrollFit);  //parseInt(ScrollContainerHeight - ScrollFit);            
            if (ScrollGadgetSize < MinScrollGadgetSize) { ScrollInterval = ScrollInterval + 0.01; }
        }

        ip_GridProps[GridID].scrollXInterval = ScrollInterval;

        if (ip_GridProps[GridID].cols < ip_GridProps[GridID].loadedCols) { ScrollGadgetSize = ScrollContainerWidth; }

        $("#" + GridID + "_q4_scrollbar_gadget_x").css('width', ScrollGadgetSize + 'px');

        var IPScrollIncrementKey = 'bef7bea9-3157-4434-8cbb-22320d0d294e';
        var scrollLoading = false;
        var startScroll = true;
        var dragPos = 0;
        var ScrollX = ip_GridProps[GridID].frozenCols;
        var ScrollCount = 0;        
        var ScrollDelay = parseInt(ip_GridProps[GridID].cols / 20);
        var scrollDiff = 0;
        var scrollDirection = 0;

        if (ip_GridProps[GridID].scrollX == -1) { ip_GridProps[GridID].scrollX = ip_GridProps[GridID].frozenCols; }

        ip_GridProps[GridID].events.ip_InitScrollX_Large_MouseDown = ip_UnBindEvent('#' + GridID + '_q4_scrollbar_gadget_x', 'mousedown', ip_GridProps[GridID].events.ip_InitScrollX_Large_MouseDown);
        $('#' + GridID + '_q4_scrollbar_gadget_x').mousedown(ip_GridProps[GridID].events.ip_InitScrollX_Large_MouseDown = function (event) { event.stopPropagation(); });

        //actual scroll operation
        $("#" + GridID + "_q4_scrollbar_gadget_x").draggable({
            axis: "x",
            containment: "parent",
            start: function (event, ui) {

                $('#' + GridID + '_q4_scrollbar_gadget_x').addClass('ip_grid_scrollBar_gadget_scrolling');

                startScroll = true;

                //Scroll on a new thread, if we have more than 40 rows as the browser cannot handle the realtime scroll
                if (ip_GridProps[GridID].loadedCols >= loadedColsThreshold.min) {

                    var timeout = null;

                    ip_GridProps[GridID].events.ip_InitScrollX_Large_MouseMove = ip_UnBindEvent(document, 'mousemove', ip_GridProps[GridID].events.ip_InitScrollX_Large_MouseMove);                   
                    $(document).on('mousemove', ip_GridProps[GridID].events.ip_InitScrollX_Large_MouseMove = function () {

                        if (timeout !== null) { clearTimeout(timeout); }

                        if (scrollDiff > 5) {

                            timeout = setTimeout(function () {

                                clearTimeout(timeout);

                                if (!ip_GridProps[GridID].scrolling) {

                                    startScroll = true;
                                    ip_ScrollToX(GridID, ScrollX, false, 'all', 'none');


                                }

                            }, ScrollDelay);
                        }

                    });


                }
            },
            drag: function (event, ui) {

                if (ui.position.left > dragPos || ui.position.left < dragPos) { ScrollX = parseInt((ui.position.left * ScrollInterval) + ip_GridProps[GridID].frozenCols); }

                //Make sure we dont scroll beyond our grid range
                if (ScrollX >= ip_GridProps[GridID].cols) { ScrollX = ip_GridProps[GridID].cols - 1; }
                else if (ScrollX < ip_GridProps[GridID].frozenCols) { ScrollX = ip_GridProps[GridID].frozenCols; }

                scrollDirection = 1;
                scrollDiff = ScrollX - ip_GridProps[GridID].scrollX;
                if (scrollDiff < 0) { scrollDiff = ip_GridProps[GridID].scrollX - ScrollX; scrollDirection = -1; }

                if (!startScroll && ip_GridProps[GridID].loadedCols <= loadedColsThreshold.max) {

                    if (scrollDiff <= 5 || ip_GridProps[GridID].loadedCols <= loadedColsThreshold.min) { //loadedRowsThreshold

                        setTimeout(function () {
                            if (!ip_GridProps[GridID].scrolling) {
    
                                ip_ScrollToX(GridID, ScrollX, false, 'all', 'none');

                            }
                        }, 10);

                    }
                    else {

                        //Do a small scroll and then allow the threaded on to kick in for the large scroll
                        setTimeout(function () {
                            if (!ip_GridProps[GridID].scrolling) {
                             
                                ip_ScrollToX(GridID, ip_GridProps[GridID].scrollX + (scrollDirection * 5), false, 'none', 'none');

                            }
                        }, 5);
                        scrollLoading = true;
                    }
                }
                else { scrollLoading = true; }

                dragPos = ui.position.left;
                startScroll = false;

            },
            stop: function (event, ui) {

                if (ui.position.left > dragPos || ui.position.left < dragPos) { ScrollX = parseInt((ui.position.left * ScrollInterval) + ip_GridProps[GridID].frozenCols); }

                if (ScrollX >= ip_GridProps[GridID].cols) { ScrollX = ip_GridProps[GridID].cols - 1; }

                if (ip_GridProps[GridID].scrollX != ScrollX) {
                    ip_ScrollToX(GridID, ScrollX, true, 'all', 'none');
                }

                ip_GridProps[GridID].events.ip_InitScrollX_Large_MouseMove = ip_UnBindEvent(document, 'mousemove', ip_GridProps[GridID].events.ip_InitScrollX_Large_MouseMove);

                dragPos = ui.position.left;

                $('#' + GridID + '_q4_scrollbar_gadget_x').removeClass('ip_grid_scrollBar_gadget_scrolling');
            }

        });


    }
    else {
        //hide scroll bar
        $('#' + GridID + '_q4_scrollbar_container_x').css('display', 'none');
    }

    return true;
}

function ip_InitScrollY_Large(GridID, ScrollInterval, Show, Placement) {

    var GridLeft = $('#' + GridID).position().left;
    //var GridContainerTop = $('#' + GridID).parent().position().top;
    var GridTop = $('#' + GridID + '_table').position().top;
    var ShowRowIndecator = false;
    var MinScrollGadgetSize = 35; // 35;
    var ScrollBarX_height = parseInt($('#' + GridID + '_q4_scrollbar_container_x').height());
    var ScrollBarX_width = parseInt($('#' + GridID + '_q4_scrollbar_container_y').width());
    var GridTableHeight = parseInt($('#' + GridID + '_table').height());
    var FourthQuadTableHeight = parseInt($('#' + GridID + '_q4_table').outerHeight());
    var FourthQuadContainerHeight = parseInt($('#' + GridID + '_q4_container').outerHeight());
    var FirstQuadTableWidth = parseInt($('#' + GridID + '_q1_table').outerWidth());
    var FirstQuadTableHeight = parseInt($('#' + GridID + '_q1_table').outerHeight());
    var GridWidth = parseInt(ip_GridProps[GridID].dimensions.gridWidth);
    var GridHeight = parseInt(ip_GridProps[GridID].dimensions.gridHeight);
    var ScrollContainerHeight = GridHeight - FirstQuadTableHeight - ScrollBarX_height;
    

    ip_GridProps[GridID].dimensions.scrollHeight = ScrollContainerHeight;

    if (ScrollContainerHeight < 0) { ScrollContainerHeight = GridHeight; ip_RaiseEvent(GridID, 'warning', arguments.callee.caller, 'The grid may not be able to show properly because your frozen rows[s] are taller than the browser allows for!');  }
    if (ip_GridProps[GridID].loadedRows < ip_GridProps[GridID].frozenRows) { ScrollContainerHeight = ScrollBarX_height; ip_RaiseEvent(GridID, 'warning', arguments.callee.caller, 'The grid may not be able to show properly because your frozen columns[s] dont have enough space to fit!'); }

    //Initialize scrollbars
    if (Show) {

        //if (ScrollContainerHeight > FourthQuadTableHeight) { ScrollContainerHeight = FourthQuadTableHeight; }
        if (ScrollContainerHeight < MinScrollGadgetSize) { MinScrollGadgetSize = ScrollContainerHeight - 1; }
        if (ip_GridProps[GridID].loadedRows <= ip_GridProps[GridID].rows) { ScrollGadgetSize = ScrollContainerHeight; }

        $('#' + GridID + '_q4_scrollbar_container_y').css('display', '');
        $('#' + GridID + '_q4_scrollbar_container_y').css('height', ScrollContainerHeight + 'px');

        var ScrollFit = MinScrollGadgetSize - 1;
        var ScrollGadgetSize = 0;

        while (ScrollGadgetSize < MinScrollGadgetSize) {

            ScrollFit = ((ip_GridProps[GridID].rows - (ip_GridProps[GridID].loadedRows - (ip_GridProps[GridID].loadedRows - ip_GridProps[GridID].frozenRows - 1))) / ScrollInterval); //FourthQuadTableHeight - ScrollContainerHeight
            ScrollGadgetSize = parseInt(ScrollContainerHeight - GridTop - ScrollFit);
            if (ScrollGadgetSize < MinScrollGadgetSize) { ScrollInterval = ScrollInterval + 0.01; }

        }

     

        ip_GridProps[GridID].scrollYInterval = ScrollInterval;

        if (ip_GridProps[GridID].rows < ip_GridProps[GridID].loadedRows) { ScrollGadgetSize = FourthQuadContainerHeight; }
        if (ip_GridProps[GridID].scrollY == -1) { ip_GridProps[GridID].scrollY = ip_GridProps[GridID].frozenRows; }

        //Position grid resizer
        if (ip_GridProps[GridID].showGridResizerX) {
            $("#" + GridID + "_gridResizer").css('width', ScrollBarX_width + 'px');
            $("#" + GridID + "_gridResizer").css('left', ((GridWidth + GridLeft) - ScrollBarX_width) + 'px'); //GridContainerLeft
            $("#" + GridID + "_gridResizer").css('top', (GridTop) + 'px'); //GridContainerTop
            $("#" + GridID + "_gridResizer").css('height', ip_GridProps[GridID].dimensions.columnSelectorHeight + 'px');
            $("#" + GridID + "_gridResizerLine").css('height', (GridHeight - ip_GridProps[GridID].dimensions.columnSelectorHeight) + 'px');
        }
        else { $("#" + GridID + "_gridResizer").hide(); }

        $("#" + GridID + "_q4_scrollbar_container_y").css('height', FourthQuadContainerHeight + 'px');
        $("#" + GridID + "_q4_scrollbar_gadget_y").css('height', ScrollGadgetSize + 'px');
        $('#' + GridID + '_q4_scrollbar_gadget_scrollIndecator_y').css('top', ((ScrollGadgetSize / 2) - ($('#' + GridID + '_q4_scrollbar_gadget_scrollIndecator_y').height() / 1.5)) + 'px');

        var ScrollGadgetTop = (ip_GridProps[GridID].scrollY - ip_GridProps[GridID].frozenRows) / ip_GridProps[GridID].scrollYInterval;
        var IPScrollIncrementKey = 'bef7bea9-3157-4434-8cbb-22320d0d294e';
        var scrollLoading = false;
        var dragPos = 0;
        var ScrollY = ip_GridProps[GridID].frozenRows;      
        var ScrollIndecatorVisible = false;

        ip_GridProps[GridID].events.ip_InitScrollY_Large_MouseDown = ip_UnBindEvent('#' + GridID + '_q4_scrollbar_gadget_y', 'mousedown', ip_GridProps[GridID].events.ip_InitScrollY_Large_MouseDown);
        $('#' + GridID + '_q4_scrollbar_gadget_y').mousedown(ip_GridProps[GridID].events.ip_InitScrollY_Large_MouseDown = function (e) {

            e.stopPropagation();
            e.preventDefault()
        });
        $('#' + GridID + '_q4_scrollbar_gadget_y').css('top', ScrollGadgetTop + 'px');

        //Dragging the scroll gadget        
        $("#" + GridID + "_q4_scrollbar_gadget_y").draggable({
            axis: "y",
            containment: "parent",
            start: function (event, ui) {
                
                if (ShowRowIndecator) { $('#' + GridID + '_q4_scrollbar_gadget_scrollIndecator_y').show() }
                $('#' + GridID + '_q4_scrollbar_gadget_y').addClass('ip_grid_scrollBar_gadget_scrolling');
                
            },
            drag: function (event, ui) {

                if (ui.position.top > dragPos || ui.position.top < dragPos) { ScrollY = parseInt((ui.position.top * ScrollInterval) + ip_GridProps[GridID].frozenRows); }
                if (ScrollY < 0) { ScrollY = 0; } else if (ScrollY >= ip_GridProps[GridID].rows) { ScrollY = ip_GridProps[GridID].rows - 1; }

                //ScrollY = parseInt(ScrollY);

                if (ip_GridProps[GridID].scrolling || scrollLoading) { return true; }

                //if (!scrollLoading && !ip_GridProps[GridID].scrolling) { // && !ip_GridProps[GridID].rowData[ScrollY].hide

                    scrollLoading = true;
                    clearTimeout(ip_GridProps[GridID].events.ip_InitScrollY_Large_scrollTimeout);

                    var scrollDirection = 1;
                    var scrollDiff = ScrollY - ip_GridProps[GridID].scrollY;                    
                    if (scrollDiff < 0) { scrollDiff = ip_GridProps[GridID].scrollY - ScrollY; scrollDirection = -1; }                                           
                    
                    ip_GridProps[GridID].events.ip_InitScrollY_Large_scrollTimeout = setTimeout(function () {

                        //if (!ip_GridProps[GridID].scrolling ) {

             
                        ip_ScrollToY(GridID, ScrollY, false, 'all', 'none');

                        //}                        

                        scrollLoading = false;                        

                    }, scrollDiff < 15 ? 15 : 50); //For smaller scrolls we have faster feedback
                          
                //}

                //Show scroll indecator if turned on or rows are hidden
                if (ShowRowIndecator) { $('#' + GridID + '_q4_scrollbar_gadget_scrollIndecator_y').text(ScrollY + (ip_GridProps[GridID].rowData[ScrollY].hide ? ' (hidden)' : '') + " ..."); }
                else if(ip_GridProps[GridID].rowData[ScrollY].hide)
                {
                    if (!ScrollIndecatorVisible) { $('#' + GridID + '_q4_scrollbar_gadget_scrollIndecator_y').show(); ScrollIndecatorVisible = true; }                                      
                    $('#' + GridID + '_q4_scrollbar_gadget_scrollIndecator_y').text(ScrollY + '*');
                }
                else if (ScrollIndecatorVisible) { $('#' + GridID + '_q4_scrollbar_gadget_scrollIndecator_y').hide(); ScrollIndecatorVisible = false; }

                dragPos = ui.position.top;
     
  
            },
            stop: function (event, ui) {

                if (ui.position.top > dragPos || ui.position.top < dragPos) { ScrollY = (ui.position.top * ScrollInterval) + ip_GridProps[GridID].frozenRows; }

                if (ScrollY >= ip_GridProps[GridID].rows) { ScrollY = ip_GridProps[GridID].rows - 1; }

                if (ip_GridProps[GridID].scrollY != ScrollY) { ip_ScrollToY(GridID, ScrollY, true, 'all', 'none'); }


                ip_GridProps[GridID].events.ip_InitScrollY_Large_MouseMove = ip_UnBindEvent(document, 'mousemove', ip_GridProps[GridID].events.ip_InitScrollY_Large_MouseMove);

                $('#' + GridID + '_q4_scrollbar_gadget_scrollIndecator_y').fadeOut('fast');
                $('#' + GridID + '_q4_scrollbar_gadget_y').removeClass('ip_grid_scrollBar_gadget_scrolling');

                dragPos = ui.position.top;

            }

        });

        
  

    }
    else {
        //hide scroll bar
        $('#' + GridID + '_q4_scrollbar_container_y').css('display', 'none');
    }

    return true;
}

function ip_InitScrollY(GridID, ScrollInterval, Show, Placement) {

    var dragPos = 0;
    var ScrollY = 0;
    var ScrollCount = 0;


    var ScrollBarX_height = parseInt($('#' + GridID + '_q4_scrollbar_container_x').height());
    var GridTableHeight = parseInt($('#' + GridID + '_table').height());
    var FourthQuadTableHeight = parseInt($('#' + GridID + '_q4_table').outerHeight());
    var FourthQuadContainerHeight = parseInt($('#' + GridID + '_q4_container').outerHeight());
    var FirstQuadTableWidth = parseInt($('#' + GridID + '_q1_table').outerWidth());
    var FirstQuadTableHeight = parseInt($('#' + GridID + '_q1_table').outerHeight());
    var GridWidth = parseInt(ip_GridProps[GridID].dimensions.gridWidth);
    var GridHeight = parseInt(ip_GridProps[GridID].dimensions.gridHeight);
    var ScrollContainerHeight = GridHeight - FirstQuadTableHeight - ScrollBarX_height;

    //Initialize scrollbars
    if (Show) {


        if (ScrollContainerHeight > FourthQuadTableHeight) { ScrollContainerHeight = ScrollContainerHeight; }

        $('#' + GridID + '_q4_scrollbar_container_y').css('display', '');
        $('#' + GridID + '_q4_scrollbar_container_y').css('height', ScrollContainerHeight + 'px');

        var ScrollFit = 0;
        var ScrollGadgetSize = 0;

        while (ScrollGadgetSize < 10) {
            ScrollFit = ((FourthQuadTableHeight - ScrollContainerHeight) / ScrollInterval);
            ScrollGadgetSize = parseInt(ScrollContainerHeight - ScrollFit);
            if (ScrollGadgetSize < 10) { ScrollInterval++; }
        }

        if (FourthQuadTableHeight < FourthQuadContainerHeight) { ScrollGadgetSize = FourthQuadContainerHeight; }


        $("#" + GridID + "_q4_scrollbar_container_y").css('height', FourthQuadContainerHeight + 'px');
        $("#" + GridID + "_q4_scrollbar_gadget_y").css('height', ScrollGadgetSize + 'px');

        //alert('FourthQuadTableHeight : ' + FourthQuadTableHeight + '\nFourthQuadContainerHeight : ' + FourthQuadContainerHeight + '\nScrollGadgetSize : ' + ScrollGadgetSize);

        $("#" + GridID + "_q4_scrollbar_gadget_y").draggable({
            axis: "y",
            containment: "parent",
            drag: function (event, ui) {

                if (ui.position.top > dragPos || ui.position.top < dragPos) { ScrollY = ui.position.top * ScrollInterval; }

                ip_GridProps[GridID].scrollY = ScrollY;
                $('#' + GridID + '_q3_container').scrollTop(ScrollY);
                $('#' + GridID + '_q4_container').scrollTop(ScrollY);

                dragPos = ui.position.top;

                $('#ScrollCount').html(dragPos);
            }



        }, 50);

    }
    else {
        //hide scroll bar
        $('#' + GridID + '_q4_scrollbar_container_y').css('display', 'none');
    }

}

function ip_ScrollToY(GridID, ScrollToRow, SetScrollGadget, reloadRanges, hideRanges) {
    
    //if (!SetScrollGadget && ip_GridProps[GridID].rowData[ScrollToRow].hide) { return false; }
    if (ip_GridProps[GridID].scrolling) { return false; }

    ScrollToRow = parseInt(ScrollToRow);

    if (ip_GridProps[GridID].scrollY != ScrollToRow && !ip_GridProps[GridID].scrolling) { 

        ip_GridProps[GridID].scrolling = true;
        
        clearTimeout(ip_GridProps[GridID].timeouts.scrollYCompleteTimeout);


        //Valide the scroll to row
        var direction = (ip_GridProps[GridID].scrollY < ScrollToRow ? 'down' : 'up');
        if (SetScrollGadget) { ScrollToRow = ip_NextNonHiddenRow(GridID, ScrollToRow, null, ip_GridProps[GridID].scrollY, direction); }

        //while (SetScrollGadget && ScrollToRow > ip_GridProps[GridID].frozenRows && ScrollToRow < ip_GridProps[GridID].rows && ip_GridProps[GridID].rowData[ScrollToRow].hide) { if (direction == 'down') { ScrollToRow++; } else { ScrollToRow--; } }
        
        if (ScrollToRow >= ip_GridProps[GridID].rows) { ScrollToRow = ip_GridProps[GridID].rows - 1; }
        else if (ScrollToRow < ip_GridProps[GridID].frozenRows) { ScrollToRow = ip_GridProps[GridID].frozenRows; }

        var ForceLarge = false;
        var visibleRows = ip_GridProps[GridID].loadedRows - ip_GridProps[GridID].frozenRows;
        var startScrollLoopAt = ip_GridProps[GridID].scrollY;
        var scrollDiff = Math.abs(ScrollToRow - startScrollLoopAt);
        var SetRows = ScrollToRow;
        var MaxSetRows = ScrollToRow + ip_GridProps[GridID].loadedRows;
        var SetRowsDepth = scrollDiff;

               
        

        if (direction == 'down') {

            //-- SCROLL DOWN ---------

            ip_GridProps[GridID].scrollY = ScrollToRow;            

            //Scroll Down
            if (scrollDiff >= ip_GridProps[GridID].loadedRows || ForceLarge) { //ip_GridProps[GridID].loadedRows / 3
                

                ip_AddRow(GridID, { row: ScrollToRow, appendTo: 'end', fullRerender: true });
                SetRowsDepth = 1;
     
            }
            else {

                //SCROLL DOWN using smaller intervals (offers better performance on small scrolls)
                var AddRow = startScrollLoopAt + visibleRows;

                for (var row = startScrollLoopAt; row < ScrollToRow; row++) {
                                        
                    ip_RemoveRow(GridID, { row: row });
                    ip_GridProps[GridID].loadedRows--;

                    if (ip_GridProps[GridID].dimensions.accumulativeScrollHeight < ip_GridProps[GridID].dimensions.scrollHeight && AddRow < ip_GridProps[GridID].rows) {
                                       
                        ip_GridProps[GridID].loadedRows++;
                        ip_AddRow(GridID, { row: AddRow, appendTo: 'end' });
                        AddRow++;

                    }
                    
                }                              

                //Full in the gaps that may be caused due to varying row height
                while (ip_GridProps[GridID].dimensions.accumulativeScrollHeight < ip_GridProps[GridID].dimensions.scrollHeight && AddRow < ip_GridProps[GridID].rows) {
                                        
                    ip_GridProps[GridID].loadedRows++;
                    ip_AddRow(GridID, { row: AddRow, appendTo: 'end' });                   
                    AddRow++;

                }

                SetRows = ScrollToRow + (scrollDiff * (scrollDiff < 0 ? -1 : 1));
            }
        }
        else {

            //-- SCROLL UP ---------
            ip_GridProps[GridID].scrollY = ScrollToRow;

            if (scrollDiff >= ip_GridProps[GridID].loadedRows || ForceLarge) {
                    
                ip_GridProps[GridID].loadedRows++;
                ip_AddRow(GridID, { row: ScrollToRow, appendTo: 'end', fullRerender: true });
                SetRowsDepth = 1;

            }
            else {
  
                //SCROLL UP : using smaller intervals (offers better performance on small scrolls)

                var RemoveRow = startScrollLoopAt + visibleRows - 1;

                
                for (var row = startScrollLoopAt - 1; row >= ScrollToRow; row--) {
                    
                    ip_GridProps[GridID].loadedRows++;
                    ip_AddRow(GridID, { row: row, appendTo: 'start' });

                    if (ip_GridProps[GridID].dimensions.accumulativeScrollHeight - ip_GridProps[GridID].rowData[RemoveRow].height >= ip_GridProps[GridID].dimensions.scrollHeight && RemoveRow > ip_GridProps[GridID].frozenRows) {

                        ip_RemoveRow(GridID, { row: RemoveRow });
                        ip_GridProps[GridID].loadedRows--;
                        RemoveRow--;

                    }

                }

                //Full in the gaps that may be caused due to varying row height
                while (ip_GridProps[GridID].dimensions.accumulativeScrollHeight - ip_GridProps[GridID].rowData[RemoveRow].height > ip_GridProps[GridID].dimensions.scrollHeight && RemoveRow > ip_GridProps[GridID].frozenRows) {

                    
                    ip_GridProps[GridID].loadedRows--;
                    ip_RemoveRow(GridID, { row: RemoveRow });
                    RemoveRow--;

                }
                

                SetRows = ScrollToRow + (scrollDiff * (scrollDiff < 0 ? -1 : 1));
            }

            //ip_GridProps[GridID].scrollY = ScrollToRow;
        }

        if (SetRows > MaxSetRows) { SetRows = MaxSetRows; }
        ip_SetRowColSpan(GridID, ScrollToRow, SetRows, null, null, SetRowsDepth);

            
        if (reloadRanges != null && reloadRanges != 'none') { ip_RePoistionRanges(GridID, reloadRanges, false, true); }

        ////set scroll bar position
        if (SetScrollGadget == true) {
            var ScrollGadgetTop = (ScrollToRow - ip_GridProps[GridID].frozenRows) / ip_GridProps[GridID].scrollYInterval;
            $('#' + GridID + '_q4_scrollbar_gadget_y').css('top', ScrollGadgetTop + 'px');
        }

        //Raise scroll completed event
        ip_GridProps[GridID].timeouts.scrollYCompleteTimeout = setTimeout(function () {

            ip_RaiseEvent(GridID, 'ip_ScrollComplete', '', { ScrollComplete: { Inputs: null, Effected: { scrollY: ScrollToRow, direction: direction } } });

        }, 250);

        ip_GridProps[GridID].scrolling = false;

        return true;
    }

    return false;
}

function ip_ScrollToX(GridID, ScrollToCol, SetScrollGadget, reloadRanges, hideRanges) {

    if (ip_GridProps[GridID].scrollX != ScrollToCol && !ip_GridProps[GridID].scrolling) {

        ip_GridProps[GridID].scrolling = true;

        var visibleCols = ip_GridProps[GridID].loadedCols - ip_GridProps[GridID].frozenCols;
        var col1 = null;
        var col2 = null;
        var startScrollLoopAt = ip_GridProps[GridID].scrollX;
        var scrollDiff = ScrollToCol - startScrollLoopAt;
        var SetCols = ScrollToCol;
        var MaxSetCols = ScrollToCol + ip_GridProps[GridID].loadedCols;
        var SetColsDepth = scrollDiff;
        var Grid = $('#' + GridID);

        ScrollToCol = parseInt(ScrollToCol);


        //Valide the scroll to row
        if (ScrollToCol >= ip_GridProps[GridID].cols) { ScrollToCol = ip_GridProps[GridID].cols - 1; }
        else if (ScrollToCol < ip_GridProps[GridID].frozenCols) { ScrollToCol = ip_GridProps[GridID].frozenCols; }


        if (ip_GridProps[GridID].scrollX < ScrollToCol) {

            ip_GridProps[GridID].scrollX = ScrollToCol; //needed to we append cols at end not begining

            //Scroll RIGHT
            if (ScrollToCol > startScrollLoopAt + ip_GridProps[GridID].loadedCols) {

                //Scroll right using larger intervals (offers excellent performance on large scrolls)
                ip_AddCol(GridID, { col: ScrollToCol, appendTo: 'end', fullRerender: true });
                SetColsDepth = 1;

            }
            else {

                var AddCol = startScrollLoopAt + visibleCols;

                //Scroll right using smaller intervals (offers better performance on small scrolls)
                for (var col = startScrollLoopAt; col < ScrollToCol; col++) {

                    $(Grid).ip_RemoveCol({ col: col });
                    ip_GridProps[GridID].loadedCols--;

                    if (ip_GridProps[GridID].dimensions.accumulativeScrollWidth < ip_GridProps[GridID].dimensions.scrollWidth && AddCol < ip_GridProps[GridID].cols) {

                        ip_GridProps[GridID].loadedCols++;
                        ip_AddCol(GridID, { col: AddCol, appendTo: 'end' });
                        AddCol++;

                    }
                }

                //Full in the gaps that may be caused due to varying col width
                while (ip_GridProps[GridID].dimensions.accumulativeScrollWidth < ip_GridProps[GridID].dimensions.scrollWidth && AddCol < ip_GridProps[GridID].cols) {

                    ip_GridProps[GridID].loadedCols++;
                    ip_AddCol(GridID, { col: AddCol, appendTo: 'end' });
                    AddCol++;

                }

                SetCols = ScrollToCol + (scrollDiff * (scrollDiff < 0 ? -1 : 1));
            }
        }
        else if (ip_GridProps[GridID].scrollX > ScrollToCol) {

            ip_GridProps[GridID].scrollX = ScrollToCol; //needed to we append cols at end not begining

            //Scroll Left
            if (ScrollToCol < startScrollLoopAt - visibleCols) {

                //Scroll Left using larger intervals (offers excellent performance on large scrolls)
                ip_AddCol(GridID, { col: ScrollToCol, appendTo: 'end', fullRerender:true });
                SetColsDepth = 1;

            }
            else {

                var RemoveCol = startScrollLoopAt + visibleCols - 1;

                //Scroll if using smaller intervals (offers better performance on small scrolls)
                for (var col = startScrollLoopAt - 1; col >= ScrollToCol; col--) {

                    ip_GridProps[GridID].loadedCols++;
                    ip_AddCol(GridID, { col: col, appendTo: 'start' });
                    var test = ip_GridProps[GridID].dimensions.accumulativeScrollWidth;
                    if (ip_GridProps[GridID].dimensions.accumulativeScrollWidth - ip_ColWidth(GridID, RemoveCol, true) >= ip_GridProps[GridID].dimensions.scrollWidth && RemoveCol > ip_GridProps[GridID].frozenCols) {

                        ip_RemoveCol(GridID, { col: RemoveCol });
                        ip_GridProps[GridID].loadedCols--;
                        RemoveCol--;

                    }    

                }

                //Full in the gaps that may be caused due to varying col width
                while (ip_GridProps[GridID].dimensions.accumulativeScrollWidth - ip_ColWidth(GridID, RemoveCol, true) > ip_GridProps[GridID].dimensions.scrollWidth && RemoveCol > ip_GridProps[GridID].frozenCols) {
                                        
                    ip_RemoveCol(GridID, { col: RemoveCol });
                    ip_GridProps[GridID].loadedCols--;
                    RemoveCol--;

                }

                SetCols = ScrollToCol + (scrollDiff * (scrollDiff < 0 ? -1 : 1));
            }
        }

        if (SetCols > MaxSetCols) { SetCols = MaxSetCols; }
        ip_SetRowColSpan(GridID, null, null, ScrollToCol, SetCols, SetColsDepth);

        if (reloadRanges != null && reloadRanges != 'none') {
            clearTimeout(ip_GridProps[GridID].events.scrollToXY_RangeTimeout);
            ip_GridProps[GridID].events.scrollToXY_RangeTimeout = setTimeout(function () { ip_RePoistionRanges(GridID, reloadRanges, true, false); }, 15);
        }

        //set scroll bar position
        if (SetScrollGadget == true || SetScrollGadget == null) {
            var ScrollGadgetLeft = (ScrollToCol - ip_GridProps[GridID].frozenCols) / ip_GridProps[GridID].scrollXInterval;
            $('#' + GridID + '_q4_scrollbar_gadget_x').css('left', ScrollGadgetLeft + 'px');
        }

        ip_GridProps[GridID].scrollX = ScrollToCol;
   
        ip_GridProps[GridID].scrolling = false;
    }

}

function ip_EnableScrollAnimate(GridID) {
    //Turns on scrolling when the mouse moves into a hot zone 
        
    if (!ip_GridProps[GridID].scrollAnimate) {

        ip_GridProps[GridID].scrollAnimate = true;

        var GridTop = $('#' + GridID + '_table').offset().top;
        var GridLeft = $('#' + GridID + '_table').offset().left;
        var GridHeight = ip_GridProps[GridID].dimensions.gridHeight;
        var GridWidth = ip_GridProps[GridID].dimensions.gridWidth;

        ip_GridProps[GridID].events.scrollAnimate_MouseMove = ip_UnBindEvent('#' + GridID, 'mousemove', ip_GridProps[GridID].events.scrollAnimate_MouseMove);        
        $('#' + GridID).mousemove(ip_GridProps[GridID].events.scrollAnimate_MouseMove = function (e) {

            
            ip_ScrollAnimate(GridID, GridTop, GridHeight, GridLeft, GridWidth, e, 10);

        });

        ip_GridProps[GridID].events.scrollAnimate_MouseUp = ip_UnBindEvent(document, 'mouseup', ip_GridProps[GridID].events.scrollAnimate_MouseUp);        
        $(document).mouseup(ip_GridProps[GridID].events.scrollAnimate_MouseUp = function (e) {
            ip_DisableScrollAnimate(GridID);
            ip_GridProps[GridID].events.scrollAnimate_MouseMove = ip_UnBindEvent('#' + GridID, 'mousemove', ip_GridProps[GridID].events.scrollAnimate_MouseMove);
            ip_GridProps[GridID].events.scrollAnimate_MouseUp = ip_UnBindEvent(document, 'mouseup', ip_GridProps[GridID].events.scrollAnimate_MouseUp);
        });

    }
    
}

function ip_ScrollAnimate(GridID, GridTop, GridHeight, GridLeft, GridWidth, mouseArgs, speed) {
    
    ip_GridProps[GridID].scrollAnimate = true;
    
    var AreaZone = ip_GridProps[GridID].dimensions.defaultRowHeight;
    var GridBottomZone = GridTop + GridHeight - AreaZone;
    var GridTopZone = GridTop + AreaZone;
    var GridRightZone = GridLeft + GridWidth - AreaZone;
    var GridLeftZone = GridLeft + AreaZone;
    
    speed = (speed == null ? 0 : speed);

    if (mouseArgs.pageY <= GridTopZone && ip_GridProps[GridID].scrollY > ip_GridProps[GridID].frozenRows) {
        //Scroll top

        if (ip_GridProps[GridID].timeouts.scrollAnimateInterval != null) { return; }

        ip_GridProps[GridID].timeouts.scrollAnimateInterval = setInterval(function () {

            ip_ScrollToY(GridID, ip_GridProps[GridID].scrollY - (1), true, 'all', 'none');

        }, speed);

        return;

    }
    else if (mouseArgs.pageY >= GridBottomZone && ip_GridProps[GridID].scrollY < ip_GridProps[GridID].rows - 1) {
        //Scroll bottom

        if (ip_GridProps[GridID].timeouts.scrollAnimateInterval != null) { return; }

        ip_GridProps[GridID].timeouts.scrollAnimateInterval = setInterval(function () {

            ip_ScrollToY(GridID, ip_GridProps[GridID].scrollY + (1), true, 'all', 'none');

        }, speed);
        
        return;

    }
    else if (mouseArgs.pageX <= GridLeftZone && ip_GridProps[GridID].scrollX > ip_GridProps[GridID].frozenCols) {
        //Scroll left

        if (ip_GridProps[GridID].timeouts.scrollAnimateInterval != null) { return; }

        ip_GridProps[GridID].timeouts.scrollAnimateInterval = setInterval(function () {

            ip_ScrollToX(GridID, ip_GridProps[GridID].scrollX - (1), true, 'all', 'none');

        }, speed);

        return;

    }
    else if (mouseArgs.pageX >= GridRightZone && ip_GridProps[GridID].scrollX < ip_GridProps[GridID].cols - 1) {
        //Scroll right

        if (ip_GridProps[GridID].timeouts.scrollAnimateInterval != null) { return; }

        ip_GridProps[GridID].timeouts.scrollAnimateInterval = setInterval(function () {

            ip_ScrollToX(GridID, ip_GridProps[GridID].scrollX + (1), true, 'all', 'none');

        }, speed);

        return;
    }


    ip_DisableScrollAnimate(GridID);
    var AnimateSession = 'bef7bea9-3157-4434-8cbb-22320d0d294e';
}

function ip_DisableScrollAnimate(GridID) {

    clearInterval(ip_GridProps[GridID].timeouts.scrollAnimateInterval);
    ip_GridProps[GridID].timeouts.scrollAnimateInterval = null;
    ip_GridProps[GridID].scrollAnimate = false;

}

function ip_ScrollCell(GridID, direction, startRow, startCol, endRow, endCol, scrollIncrement, reloadRanges) {

    endRow = (endRow == null ? startRow : endRow);
    endCol = (endCol == null ? startCol : endCol);

    scrollIncrement = (scrollIncrement < 0 ? scrollIncrement * -1 : scrollIncrement);

    //Dont scrol if we are in a frozen zone
    if (startRow == endRow && (direction == 'up' || direction == 'down')) {
        var Quad = ip_GetQuad(GridID, startRow, null);
        if (Quad == 1 || Quad == 2) { return }
    }
    else if (endCol == endRow && (startCol == 'left' || direction == 'right')) {
        var Quad = ip_GetQuad(GridID, null, startCol);
        if (Quad == 1 || Quad == 3) { return }
    }


    if (direction == 'up') {

        var startPos = ip_CellRelativePosition(GridID, startRow + 1, startCol, true);
        var endPos = ip_CellRelativePosition(GridID, endRow + 1, endCol, true);

        if ((startPos.localTop <= 1 || endPos.localTop <= 1)) {
            ip_ScrollToY(GridID, ip_GridProps[GridID].scrollY - scrollIncrement, true, reloadRanges, 'none');

            return 'up';
        }
    }
    else if (direction == 'down' && endRow < ip_GridProps[GridID].rows - 1) {

        var QuadHeight = ip_GridProps[GridID].dimensions.scrollHeight
        var startPos = ip_CellRelativePosition(GridID, startRow, startCol, true);
        var endPos = ip_CellRelativePosition(GridID, endRow + 1, endCol, true); //+ 1

        if (startPos.localBottom >= QuadHeight || endPos.localBottom >= QuadHeight) { // || endPos.localTop <= 0
            ip_ScrollToY(GridID, ip_GridProps[GridID].scrollY + scrollIncrement, true, reloadRanges, 'none');

            return 'down';
        }
    }
    else if (direction == 'left') {

        var startPos = ip_CellRelativePosition(GridID, startRow, startCol + 1, true);
        var endPos = ip_CellRelativePosition(GridID, endRow, endCol + 1, true);

        if (startPos.localLeft <= 1 || endPos.localLeft <= 1) {

            ip_ScrollToX(GridID, ip_GridProps[GridID].scrollX - scrollIncrement, true, reloadRanges, 'all');
            return 'left';
        }
    }
    else if (direction == 'right' && endRow < ip_GridProps[GridID].rows - 1) {

        var QuadWidth = ip_GridProps[GridID].dimensions.scrollWidth; 
        var startPos = ip_CellRelativePosition(GridID, startRow, startCol, true);
        var endPos = ip_CellRelativePosition(GridID, endRow, endCol, true); //+ 1

        if (startPos.localRight >= QuadWidth || endPos.localRight >= QuadWidth) { // || endPos.localLeft <= 0

            ip_ScrollToX(GridID, ip_GridProps[GridID].scrollX + scrollIncrement, true, reloadRanges, 'none');
            return 'right';

        }

    }


}


//----- ADD/REMOVE INSERT/DELETE ROWS/COLS ------------------------------------------------------------------------------------------------------------------------------------

function ip_AddRow(GridID, options) {

    var options = $.extend({

        row: ip_GridProps[GridID].rows, //Row to add
        count: 1, //number of rows to add from the row
        visualize: true,
        appendTo: 'end', //'start' or 'end',
        fullRerender: false //Automatically reloads all rows in the visible grid

    }, options);

    
    //Validate row
    if (options.row >= ip_GridProps[GridID].rows) { options.row = ip_GridProps[GridID].rows; }
    if (options.fullRerender) { options.count = 1; }

    var AddedNewRows = false;
    var TotalRows = 0;
    var RowsQ3 = new Array();
    var RowsQ4 = new Array();
    var optionsQ3 = null;
    var optionsQ4 = null;
    var loadedScrollable = ip_LoadedRowsCols(GridID);

    var Quad = ip_GetQuad(GridID, options.row, -1);
    var Quad2 = (Quad == 1 ? 2 : 4);

    var addCount = options.row + options.count;
    var TransactionID = null;

    if (options.fullRerender) {



        if (Quad == 1) { ip_RaiseEvent(GridID, 'warning',arguments.callee.caller, 'Full re-render only works in scrollable zone'); return false; }

        //Reset the accumulative scroll height as this will be a full re-render
        ip_GridProps[GridID].dimensions.accumulativeScrollHeight = 0;

        //clear rows
        if (thisBrowser.name == 'ie' && thisBrowser.version < 10) {
            $('#' + GridID + '_q3_table tbody tr').remove();
            $('#' + GridID + '_q4_table tbody tr').remove();
        }
        else {
            document.getElementById(GridID + '_q3_table_tbody').innerHTML = '';
            document.getElementById(GridID + '_q4_table_tbody').innerHTML = '';
        }

        if (ip_GridProps[GridID].rows == ip_GridProps[GridID].frozenRows) { return false; }
    } 

    for (var row = options.row; row < addCount; row++) {




        if (row >= ip_GridProps[GridID].rows) {

            //We are adding a NEW row, so make sure its added to the data object
            //Add to undo stack
            if (TransactionID == null) { TransactionID = ip_GenerateTransactionID(); }

            var AddRange = { startRow: options.row, startCol: 0, endRow: options.row, endCol: (ip_GridProps[GridID].cols - 1) };
            var FunctionUndoData = ip_AddUndo(GridID, 'ip_AddRow', TransactionID, 'function', AddRange, AddRange, { row: AddRange.startRow, col: AddRange.startCol }, null, function () { ip_RemoveRow(GridID, { row: options.row, count: 1, mode:'destroy', render: false, raiseEvent: false, createUndo: false }); });

            ip_GridProps[GridID].rows++;
            ip_ValidateData(GridID, ip_GridProps[GridID].rows, -1);
            if (ip_GridProps[GridID].dimensions.accumulativeScrollHeight  <= ip_GridProps[GridID].dimensions.scrollHeight) { ip_GridProps[GridID].loadedRows++; }
            loadedScrollable = ip_LoadedRowsCols(GridID);
            AddedNewRows = true;
        }

        //Do not visualize row if it is beyond loaded rows
        if (row > loadedScrollable.rowTo_scroll && !options.fullRerender) { options.visualize = false; } //OLD WAY // + 2
        if (options.visualize) {

            optionsQ3 = {
                showColSelector: false,
                showRowSelector: ip_GridProps[GridID].showRowSelector,
                id: GridID + '_q' + Quad,
                cols: ip_GridProps[GridID].frozenCols,
                Quad: Quad,
                GridID: GridID,
                row: row,
                startCol: 0,
                loadedScrollable: loadedScrollable
            }

            optionsQ4 = {
                showColSelector: false,
                showRowSelector: false,
                id: GridID + '_q' + Quad2,
                cols: ip_GridProps[GridID].cols - ip_GridProps[GridID].frozenCols,
                Quad: Quad2,
                GridID: GridID,
                row: row,
                startCol: ip_GridProps[GridID].scrollX,
                loadedScrollable: loadedScrollable
            }


            ip_CreateGridQuadRow(optionsQ3, RowsQ3);
            ip_CreateGridQuadRow(optionsQ4, RowsQ4);

            //Automatically calulcate the amout of rows to add for a full rerender
            if (Quad == 3) { ip_GridProps[GridID].dimensions.accumulativeScrollHeight += ip_RowHeight(GridID, row, true); }            
            if (options.fullRerender && ip_GridProps[GridID].dimensions.accumulativeScrollHeight < ip_GridProps[GridID].dimensions.scrollHeight && row < ip_GridProps[GridID].rows - 1) { addCount++; }


            TotalRows++;
        }
    }

    if (options.fullRerender) { ip_GridProps[GridID].loadedRows = TotalRows + ip_GridProps[GridID].frozenRows; }

    if (RowsQ3.length > 0) {

        if (options.fullRerender) {

            if (thisBrowser.name == 'ie' && thisBrowser.version < 10) {

                $('#' + GridID + '_q' + Quad + '_table_tbody').append(RowsQ3.join(''));
                $('#' + GridID + '_q' + Quad2 + '_table_tbody').append(RowsQ4.join(''));

            }
            else {

                var Q3R = document.getElementById(GridID + '_q' + Quad + '_table_tbody');
                var Q4R = document.getElementById(GridID + '_q' + Quad2 + '_table_tbody');
                if (Q3R != null) { Q3R.innerHTML = RowsQ3.join(''); }
                if (Q4R != null) { Q4R.innerHTML = RowsQ4.join(''); }

            }

        }
        else {
            if (options.appendTo == 'end') { //options.row > ip_GridProps[GridID].scrollY

                $('#' + GridID + '_q' + Quad + '_table_tbody').append(RowsQ3.join(''));
                $('#' + GridID + '_q' + Quad2 + '_table_tbody').append(RowsQ4.join(''));

            }
            else {
                $('#' + GridID + '_q' + Quad + '_table_tbody').prepend(RowsQ3.join(''));
                $('#' + GridID + '_q' + Quad2 + '_table_tbody').prepend(RowsQ4.join(''));
            }
        }
        
    }

    if (AddedNewRows) {

        $('#' + GridID).ip_Scrollable();

        //Raise event
        var Effected = { count: options.count };
        ip_RaiseEvent(GridID, 'ip_AddRow', TransactionID, { AddRow: { Inputs: options, Effected: Effected } });

    }

    return true;
}

function ip_RemoveRow(GridID, options) {

    var options = $.extend({

        row: ip_GridProps[GridID].rows - 1,
        count: 1,
        mode: 'hide', //hide makes row invisible but does not destroy it from the datasource
        render: true,
        raiseEvent: true,
        createUndo: true
        
    }, options);


    if (options.row == null && ip_GridProps[GridID].selectedRow.length > 0) { options.row = ip_GridProps[GridID].selectedRow[0]; }
    else if (options.row == null) { options.row = ip_GridProps[GridID].rows; }
    if (options.count == null && ip_GridProps[GridID].selectedRow.length > 0) { options.count = ip_GridProps[GridID].selectedRow.length; }
    else if (options.count == null) { options.count = 1; }
    
    var Quad = ip_GetQuad(GridID, options.row, -1);
    var Quad2 = (Quad == 1 ? 2 : 4);


    //Validate count
    if ((options.row + options.count) > ip_GridProps[GridID].rows) { options.count = ip_GridProps[GridID].rows - options.row; }

    //Remove row from asthetical view only - mainly used for scrolling
    if (options.mode == 'hide') {
        
        for (var c = 0; c < options.count; c++) {

            var Q3R = document.getElementById(GridID + '_q' + Quad + '_gridRow_' + options.row);
            var Q4R = document.getElementById(GridID + '_q' + Quad2 + '_gridRow_' + options.row);

            if (Q3R != null && Q4R != null) {
                
                //Visual area

                if (thisBrowser.name == 'ie' && thisBrowser.version < 10) {

                    Q3R.parentNode.removeChild(Q3R);
                    Q4R.parentNode.removeChild(Q4R);

                }
                else {

                    Q3R.outerHTML = ''
                    Q4R.outerHTML = ''
                }

                ip_GridProps[GridID].dimensions.accumulativeScrollHeight -= ip_RowHeight(GridID, options.row, true);
            }

        }

    }
    else {


        //Perminantly destroy ROW
        if (options.row < ip_GridProps[GridID].rows) {

            var TransactionID = ip_GenerateTransactionID();
            var Effected = {  row: options.row , count: options.count, rowData:null };
            
            //Move formulas - but exclude the range that we are moving
            var fxList = ip_ChangeFormulaOrigin(GridID, { rowDiff: -options.count, range: { startRow: (options.row + options.count) } });


            for (var c = 0; c < options.count; c++) {

                if (options.createUndo) {

                    var RemoveRange = { startRow: options.row, startCol: -1, endRow: options.row, endCol: (ip_GridProps[GridID].cols - 1) };
                    var CellUndoData = ip_AddUndo(GridID, 'ip_RemoveRow', TransactionID, 'CellData', RemoveRange, (c == 0 ? { startRow: RemoveRange.startRow, startCol: RemoveRange.startCol, endRow: RemoveRange.endRow + options.count - 1, endCol: RemoveRange.endCol } : null ), ( c == 0 ?  { row: RemoveRange.startRow, col: RemoveRange.startCol } : null), null, null);
                    var MergeUndoData = ip_AddUndo(GridID, 'ip_RemoveRow', TransactionID, 'MergeData', RemoveRange);
                    var RowUndoData = ip_AddUndo(GridID, 'ip_RemoveRow', TransactionID, 'RowData', RemoveRange, null, null, null, ip_CloneRow(GridID, options.row));
                    var FunctionUndoData = ip_AddUndo(GridID, 'ip_RemoveRow', TransactionID, 'function', RemoveRange, null, null, null, function () { ip_InsertRow(GridID, { row: options.row, count: 1, render: false, raiseEvent: false, createUndo:false }); });
                }


                //Clear out any merges that are directly on the row so that they can be adjusted and added back after the row removal has taken place
                var merges = ip_ValidateRangeMergedCells(GridID, options.row, 0, options.row, ip_GridProps[GridID].cols - 1);//Get merges that fall within zone            
                for (var m = 0; m < merges.merges.length; m++) {
                    //Create merge undo stack
                    ip_AddUndoTransactionData(GridID, MergeUndoData, ip_mergeObject(merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol, merges.merges[m].rowSpan, merges.merges[m].colSpan));
                    ip_ResetCellMerge(GridID, merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol);
                }
                                
                //Create cell undo stack
                for (var cl = 0; cl < ip_GridProps[GridID].rowData[options.row].cells.length; cl++) { ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, options.row, cl)); }

                //Move all other merges up - ignoring index alignment as this will happen naturally with the splice
                ip_ReshuffelMergesRow(GridID, options.row + 1, ip_GridProps[GridID].rows - 1, -1, false);
                

                //Remove the col from the data object
                ip_GridProps[GridID].rows--;
                ip_GridProps[GridID].rowData.splice(options.row, 1);

                //Bring back any merges that were directly on our row - making the adjustment for the removed row
                for (var m = 0; m < merges.merges.length; m++) { ip_AddRowToMerge(GridID, merges.merges[m], -1); }

                //Manage frozen rows & ScrollY
                if (ip_GridProps[GridID].rows <= ip_GridProps[GridID].frozenRows) { ip_GridProps[GridID].frozenRows = ip_GridProps[GridID].rows; }
                if (ip_GridProps[GridID].scrollY >= ip_GridProps[GridID].rows) { ip_GridProps[GridID].scrollY = ip_GridProps[GridID].rows - 1; }
                if (ip_GridProps[GridID].scrollY < ip_GridProps[GridID].frozenRows) { ip_GridProps[GridID].scrollY = ip_GridProps[GridID].frozenRows; }

            }


            var excludeRangeChange = (options.row == 0 ? null : ip_rangeObject(0, 0, options.row - 1, ip_GridProps[GridID].cols - 1));
            for (var key in fxList) {

                var fx = ip_GridProps[GridID].indexedData.formulaData[key];                                
                ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: ip_MoveFormulaOrigon(GridID, ip_GridProps[GridID].rowData[fx.row].cells[fx.col].formula, fx.row, fx.col, (fx.row - options.count), fx.col, excludeRangeChange) });
                ip_AppendEffectedRowData(GridID, Effected, { row: fx.row, col: fx.col, formula: ip_GridProps[GridID].rowData[fx.row].cells[fx.col].formula });

            }

            //Recalculate formulas                     
            //Effected.rowData = ip_ReCalculateFormulas(GridID, { range: [{ startRow: options.row, startCol: 0, endRow: ip_GridProps[GridID].rows - 1, endCol: ip_GridProps[GridID].cols - 1 }], transactionID: TransactionID, render: false, raiseEvent: false, createUndo: options.createUndo }).Effected.rowData;



            if (options.render) {
                ip_ReRenderRows(GridID);
                $('#' + GridID).ip_Scrollable();
            }

            //Select rows
            if ((options.row + options.count) < ip_GridProps[GridID].rows - 1) { $('#' + GridID).ip_SelectRow({ row: options.row, count: options.count, unselect: false }); }
            else { $('#' + GridID).ip_UnselectRow(); }

            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_RemoveRow', TransactionID, { RemoveRow: { Inputs: options, Effected: Effected } }); }
        }
    }


}

function ip_AddCol(GridID, options) {

    var options = $.extend({

        col: ip_GridProps[GridID].cols, //By default append a column onto the end
        count: 1,
        visualize: true,
        appendTo: 'end', //'start' or 'end',
        fullRerender: false //Automatically reloads all rows in the visible grid

    }, options);

    //Validate column
    if (options.col >= ip_GridProps[GridID].cols) { options.col = ip_GridProps[GridID].cols; }
    if (options.fullRerender) { options.count = 1; }
    
    var AddedNewCols = false;
    var Quad = ip_GetQuad(GridID, -1, options.col);
    var Quad2 = (Quad == 2 ? 4 : 3);
    var NewColQ2 = null;
    var NewColQ4 = null;
    var optionsQ2 = null;
    var optionsQ4 = null;
    var addCount = options.col + options.count;
    var loadedScrollable = ip_LoadedRowsCols(GridID);
    var TotalCols = 0;
    var TransactionID = null;

    if (options.fullRerender) {

        if (Quad == 1) { ip_RaiseEvent(GridID, 'warning',arguments.callee.caller, 'Full re-render only works in scrollable zone'); return false; }

        //Reset the accumulative scroll height as this will be a full re-render
        ip_GridProps[GridID].dimensions.accumulativeScrollWidth = 0;

        //clear columns
        $('#' + GridID + '_q2_table .ip_grid_columnSelectorCell').remove();
        $('#' + GridID + '_q2_table .ip_grid_cell').remove();
        $('#' + GridID + '_q4_table .ip_grid_columnSelectorCell').remove();
        $('#' + GridID + '_q4_table .ip_grid_cell').remove();

        if (ip_GridProps[GridID].cols == ip_GridProps[GridID].frozenCols) { return false; }

    }

    for (var col = options.col; col < addCount; col++) {

        
        if (col >= ip_GridProps[GridID].cols) {

            //We are adding a NEW column, so make sure its added to the data object

            //Add to undo stack
            if (TransactionID == null) { TransactionID = ip_GenerateTransactionID(); }

            var AddRange = { startRow: 0, startCol: options.col, endRow: (ip_GridProps[GridID].rows - 1), endCol: options.col };
            var FunctionUndoData = ip_AddUndo(GridID, 'ip_AddCol', TransactionID, 'function', AddRange, AddRange, { row: AddRange.startRow, col: AddRange.startCol }, null, function () { ip_RemoveCol(GridID, { col: options.col, count: 1, mode: 'destroy', render: false, raiseEvent: false, createUndo: false }); });


            ip_GridProps[GridID].cols++;
            ip_ValidateData(GridID, -1, ip_GridProps[GridID].cols);
            if (ip_GridProps[GridID].dimensions.accumulativeScrollWidth <= ip_GridProps[GridID].dimensions.scrollWidth) { ip_GridProps[GridID].loadedCols++; }
            loadedScrollable = ip_LoadedRowsCols(GridID);
            AddedNewCols = true;
        }
        

        //Do not visualize if we are beyond loaded rows
        if (col > loadedScrollable.colTo_scroll + 1 && !options.fullRerender) { options.visualize = false; } //OLD WAY
   
        if (options.visualize) {

            //Add frozen rows
            for (var row = -1; row <= loadedScrollable.rowTo_frozen; row++) {

                optionsQ2 = {
                    showColSelector: ip_GridProps[GridID].showColSelector,
                    showRowSelector: false,
                    id: GridID + '_q' + Quad,
                    Quad: Quad,
                    GridID: GridID,
                    col: col,
                    row: row,
                    cellType: (row == -1 ? 'ColSelector' : '')
                }

                NewColQ2 = ip_CreateGridQuadCell(optionsQ2);

                if (options.appendTo == 'end') {

                    $('#' + GridID + '_q' + Quad + '_gridRow_' + row).append(NewColQ2);

                }
                else {

                    var selectorID = '';

                    if (row == -1) { selectorID = GridID + '_q' + Quad + '_columnSelectorCell_-1'; }
                    else { selectorID = GridID + '_q' + Quad + '_rowSelecterCell_' + row; }

                    $(NewColQ2).insertAfter($('#' + selectorID));

                }

            }

            //Add scollable rows
            for (var row = -1; row <= loadedScrollable.rowTo_scroll; row++) {

                optionsQ4 = {
                    showColSelector: false,
                    showRowSelector: false,
                    id: GridID + '_q' + Quad2,
                    Quad: Quad2,
                    GridID: GridID,
                    col: col,
                    row: row,
                    cellType: (row == -1 ? 'ColSelector' : '')
                }

                NewColQ4 = ip_CreateGridQuadCell(optionsQ4);

                if (options.appendTo == 'end') {  

                    //append at end of row
                    $('#' + GridID + '_q' + Quad2 + '_gridRow_' + row).append(NewColQ4);

                }
                else {


                    var selectorID = '';
                    if (row == -1) { selectorID = GridID + '_q' + Quad2 + '_columnSelectorCell_-1'; } //append after row selector cell
                    else { selectorID = GridID + '_q' + Quad2 + '_rowSelecterCell_' + row; }

                    $(NewColQ4).insertAfter($('#' + selectorID));

                }
                
                if (row == -1) { row = loadedScrollable.rowFrom_scroll - 1; }

            }
            

            //Automatically calulcate the amout of cols to add for a full rerender
            if (Quad == 2) { ip_GridProps[GridID].dimensions.accumulativeScrollWidth += ip_ColWidth(GridID, col, true); }
            if (options.fullRerender && ip_GridProps[GridID].dimensions.accumulativeScrollWidth < ip_GridProps[GridID].dimensions.scrollWidth && col < ip_GridProps[GridID].cols - 1) { addCount++; }

            TotalCols++;

        }
    }

    if (options.fullRerender) { ip_GridProps[GridID].loadedCols = TotalCols + ip_GridProps[GridID].frozenCols; }

    if (AddedNewCols) {

        $('#' + GridID).ip_Scrollable({ initY: false });

        //Raise event
        var Effected = { count: options.count };
        ip_RaiseEvent(GridID, 'ip_AddCol', TransactionID, { AddCol: { Inputs: options, Effected: Effected } });

    }

    return true;
}

function ip_RemoveCol(GridID, options) {

    var options = $.extend({

        col: ip_GridProps[GridID].cols - 1,
        mode: 'hide', //hide makes row invisible but does not destroy it from the datasource
        count: 1,
        render: true,
        raiseEvent: true,
        createUndo: true

    }, options);

    if (options.col == null && ip_GridProps[GridID].selectedColumn.length > 0) { options.col = ip_GridProps[GridID].selectedColumn[0]; }
    else if (options.col == null) { options.col = ip_GridProps[GridID].cols; }
    if (options.count == null && ip_GridProps[GridID].selectedColumn.length > 0) { options.count = ip_GridProps[GridID].selectedColumn.length; }
    else if (options.count == null) { options.count = 1; }

    var Quad = ip_GetQuad(GridID, -1, options.col);
    var Quad2 = (Quad == 1 ? 3 : 4);
    var loadedScrollable = ip_LoadedRowsCols(GridID);



    //Validate count
    if ((options.col + options.count) > ip_GridProps[GridID].cols) { options.count = ip_GridProps[GridID].cols - options.col; }

    if (options.mode == 'hide') {

        for (var c = 0; c < options.count; c++) {

            //Remove row from asthetical view only
            //Remove frozen columns
            for (var row = -1; row < loadedScrollable.rowLoaded_frozen; row++) {

                var ID = '';

                if (row == -1) { ID = GridID + '_q' + Quad + '_columnSelectorCell_' + options.col; }
                else { ID = GridID + '_cell_' + row + '_' + options.col; }

                var cell = document.getElementById(ID);

                if (cell != null) {

                    if (thisBrowser.name == 'ie' && thisBrowser.version < 10) {
                        cell.parentNode.removeChild(cell);
                    } else {
                        cell.outerHTML = '';
                    }

                }


            }

            //Remove scroll columns
            for (var row = -1; row < loadedScrollable.rowFrom_scroll + loadedScrollable.rowLoaded_scroll; row++) {

                var ID = '';

                if (row == -1) { ID = GridID + '_q' + Quad2 + '_columnSelectorCell_' + options.col; }
                else { ID = GridID + '_cell_' + row + '_' + options.col; }

                var cell = document.getElementById(ID);
                if (cell != null) {

                    if (thisBrowser.name == 'ie' && thisBrowser.version < 10) {
                        cell.parentNode.removeChild(cell);
                        //$('#' + ID).remove();
                    } else {
                        cell.outerHTML = '';
                    }

                }


                if (row == -1) {
                    row = ip_GridProps[GridID].scrollY - 1;
                    if (cell != null) { ip_GridProps[GridID].dimensions.accumulativeScrollWidth -= ip_ColWidth(GridID, options.col, true); }
                }

            }

        }

    }
    else {
        //Perminantly destroy column
        if (options.col < ip_GridProps[GridID].cols) {

            var TransactionID = ip_GenerateTransactionID();
            var Effected = { col: options.col, count: options.count, rowData: null };


            //Move formulas - but exclude the range that we are moving
            var fxList = ip_ChangeFormulaOrigin(GridID, { colDiff: -options.count, range: { startCol: (options.col + options.count) } });

            for (var c = 0; c < options.count; c++) {

                if (options.createUndo) {
                    var RemoveRange = { startRow: -1, startCol: options.col, endRow: (ip_GridProps[GridID].rows - 1), endCol: (options.col) };
                    var CellUndoData = ip_AddUndo(GridID, 'ip_RemoveCol', TransactionID, 'CellData', RemoveRange, (c == 0 ? { startRow: RemoveRange.startRow, startCol: RemoveRange.startCol, endRow: RemoveRange.endRow, endCol: (RemoveRange.endCol + options.count - 1) } : null), (c == 0 ? { row: RemoveRange.startRow, col: RemoveRange.startCol } : null));
                    var MergeUndoData = ip_AddUndo(GridID, 'ip_RemoveCol', TransactionID, 'MergeData', RemoveRange);
                    var ColUndoData = ip_AddUndo(GridID, 'ip_RemoveCol', TransactionID, 'ColData', RemoveRange, null, null, null, ip_CloneCol(GridID, (options.col)));
                    var FunctionUndoData = ip_AddUndo(GridID, 'ip_RemoveCol', TransactionID, 'function', RemoveRange, null, null, null, function () { ip_InsertCol(GridID, { col: options.col, count: 1, render: false, raiseEvent: false, createUndo:false }); });
                }
              

                //REMOVE COL                
                //Clear out any merges that are DIRECTLY on the column so that they can be adjusted and added back after the column removal
                var merges = ip_ValidateRangeMergedCells(GridID, 0, options.col, ip_GridProps[GridID].rows - 1, options.col);//Get merges that fall within zone
                for (var m = 0; m < merges.merges.length; m++) {
                    //Create merge undo stack
                    ip_AddUndoTransactionData(GridID, MergeUndoData, ip_mergeObject(merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol, merges.merges[m].rowSpan, merges.merges[m].colSpan)); 
                    ip_ResetCellMerge(GridID, merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol);
                }

                //Move all other merges down - ignoring index alignment as this will happen naturally with the splice
                ip_ReshuffelMergesCol(GridID, options.col + 1, ip_GridProps[GridID].cols - 1, -1, false);

                ip_GridProps[GridID].cols--;

                //Reduce accumulative scroll width
                if (ip_IsColLoaded(GridID, options.col, loadedScrollable, true)) { ip_GridProps[GridID].dimensions.accumulativeScrollWidth -= ip_ColWidth(GridID, options.col, true); }

                //Remove on a data level                
                for (var row = 0; row < ip_GridProps[GridID].rowData.length; row++) {
                    //Create cell undo stack
                    ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, row, options.col));
                    ip_GridProps[GridID].rowData[row].cells.splice(options.col, 1);
                }

                //Remove the col from the coldata object                
                ip_GridProps[GridID].colData.splice(options.col, 1);

                //Bring back any merges that were directly on our column - making the adjustment for the removed column
                for (var m = 0; m < merges.merges.length; m++) { ip_AddColumnToMerge(GridID, merges.merges[m], -1); }

                //Manage frozen rows & ScrollY
                if (ip_GridProps[GridID].cols <= ip_GridProps[GridID].frozenCols) { ip_GridProps[GridID].frozenCols = ip_GridProps[GridID].cols; }
                if (ip_GridProps[GridID].scrollX >= ip_GridProps[GridID].cols) { ip_GridProps[GridID].scrollX = ip_GridProps[GridID].cols - 1; }
                if (ip_GridProps[GridID].scrollX < ip_GridProps[GridID].frozenCols) { ip_GridProps[GridID].scrollX = ip_GridProps[GridID].frozenCols; }

            }


            //Shift formula columns e.g. e0 = d0
            var excludeRangeChange = (options.col == 0 ? null : ip_rangeObject(0, 0, ip_GridProps[GridID].rows - 1, options.col - 1));
            for (var key in fxList) {

                var fx = ip_GridProps[GridID].indexedData.formulaData[key];
                ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: ip_MoveFormulaOrigon(GridID, ip_GridProps[GridID].rowData[fx.row].cells[fx.col].formula, fx.row, fx.col, fx.row, (fx.col - options.count), excludeRangeChange) });
                ip_AppendEffectedRowData(GridID, Effected, { row: fx.row, col: fx.col, formula: ip_GridProps[GridID].rowData[fx.row].cells[fx.col].formula });

            }

            //Recalculate formulas                     
            //Effected.rowData = ip_ReCalculateFormulas(GridID, { range: [{ startCol: options.col, startRow: 0, endCol: ip_GridProps[GridID].cols - 1, endRow: ip_GridProps[GridID].rows - 1 }], transactionID: TransactionID, render: false, raiseEvent: false, createUndo: options.createUndo }).Effected.rowData;


            //Render
            if (options.render) {
                ip_ReRenderCols(GridID);
                $('#' + GridID).ip_Scrollable({ initY: false });
            }


            //Select rows
            if ((options.col + options.count) < ip_GridProps[GridID].cols - 1) { $('#' + GridID).ip_SelectColumn({ col: options.col, count: options.count, unselect: false }); }
            else { $('#' + GridID).ip_UnselectColumn(); }

            //Raise event
            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_RemoveCol', TransactionID, { RemoveCol: { Inputs: options, Effected: Effected } }); }
        }

    }

}

function ip_InsertCol(GridID, options){

    var options = $.extend({

        col: ip_GridProps[GridID].cols, //By default append a column onto the end
        count: 1,
        appendTo: 'before', //'before' or 'after',
        render: true,
        raiseEvent: true,
        createUndo: true
        

    }, options);

    var TransactionID = ip_GenerateTransactionID();
    
    if (options.col == null && ip_GridProps[GridID].selectedColumn.length > 0) { options.col = ip_GridProps[GridID].selectedColumn[0]; }
    else if(options.col == null) { options.col = ip_GridProps[GridID].cols; }

    if (options.count == null && ip_GridProps[GridID].selectedColumn.length > 0) { options.count = ip_GridProps[GridID].selectedColumn.length; }
    else if (options.count == null) { options.count = 1; }

    if (options.col < ip_GridProps[GridID].cols) {
            
        var ColToSplice = (options.appendTo == 'before' ? options.col : options.col + 1);
        var Effected = { col: options.col, count: options.count, appendTo: options.appendTo, rowData: null };

        //Add to undostack   
        if (options.createUndo) {            
            var AddRange = { startRow: 0, startCol: ColToSplice, endRow: (ip_GridProps[GridID].rows - 1), endCol: (ColToSplice + options.count - 1) };
            var FunctionUndoData = ip_AddUndo(GridID, 'ip_InsertCol', TransactionID, 'function', AddRange, AddRange, { row: AddRange.startRow, col: AddRange.startCol }, null, function () { ip_RemoveCol(GridID, { col: ColToSplice, count: options.count, mode: 'destroy', render: false, raiseEvent: false, createUndo: false }); });
        }


        //Move formulas - but exclude the range that we are moving
        var fxList = ip_ChangeFormulaOrigin(GridID, { colDiff: options.count, range: { startCol: ColToSplice } });



        //INSERT COLUMN
        //Clear out any merges that are directly on the col so that they can be adjusted and added back after the col insert has taken place
        var mergesCol = (options.appendTo == 'before' ? options.col - 1 : options.col + 1);
        var merges = ip_ValidateRangeMergedCells(GridID, 0, mergesCol, ip_GridProps[GridID].rows - 1, mergesCol);//Get merges that fall where we want to insert col                         
        ip_ReshuffelMergesCol(GridID, (options.appendTo == 'before' ? options.col : options.col + 1), ip_GridProps[GridID].cols - 1, options.count, false); //Move all other merges down - ignoring index alignment as this will happen naturally with the splice
            
            
        //Allow for inserting multiple columns
        for (var i = 0; i < options.count; i++) {

            //var ColToSplice = (options.appendTo == 'before' ? options.col : options.col + 1);
            var Quad = ip_GetQuad(GridID, -1, ColToSplice);
            var Quad2 = (Quad == 2 ? 4 : 3);

            //Add column to data object
            ip_GridProps[GridID].cols++;

            var ColObj = ip_colObject();
            ip_GridProps[GridID].colData.splice(ColToSplice, 0, ColObj);
            for (var row = 0; row < ip_GridProps[GridID].rows; row++) {
                var CellObj = ip_cellObject(null, null, null, row, ColToSplice);
                ip_GridProps[GridID].rowData[row].cells.splice(ColToSplice, 0, CellObj);
            }


        }

        //Bring back any merges that were directly on our col - making the adjustment for the removed col
        for (var m = 0; m < merges.merges.length; m++) {

            var MergeStartCol = merges.merges[m].mergedWithCol;
            var MergeEndCol = merges.merges[m].mergedWithCol + merges.merges[m].colSpan - 1;

            if (mergesCol != MergeStartCol && options.appendTo == 'after') { ip_ResetCellMerge(GridID, merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol, 0, options.count); ip_AddColumnToMerge(GridID, merges.merges[m], options.count); }
            else if (mergesCol != MergeEndCol && options.appendTo == 'before') { ip_ResetCellMerge(GridID, merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol, 0, options.count); ip_AddColumnToMerge(GridID, merges.merges[m], options.count); }

        }

        

        //Shift formula columns e.g. d0 = e0
        var excludeRangeChange = (ColToSplice == 0 ? null : ip_rangeObject(0, 0, ip_GridProps[GridID].rows - 1, ColToSplice - 1));
        for (var key in fxList) {

            
            var fx = ip_GridProps[GridID].indexedData.formulaData[key];
            ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: ip_MoveFormulaOrigon(GridID, ip_GridProps[GridID].rowData[fx.row].cells[fx.col].formula, fx.row, fx.col, fx.row, (fx.col + options.count), excludeRangeChange) });
            ip_AppendEffectedRowData(GridID, Effected, { row: fx.row, col: fx.col, formula: ip_GridProps[GridID].rowData[fx.row].cells[fx.col].formula });


        }
        

        //Commented out because instead of recalculating formulas, we actually change the formula in the above code
        //Recalculate formulas  
        //var range = ip_rangeObject(0, ColToSplice, ip_GridProps[GridID].rows - 1, ip_GridProps[GridID].cols - 1);
        //Effected.rowData = ip_ReCalculateFormulas(GridID, { range: [range], transactionID: TransactionID, render: false, raiseEvent: false, createUndo: false }).Effected.rowData; //{ startCol: ColToSplice, startRow: 0, endCol: ip_GridProps[GridID].cols - 1, endRow: ip_GridProps[GridID].rows - 1 }
                               
        //rerender
        if (options.render) {

            ip_ReRenderCols(GridID);
            $('#' + GridID).ip_Scrollable({ initY: false });

        }


        //Select cols
        $('#' + GridID).ip_SelectColumn({ col: ColToSplice, count: options.count, unselect: false }); 


        //Raise event
        if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_InsertCol', TransactionID, { InsertCol: { Inputs: options, Effected: Effected } });  }
    }
    else {

        //Add column onto end - not add col handles all the logisics independantly
        ip_AddCol(GridID, { count: options.count });  
                
    }

}

function ip_InsertRow(GridID, options) {

    var options = $.extend({

        row: ip_GridProps[GridID].rows, //By default append a row on to the end
        count: 1,
        appendTo: 'before', //'before' or 'after',
        render: true,
        raiseEvent: true,
        createUndo: true

    }, options);

    var TransactionID = ip_GenerateTransactionID();

    if (options.row == null && ip_GridProps[GridID].selectedRow.length > 0) { options.row = ip_GridProps[GridID].selectedRow[0]; }
    else if (options.row == null) { options.row = ip_GridProps[GridID].rows; }

    if (options.count == null && ip_GridProps[GridID].selectedRow.length > 0) { options.count = ip_GridProps[GridID].selectedRow.length; }
    else if (options.count == null) { options.count = 1; }

    if (options.row < ip_GridProps[GridID].rows) {

        //INSERT ROW
        var RowToSplice = (options.appendTo == 'before' ? options.row : options.row + 1);
        var Effected = { row: options.row, count: options.count, appendTo: options.appendTo, rowData: null };

        //Add to undostack    
        if (options.createUndo) {            
            var AddRange = { startRow: RowToSplice, startCol: 0, endRow: (RowToSplice + options.count - 1), endCol: (ip_GridProps[GridID].cols - 1) };
            var FunctionUndoData = ip_AddUndo(GridID, 'ip_InsertRow', TransactionID, 'function', AddRange, AddRange, { row: AddRange.startRow, col: AddRange.startCol }, null, function () { ip_RemoveRow(GridID, { row: RowToSplice, count: options.count, mode: 'destroy', render: false, raiseEvent: false, createUndo: false }); });
        }

        //Move formulas - but exclude the range that we are moving
        var fxList = ip_ChangeFormulaOrigin(GridID, { rowDiff: options.count, range: { startRow: RowToSplice } });


        //Get merges on row and shuffle them down
        var mergesRow = (options.appendTo == 'before' ? options.row - 1 : options.row + 1);
        var merges = ip_ValidateRangeMergedCells(GridID, mergesRow, 0, mergesRow, ip_GridProps[GridID].cols - 1);//Get merges that fall where we want to insert row                         
        ip_ReshuffelMergesRow(GridID, (options.appendTo == 'before' ? options.row : options.row + 1), ip_GridProps[GridID].rows - 1, options.count, false); //Move all other merges down - ignoring index alignment as this will happen naturally with the splice



        //Allow for inserting multiple columns
        for (var i = 0; i < options.count; i++) {

            var Quad = ip_GetQuad(GridID, RowToSplice, -1);
            var Quad2 = (Quad == 3 ? 4 : 2);

            //Add column to data object
            ip_GridProps[GridID].rows++;

            var RowObj = ip_rowObject(ip_GridProps[GridID].dimensions.defaultRowHeight, null, RowToSplice, ip_GridProps[GridID].cols);  //({ cols: ip_GridProps[GridID].cols, row: row, height: ip_GridProps[GridID].dimensions.defaultRowHeight });                
            ip_GridProps[GridID].rowData.splice(RowToSplice, 0, RowObj);

            //The code below is commented out because it is not handled on the server side
            //Increase the amount of frozen columns if the col to add is inside fozen columns
            //if (Quad == 1 && ip_GridProps[GridID].frozenRows > 0) {
            //    ip_GridProps[GridID].frozenRows++;
            //    ip_GridProps[GridID].scrollY++;
            //}
        }

        ////Bring back any merges that were directly on our row - making the adjustment for the removed row
        for (var m = 0; m < merges.merges.length; m++) {

            var MergeStartRow = merges.merges[m].mergedWithRow;
            var MergeEndRow = merges.merges[m].mergedWithRow + merges.merges[m].rowSpan - 1;

            if (mergesRow != MergeStartRow && options.appendTo == 'after') { ip_ResetCellMerge(GridID, merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol, options.count); ip_AddRowToMerge(GridID, merges.merges[m], options.count); }
            else if (mergesRow != MergeEndRow && options.appendTo == 'before') { ip_ResetCellMerge(GridID, merges.merges[m].mergedWithRow, merges.merges[m].mergedWithCol, options.count); ip_AddRowToMerge(GridID, merges.merges[m], options.count); }

        }

        var excludeRangeChange = (RowToSplice == 0 ? null : ip_rangeObject(0, 0, RowToSplice - 1, ip_GridProps[GridID].cols - 1));
        for (var key in fxList) {

            var fx = ip_GridProps[GridID].indexedData.formulaData[key];
            ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: ip_MoveFormulaOrigon(GridID, ip_GridProps[GridID].rowData[fx.row].cells[fx.col].formula, fx.row, fx.col, (fx.row + options.count), fx.col, excludeRangeChange) });
            ip_AppendEffectedRowData(GridID, Effected, { row: fx.row, col: fx.col, formula: ip_GridProps[GridID].rowData[fx.row].cells[fx.col].formula });

        }

        //Recalculate formulas                     
        //Effected.rowData = ip_ReCalculateFormulas(GridID, { range: [{ startRow: RowToSplice, startCol: 0, endRow: ip_GridProps[GridID].rows - 1, endCol: ip_GridProps[GridID].cols - 1 }], transactionID: TransactionID, render: false, raiseEvent: false, createUndo: false }).Effected.rowData;


        //ReRender the grid
        if (options.render) {
            ip_ReRenderRows(GridID, 'frozen');
            $('#' + GridID).ip_Scrollable();
            ip_ReRenderRows(GridID, 'scroll');
        }

        //Select rows
        $('#' + GridID).ip_SelectRow({ row: RowToSplice, count: options.count, unselect: false });

        //Raise event
        if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_InsertRow', TransactionID, { InsertRow: { Inputs: options, Effected: Effected } });  }

    }
    else {

        //Add column onto end - not add col handles all the logisics independantly            
        ip_AddRow(GridID, { count: options.count });
    }



}


//----- FROZEN ROWS/COLS ------------------------------------------------------------------------------------------------------------------------------------

function ip_FrozenDone(GridID, FrozenHandle) {

    //$('#' + GridID + '_columnFrozenHandleLine').hide();
    //$('#' + GridID + '_rowFrozenHandleLine').hide();
    //$(FrozenHandle).hide();
    ip_HideFloatingHandles(GridID);
    ip_GridProps[GridID].resizing = false;

}

function ip_ShowColumnFrozenHandle(GridID, AlignOnly) {

    ip_HideFloatingHandles(GridID);

    var FrozenHandle = $('#' + GridID + '_columnFrozenHandle');
    var col = (ip_GridProps[GridID].scrollX < ip_GridProps[GridID].frozenCols ? ip_GridProps[GridID].frozenCols : ip_GridProps[GridID].scrollX);
    var Quad1TableWidth = ip_TableWidth(GridID + '_q1_table');
    var Quad = 2;
    var pos = ip_CellRelativePosition(GridID, -1, col);
    var columnSelector = $('#' + GridID + '_q1_columnSelectorCell_-1'); // $('#' + GridID + '_q' + Quad + '_columnSelectorCell_' + col);
    var FrozenHandleLine = $('#' + GridID + '_columnFrozenHandleLine');

    if (columnSelector.length > 0) { //This will be 0 if we have no cols

        var gridTop = $('#' + GridID + '_table').position().top;
        var columnSelectorBorderWidth = parseInt($(columnSelector).css('border-left-width').replace('px', '')) + parseInt($(columnSelector).css('border-right-width').replace('px', ''));
        var positionLeft = Quad1TableWidth - (col == 0 ? $(FrozenHandle).width() : ($(FrozenHandle).width() / 2) + 1); //- columnSelectorBorderWidth //pos.localLeft + Quad1TableWidth - columnSelectorBorderWidth - (col == 0 ? $(FrozenHandle).width() : 1);

        $(FrozenHandle).show();
        $(FrozenHandle).css('top', (columnSelectorBorderWidth + gridTop));
        $(FrozenHandle).css('height', ($(columnSelector).height() - (columnSelectorBorderWidth / 2)) + 'px'); //IE patch
        $(FrozenHandle).css('left', positionLeft + 'px');

        $(FrozenHandleLine).css('height', ip_GridProps[GridID].dimensions.gridHeight + 'px');

        if (!AlignOnly) {


            $(FrozenHandle).draggable({
                snap: ".ip_grid_columnSelectorCell,.ip_grid_cell,.ip_grid_columnSelectorCellCorner,.ip_grid_rowSelecterCell",
                snapTolerance: 15,
                snapMode: "outer",
                axis: "x",
                containment: "parent",
                start: FrozenStart = function (event, ui) {

                    //if (startDragLeft == null) { startDragLeft = ui.position.left; }
                    ip_GridProps[GridID].resizing = true;
                    $(FrozenHandle).css('pointer-events', 'none');
                    $(FrozenHandleLine).show();

                },
                drag: FrozenDrag = function (event, ui) {


                },
                stop: FrozenEnd = function (event, ui) {

                    var snappedTo = $.map($(this).data('draggable').snapElements, function (element) {
                        var col = $(element.item).attr('col');
                        if (col && element.snapping) { return parseInt(col); }
                    });

                    var uniqueNumsIndex = {};
                    var uniqueNums = 0;
                    var col = null;
                    for (e in snappedTo) { if (col == null || col > snappedTo[e]) { col = snappedTo[e]; } }
                    if (col == null) { col = -1; }
                    col++; 

                    ip_FrozenDone(GridID, FrozenHandle);

                    $('#' + GridID).ip_FrozenRowsCols({ cols: col })
                    $(FrozenHandle).css('pointer-events', 'all');
                    ip_ShowColumnFrozenHandle(GridID, true);
                }

            });
        }
    }
}

function ip_ShowRowFrozenHandle(GridID, AlignOnly) {

    ip_HideFloatingHandles(GridID);

    var FrozenHandle = $('#' + GridID + '_rowFrozenHandle');
    var row = (ip_GridProps[GridID].scrollY < ip_GridProps[GridID].frozenRows ? ip_GridProps[GridID].frozenRows : ip_GridProps[GridID].scrollY);
    var Quad1TableHeight = $('#' + GridID + '_q1_table').outerHeight();
    var Quad = 3;
    var pos = ip_CellRelativePosition(GridID, row, -1);
    var rowSelector = $('#' + GridID + '_q1_columnSelectorCell_-1'); //$('#' + GridID + '_q' + Quad + '_rowSelecterCell_' + row); //MyGrid_q3_rowSelecterCell_3
    var FrozenHandleLine = $('#' + GridID + '_rowFrozenHandleLine');

    if (rowSelector.length > 0) { //This will be 0 if we have no rows

        var gridTop = $('#' + GridID + '_table').position().top;
        var rowSelectorBorderHeight = parseInt($(rowSelector).css('border-bottom-width').replace('px', ''));
        var positionTop = Quad1TableHeight + gridTop - (row == 0 || ip_GridProps[GridID].frozenRows == 0 ? $(FrozenHandle).height() : $(FrozenHandle).height() / 2);// is an ie patch

        $(FrozenHandle).show();

        $(FrozenHandle).css('left', rowSelectorBorderHeight); //0.5 is an ie patch
        $(FrozenHandle).css('width', ($(rowSelector).width()) + 'px');
        $(FrozenHandle).css('top', positionTop + 'px');
        $(FrozenHandleLine).css('width', ip_GridProps[GridID].dimensions.gridWidth + 'px');

        if (!AlignOnly) {

            $(FrozenHandle).draggable({
                snap: ".ip_grid_row,.ip_grid_columnSelectorRow",
                snapTolerance: 10,
                snapMode: "outer",
                axis: "y",
                containment: "parent",
                start: FrozenStart = function (event, ui) {

                    ip_GridProps[GridID].resizing = true;
                    $(FrozenHandle).css('pointer-events', 'none');
                    $(FrozenHandleLine).show();

                },
                drag: FrozenDrag = function (event, ui) {


                },
                stop: FrozenEnd = function (event, ui) {

                    var snappedTo = $.map($(this).data('draggable').snapElements, function (element) {
                        var row = $(element.item).attr('row');
                        if (row && element.snapping) { return parseInt(row); }
                    });

                    var uniqueNumsIndex = {};
                    var uniqueNums = 0;
                    var row = null;
                    for (e in snappedTo) { if (row == null || row > snappedTo[e]) { row = snappedTo[e]; }  }
                    if (row == null) { row = -1; }
                    else { row++; }
                    ip_FrozenDone(GridID, FrozenHandle);

                    $('#' + GridID).ip_FrozenRowsCols({ rows: row });
                    $(FrozenHandle).css('pointer-events', 'all');
                    ip_ShowRowFrozenHandle(GridID, true);

                }

            });
        }
    }
}

function ip_HideFloatingHandles(GridID) {


    $('#' + GridID + '_rowResizer').hide();
    $('#' + GridID + '_columnResizer').hide();
    $('#' + GridID + '_rowFrozenHandleLine').hide();
    $('#' + GridID + '_columnFrozenHandleLine').hide();

}


//----- RESIZE ROWS/COLS/GRID ------------------------------------------------------------------------------------------------------------------------------------

function ip_resizeDone(GridID, Resizer) {

    $('#' + GridID + '_columnLine').hide();
    $('#' + GridID + '_rowLine').hide();

    $(Resizer).hide();
    ip_GridProps[GridID].resizing = false;

}

function ip_ShowColumnResizer(GridID, col) {

    ip_HideFloatingHandles(GridID);

    var Resizer = $('#' + GridID + '_columnResizer');

    var Quad1TableWidth = ip_TableWidth(GridID + '_q1_table');
    var Quad = (col >= ip_GridProps[GridID].frozenCols ? 2 : 1);
    var pos = ip_CellRelativePosition(GridID, -1, col);
    var columnSelector = $('#' + GridID + '_q' + Quad + '_columnSelectorCell_' + col);
    var ResizerLine = $('#' + GridID + '_columnLine');
    var GridTop = $('#' + GridID + '_table').position().top;

    $(Resizer).show();

    var ResizerWidth = $(Resizer).width();
    var columnSelectorBorderWidth = parseInt($(columnSelector).css('border-left-width').replace('px', ''));
    var columnSelectorBorderHeight = parseInt($(columnSelector).css('border-top-width').replace('px', ''));

    var positionLeft = pos.localLeft + (Quad == 2 ? Quad1TableWidth : 0) + $(columnSelector).width() - ResizerWidth + columnSelectorBorderWidth + $(ResizerLine).position().left + 1;

    $(ResizerLine).css('height', ip_GridProps[GridID].dimensions.gridHeight + 'px');
    $(Resizer).css('top', (columnSelectorBorderWidth + GridTop));
    $(Resizer).css('height', ($(columnSelector).height() + (columnSelectorBorderHeight * 2)) + 'px');
    $(Resizer).css('left', positionLeft + 'px');


    ip_GridProps[GridID].events.colResizer_dblClick = ip_UnBindEvent(Resizer, 'dblclick', ip_GridProps[GridID].events.colResizer_dblClick);    
    $(Resizer).dblclick(ip_GridProps[GridID].events.colResizer_dblClick = function () {

        var columnSize = '';
        var columnsToResize = new Array();


        for (var i = 0; i < ip_GridProps[GridID].selectedColumn.length; i++) {
            columnsToResize[i] = ip_GridProps[GridID].selectedColumn[i];
        }

        if (jQuery.inArray(-1, columnsToResize) != -1) {

            columnsToResize = new Array();
            columnsToResize[0] = -1;

        }
        else if (jQuery.inArray(parseInt(col), columnsToResize) == -1) {

            columnsToResize = new Array();
            columnsToResize[0] = col;

        }


        $('#' + GridID).ip_ResizeColumn({ columns: columnsToResize, size: columnSize });

        ip_resizeDone(GridID, Resizer);

    });

    ip_GridProps[GridID].events.colResizer_MouseDown = ip_UnBindEvent(Resizer, 'mousedown', ip_GridProps[GridID].events.colResizer_MouseDown);    
    $(Resizer).mousedown(ip_GridProps[GridID].events.colResizer_MouseDown = function () {

        ip_GridProps[GridID].resizing = true;
        $(ResizerLine).show();

        ip_GridProps[GridID].events.colResizer_MouseUp = ip_UnBindEvent('#' + GridID, 'mouseup', ip_GridProps[GridID].events.colResizer_MouseUp);        
        $('#' + GridID).mouseup(ip_GridProps[GridID].events.colResizer_MouseUp = function () {

            ip_GridProps[GridID].resizing = false;

        });

    });
    
    ip_GridProps[GridID].events.colResizer_MouseLeave = ip_UnBindEvent(Resizer, 'mouseleave', ip_GridProps[GridID].events.colResizer_MouseLeave);    
    $(Resizer).mouseleave(ip_GridProps[GridID].events.colResizer_MouseLeave = function () {
        if (!ip_GridProps[GridID].resizing) {

            ip_resizeDone(GridID, Resizer);

        }
    });
    
    var minDrag = pos.localLeft + (Quad == 2 ? Quad1TableWidth : 0);// + ResizerWidth - columnSelectorBorderWidth;
    var startWidth = $(columnSelector).width();
    var startDragLeft = null;

    $(Resizer).draggable({
        axis: "x",
        containment: "parent",
        start: resizerStart = function (event, ui) {

            if (startDragLeft == null) { startDragLeft = ui.position.left; }
            ip_GridProps[GridID].resizing = true;


        },
        drag: resizerDrag = function (event, ui) {


            if (ui.position.left < minDrag) { return false; }

        },
        stop: resizerEnd = function (event, ui) {

            var columnSize = (startWidth + ((ui.position.left < minDrag ? minDrag : ui.position.left) - startDragLeft));
            var columnsToResize = new Array();

            for (var i = 0; i < ip_GridProps[GridID].selectedColumn.length; i++) {
                columnsToResize[i] = parseInt(ip_GridProps[GridID].selectedColumn[i]);
            }

            if (jQuery.inArray(-1, columnsToResize) != -1) {

                columnsToResize = new Array();
                columnsToResize[0] = -1;

            }
            else if (jQuery.inArray(parseInt(col), columnsToResize) == -1) {

                columnsToResize = new Array();
                columnsToResize[0] = col;

            }

            $('#' + GridID).ip_ResizeColumn({ columns: columnsToResize, size: columnSize });

            ip_resizeDone(GridID, Resizer);


        }

    });


}

function ip_ShowRowResizer(GridID, row) {

    ip_HideFloatingHandles(GridID);

    var Resizer = $('#' + GridID + '_rowResizer');

    var Quad1TableHeight = $('#' + GridID + '_q1_table').outerHeight();
    var pos = ip_CellRelativePosition(GridID, row, -1);
    var Quad = (row >= ip_GridProps[GridID].frozenRows ? 3 : 1);
    var rowSelector = $('#' + GridID + '_q' + Quad + '_rowSelecterCell_' + row);

    var ResizerLine = $('#' + GridID + '_rowLine');

    var ResizerHeight = $(Resizer).height();
    var rowSelectorBorderWidth = parseInt($(rowSelector).css('border-bottom-width').replace('px', ''));

    var cellHeight = $(rowSelector).height();
    var cellOffsetTop = $(rowSelector).offset().top;
    var resizerOffsetTop = cellOffsetTop + cellHeight - ResizerHeight;
    

    $(Resizer).show();
    $(Resizer).offset({ top: resizerOffsetTop });
    $(Resizer).css('left', rowSelectorBorderWidth);
    $(Resizer).css('width', ($(rowSelector).width() + rowSelectorBorderWidth) + 'px');

    $(ResizerLine).css('width', ip_GridProps[GridID].dimensions.gridWidth + 'px');

    ip_GridProps[GridID].events.rowResizer_dblClick = ip_UnBindEvent(Resizer, 'dblclick', ip_GridProps[GridID].events.rowResizer_dblClick);    
    $(Resizer).dblclick(ip_GridProps[GridID].events.rowResizer_dblClick = function () {

        var rowSize = '';
        var rowsToResize = new Array();


        for (var i = 0; i < ip_GridProps[GridID].selectedRow.length; i++) {
            rowsToResize[i] = ip_GridProps[GridID].selectedRow[i];
        }

        if (jQuery.inArray(-1, rowsToResize) != -1) {

            rowsToResize = new Array();
            rowsToResize[0] = -1;

        }
        else if (jQuery.inArray(parseInt(row), rowsToResize) == -1) {

            rowsToResize = new Array();
            rowsToResize[0] = row;

        }


        $('#' + GridID).ip_ResizeRow({ rows: rowsToResize, size: rowSize });

        ip_resizeDone(GridID, Resizer);

    });

    ip_GridProps[GridID].events.rowResizer_MouseDown = ip_UnBindEvent(Resizer, 'mousedown', ip_GridProps[GridID].events.rowResizer_MouseDown);    
    $(Resizer).mousedown(ip_GridProps[GridID].events.rowResizer_MouseDown = function () {

        ip_GridProps[GridID].resizing = true;
        $(ResizerLine).show();

        ip_GridProps[GridID].events.rowResizer_MouseUp = ip_UnBindEvent('#' + GridID, 'mouseup', ip_GridProps[GridID].events.rowResizer_MouseUp);        
        $('#' + GridID).mouseup(ip_GridProps[GridID].events.rowResizer_MouseUp = function () {

            ip_GridProps[GridID].resizing = false;
            //ip_resizeDone(GridID, Resizer); //COMMENTED OUT BECAUSE IT CANCELS DOUBLE CLICK EVENT

        });

    });

    ip_GridProps[GridID].events.rowResizer_MouseLeave = ip_UnBindEvent(Resizer, 'mouseleave', ip_GridProps[GridID].events.rowResizer_MouseLeave);    
    $(Resizer).mouseleave(ip_GridProps[GridID].events.rowResizer_MouseLeave = function () {
        if (!ip_GridProps[GridID].resizing) {

            ip_resizeDone(GridID, Resizer);

        }
    });


    var minDrag = pos.localTop + (Quad == 3 ? Quad1TableHeight : 0) - ResizerHeight + 1; //+ ResizerHeight - rowSelectorBorderWidth;
    var startHeight = $(rowSelector).height();
    var startDragTop = null;

    $(Resizer).draggable({

        axis: "y",
        containment: "parent",
        start: resizerStart = function (event, ui) {

            if (startDragTop == null) { startDragTop = ui.position.top; }
            ip_GridProps[GridID].resizing = true;


        },
        drag: resizerDrag = function (event, ui) {

            if (ui.position.top < minDrag) { return false; }

        },
        stop: resizerEnd = function (event, ui) {

            var rowSize = (startHeight + ((ui.position.top < minDrag ? minDrag : ui.position.top) - startDragTop));
            var rowsToResize = new Array();

            for (var i = 0; i < ip_GridProps[GridID].selectedRow.length; i++) {
                rowsToResize[i] = parseInt(ip_GridProps[GridID].selectedRow[i]);
            }

            if (jQuery.inArray(-1, rowsToResize) != -1) {
                rowsToResize = new Array();
                rowsToResize[0] = -1;
            }
            else if (jQuery.inArray(parseInt(row), rowsToResize) == -1) {
                rowsToResize = new Array();
                rowsToResize[0] = row;
            }

            $('#' + GridID).ip_ResizeRow({ rows: rowsToResize, size: rowSize });

            //CONTINUE FROM HERE:
            ip_resizeDone(GridID, Resizer);


        }

    });



}

function ip_ShowGridResizerHandle(GridID) {

    var Resizer = $("#" + GridID + "_gridResizer");
    var containment = $(document.body); //$('#' + GridID).parent();


    //Deals with: reisizing the grid
    $(Resizer).unbind('dblclick');
    $(Resizer).on('dblclick', function (e) {

        var parentWidth = $('#' + GridID).parent().width();
        var gridWidth = $('#' + GridID).width();
        var newWidth = (gridWidth < parentWidth ? parentWidth : Math.ceil(parentWidth / 2));
       
        ip_ResizeGrid(GridID, newWidth);

    });

    
    //'.panes.main'
    $(Resizer).draggable({
        //grid: [ 10, 10 ],
        snap: $(containment), snapMode: "outer", //$(Resizer).parent().attr('class').replace(' ', '.')
        axis: "x",
        containment: containment,
        start: resizerStart = function (event, ui) {

            ip_GridProps[GridID].resizing = true;
            $(Resizer).addClass('resizing');
            $("#" + GridID + "_gridResizerLine").show();

        },
        drag: resizerDrag = function (event, ui) {



        },
        stop: resizerEnd = function (event, ui) {

            $(Resizer).removeClass('resizing');

            var newWidth = ($(Resizer).width() + ui.position.left) - $('#' + GridID).position().left;

            if (newWidth > $(containment).width()) { newWidth = $(containment).width(); }

            var Error = ip_ResizeGrid(GridID, newWidth);

            ip_GridProps[GridID].resizing = false;

            if (Error != '') {

                $(Resizer).css('left', (ip_GridProps[GridID].dimensions.gridWidth - parseInt($('#' + GridID + '_q4_scrollbar_container_y').width())) + 'px');
                ip_RaiseEvent(GridID, 'warning', null, Error);

            }
        }

    });

}

function ip_ResizeGrid(GridID, newWidth, raiseEvent) {

    var Error = '';
    var FirstQuadTableWidth = parseInt($('#' + GridID + '_q1_table').outerWidth());

    if (raiseEvent == null) { raiseEvent = true; }
    if (newWidth == null) { newWidth = $('#' + GridID).width(); }

    $("#" + GridID + "_gridResizerLine").hide();

    if (newWidth > FirstQuadTableWidth) {

        
        ip_GridProps[GridID].dimensions.gridWidth = newWidth;        
        $('#' + GridID).width(newWidth);
        $('#' + GridID).ip_Scrollable();
        ip_ReRenderCols(GridID, 'scroll');
        
    }
    else { Error = 'Your grid size may not be smaller than frozen columns, please re-adjust the width so that it is beyond the frozen columns'; }



    if (Error == '') {

        if (raiseEvent) {
            var Effected = { width: newWidth }
            ip_RaiseEvent(GridID, 'ip_ResizeGrid', ip_GenerateTransactionID(), { ResizeGrid: { Inputs: null, Effected: Effected } });
        }
    }

    return Error;
}


//----- MOVE ROWS/COLS ------------------------------------------------------------------------------------------------------------------------------------

function ip_MoveColumn(GridID, options) {
    var options = $.extend({

        col: 0,  //Int, from col
        toCol: 0, //Int, to col
        count: 1,  //count
        prevalidated: null, //{ error: '', containsMerges: null, valid: true } ---- allows us to validate in a previous method and pass the results through for performance reasons
        render: true,
        raiseEvent: true,
        createUndo: true

    }, options);

    
    var col = options.col;
    var toCol = options.toCol;
    var count = options.count;
    var colDiff = toCol - col;
    var validateResult = (options.prevalidated == null ? ip_ValidateColumnMove(GridID, col, toCol, count) : options.prevalidated);
    var TransactionID = ip_GenerateTransactionID();
    var Effected = { col: col, toCol: toCol, count: count, rowData: null };
    var toRange = { startCol: toCol, endCol: toCol + count - 1, startRow: 0, endRow: ip_GridProps[GridID].rows - 1 };
    var moveRange = { startCol: col, endCol: col + count - 1, startRow: 0, endRow: ip_GridProps[GridID].rows - 1 };

    //Only alow the move if our range does not contain an overlap
    if (validateResult.valid) {

        if (col >= 0 && col < ip_GridProps[GridID].cols && toCol >= 0 && toCol < ip_GridProps[GridID].cols) {

            if (options.createUndo) {
                var MoveRange = { startRow: -1, startCol: col, endRow: (ip_GridProps[GridID].rows - 1), endCol: (col + count - 1) };
                var FunctionUndoData = ip_AddUndo(GridID, 'ip_MoveCol', TransactionID, 'function', MoveRange, MoveRange, { row: MoveRange.startRow, col: MoveRange.startCol }, null, function () { ip_MoveColumn(GridID, { col: toCol, toCol: col, count: count, render: false, raiseEvent: false, createUndo: false }); });
            }
                               

            //Remove merges
            for (var m = 0; m < validateResult.containsMerges.merges.length; m++) { ip_ResetCellMerge(GridID, validateResult.containsMerges.merges[m].mergedWithRow, validateResult.containsMerges.merges[m].mergedWithCol); }

            //Reshuffle exsting merges appropriatly
            if (colDiff > 0) { ip_ReshuffelMergesCol(GridID, col, toCol + count - 1, -count, false); } //Right
            else { ip_ReshuffelMergesCol(GridID, toCol, col, count, false); } //left   

            ////Move column object
            var moveColData = new Array();
            for (var i = 0; i < count; i++) { moveColData[i] = ip_GridProps[GridID].colData[col + i]; }
            ip_GridProps[GridID].colData.splice(col, count);
            for (var i = 0; i < moveColData.length; i++) { ip_GridProps[GridID].colData.splice(toCol + i, 0, moveColData[i]); }

            //Move formulas - but exclude the range that we are moving
            //ip_ChangeFormulaOrigin(GridID, { colDiff: (count * (colDiff < 0 ? 1 : -1)), range: { startCol: (col <= toCol ? col + 1 : toCol), endCol: (col >= toCol ? col - 1 : toCol) } });

            //--NEW CODE APPROACH            
            Effected.rowData = ip_ChangeFormulasForColumMove(GridID, { count: count, fromCol: col, toCol: toCol }).rowData;
            //--END OF NEW CODE

            //realign the columns for each row
            for (var row = 0; row < ip_GridProps[GridID].rows; row++) {

                var cells = ip_GridProps[GridID].rowData[row].cells;

                var moveData = new Array();
                for (var i = 0; i < count; i++) {
                    moveData[i] = cells[col + i];
                    //ip_ChangeFormulaOrigin(GridID, { fromCol: col + i, toCol: toCol + i, fromRow: row });
                }

                //First remove the cells we want to move as they are now stored in moveData
                cells.splice(col, count);
                
                //Next add moved columns
                for (var i = 0; i < moveData.length; i++) { cells.splice(toCol + i, 0, moveData[i]); }
                
            }

            ////Add new merges
            for (var n = 0; n < validateResult.containsMerges.merges.length; n++) {

                var startRow = validateResult.containsMerges.merges[n].mergedWithRow;
                var startCol = validateResult.containsMerges.merges[n].mergedWithCol + colDiff;
                var endRow = validateResult.containsMerges.merges[n].mergedWithRow + validateResult.containsMerges.merges[n].rowSpan - 1;
                var endCol = startCol + validateResult.containsMerges.merges[n].colSpan - 1;

                ip_SetCellMerge(GridID, startRow, startCol, endRow, endCol);
            }

            //Recalculate
            Effected.rowData = Effected.rowData.concat(ip_ReCalculateFormulas(GridID, { range: [moveRange, toRange], transactionID: TransactionID, render: false, raiseEvent: false, createUndo:false }).Effected.rowData);


            if (options.render) {
                
                ip_ReRenderCols(GridID, 'frozen');
                $('#' + GridID).ip_Scrollable();
                ip_ReRenderCols(GridID, 'scroll');
            }

            //Raise event    
            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_MoveCol', TransactionID, { MoveCol: { Inputs: options, Effected: Effected } });  }

            return true;

        }

    }


    ip_RaiseEvent(GridID, 'warning', arguments.callee.caller, validateResult.error);
    return false;
}

function ip_ChangeFormulasForColumMove(GridID, options) {
//Specialized method for changing the formulas when moving columns. It preempts index changes and sets them premove for move
    var options = $.extend({
        
        fromCol: null,
        toCol : null,
        count: 1,

    }, options);

    
    var Effected = { rowData: [] }
    var count = options.count;
    var fromCol = options.fromCol;
    var toCol = options.toCol;
    var colDiff = toCol - fromCol;
    var toRange = { startCol: toCol, endCol: toCol + count - 1, startRow: 0, endRow: ip_GridProps[GridID].rows - 1 };
    var moveRange = { startCol: fromCol, endCol: fromCol + count - 1, startRow: 0, endRow: ip_GridProps[GridID].rows - 1 };
    var shuffleRange;
    

    if (colDiff < 0) { shuffleRange = { startCol: toCol, endCol: fromCol - 1, startRow: 0, endRow: ip_GridProps[GridID].rows - 1 } }
    else { shuffleRange = { startCol: fromCol + count, endCol: toCol + count - 1, startRow: 0, endRow: ip_GridProps[GridID].rows - 1 } }

    var fxInMoveListAA = ip_FormulasInRange(GridID, moveRange, true, true);
    var fxInShuffleListBB = ip_FormulasInRange(GridID, shuffleRange, true, true);
    
    //MOVE LIST
    for (var key in fxInMoveListAA) {

        
        var fx = ip_GridProps[GridID].indexedData.formulaData[key];
        //if (!fx) { continue; }
        var formula = fx.formula;        
        var newFormula = "";

        newFormula = ip_MoveFormulaOrigon(GridID, formula, fx.row, fx.col, fx.row, fx.col + colDiff, null, moveRange, 0, (colDiff < 0 ? count : -count), shuffleRange, moveRange);


        var newFX = ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: newFormula, isRowColShuffel: true });

        //Move the indexes as they will be processed by the col move to a new position
        if (newFX.col >= moveRange.startCol && newFX.col <= moveRange.endCol) { newFX.col += colDiff; } //The "move" columns        
        else if (colDiff < 0 && newFX.col >= toCol && newFX.col <= fromCol) { newFX.col += count; }
        else if (colDiff > 0 && newFX.col <= toCol && newFX.col >= fromCol) { newFX.col -= count; }

     
        if (formula != newFormula) { ip_AppendEffectedRowData(GridID, Effected, { row: newFX.row, col: newFX.col, formula: newFormula }); }
    }

    //SHUFFLe LIST
    for (var key in fxInShuffleListBB) {

        if (fxInMoveListAA[key] == null && ip_GridProps[GridID].indexedData.formulaData[key] != null) {

            
            var fx = ip_GridProps[GridID].indexedData.formulaData[key];
            var formula = fx.formula;
            var newFormula = ip_MoveFormulaOrigon(GridID, formula, fx.row, fx.col, fx.row, fx.col + (colDiff < 0 ? count : -count), null, shuffleRange, null, null, null, shuffleRange);
            var newFX = ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: newFormula });

            if (colDiff > 0 && newFX.col >= shuffleRange.startCol && newFX.col <= shuffleRange.endCol) { newFX.col -= count; } //Move LEFT
            else if (colDiff < 0 && newFX.col >= shuffleRange.startCol && newFX.col <= shuffleRange.endCol) { newFX.col += count; } //Move LEFT

            if (formula != newFormula) { ip_AppendEffectedRowData(GridID, Effected, { row: newFX.row, col: newFX.col, formula: newFormula }); }
        }
    }

    return Effected;
}

function ip_MoveRow(GridID, options) {

    var options = $.extend({

        row: 0,
        toRow: 0,
        count: 1,
        revalidated: null, //{ error: '', containsMerges: null, valid: true } ---- allows us to validate in a previous method and pass the results through for performance reasons
        render: true,
        raiseEvent: true,
        createUndo: true

    }, options);
    
    var row = options.row;
    var toRow = options.toRow;
    var count = options.count;
    var rowDiff = toRow - row;
    var validateResult = (options.prevalidated == null ? ip_ValidateRowMove(GridID, row, toRow, count) : options.prevalidated);
    var TransactionID = ip_GeneratePublicKey();
    var Effected = { row: row, toRow: toRow, count: count, rowData:null };
    var moveRange = { startRow: row, endRow: row + count - 1, startCol: 0, endCol: ip_GridProps[GridID].cols - 1 };
    var toRange = { startRow: toRow, endRow: toRow + count - 1, startCol: 0, endCol: ip_GridProps[GridID].cols - 1 };

    //Only alow the move if our range does not contain an overlap
    if (validateResult.valid) {

        if (row >= 0 && row < ip_GridProps[GridID].rows && toRow >= 0 && toRow < ip_GridProps[GridID].rows) {

            if (options.createUndo) {
                var MoveRange = { startRow: row, startCol: -1, endRow: (row + count - 1), endCol: ip_GridProps[GridID].cols - 1 };
                var FunctionUndoData = ip_AddUndo(GridID, 'ip_MoveRow', TransactionID, 'function', MoveRange, MoveRange, { row: MoveRange.startRow, col: MoveRange.startCol }, null, function () { ip_MoveRow(GridID, { row: toRow, toRow: row, count: count, render: false, raiseEvent: false, createUndo: false }); });
            }

            //Move formulas - but exclude the range that we are moving
            //ip_ChangeFormulaOrigin(GridID, { rowDiff: (count * (rowDiff < 0 ? 1 : -1)), range: { startRow: (row <= toRow ? row + 1 : toRow), endRow: (row >= toRow ? row - 1 : toRow) } });
            Effected.rowData = ip_ChangeFormulasForRowMove(GridID, { count: count, fromRow: row, toRow: toRow }).rowData;

            //Remove merges
            for (var m = 0; m < validateResult.containsMerges.merges.length; m++) { ip_ResetCellMerge(GridID, validateResult.containsMerges.merges[m].mergedWithRow, validateResult.containsMerges.merges[m].mergedWithCol); }

            //Reshuffle exsting merges appropriatly
            if (rowDiff > 0) { ip_ReshuffelMergesRow(GridID, row, toRow + count - 1, -count, false); } //down
            else { ip_ReshuffelMergesRow(GridID, toRow, row, count, false); } //up   

            //First store rows we want to move
            var moveData = new Array();
            for (var r = 0; r < count; r++) { moveData[r] = ip_GridProps[GridID].rowData[row + r]; }

            //Remove rows we have stored
            ip_GridProps[GridID].rowData.splice(row, count);

            //Add rows we have stored in moveData to correct location
            for (var r = 0; r < count; r++) {
                ip_GridProps[GridID].rowData.splice(toRow + r, 0, moveData[r]);
                //for (var c = 0; c < ip_GridProps[GridID].cols; c++) { ip_ChangeFormulaOrigin(GridID, { toRow: (toRow + r), toCol: c }); }
            }

            ////Add new merges
            for (var n = 0; n < validateResult.containsMerges.merges.length; n++) {
                var startRow = validateResult.containsMerges.merges[n].mergedWithRow + rowDiff;
                var startCol = validateResult.containsMerges.merges[n].mergedWithCol;
                var endRow = startRow + validateResult.containsMerges.merges[n].rowSpan - 1;
                var endCol = validateResult.containsMerges.merges[n].mergedWithCol + validateResult.containsMerges.merges[n].colSpan - 1;

                ip_SetCellMerge(GridID, startRow, startCol, endRow, endCol);
            }
   
            //Recalculate
            Effected.rowData = Effected.rowData.concat(ip_ReCalculateFormulas(GridID, { range: [moveRange, toRange], transactionID: TransactionID, render: false, raiseEvent: false, createUndo:false }).Effected.rowData);
   
            if (options.render) {
                ip_ReRenderRows(GridID, 'frozen');
                $('#' + GridID).ip_Scrollable();
                ip_ReRenderRows(GridID, 'scroll');
            }

            //Raise event  
            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_MoveRow', TransactionID, { MoveRow: { Inputs: options, Effected: Effected } });    }

            return true;
        }

    }


    ip_RaiseEvent(GridID, 'warning', arguments.callee.caller, validateResult.error);
    return false;

}

function ip_ChangeFormulasForRowMove(GridID, options) {
    //Specialized method for changing the formulas when moving columns. It preempts index changes and sets them premove for move
    var options = $.extend({

        fromRow: null,
        toRow: null,
        count: 1,

    }, options);


    var Effected = { rowData: [] }
    var count = options.count;
    var fromRow = options.fromRow;
    var toRow = options.toRow;
    var rowDiff = toRow - fromRow;
    var toRange = { startRow: toRow, endRow: toRow + count - 1, startCol: 0, endCol: ip_GridProps[GridID].cols - 1 };
    var moveRange = { startRow: fromRow, endRow: fromRow + count - 1, startCol: 0, endCol: ip_GridProps[GridID].cols - 1 };
    var shuffleRange;


    if (rowDiff < 0) { shuffleRange = { startRow: toRow, endRow: fromRow - 1, startCol: 0, endCol: ip_GridProps[GridID].cols - 1 } }
    else { shuffleRange = { startRow: fromRow + count, endRow: toRow + count - 1, startCol: 0, endCol: ip_GridProps[GridID].cols - 1 } }

    var fxInMoveList = ip_FormulasInRange(GridID, moveRange, true, true);
    var fxInShuffleList = ip_FormulasInRange(GridID, shuffleRange, true, true);


    //MOVE LIST
    for (var key in fxInMoveList) {


        var fx = ip_GridProps[GridID].indexedData.formulaData[key];
        var formula = fx.formula;
        var newFormula = "";

        newFormula = ip_MoveFormulaOrigon(GridID, formula, fx.row, fx.col, fx.row + rowDiff, fx.col, null, moveRange, (rowDiff < 0 ? count : -count), 0,shuffleRange, moveRange);
    
        var newFX = ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: newFormula, isRowColShuffel:true });

        //Move the indexes as they will be processed by the row move to a new position
        if (newFX.row >= moveRange.startRow && newFX.row <= moveRange.endRow) { newFX.row += rowDiff; } //The "move" columns            
        else if (rowDiff < 0 && newFX.row >= toRow && newFX.row <= fromRow) { newFX.row += count; }
        else if (rowDiff > 0 && newFX.row <= toRow && newFX.row >= fromRow) { newFX.row -= count; }
        
        if (formula != newFormula) { ip_AppendEffectedRowData(GridID, Effected, { row: newFX.row, col: newFX.col, formula: newFormula }); }
    }

    //SHUFFLe LIST
    for (var key in fxInShuffleList) {

        if (fxInMoveList[key] == null && ip_GridProps[GridID].indexedData.formulaData[key] != null) {


            var fx = ip_GridProps[GridID].indexedData.formulaData[key];
            var formula = fx.formula;
            var newFormula = ip_MoveFormulaOrigon(GridID, formula, fx.row, fx.col, fx.row + (rowDiff < 0 ? count : -count), fx.col, null, shuffleRange, null, null, null, shuffleRange);
            var newFX = ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: newFormula });

            if (rowDiff > 0 && newFX.row >= shuffleRange.startRow && newFX.row <= shuffleRange.endRow) { newFX.row -= count; } //Move LEFT
            else if (rowDiff < 0 && newFX.row >= shuffleRange.startRow && newFX.row <= shuffleRange.endRow) { newFX.row += count; } //Move LEFT

            if (formula != newFormula) { ip_AppendEffectedRowData(GridID, Effected, { row: newFX.row, col: newFX.col, formula: newFormula }); }
        }
    }

    return Effected;
}

function ip_ShowColumnMove(GridID, col) {

    ip_GridProps[GridID].resizing = true;

    ip_GridProps[GridID].selectedColumn.sort(function (a, b) { return a - b });

    $('#' + GridID).ip_RemoveRangeHighlight();

    var RangeObjects = $('#' + GridID + " .ip_grid_cell_rangeselector_selected");
    $(RangeObjects).addClass('ip_grid_cell_rangeselector_move');
    $(RangeObjects).each(function () { this.startPos = $(this).position().left; });
        


    //Create A move Startegy
    var MoveStrategy = [];
    var prevC = -2;
    for (var c = 0; c < ip_GridProps[GridID].selectedColumn.length; c++) {

        var cc = ip_GridProps[GridID].selectedColumn[c];

        if (cc - prevC != 1) { MoveStrategy[MoveStrategy.length] = { col: cc, count: 1, toCol: cc } }
        else { MoveStrategy[MoveStrategy.length - 1].count++; }

        prevC = cc;

    }


    //Load ONLY column ranges
    var MouseCursorStartX = -1;
    var OldMoveAmt = 0;
    $('#' + GridID).mousemove(ip_GridProps[GridID].events.moveRange_mouseMove = function (e) {

        

        if (MouseCursorStartX == -1) { MouseCursorStartX = e.pageX; }
        var MouseMoveX = e.pageX - MouseCursorStartX;
        $(RangeObjects).each(function () { $(this).css('left', (this.startPos + MouseMoveX) + 'px'); });

        //Highlight the move
        ip_SetHoverCell(GridID, '.ip_grid_cell_rangeselector_move', e);
        var endCol = parseInt($(ip_GridProps[GridID].hoverCell).attr('col')); //ip_GridProps[GridID].hoverColumnIndex;
        var moveAmt = endCol - col;

        if (OldMoveAmt != moveAmt) {
            for (var m = 0; m < MoveStrategy.length; m++) { $('#' + GridID).ip_RangeHighlight({ multiselect: (m == 0 ? false : true), highlightType: 'ip_grid_cell_rangeHighlight_move', range: { startRow: 0, startCol: (MoveStrategy[m].col + moveAmt), endRow: ip_GridProps[GridID].rows - 1, endCol: (MoveStrategy[m].col + moveAmt + MoveStrategy[m].count - 1) } }); }
            OldMoveAmt = moveAmt;
        }

    });

    //Do column move
    $(document).mouseup(ip_GridProps[GridID].events.moveRange_mouseUp = function (e) {

        ip_GridProps[GridID].resizing = false;

        ip_SetHoverCell(GridID, '.ip_grid_cell_rangeselector_move', e); //Fix used for IE

        ip_GridProps[GridID].events.moveRange_mouseUp = ip_UnBindEvent(document, 'mouseup', ip_GridProps[GridID].events.moveRange_mouseUp);
        ip_GridProps[GridID].events.moveRange_mouseMove = ip_UnBindEvent('#' + GridID, 'mousemove', ip_GridProps[GridID].events.moveRange_mouseMove);

        $('#' + GridID).ip_RemoveRangeHighlight();

        var endCol = parseInt($(ip_GridProps[GridID].hoverCell).attr('col')); //ip_GridProps[GridID].hoverColumnIndex;
        var moveAmt = endCol - col;
        var error = '';


        if (moveAmt != 0) {


            //Update move strategy
            for (var m = 0; m < MoveStrategy.length; m++) { MoveStrategy[m].toCol = MoveStrategy[m].col + moveAmt; }

            //Validate the move
            for (var m = 0; m < MoveStrategy.length; m++) {
                MoveStrategy[m].prevalidated = ip_ValidateColumnMove(GridID, MoveStrategy[m].col, MoveStrategy[m].toCol, MoveStrategy[m].count);
                var error = MoveStrategy[m].prevalidated.error;
                if (error != '') { m = MoveStrategy.length; }
            }

            if (error == '') {

                //Do the actaul move
                if (moveAmt > 0) { for (var m = MoveStrategy.length - 1; m >= 0; m--) { ip_MoveColumn(GridID, MoveStrategy[m]); } }
                else { for (var m = 0; m < MoveStrategy.length; m++) { ip_MoveColumn(GridID, MoveStrategy[m]); } }

                //Select the columns we've moved
                for (var m = 0; m < MoveStrategy.length; m++) { $('#' + GridID).ip_SelectColumn({ multiselect: (m == 0 ? false : true), unselect: false, col: MoveStrategy[m].toCol, count: MoveStrategy[m].count }); }

            }
            else {

                //ReSelect the columns
                for (var m = 0; m < MoveStrategy.length; m++) { $('#' + GridID).ip_SelectColumn({ multiselect: (m == 0 ? false : true), col: MoveStrategy[m].col, count: MoveStrategy[m].count }); }
                ip_RaiseEvent(GridID, 'warning', null, error);
            }

        }
        else {

            //Reselect columns
            for (var m = 0; m < MoveStrategy.length; m++) { $('#' + GridID).ip_SelectColumn({ multiselect: (m == 0 ? false : true), col: MoveStrategy[m].col, count: MoveStrategy[m].count }); }
        }

    });

}

function ip_ShowRowMove(GridID, row) {

    ip_GridProps[GridID].resizing = true;

    ip_GridProps[GridID].selectedRow.sort(function (a, b) { return a - b });

    $('#' + GridID).ip_RemoveRangeHighlight();

    var RangeObjects = $('#' + GridID + " .ip_grid_cell_rangeselector_selected");
    $(RangeObjects).addClass('ip_grid_cell_rangeselector_move');
    $(RangeObjects).each(function () { this.startPos = $(this).position().top; });


    //Create A move Startegy
    var MoveStrategy = [];
    var prevR = -2;
    for (var r = 0; r < ip_GridProps[GridID].selectedRow.length; r++) {

        var rr = ip_GridProps[GridID].selectedRow[r];

        if (rr - prevR != 1) { MoveStrategy[MoveStrategy.length] = { row: rr, count: 1, toRow: rr } }
        else { MoveStrategy[MoveStrategy.length - 1].count++; }

        prevR = rr;

    }


    //Load ONLY column ranges
    var MouseCursorStartY = -1;
    var OldMoveAmt = 0;
    $('#' + GridID).mousemove(ip_GridProps[GridID].events.moveRange_mouseMove = function (e) {

        if (MouseCursorStartY == -1) { MouseCursorStartY = e.pageY; }
        var MouseMoveY = e.pageY - MouseCursorStartY;
        $(RangeObjects).each(function () { $(this).css('top', (this.startPos + MouseMoveY) + 'px'); });

        //Highlight the move
        ip_SetHoverCell(GridID, '.ip_grid_cell_rangeselector_move', e);
        var endRow = parseInt($(ip_GridProps[GridID].hoverCell).attr('row')); //ip_GridProps[GridID].hoverColumnIndex;
        var moveAmt = endRow - row;

        if (OldMoveAmt != moveAmt) {
            for (var m = 0; m < MoveStrategy.length; m++) { $('#' + GridID).ip_RangeHighlight({ multiselect: (m == 0 ? false : true), highlightType: 'ip_grid_cell_rangeHighlight_move', range: { startRow: (MoveStrategy[m].row + moveAmt), startCol: 0, endRow: (MoveStrategy[m].row + moveAmt + MoveStrategy[m].count - 1), endCol: ip_GridProps[GridID].cols - 1 } }); }
            OldMoveAmt = moveAmt;
        }

    });

    //Do column move
    $(document).mouseup(ip_GridProps[GridID].events.moveRange_mouseUp = function (e) {

        ip_GridProps[GridID].resizing = false;

        ip_SetHoverCell(GridID, '.ip_grid_cell_rangeselector_move', e); //Fix used for IE

        ip_GridProps[GridID].events.moveRange_mouseUp = ip_UnBindEvent(document, 'mouseup', ip_GridProps[GridID].events.moveRange_mouseUp);
        ip_GridProps[GridID].events.moveRange_mouseMove = ip_UnBindEvent('#' + GridID, 'mousemove', ip_GridProps[GridID].events.moveRange_mouseMove);

        $('#' + GridID).ip_RemoveRangeHighlight();

        var endRow = parseInt($(ip_GridProps[GridID].hoverCell).attr('row')); //ip_GridProps[GridID].hoverColumnIndex;
        var moveAmt = endRow - row;
        var error = '';


        if (moveAmt != 0) {


            //Update move strategy
            for (var m = 0; m < MoveStrategy.length; m++) { MoveStrategy[m].toRow= MoveStrategy[m].row + moveAmt; }

            //Validate the move
            for (var m = 0; m < MoveStrategy.length; m++) {
                MoveStrategy[m].prevalidated = ip_ValidateRowMove(GridID, MoveStrategy[m].row, MoveStrategy[m].toRow, MoveStrategy[m].count);
                var error = MoveStrategy[m].prevalidated.error;
                if (error != '') { m = MoveStrategy.length; }
            }

            if (error == '') {

                //Do the actaul move
                if (moveAmt > 0) { for (var m = MoveStrategy.length - 1; m >= 0; m--) { ip_MoveRow(GridID, MoveStrategy[m]); } }
                else { for (var m = 0; m < MoveStrategy.length; m++) { ip_MoveRow(GridID, MoveStrategy[m]); } }

                //Select the columns we've moved
                for (var m = 0; m < MoveStrategy.length; m++) { $('#' + GridID).ip_SelectRow({ multiselect: (m == 0 ? false : true), unselect:false, row: MoveStrategy[m].toRow, count: MoveStrategy[m].count }); }

            }
            else {

                //ReSelect the columns
                for (var m = 0; m < MoveStrategy.length; m++) { $('#' + GridID).ip_SelectRow({ multiselect: (m == 0 ? false : true), row: MoveStrategy[m].row, count: MoveStrategy[m].count }); }
                ip_RaiseEvent(GridID, 'warning', null, error);
            }

        }
        else {

            //Reselect columns
            for (var m = 0; m < MoveStrategy.length; m++) { $('#' + GridID).ip_SelectRow({ multiselect: (m == 0 ? false : true), row: MoveStrategy[m].row, count: MoveStrategy[m].count }); }
        }

    });



}

function ip_ValidateColumnMove(GridID, col, toCol, count) {
    //Validates if the column merge may occur, also gets the merges that are in those columns

    var result = { error: '', containsMerges: null, valid: true }
    var colDiff = toCol - col;
    var error = '';

    if (toCol < 0 || toCol >= ip_GridProps[GridID].cols) { error = 'When moving columns, make sure the move stays within the range of the sheet'; }

    //Validate from merges 
    result.containsMerges = ip_ValidateRangeMergedCells(GridID, 0, col, ip_GridProps[GridID].rows - 1, col + count - 1);
    if (result.containsMerges.containsOverlap == true) { error = 'When moving columns that contain merges please take the full merge, or unmerge your cells' }
    for (var i = 0; i < result.containsMerges.merges.length; i++) {

        if (result.containsMerges.merges[i].containsOverlap) {
            var startRow = result.containsMerges.merges[i].mergedWithRow;
            var startCol = result.containsMerges.merges[i].mergedWithCol;
            var endRow = startRow + result.containsMerges.merges[i].rowSpan - 1;
            var endCol = startCol + result.containsMerges.merges[i].colSpan - 1;

            $('#' + GridID).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, range: { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol } });
        }
    }

    ////Validate to merges 
    var validateToMerges = ip_ValidateRangeMergedCells(GridID, 0, col + colDiff, ip_GridProps[GridID].rows - 1, col + colDiff + count - 1);
    validateToMerges.containsOverlap = false;
    for (var i = 0; i < validateToMerges.merges.length; i++) {

        if (validateToMerges.merges[i].containsOverlap) {

            var startRow = validateToMerges.merges[i].mergedWithRow;
            var startCol = validateToMerges.merges[i].mergedWithCol;
            var endRow = startRow + validateToMerges.merges[i].rowSpan - 1;
            var endCol = startCol + validateToMerges.merges[i].colSpan - 1;

            if (colDiff > 0 && endCol >= toCol + count) { // + count

                validateToMerges.containsOverlap = true;
                $('#' + GridID).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, range: { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol } });

            }
            else if (colDiff < 0 && startCol < toCol) {

                validateToMerges.containsOverlap = true;
                $('#' + GridID).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, range: { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol } });

            }
        }
    }
    if (validateToMerges.containsOverlap == true) { error += (error != '' ? ', you also c' : 'C') + 'annot move columns on top cells that contain merges' }

    result.error = error;
    result.valid = (error == '' ? true : false);

    return result;
}

function ip_ValidateRowMove(GridID, row, toRow, count) {
    //Validates if the column merge may occur, also gets the merges that are in those columns

    var result = { error: '', containsMerges: null, valid: true }
    var rowDiff = toRow - row;
    var error = '';

    if (toRow < 0 || toRow >= ip_GridProps[GridID].rows) { error = 'When moving rows, make sure the move stays within the range of the sheet'; }

    //Validate from merges 
    result.containsMerges = ip_ValidateRangeMergedCells(GridID, row, 0, row + count - 1, ip_GridProps[GridID].cols - 1);
    if (result.containsMerges.containsOverlap == true) { error = 'When moving rows that contain merges please take the full merge, or unmerge your cells' }
    for (var i = 0; i < result.containsMerges.merges.length; i++) {

        if (result.containsMerges.merges[i].containsOverlap) {
            var startRow = result.containsMerges.merges[i].mergedWithRow;
            var startCol = result.containsMerges.merges[i].mergedWithCol;
            var endRow = startRow + result.containsMerges.merges[i].rowSpan - 1;
            var endCol = startCol + result.containsMerges.merges[i].colSpan - 1;

            $('#' + GridID).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, range: { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol } });
        }
    }

    ////Validate to merges 
    var validateToMerges = ip_ValidateRangeMergedCells(GridID, row + rowDiff, 0, row + rowDiff + count - 1, ip_GridProps[GridID].cols - 1);
    validateToMerges.containsOverlap = false;
    for (var i = 0; i < validateToMerges.merges.length; i++) {

        if (validateToMerges.merges[i].containsOverlap) {

            var startRow = validateToMerges.merges[i].mergedWithRow;
            var startCol = validateToMerges.merges[i].mergedWithCol;
            var endRow = startRow + validateToMerges.merges[i].rowSpan - 1;
            var endCol = startCol + validateToMerges.merges[i].colSpan - 1;

            if (rowDiff > 0 && endRow >= toRow + count) { // + count

                validateToMerges.containsOverlap = true;
                $('#' + GridID).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, range: { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol } });

            }
            else if (rowDiff < 0 && startRow < toRow) {

                validateToMerges.containsOverlap = true;
                $('#' + GridID).ip_RangeHighlight({ fadeOut: true, expireTimeout: 3000, highlightType: 'ip_grid_cell_rangeHighlight_alert', multiselect: true, range: { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol } });

            }
        }
    }
    if (validateToMerges.containsOverlap == true) { error += (error != '' ? ', you also c' : 'C') + 'annot move rows on top cells that contain merges' }

    result.error = error;
    result.valid = (error == '' ? true : false);

    return result;
}


//----- HIDDEN ROWS/COLS ------------------------------------------------------------------------------------------------------------------------------------

function ip_NextNonHiddenRow(GridID, startRow, endRow, defaultRow, direction, scrollArea) {

    if (startRow < 0) { startRow = 0; }
    if (scrollArea && startRow < ip_GridProps[GridID].scrollY) { startRow = ip_GridProps[GridID].scrollY; }
    if (startRow >= ip_GridProps[GridID].rows) { return defaultRow; }
    if (!ip_GridProps[GridID].rowData[startRow].hide) { return startRow; }

    var currentRow = startRow;

    if (direction == 'down') {

        if (endRow == null) { endRow = ip_GridProps[GridID].rows - 1; }

        while (currentRow <= endRow) {

            if (!ip_GridProps[GridID].rowData[currentRow].hide) { return currentRow; }
            currentRow++;

        }
    }
    else {

        if (endRow == null) { endRow = 0; }

        while (currentRow >= endRow) {

            if (!ip_GridProps[GridID].rowData[currentRow].hide) { return currentRow; }
            currentRow--;

        }
    }

    return defaultRow;
}

function ip_NextNonHiddenCol(GridID, startCol, endCol, defaultCol, direction, scrollArea) {

    if (startCol < 0) { startCol = 0; }
    if (scrollArea && startCol < ip_GridProps[GridID].scrollX) { startCol = ip_GridProps[GridID].scrollX; }
    if (!ip_GridProps[GridID].colData[startCol].hide) { return startCol; }

    var currentCol = startCol;

    if (direction == 'right') {

        if (endCol == null) { endCol = ip_GridProps[GridID].cols - 1; }

        while (currentCol <= endCol) {

            if (!ip_GridProps[GridID].colData[currentCol].hide) { return currentCol; }
            currentCol++;

        }
    }
    else {

        if (endCol == null) { endCol = 0; }

        while (currentCol >= endCol) {

            if (!ip_GridProps[GridID].colData[currentCol].hide) { return currentCol; }
            currentCol--;

        }
    }

    return defaultCol;
}

function ip_CalculateHiddenRowSpan(GridID, startRow, endRow) {

    var rowSpan = 0;

    for (var r = startRow; r <= endRow; r++)
    {
        if (!ip_GridProps[GridID].rowData[r].hide) { rowSpan++; }
    }

    return rowSpan;
}

function ip_CalculateHiddenColSpan(GridID, startCol, endCol) {

    var colSpan = 0;

    for (var c = startCol; c <= endCol; c++) {
        if (!ip_GridProps[GridID].colData[c].hide) { colSpan++; }
    }

    return colSpan;
}


//----- MERGE WITH ROWS/COLS ------------------------------------------------------------------------------------------------------------------------------------

function ip_RemoveAllMerges(GridID) {
    
    for (key in ip_GridProps[GridID].mergeData) {
        var startRow = ip_GridProps[GridID].mergeData[key].mergedWithRow;
        var startCol = ip_GridProps[GridID].mergeData[key].mergedWithCol;
        var endRow = startRow + ip_GridProps[GridID].mergeData[key].rowSpan - 1;
        var endCol = startCol + ip_GridProps[GridID].mergeData[key].colSpan - 1;
        ip_ResetMerge(GridID, startRow, startCol, endRow, endCol);
    }

}

function ip_ResetCellMerge(GridID, startRow, startCol, rowIndexSweepDepth, colIndexSweepDepth) {
    //Removes merge - but validates if merge exists for that cell first

    if (rowIndexSweepDepth == null) { rowIndexSweepDepth = 0; }
    if (colIndexSweepDepth == null) { colIndexSweepDepth = 0; }

    var merge = ip_GridProps[GridID].rowData[startRow].cells[startCol].merge;


    if (merge != null) {

        //Get the root merge
        merge = ip_GridProps[GridID].rowData[merge.mergedWithRow].cells[merge.mergedWithCol].merge;

        var endRow = merge.mergedWithRow + merge.rowSpan - 1 + rowIndexSweepDepth;
        var endCol = merge.mergedWithCol + merge.colSpan - 1 + colIndexSweepDepth;

        startRow = merge.mergedWithRow;
        startCol = merge.mergedWithCol;

        ip_ResetMerge(GridID, startRow, startCol, endRow, endCol);

    }

}

function ip_ResetMerge(GridID, startRow, startCol, endRow, endCol) {
    
    //Removes merge - without validating if merge exists
    delete ip_GridProps[GridID].mergeData[startRow + '-' + startCol];
    

    for (var r = startRow; r <= endRow; r++) {

        for (var c = startCol; c <= endCol; c++) {

            if (ip_GridProps[GridID].rowData[r].cells[c].merge != null) { delete ip_GridProps[GridID].rowData[r].cells[c].merge; }

            //Clear out the indexed merge for row
            if (ip_GridProps[GridID].colData[c].containsMerges != null) {
                for (var ci = ip_GridProps[GridID].colData[c].containsMerges.length - 1; ci >= 0 ; ci--) {

                    var indexedMerge = ip_GridProps[GridID].colData[c].containsMerges[ci];
                    if (indexedMerge.mergedWithRow == startRow && indexedMerge.mergedWithCol == startCol) {

                        //Remove index
                        ip_GridProps[GridID].colData[c].containsMerges.splice(ci, 1);
                        ci = 0;

                    }

                }

                if (ip_GridProps[GridID].colData[c].containsMerges.length == 0) { ip_GridProps[GridID].colData[c].containsMerges = null; }
            }
        }

        //Clear out the indexed mergeds for row
        if (ip_GridProps[GridID].rowData[r].containsMerges != null) {
            for (var ri = ip_GridProps[GridID].rowData[r].containsMerges.length - 1; ri >= 0 ; ri--) {

                var indexedMerge = ip_GridProps[GridID].rowData[r].containsMerges[ri];
                if (indexedMerge.mergedWithRow == startRow && indexedMerge.mergedWithCol == startCol) {
                    //Remove index
                    ip_GridProps[GridID].rowData[r].containsMerges.splice(ri, 1);
                    ri = 0;
                }

            }

            if (ip_GridProps[GridID].rowData[r].containsMerges.length == 0) { ip_GridProps[GridID].rowData[r].containsMerges = null; }
        }
    }

}

function ip_SetCellMerge(GridID, rngStartRow, rngStartCol, rngEndRow, rngEndCol, CellUndoData, MergeUndoData) {

    var arrRange = ip_ValidateFrozenMerge(GridID, rngStartRow, rngStartCol, rngEndRow, rngEndCol);

    for (var rng = 0; rng < arrRange.length; rng++) {

        var startRow = arrRange[rng].startRow;
        var startCol = arrRange[rng].startCol;
        var endRow = arrRange[rng].endRow;
        var endCol = arrRange[rng].endCol;

        var rCount = 0;
        var MergeRow = startRow;
        var MergeCol = startCol;
        var RowSpan = endRow - startRow + 1;
        var ColSpan = endCol - startCol + 1;
        var merge = null;

        //validate range
        if (startRow < 0) { startRow = 0; }
        if (startCol < 0) { startCol = 0; }
        if (endRow >= ip_GridProps[GridID].rows) { endRow = ip_GridProps[GridID].rows - 1; }
        if (endCol >= ip_GridProps[GridID].cols) { endCol = ip_GridProps[GridID].cols - 1; }
        if (endRow < 0) { return false; }
        if (endCol < 0) { return false; }
        if (startRow >= ip_GridProps[GridID].rows) { return false; }
        if (startCol >= ip_GridProps[GridID].cols) { return false; }

        var merge = ip_GridProps[GridID].rowData[startRow].cells[startCol].merge;

        //If we have an existing merge - overwrite it, else ignore it as it already exists
        if (merge == null || merge.rowSpan != RowSpan || merge.colSpan != ColSpan) {

            //Do the frozen row/cols check
            if (startRow < ip_GridProps[GridID].frozenRows && endRow >= ip_GridProps[GridID].frozenRows) { endRow = ip_GridProps[GridID].frozenRows - 1; }
            if (startCol < ip_GridProps[GridID].frozenCols && endCol >= ip_GridProps[GridID].frozenCols) { endCol = ip_GridProps[GridID].frozenCols - 1; }


            //Clear any existing merges
            var validateMerges = ip_ValidateRangeMergedCells(GridID, startRow, startCol, endRow, endCol);
            for (m = 0; m < validateMerges.merges.length; m++) {
                //Add to undo stack
                ip_AddUndoTransactionData(GridID, MergeUndoData, ip_mergeObject(validateMerges.merges[m].mergedWithRow, validateMerges.merges[m].mergedWithCol, validateMerges.merges[m].rowSpan, validateMerges.merges[m].colSpan));
                ip_ResetCellMerge(GridID, validateMerges.merges[m].mergedWithRow, validateMerges.merges[m].mergedWithCol);
            }

            //Create the indexed merge NBNB stored by reference so ALL indexes for that merge share the same object
            var merge = ip_mergeObject(); 
            merge.mergedWithRow = MergeRow;
            merge.mergedWithCol = MergeCol;
            merge.rowSpan = RowSpan;
            merge.colSpan = ColSpan;


            ip_GridProps[GridID].mergeData[merge.mergedWithRow + '-' + merge.mergedWithCol] = merge;

            for (var r = startRow; r <= endRow  ; r++) {

                var cCount = 0;

                //Index the merge with row
                if (ip_GridProps[GridID].rowData[r].containsMerges == null) { ip_GridProps[GridID].rowData[r].containsMerges = new Array(); };
                ip_GridProps[GridID].rowData[r].containsMerges[ip_GridProps[GridID].rowData[r].containsMerges.length] = merge;

                for (var c = startCol; c <= endCol ; c++) {

                    //Add to undostack
                    ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, r, c));

                    //Index the merged with column
                    if (ip_GridProps[GridID].colData[c].containsMerges == null) { ip_GridProps[GridID].colData[c].containsMerges = new Array(); }
                    if (r == startRow) { ip_GridProps[GridID].colData[c].containsMerges[ip_GridProps[GridID].colData[c].containsMerges.length] = merge; }

                    //create an instance of the merge object if we dont have one
                    if (ip_GridProps[GridID].rowData[r].cells[c].merge == null) { ip_GridProps[GridID].rowData[r].cells[c].merge = ip_mergeObject(); }

                    //Reset the values of the cells we merged
                    if (r != MergeRow || c != MergeCol) { ip_SetValue(GridID, r, c, null); }

                    ip_GridProps[GridID].rowData[r].cells[c].merge.mergedWithRow = MergeRow;
                    ip_GridProps[GridID].rowData[r].cells[c].merge.mergedWithCol = MergeCol;
                    ip_GridProps[GridID].rowData[r].cells[c].merge.rowSpan = RowSpan - rCount;
                    ip_GridProps[GridID].rowData[r].cells[c].merge.colSpan = ColSpan - cCount;

                    cCount++;
                }

                rCount++;
            }

            //return true;
        }
    }

    return true;
}

function ip_ValidateFrozenMerge(GridID, rngStartRow, rngStartCol, rngEndRow, rngEndCol) {
//Validates the range to see if spans frozen rows and scroll rows, creates independant merges and returns this

    var arrRange = new Array();

    if (rngStartRow < ip_GridProps[GridID].frozenRows && rngEndRow >= ip_GridProps[GridID].frozenRows && rngStartCol < ip_GridProps[GridID].frozenCols && rngEndCol >= ip_GridProps[GridID].frozenCols) {

        //FROZEN ROWS AND COLS, SCROLL ROWS AND COLS

        //quad 1
        arrRange[arrRange.length] = { startRow: rngStartRow, startCol: rngStartCol, endRow: ip_GridProps[GridID].frozenRows - 1, endCol: ip_GridProps[GridID].frozenCols - 1 }

        //quad 2
        arrRange[arrRange.length] = { startRow: rngStartRow, startCol: ip_GridProps[GridID].frozenCols, endRow: ip_GridProps[GridID].frozenRows - 1, endCol: rngEndCol }

        //quad 3
        arrRange[arrRange.length] = { startRow: ip_GridProps[GridID].frozenRows, startCol: rngStartCol, endRow: rngEndRow, endCol: ip_GridProps[GridID].frozenCols - 1 }

        //quad 4
        arrRange[arrRange.length] = { startRow: ip_GridProps[GridID].frozenRows, startCol: ip_GridProps[GridID].frozenCols, endRow: rngEndRow, endCol: rngEndCol }

    }
    else if (rngStartRow < ip_GridProps[GridID].frozenRows && rngEndRow >= ip_GridProps[GridID].frozenRows) { // && (rngStartCol < ip_GridProps[GridID].frozenCols && rngEndCol < ip_GridProps[GridID].frozenCols || rngStartCol >= ip_GridProps[GridID].frozenCols && rngEndCol >= ip_GridProps[GridID].frozenCols)

        //upper quads
        arrRange[arrRange.length] = { startRow: rngStartRow, endRow: ip_GridProps[GridID].frozenRows - 1, startCol: rngStartCol, endCol: rngEndCol }

        //lower quads
        arrRange[arrRange.length] = { startRow: ip_GridProps[GridID].frozenRows, endRow: rngEndRow, startCol: rngStartCol, endCol: rngEndCol }

    }
    else if (rngStartCol < ip_GridProps[GridID].frozenCols && rngEndCol >= ip_GridProps[GridID].frozenCols) { // && (rngStartCol < ip_GridProps[GridID].frozenCols && rngEndCol < ip_GridProps[GridID].frozenCols || rngStartCol >= ip_GridProps[GridID].frozenCols && rngEndCol >= ip_GridProps[GridID].frozenCols)

        //left quads
        arrRange[arrRange.length] = { startRow: rngStartRow, endRow: rngEndRow, startCol: rngStartCol, endCol: ip_GridProps[GridID].frozenCols - 1 }

        //right quads
        arrRange[arrRange.length] = { startRow: rngStartRow, endRow: rngEndRow, startCol: ip_GridProps[GridID].frozenCols, endCol: rngEndCol }

    }
    else { arrRange[arrRange.length] = { startRow: rngStartRow, startCol: rngStartCol, endRow: rngEndRow, endCol: rngEndCol }  }

    
    return arrRange;
}

function ip_SetRowColSpan(GridID, startRow, endRow, startCol, endCol, depth, LoadedRowsCols) {


    var Indexed = new Array();
    
    
    if (depth == null) { depth = 0; }

    ////Validate
    if (endRow >= ip_GridProps[GridID].rows) { endRow = ip_GridProps[GridID].rows - 1; }
    if (endCol >= ip_GridProps[GridID].cols) { endCol = ip_GridProps[GridID].cols - 1; }

    //ROWS
    if (startRow != null) {

        for (var r = startRow; r <= endRow; r++) {

            //Check if row contains merges
            if (ip_GridProps[GridID].rowData[r].containsMerges != null) {

                if (LoadedRowsCols == null) { LoadedRowsCols = ip_LoadedRowsCols(GridID, false); }

                //Foreach range
                for (var m = 0; m < ip_GridProps[GridID].rowData[r].containsMerges.length; m++) {

                    var merge = ip_GridProps[GridID].rowData[r].containsMerges[m];
                    var IndexedKey = merge.mergedWithRow + '-' + merge.mergedWithCol;

                    if (Indexed[IndexedKey] == null) {

                        //Check if merge is in the visible area
                        var mergeLoadedProperties = ip_GetMergeLoadedProperties(GridID, merge, LoadedRowsCols);

                        if (mergeLoadedProperties != null) {


                            for (var rMerged = mergeLoadedProperties.startRow; rMerged <= mergeLoadedProperties.endRow; rMerged++) {
      
                                $('#' + GridID + '_cell_' + rMerged + '_' + mergeLoadedProperties.startCol).addClass('ip_grid_MergedCell'); //.hide();
                 
                            }


                            //Calculate the rows span, factoring in hidden rows
                            var rowSpan = ip_CalculateHiddenRowSpan(GridID, mergeLoadedProperties.startRow, mergeLoadedProperties.endRow);
                            $('#' + GridID + '_cell_' + mergeLoadedProperties.startRow + '_' + mergeLoadedProperties.startCol).attr('rowspan', rowSpan);

                            var colSpan = ip_CalculateHiddenColSpan(GridID, mergeLoadedProperties.startCol, mergeLoadedProperties.endCol);
                            $('#' + GridID + '_cell_' + mergeLoadedProperties.startRow + '_' + mergeLoadedProperties.startCol).attr('colspan', colSpan);

                            //Show the corner cell                            
                            $('#' + GridID + '_cell_' + mergeLoadedProperties.startRow + '_' + mergeLoadedProperties.startCol).removeClass('ip_grid_MergedCell');

                        }

                    }

                    Indexed[IndexedKey] = true;
                }

            }

        }
    }

    //COLS
    if (startCol != null) {

        for (var c = startCol; c <= endCol; c++) {

            //Check if row contains merges
            if (ip_GridProps[GridID].colData[c].containsMerges != null) {

                if (LoadedRowsCols == null) { LoadedRowsCols = ip_LoadedRowsCols(GridID, false); }

                //Foreach range
                for (var m = 0; m < ip_GridProps[GridID].colData[c].containsMerges.length; m++) {

                    var merge = ip_GridProps[GridID].colData[c].containsMerges[m];
                    var IndexedKey = merge.mergedWithRow + '-' + merge.mergedWithCol;

                    if (Indexed[IndexedKey] == null) {

                        //Check if merge is in the visible area
                        var mergeLoadedProperties = ip_GetMergeLoadedProperties(GridID, merge, LoadedRowsCols);

                        if (mergeLoadedProperties != null) {


                            for (var cMerged = mergeLoadedProperties.startCol; cMerged <= mergeLoadedProperties.endCol; cMerged++) {

                                $('#' + GridID + '_cell_' + mergeLoadedProperties.startRow + '_' + cMerged).addClass('ip_grid_MergedCell'); //.hide();
                         
                            }

                            var colSpan = ip_CalculateHiddenColSpan(GridID, mergeLoadedProperties.startCol, mergeLoadedProperties.endCol);
                            $('#' + GridID + '_cell_' + mergeLoadedProperties.startRow + '_' + mergeLoadedProperties.startCol).attr('colspan', colSpan);

                            var rowSpan = ip_CalculateHiddenRowSpan(GridID, mergeLoadedProperties.startRow, mergeLoadedProperties.endRow);
                            $('#' + GridID + '_cell_' + mergeLoadedProperties.startRow + '_' + mergeLoadedProperties.startCol).attr('rowspan', rowSpan);

                            //Show the corner cell
                            $('#' + GridID + '_cell_' + mergeLoadedProperties.startRow + '_' + mergeLoadedProperties.startCol).removeClass('ip_grid_MergedCell');

                        }

                    }

                    Indexed[IndexedKey] = true;
                }

            }

        }
    }
}

function ip_GetNextMergedCell(GridID, row, col, direction) {

    var cellData = ip_CellData(GridID, row, col);
    var cell = { row: row, col: col }

    if (cellData != null) {

        //There is a merge...
        if (cellData.merge != null) {

            //Check for hidden rows
            cell.row = ip_NextNonHiddenRow(GridID, cellData.merge.mergedWithRow, cellData.merge.mergedWithRow + cellData.merge.rowSpan - 1, row, 'down');
            //--this
            cell.col = ip_NextNonHiddenCol(GridID, cellData.merge.mergedWithCol, cellData.merge.mergedWithCol + cellData.merge.colSpan - 1, col, 'right');

            if (cellData.merge.mergedWithRow != null) {

                var RowAfterMerge = cellData.merge.mergedWithRow + cellData.merge.rowSpan + (row - cellData.merge.mergedWithRow);

                if (direction == 'down' && (row > cellData.merge.mergedWithRow && row < RowAfterMerge)) { cell.row = RowAfterMerge; }
                else if (row > cellData.merge.mergedWithRow && row < RowAfterMerge) { cell.row = cellData.merge.mergedWithRow; }

            }

            if (cellData.merge.mergedWithCol != null) {

                var ColAfterMerge = cellData.merge.mergedWithCol + cellData.merge.colSpan + (col - cellData.merge.mergedWithCol);

                if (direction == 'right' && (col > cellData.merge.mergedWithCol && col < ColAfterMerge)) { cell.col = ColAfterMerge; }
                else if (col > cellData.merge.mergedWithCol && col < ColAfterMerge) { cell.col = cellData.merge.mergedWithCol; }

            }


        }

    }

    return cell;
}

function ip_GetMergeLoadedProperties(GridID, merge, LoadedRowsCols) {
    //Returns null if merge is not loaded into visible area
    //Loads up the visible properties of a merge
    
    var MergeStartRow = merge.mergedWithRow;
    var MergeEndRow = merge.mergedWithRow + merge.rowSpan - 1;
    var MergeStartCol = merge.mergedWithCol;
    var MergeEndCol = merge.mergedWithCol + merge.colSpan - 1;
    var Quad = ip_GetQuad(GridID, MergeStartRow, MergeStartCol); //MergeEndRow

    LoadedRowsCols = (LoadedRowsCols == null ? ip_LoadedRowsCols(GridID, false) : LoadedRowsCols);

    
    if(Quad == 1 || Quad == 2){

        //Validate rows in frozen area
        if(MergeStartRow < LoadedRowsCols.rowFrom_frozen && MergeEndRow < LoadedRowsCols.rowFrom_frozen) { return null; }
        if(MergeStartRow > LoadedRowsCols.rowTo_frozen   && MergeEndRow > LoadedRowsCols.rowTo_frozen) { return null; }

    }
    else if (Quad == 3 || Quad == 4){

        //Validate rows in scroll area
        if(MergeStartRow < LoadedRowsCols.rowFrom_scroll && MergeEndRow < LoadedRowsCols.rowFrom_scroll) { return null; }
        if(MergeStartRow > LoadedRowsCols.rowTo_scroll   && MergeEndRow > LoadedRowsCols.rowTo_scroll) { return null; }

    }

    
    if(Quad == 1 || Quad == 3){

        //Validate cols in frozen area
        if(MergeStartCol < LoadedRowsCols.colFrom_frozen && MergeEndCol < LoadedRowsCols.colFrom_frozen) { return null; }
        if(MergeStartCol > LoadedRowsCols.colTo_frozen   && MergeEndCol > LoadedRowsCols.colTo_frozen) { return null; }

    }
    else if (Quad == 2 || Quad == 4){

        //Validate cols in scroll area
        if(MergeStartCol < LoadedRowsCols.colFrom_scroll && MergeEndCol < LoadedRowsCols.colFrom_scroll) { return null; }
        if(MergeStartCol > LoadedRowsCols.colTo_scroll   && MergeEndCol > LoadedRowsCols.colTo_scroll) { return null; }

    }

    var mergeLoadedProperties = { 
        startRow: MergeStartRow,
        startCol: MergeStartCol,
        endRow: MergeEndRow,
        endCol: MergeEndCol,
        rowSpan: merge.rowSpan,
        colSpan: merge.colSpan
    }

    if (Quad == 4) {

        //Rows
        if (MergeStartRow < LoadedRowsCols.rowFrom_scroll) { mergeLoadedProperties.startRow = LoadedRowsCols.rowFrom_scroll; }
        if (MergeEndRow > LoadedRowsCols.rowTo_scroll) { mergeLoadedProperties.endRow = LoadedRowsCols.rowTo_scroll; }

        //Cols
        if (MergeStartCol < LoadedRowsCols.colFrom_scroll) { mergeLoadedProperties.startCol = LoadedRowsCols.colFrom_scroll }
        if (MergeEndCol > LoadedRowsCols.colTo_scroll) { mergeLoadedProperties.endCol = LoadedRowsCols.colTo_scroll; }

    }
    else if (Quad == 3) {

        //Rows
        if (MergeStartRow < LoadedRowsCols.rowFrom_scroll) { mergeLoadedProperties.startRow = LoadedRowsCols.rowFrom_scroll; }
        if (MergeEndRow > LoadedRowsCols.rowTo_scroll) { mergeLoadedProperties.endRow = LoadedRowsCols.rowTo_scroll; }

        //Cols
        if (MergeStartCol < LoadedRowsCols.colFrom_frozen) { mergeLoadedProperties.startCol = LoadedRowsCols.colFrom_frozen; }
        if (MergeEndCol > LoadedRowsCols.colTo_frozen) { mergeLoadedProperties.endCol = LoadedRowsCols.colTo_frozen; }

    }
    else if (Quad == 2) {

        //Rows
        if (MergeStartRow < LoadedRowsCols.rowFrom_frozen) { mergeLoadedProperties.startRow = LoadedRowsCols.rowFrom_frozen; }
        if (MergeEndRow > LoadedRowsCols.rowTo_frozen) { mergeLoadedProperties.endRow = LoadedRowsCols.rowTo_frozen; }

        //Cols
        if (MergeStartCol < LoadedRowsCols.colFrom_scroll) { mergeLoadedProperties.startCol = LoadedRowsCols.colFrom_scroll }
        if (MergeEndCol > LoadedRowsCols.colTo_scroll) { mergeLoadedProperties.endCol = LoadedRowsCols.colTo_scroll; }

    }

    //Factor in hidden rows/cols
    mergeLoadedProperties.startRow = ip_NextNonHiddenRow(GridID, mergeLoadedProperties.startRow, mergeLoadedProperties.endRow, mergeLoadedProperties.startRow, 'down');
    //--this
    mergeLoadedProperties.startCol = ip_NextNonHiddenCol(GridID, mergeLoadedProperties.startCol, mergeLoadedProperties.endCol, mergeLoadedProperties.startCol, 'right');

    return mergeLoadedProperties;
}

function ip_GetRangeMergedCells(GridID, startRow, startCol, endRow, endCol) {

    //This method gets the appropriate range  row/col when selecting a range that covers merged cells
    if (startRow < -1) { startRow = -1; }
    if (startCol < -1) { startCol = -1; }


    var Results = { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol, startMerge: false, endMerge: false }
    var endCellData = null;
    var startCellData = null;

    //Validate that we do not exceed grid area
    if (endRow >= ip_GridProps[GridID].rows || endCol >= ip_GridProps[GridID].cols) { return Results; }
    else if (endRow < 0 || endCol < 0) { return Results; }


    if (Results.startRow != -1) {


        //TopRight
        var startCellDataB = ip_CellData(GridID, Results.startRow, Results.endCol);
        if (startCellDataB.merge != null) {

            var Merge = ip_CellData(GridID, startCellDataB.merge.mergedWithRow, startCellDataB.merge.mergedWithCol);
            var SR = Merge.merge.mergedWithRow;
            var SC = Merge.merge.mergedWithCol;
            var ER = Merge.merge.mergedWithRow + Merge.merge.rowSpan - 1;
            var EC = Merge.merge.mergedWithCol + Merge.merge.colSpan - 1;

            if (SR < Results.startRow) { Results.startRow = SR; }
            if (SC < Results.startCol) { Results.startCol = SC; }
            if (ER > Results.endRow) { Results.endRow = ER; }
            if (EC > Results.endCol) { Results.endCol = EC; }

        }

    }

    if (Results.startCol != -1) {

        //BottomLeft
        var startCellDataC = ip_CellData(GridID, Results.endRow, Results.startCol);
        if (startCellDataC.merge != null) {

            var Merge = ip_CellData(GridID, startCellDataC.merge.mergedWithRow, startCellDataC.merge.mergedWithCol);
            var SR = Merge.merge.mergedWithRow;
            var SC = Merge.merge.mergedWithCol;
            var ER = Merge.merge.mergedWithRow + Merge.merge.rowSpan - 1;
            var EC = Merge.merge.mergedWithCol + Merge.merge.colSpan - 1;

            if (SR < Results.startRow) { Results.startRow = SR; }
            if (SC < Results.startCol) { Results.startCol = SC; }
            if (ER > Results.endRow) { Results.endRow = ER; }
            if (EC > Results.endCol) { Results.endCol = EC; }

        }

    }

    if (Results.startRow != -1 && Results.startCol != -1) {

        //TopLeft    
        var startCellDataA = ip_CellData(GridID, Results.startRow, Results.startCol);
        if (startCellDataA.merge != null) {

            var Merge = ip_CellData(GridID, startCellDataA.merge.mergedWithRow, startCellDataA.merge.mergedWithCol);
            var SR = Merge.merge.mergedWithRow;
            var SC = Merge.merge.mergedWithCol;
            var ER = Merge.merge.mergedWithRow + Merge.merge.rowSpan - 1;
            var EC = Merge.merge.mergedWithCol + Merge.merge.colSpan - 1;

            if (SR < Results.startRow) { Results.startRow = SR; }
            if (SC < Results.startCol) { Results.startCol = SC; }
            if (ER > Results.endRow) { Results.endRow = ER; }
            if (EC > Results.endCol) { Results.endCol = EC; }

            Results.startMerge = true;
        }

    }


    //BottomRight
    var startCellDataD = ip_CellData(GridID, Results.endRow, Results.endCol);
    if (startCellDataD.merge != null) {

        var Merge = ip_CellData(GridID, startCellDataD.merge.mergedWithRow, startCellDataD.merge.mergedWithCol);
        var SR = Merge.merge.mergedWithRow;
        var SC = Merge.merge.mergedWithCol;
        var ER = Merge.merge.mergedWithRow + Merge.merge.rowSpan - 1;
        var EC = Merge.merge.mergedWithCol + Merge.merge.colSpan - 1;

        if (SR < Results.startRow) { Results.startRow = SR; }
        if (SC < Results.startCol) { Results.startCol = SC; }
        if (ER > Results.endRow) { Results.endRow = ER; }
        if (EC > Results.endCol) { Results.endCol = EC; }

        Results.endMerge = true;

    }


    //Facter in hidden rows/cols
    Results.startRow = ip_NextNonHiddenRow(GridID, Results.startRow, Results.endRow, Results.startRow, 'down');
    //--this
    Results.startCol = ip_NextNonHiddenCol(GridID, Results.startCol, Results.endCol, Results.startCol, 'right');

    return Results;
}

function ip_ValidateRangeMergedCells(GridID, startRow, startCol, endRow, endCol)
{
    //This methods returns the merged cells within a range

    var Result = { containsOverlap: false, merges: new Array(), containsUnmergedCells:false, mergesIdentical: true }
    var Indexed = new Array();

    if (startRow < 0) { startRow = 0; }
    if (startCol < 0) { startCol = 0; }
    if (endRow >= ip_GridProps[GridID].rows) { endRow = ip_GridProps[GridID].rows - 1; }
    if (endCol >= ip_GridProps[GridID].cols) { endRow = ip_GridProps[GridID].cols - 1; }

    for (r = startRow; r <= endRow; r++)
    {
        if (ip_GridProps[GridID].rowData[r].containsMerges != null) { //This means we do not need to seek out merges on a cell level, massivly improving performance

            for (c = startCol; c <= endCol; c++) {

                var cellData = ip_CellData(GridID, r, c);
                if (cellData.merge != null) {

                    var merge = cellData.merge;
                    var IndexedKey = merge.mergedWithRow + '-' + merge.mergedWithCol;

                    if (Indexed[IndexedKey] == null) {

                        //Create a new instance of merge so we dont store it by ref
                        var newMerge = ip_mergeObject();
                        var containsOverlap = ip_DoesRangeOverlapMerge(GridID, merge, startRow, startCol, endRow, endCol);

                        //Mereges found - add them to the result object (once)
                        newMerge.containsOverlap = containsOverlap;
                        newMerge.mergedWithRow = merge.mergedWithRow;
                        newMerge.mergedWithCol = merge.mergedWithCol;
                        newMerge.rowSpan = ip_GridProps[GridID].rowData[merge.mergedWithRow].cells[merge.mergedWithCol].merge.rowSpan;
                        newMerge.colSpan = ip_GridProps[GridID].rowData[merge.mergedWithRow].cells[merge.mergedWithCol].merge.colSpan;
                        Result.merges[Result.merges.length] = newMerge;

                        if (containsOverlap == true) { Result.containsOverlap = true }

                    }

                    Indexed[IndexedKey] = true;

                }
            }
        }

    }

    //Calculate if the entire range is fulled up with merges and if merges are identical
    var rowSpan = -1;
    var colSpan = -1;
    var totalMergedCells = 0;
    var totalRangeCells = (endRow - startRow + 1) * (endCol - startCol + 1);
    for (var m = 0; m < Result.merges.length; m++) {
        if (rowSpan != Result.merges[m].rowSpan && rowSpan != -1) { Result.mergesIdentical = false; }
        else if (colSpan != Result.merges[m].colSpan && colSpan != -1) { Result.mergesIdentical = false; }
        totalMergedCells += Result.merges[m].rowSpan * Result.merges[m].colSpan;
        if (rowSpan == -1) { rowSpan = Result.merges[m].rowSpan; }
        if (colSpan == -1) { colSpan = Result.merges[m].colSpan; }
    }
    if (totalMergedCells < totalRangeCells) { Result.containsUnmergedCells = true; }


    return Result;
}

function ip_ReshuffelMergesCol(GridID, fromCol, toCol, colDiff, shuffeIndexes) {

    var mergeIndex = new Array();
    var ForInc = 1;

    if (fromCol != null && toCol != null && colDiff != null) {


        if (colDiff > 0) {
            //Calculate loops direction
            var tmpFromCol = fromCol;
            fromCol = toCol;
            toCol = tmpFromCol;
            ForInc = -1;

        }

        //Loop through columns
        for (var c = fromCol; (ForInc < 0 ? c >= toCol : c <= toCol) ; c += ForInc) {
            if (ip_GridProps[GridID].colData[c].containsMerges != null) {

                var merges = ip_GridProps[GridID].colData[c].containsMerges;

                for (var m = 0; m < merges.length; m++) {

                    //Adjust merges in this column
                    var startRow = merges[m].mergedWithRow;
                    var startCol = merges[m].mergedWithCol;
                    var endRow = startRow + merges[m].rowSpan - 1;
                    var endCol = startCol + merges[m].colSpan - 1;
                    var mi = startRow + '-' + startCol;


                    if (mergeIndex[mi] == null) {

                        if ((startCol >= fromCol && ForInc > 0) || (startCol >= toCol && ForInc < 0)) {
                            //The above check prevents us cutting down the middle of a merge

                            //Adjust the cell's merge
                            for (var mr = startRow; mr <= endRow; mr++) {

                                for (var mc = startCol; mc <= endCol; mc++) {

                                    if (ip_GridProps[GridID].rowData[mr].cells[mc].merge != null) {
                                        ip_GridProps[GridID].rowData[mr].cells[mc].merge.mergedWithCol = startCol + colDiff;
                                    }

                                }
                            }

                            //Adjust merge index (because the merge index is shared by refence - I can updated it once for all indexes (rows + cols))
                            ip_GridProps[GridID].colData[c].containsMerges[m].mergedWithCol = startCol + colDiff;

                            
                            delete ip_GridProps[GridID].mergeData[mi];


                            //Adjust the column index merge
                            mi = startRow + '-' + (startCol + colDiff);
                            ip_GridProps[GridID].mergeData[mi] = ip_GridProps[GridID].colData[c].containsMerges[m];
                            mergeIndex[mi] = true;
                        }


                    }



                }

                //Adjust the column index
                if (shuffeIndexes) {
                    ip_GridProps[GridID].colData[c + colDiff].containsMerges = ip_GridProps[GridID].colData[c].containsMerges;
                    ip_GridProps[GridID].colData[c].containsMerges = null;
                }

            }
        }
    }


}

function ip_ReshuffelMergesRow(GridID, fromRow, toRow, rowDiff, shuffeIndexes) {

    var mergeIndex = new Array();
    var ForInc = 1;

    if (fromRow != null && toRow != null && rowDiff != null) {

        if (rowDiff > 0) {

            //Calculate loops direction
            var tmpFromRow = fromRow;
            fromRow = toRow;
            toRow = tmpFromRow;
            ForInc = -1;

        }

        //Loop through rows
        for (var r = fromRow; (ForInc < 0 ? r >= toRow : r <= toRow) ; r += ForInc) {
            if (ip_GridProps[GridID].rowData[r].containsMerges != null) {

                var merges = ip_GridProps[GridID].rowData[r].containsMerges;

                for (var m = 0; m < merges.length; m++) {

                    //Adjust merges in this column
                    var startRow = merges[m].mergedWithRow;
                    var startCol = merges[m].mergedWithCol;
                    var endRow = startRow + merges[m].rowSpan - 1;
                    var endCol = startCol + merges[m].colSpan - 1;
                    var mi = startRow + '-' + startCol;
                                        

                    if (mergeIndex[mi] == null) {
                        if ((startRow >= fromRow && ForInc > 0) || (startRow >= toRow && ForInc < 0)) { //This check prevents us cutting down the middle of the merge

                            //Adjust the cell's merge
                            for (var mr = startRow; mr <= endRow; mr++) {

                                for (var mc = startCol; mc <= endCol; mc++) {

                                    ip_GridProps[GridID].rowData[mr].cells[mc].merge.mergedWithRow = startRow + rowDiff;

                                }
                            }

                            //Adjust merge index (because the merge index is shared by refence - I can updated it once for all indexes (rows + cols))
                            ip_GridProps[GridID].rowData[r].containsMerges[m].mergedWithRow = startRow + rowDiff;
                            delete ip_GridProps[GridID].mergeData[mi];

                            //Record what we have adjusted   
                            mi = (startRow + rowDiff) + '-' + startCol;
                            ip_GridProps[GridID].mergeData[mi] = ip_GridProps[GridID].rowData[r].containsMerges[m];
                            mergeIndex[mi] = true;

                        }

                    }

                }

                //Adjust the row index
                if (shuffeIndexes) {
                    ip_GridProps[GridID].rowData[r + rowDiff].containsMerges = ip_GridProps[GridID].rowData[r].containsMerges;
                    ip_GridProps[GridID].rowData[r].containsMerges = null;
                }
            }
        }
    }

    var test = '';
}

function ip_AddColumnToMerge(GridID, merge, count) {
    //Adds another column  to an existing merge

    if (merge != null) {

        var StartRow = merge.mergedWithRow;
        var StartCol = merge.mergedWithCol;
        var EndRow = StartRow + merge.rowSpan - 1;
        var EndCol = StartCol + merge.colSpan - 1;

        EndCol += count;

        if (EndRow >= ip_GridProps[GridID].rows) { EndRow = ip_GridProps[GridID].rows - 1; }
        if (EndCol >= ip_GridProps[GridID].cols) { EndCol = ip_GridProps[GridID].cols - 1; }
        
        //Set the new merge  
        if (EndCol >= StartCol && (StartRow != EndRow || StartCol != EndCol)) { ip_SetCellMerge(GridID, StartRow, StartCol, EndRow, EndCol); }

    }


}

function ip_AddRowToMerge(GridID, merge, count) {
//Adds another row to an existing merge

    if (merge != null) {

        var StartRow = merge.mergedWithRow;
        var StartCol = merge.mergedWithCol;
        var EndRow = StartRow + merge.rowSpan - 1;
        var EndCol = StartCol + merge.colSpan - 1;

        EndRow += count;

        if (EndRow >= ip_GridProps[GridID].rows) { EndRow = ip_GridProps[GridID].rows - 1; }
        if (EndCol >= ip_GridProps[GridID].cols) { EndCol = ip_GridProps[GridID].cols - 1; }

        //Set the new merge      
        if (EndRow >= StartRow && (StartRow != EndRow || StartCol != EndCol)) { ip_SetCellMerge(GridID, StartRow, StartCol, EndRow, EndCol); }

    }


}

function ip_GetColumnMerge(GridID, col) {

    var merge = ip_mergeObject();

    if (col > -1 && ip_GridProps[GridID].rows > 0) {

        if (ip_GridProps[GridID].rowData[ip_GridProps[GridID].rows - 1].cells[col].merge != null) {

            merge = ip_GridProps[GridID].rowData[ip_GridProps[GridID].rows - 1].cells[col].merge;
            return ip_GridProps[GridID].rowData[ip_GridProps[GridID].rows - 1].cells[merge.mergedWithCol].merge;
        }

    }

    merge.mergedWithCol = col;
    merge.colSpan = 1;
    merge.rowSpan = 1;

    return merge;
}

function ip_GetRowMerge(GridID, row) {

    var merge = ip_mergeObject();

    if (row > -1 && ip_GridProps[GridID].cols > 0) {

        if (ip_GridProps[GridID].rowData[row].cells[ip_GridProps[GridID].cols - 1].merge != null) {

            merge = ip_GridProps[GridID].rowData[row].cells[ip_GridProps[GridID].cols - 1].merge;
            return ip_GridProps[GridID].rowData[merge.mergedWithRow].cells[ip_GridProps[GridID].cols - 1].merge;
        }
    }

    merge.mergedWithRow = row;
    merge.colSpan = 1;
    merge.rowSpan = 1;

    return merge;
}

function ip_DoesRangeOverlapMerge(GridID, merge, startRow, startCol, endRow, endCol) {

    //Check if this merge overlaps the range
    var containsOverlap = false;

    if (merge.mergedWithRow < startRow) { containsOverlap = true; }
    else if (merge.mergedWithRow + merge.rowSpan - 1 > endRow) { containsOverlap = true; } // 
    else if (merge.mergedWithCol < startCol) { containsOverlap = true; }
    else if (merge.mergedWithCol + merge.colSpan - 1 > endCol) { containsOverlap = true; } // 

    return containsOverlap;
}

function ip_IsMergeInRange(GridID, merge, arrRange) {

    for (var i = 0; i < arrRange.length; i++) {

        if (merge.mergedWithRow < arrRange[i].startRow) { return false; }
        if (merge.mergedWithCol < arrRange[i].startCol) { return false; }
        if (merge.mergedWithRow > arrRange[i].endRow) { return false; }
        if (merge.mergedWithCol > arrRange[i].endCol) { return false; }

    }

    return true;
}

function ip_MergedHeight(GridID, row, col) {

    var merge = ip_GridProps[GridID].rowData[row].cells[col].merge;
    if (merge) {
                
        var defaultBorderHeight = ip_GridProps[GridID].dimensions.defaultBorderHeight;
        var gridHeight = ip_GridProps[GridID].dimensions.gridHeight;
        var height = 0;
        var seekRows = row + (ip_GridProps[GridID].loadedRows < merge.rowSpan ? ip_GridProps[GridID].loadedRows : merge.rowSpan);

        for (var r = row; r < seekRows; r++) {

            height += ip_GridProps[GridID].rowData[r].height + defaultBorderHeight;
            if (height >= gridHeight) { return height - defaultBorderHeight; }

        }

    }

    return height - defaultBorderHeight;
}


//----- IP GRID FUNCTIONAL METHODS ------------------------------------------------------------------------------------------------------------------------------------

function ip_ReRender(GridID) {
   
    //This is a temporary rerender algorythm, we want to come up with a faster one later but it will do for now
    ip_RecalculateLoadedRowsCols(GridID, false, true, true);
    ip_ReRenderCols(GridID);
    ip_ShowColumnFrozenHandle(GridID, true);
    ip_ShowRowFrozenHandle(GridID, true);

}

function ip_ReRenderCell(GridID, row, col, selectCell) {
    
    $('#' + GridID + '_cell_' + row + '_' + col).replaceWith(ip_CreateGridQuadCell({ GridID: GridID, row: row, col: col, Quad: ip_GetQuad(GridID, row, col) }));
    if (selectCell) { $('#' + GridID).ip_SelectCell({ raiseEvent: false, unselect: false }); }
}

function ip_ReRenderCells(GridID, rowData, loadedScrollable) {
    
    var loadedScrollable = (loadedScrollable == null ? ip_LoadedRowsCols(GridID, false) : loadedScrollable);

    for (var r = 0; r < rowData.length; r++) {

        for (var c = 0; c < rowData[r].cells.length; c++) {

            var row = rowData[r].cells[c].row;
            var col = rowData[r].cells[c].col;
                        
            if (ip_IsCellLoaded(GridID, row, col, loadedScrollable)) {

                ip_ReRenderCell(GridID, row, col);
                
            }
        }
    }

}

function ip_ReRenderRanges(GridID, ranges, loadedScrollable, InclColumnHeaders) {
    //renders the rows within range, only renders visible rows
    var loadedScrollable = (loadedScrollable == null ? ip_LoadedRowsCols(GridID, false) : loadedScrollable);

    for (var rng = 0; rng < ranges.length; rng++) {

        var range = ranges[rng];

        var Q1range = { 
            startRow: loadedScrollable.rowFrom_frozen,
            startCol: loadedScrollable.colFrom_frozen,
            endRow: loadedScrollable.rowTo_frozen,
            endCol: loadedScrollable.colTo_frozen
        }

        Q1range = ip_TrimRange(GridID, Q1range, range);
        if (Q1range != null) {

            for (var r = Q1range.startRow; r <= Q1range.endRow; r++) {
                for (var c = Q1range.startCol; c <= Q1range.endCol; c++) {
                    var row = r;
                    var col = c;

                    if (ip_IsCellLoaded(GridID, row, col, loadedScrollable)) {

                        ip_ReRenderCell(GridID, row, col);

                    }
                }
            }
        }

        var Q2range = {
            startRow: loadedScrollable.rowFrom_frozen,
            startCol: loadedScrollable.colFrom_scroll,
            endRow: loadedScrollable.rowTo_frozen,
            endCol: loadedScrollable.colTo_scroll
        }

        Q2range = ip_TrimRange(GridID, Q2range, range);
        if (Q2range != null) {

            for (var r = Q2range.startRow; r <= Q2range.endRow; r++) {
                for (var c = Q2range.startCol; c <= Q2range.endCol; c++) {
                    var row = r;
                    var col = c;

                    if (ip_IsCellLoaded(GridID, row, col, loadedScrollable)) {

                        ip_ReRenderCell(GridID, row, col);

                    }
                }
            }
        }

        var Q3range = {
            startRow: loadedScrollable.rowFrom_scroll,
            startCol: loadedScrollable.colFrom_frozen,
            endRow: loadedScrollable.rowTo_scroll,
            endCol: loadedScrollable.colTo_frozen
        }

        Q3range = ip_TrimRange(GridID, Q3range, range);
        if (Q3range != null) {

            for (var r = Q3range.startRow; r <= Q3range.endRow; r++) {
                for (var c = Q3range.startCol; c <= Q3range.endCol; c++) {
                    var row = r;
                    var col = c;

                    if (ip_IsCellLoaded(GridID, row, col, loadedScrollable)) {

                        ip_ReRenderCell(GridID, row, col);

                    }
                }
            }
        }

        var Q4range = {
            startRow: loadedScrollable.rowFrom_scroll,
            startCol: loadedScrollable.colFrom_scroll,
            endRow: loadedScrollable.rowTo_scroll,
            endCol: loadedScrollable.colTo_scroll
        }

        Q4range = ip_TrimRange(GridID, Q4range, range);
        if (Q4range != null) {

            for (var r = Q4range.startRow; r <= Q4range.endRow; r++) {
                for (var c = Q4range.startCol; c <= Q4range.endCol; c++) {
                    var row = r;
                    var col = c;

                    if (ip_IsCellLoaded(GridID, row, col, loadedScrollable)) {

                        ip_ReRenderCell(GridID, row, col);

                    }
                }
            }
        }
    }

    if (InclColumnHeaders) { ip_ReRenderCols(GridID, 'header')  }


}

function ip_ReRenderCols(GridID, Quad) {
    //Redraws all columns in specifc quads

    if (Quad == null) { Quad = 'all';  }

    if ((Quad == 'all' || Quad == 'frozen')) {

        //Rerender FROZEN columns
        $('#' + GridID + '_q1_table .ip_grid_columnSelectorCell').remove();
        $('#' + GridID + '_q1_table .ip_grid_cell').remove();
        $('#' + GridID + '_q3_table .ip_grid_columnSelectorCell').remove();
        $('#' + GridID + '_q3_table .ip_grid_cell').remove();

        for (var c = 0; c < ip_GridProps[GridID].frozenCols; c++) { ip_AddCol(GridID, { col: c, appendTo: 'end' }); }

        ip_ShowColumnFrozenHandle(GridID, true);

    }

    if (Quad == 'all' || Quad == 'scroll') {

        //Rerender scroll columns
        if (!ip_AddCol(GridID, { col: ip_GridProps[GridID].scrollX, appendTo: 'end', fullRerender: true })) { ip_GridProps[GridID].loadedCols--; }

    }

    //Render just the header columns
    if (Quad == 'header') {

        $('#' + GridID + '_q1_table .ip_grid_columnSelectorCell').each(function (i) {

            var col = parseInt($(this).attr('col'));
            this.outerHTML = ip_CreateGridQuadCell({ GridID: GridID, id: GridID + '_q1', col: col, row: -1, cellType: 'ColSelector', Quad: 1, showColSelector: ip_GridProps[GridID].showColSelector });


        });
        $('#' + GridID + '_q2_table .ip_grid_columnSelectorCell').each(function (i) {

            var col = parseInt($(this).attr('col'));
            this.outerHTML = ip_CreateGridQuadCell({ GridID: GridID, id: GridID + '_q2', col: col, row: -1, cellType: 'ColSelector', Quad: 2, showColSelector: ip_GridProps[GridID].showColSelector });

        });
        
    }

}

function ip_ReRenderRows(GridID, Quad) {

    //Redraws rows in specifc quads
    if (Quad == null) { Quad = 'all';  }

    if ((Quad == 'all' || Quad == 'frozen')) {


        //clear rows
        if (thisBrowser.name == 'ie' && thisBrowser.version < 10) {
            $('#' + GridID + '_q1_table tbody tr').remove();
            $('#' + GridID + '_q2_table tbody tr').remove();
        }
        else {
            document.getElementById(GridID + '_q1_table_tbody').innerHTML = '';
            document.getElementById(GridID + '_q2_table_tbody').innerHTML = '';
        }

        for (var r = 0; r < ip_GridProps[GridID].frozenRows; r++) { ip_AddRow(GridID, { row: r, appendTo: 'end' }); }

        ip_ShowRowFrozenHandle(GridID, true);
    }

    if (Quad == 'all' || Quad == 'scroll') {

        if (!ip_AddRow(GridID, { row: ip_GridProps[GridID].scrollY, appendTo: 'end', fullRerender: true })) { ip_GridProps[GridID].loadedRows--; }

    }
}

function ip_ReRenderValues(GridID, rowData, fromRow, fromCol, toRow, toCol) {

    //Only GridID is a required field
    //Will rerender values found in either rowData range, or from the specified range

    if (rowData != null) {

        for (var r = 0; r < rowData.length; r++) {

            for (var c = 0; c < rowData[r].cells.length; c++) {

                var row = rowData[r].cells[c].row;
                var col = rowData[r].cells[c].col;

                $('#' + GridID + '_cell_' + row + '_' + col + ' .ip_grid_cell_innerContent').html(ip_CellInner(GridID, row, col));

            }

        }

    }
    else {

        var loadedScrollable = ip_LoadedRowsCols(GridID, true);
        var rFr = -1;
        var cFr = -1;
        var rTo = -1;
        var cTo = -1;

        fromRow = (fromRow == null ? loadedScrollable.rowFrom_frozen : fromRow);
        toRow = (toRow == null ? loadedScrollable.rowTo_scroll : toRow);
        fromCol = (fromCol == null ? loadedScrollable.colFrom_frozen : fromCol);
        toCol = (toCol == null ? loadedScrollable.rowTo_scroll : toCol);

        //FROZEN ROWS & FROZEN COLS
        rFr = (fromRow < loadedScrollable.rowFrom_frozen ? loadedScrollable.rowFrom_frozen : fromRow);
        cFr = (fromCol < loadedScrollable.colFrom_frozen ? loadedScrollable.colFrom_frozen : fromCol);

        rTo = (toRow > loadedScrollable.rowTo_frozen ? loadedScrollable.rowTo_frozen : toRow);
        cTo = (toCol > loadedScrollable.colTo_frozen ? loadedScrollable.colTo_frozen : toCol);


        for (var row = rFr; row <= rTo ; row++) {

            for (var col = cFr; col <= cTo ; col++) {

                //Repaint cell
                //var Quad = ip_GetQuad(GridID, row, col);
                //var newCell = ip_CreateGridQuadCell({ row: row, col: col, GridID: GridID, Quad: Quad });
                //var merge = ip_GridProps[GridID].rowData[row].cells[col].merge;

                ////Manage merges - make sure we render enough rows to complete the merge
                //if (merge != null) {

                //    mergeToRow = row + merge.rowSpan - 1;
                //    mergeToCol = col + merge.colSpan - 1; 
                //    if (rTo < mergeToRow) { rTo = mergeToRow;  }
                //    if (cTo < mergeToCol) { cTo = mergeToCol; }
                //}

                $('#' + GridID + '_cell_' + row + '_' + col + ' .ip_grid_cell_innerContent').html(ip_CellInner(GridID, row, col));
                //$('#' + GridID + '_cell_' + row + '_' + col).replaceWith(newCell);
            }

        }

        //SCROLL ROWS & FROZEN COLS
        rFr = (fromRow < loadedScrollable.rowFrom_scroll ? loadedScrollable.rowFrom_scroll : fromRow);
        cFr = (fromCol < loadedScrollable.colFrom_frozen ? loadedScrollable.colFrom_frozen : fromCol);

        rTo = (toRow > loadedScrollable.rowTo_scroll ? loadedScrollable.rowTo_scroll : toRow);
        cTo = (toCol > loadedScrollable.colTo_frozen ? loadedScrollable.colTo_frozen : toCol);


        for (var row = rFr; row <= rTo ; row++) {

            for (var col = cFr; col <= cTo ; col++) {
                //Repaint cell

                //var Quad = ip_GetQuad(GridID, row, col);
                //var newCell = ip_CreateGridQuadCell({ row: row, col: col, GridID: GridID, Quad: Quad });
                //var merge = ip_GridProps[GridID].rowData[row].cells[col].merge;

                ////Manage merges - make sure we render enough rows to complete the merge
                //if (merge != null) {

                //    mergeToRow = row + merge.rowSpan - 1;
                //    mergeToCol = col + merge.colSpan - 1;
                //    if (rTo < mergeToRow) { rTo = mergeToRow; }
                //    if (cTo < mergeToCol) { cTo = mergeToCol; }
                //}

                $('#' + GridID + '_cell_' + row + '_' + col + ' .ip_grid_cell_innerContent').html(ip_CellInner(GridID, row, col));
                //$('#' + GridID + '_cell_' + row + '_' + col).replaceWith(newCell);
            }

        }

        //FROZEN ROWS & SCROLL COLS
        rFr = (fromRow < loadedScrollable.rowFrom_frozen ? loadedScrollable.rowFrom_frozen : fromRow);
        cFr = (fromCol < loadedScrollable.colFrom_scroll ? loadedScrollable.colFrom_scroll : fromCol);

        rTo = (toRow > loadedScrollable.rowTo_frozen ? loadedScrollable.rowTo_frozen : toRow);
        cTo = (toCol > loadedScrollable.colTo_scroll ? loadedScrollable.colTo_scroll : toCol);


        for (var row = rFr; row <= rTo ; row++) {

            for (var col = cFr; col <= cTo ; col++) {
                //Repaint cell

                //var Quad = ip_GetQuad(GridID, row, col);
                //var newCell = ip_CreateGridQuadCell({ row: row, col: col, GridID: GridID, Quad: Quad });
                //var merge = ip_GridProps[GridID].rowData[row].cells[col].merge;

                ////Manage merges - make sure we render enough rows to complete the merge
                //if (merge != null) {

                //    mergeToRow = row + merge.rowSpan - 1;
                //    mergeToCol = col + merge.colSpan - 1;
                //    if (rTo < mergeToRow) { rTo = mergeToRow; }
                //    if (cTo < mergeToCol) { cTo = mergeToCol; }
                //}

                $('#' + GridID + '_cell_' + row + '_' + col + ' .ip_grid_cell_innerContent').html(ip_CellInner(GridID, row, col));
                //$('#' + GridID + '_cell_' + row + '_' + col).replaceWith(newCell);
            }

        }


        //SCROLL ROWS & COLS
        rFr = (fromRow < loadedScrollable.rowFrom_scroll ? loadedScrollable.rowFrom_scroll : fromRow);
        cFr = (fromCol < loadedScrollable.colFrom_scroll ? loadedScrollable.colFrom_scroll : fromCol);

        rTo = (toRow > loadedScrollable.rowTo_scroll ? loadedScrollable.rowTo_scroll : toRow);
        cTo = (toCol > loadedScrollable.colTo_scroll ? loadedScrollable.colTo_scroll : toCol);


        for (var row = rFr; row <= rTo ; row++) {

            for (var col = cFr; col <= cTo ; col++) {
                //Repaint cell

                //var Quad = ip_GetQuad(GridID, row, col);
                //var newCell = ip_CreateGridQuadCell({ row: row, col: col, GridID: GridID, Quad: Quad });
                //var merge = ip_GridProps[GridID].rowData[row].cells[col].merge;

                ////Manage merges - make sure we render enough rows to complete the merge
                //if (merge != null) {

                //    mergeToRow = row + merge.rowSpan - 1;
                //    mergeToCol = col + merge.colSpan - 1;
                //    if (rTo < mergeToRow) { rTo = mergeToRow; }
                //    if (cTo < mergeToCol) { cTo = mergeToCol; }
                //}

                //ip_CellInner
                $('#' + GridID + '_cell_' + row + '_' + col + ' .ip_grid_cell_innerContent').html(ip_CellInner(GridID, row, col));
                //$('#' + GridID + '_cell_' + row + '_' + col).replaceWith(newCell);
            }

        }

    }

}

function ip_CopyToClipboard(GridID) {

    //window.clipboardData.setData("Text", text);



}

function ip_KeyDownLoop(GridID, e, loopCounter, row, col) {

    if (ip_GridProps[GridID].timeouts.keyDownTimeout == null) { ip_GridProps[GridID].timeouts.keyDownTimeout = 0; }

    var ctrlClick = e.ctrlKey;
    var shitClick = e.shiftKey;

    if (e.keyCode == 37 || e.keyCode == 38 || e.keyCode == 39 || e.keyCode == 40 || e.keyCode == 34 || e.keyCode == 33) {

        //Arrows 
        var row = (row == null ? parseInt($(ip_GridProps[GridID].selectedCell).attr('row')) : row);
        var col = (col == null ? parseInt($(ip_GridProps[GridID].selectedCell).attr('col')) : col);
        var rowIncrement = 0;
        var colIncrement = 0;
        var scrollIncrement = 0;
        var direction = '';


        if (e.keyCode == 37) { colIncrement--; direction = 'left'; }
        else if (e.keyCode == 38) { rowIncrement--; direction = 'up'; }
        else if (e.keyCode == 39) { colIncrement++; direction = 'right'; }
        else if (e.keyCode == 40) { rowIncrement++; direction = 'down'; }
        else if (e.keyCode == 34) { rowIncrement += ip_GridProps[GridID].loadedRows - 1; direction = 'down'; }
        else if (e.keyCode == 33) { rowIncrement -= ip_GridProps[GridID].loadedRows - 1; direction = 'up'; }


        col += colIncrement;
        row += rowIncrement;


        if (colIncrement != 0) { scrollIncrement = colIncrement; }
        else if (rowIncrement != 0) { scrollIncrement = rowIncrement; }


        if (!shitClick) {

            if (row >= 0 && col >= 0) { $('#' + GridID).ip_SelectCell({ row: row, col: col, multiselect: false, scrollIncrement: scrollIncrement, direction: direction, raiseEvent: false }); }
            
        }
        else if (shitClick) { ip_ChangeRange(GridID, ip_GridProps[GridID].selectedRange[ip_GridProps[GridID].selectedRangeIndex], null, 0, rowIncrement, colIncrement, false, scrollIncrement); }



    }

    var TimeoutInterval = (loopCounter == 0 ? 250 : (loopCounter < ip_GridProps[GridID].loadedRows ? 10 : 0 ));
    
    ip_GridProps[GridID].timeouts.keyDownTimeout = setTimeout(function () { ip_KeyDownLoop(GridID, e, (loopCounter + 1), row, col) }, TimeoutInterval);

}

function ip_OptimzeLoadedRows(GridID) {
//Adds or removes rows so that they fit perfectly in the grid visible area so that we do not have excess hidden rows

    var LoadedRowsCols = ip_LoadedRowsCols(GridID, false, true, false);

    //Full in any extra rows as a result of maing rows small            
    var AddRow = LoadedRowsCols.rowTo_scroll + 1;
    while (ip_GridProps[GridID].dimensions.accumulativeScrollHeight < ip_GridProps[GridID].dimensions.scrollHeight && AddRow < ip_GridProps[GridID].rows) {
        ip_GridProps[GridID].loadedRows++;
        ip_AddRow(GridID, { row: AddRow, appendTo: 'end' });
        AddRow++;
    }

    //Remove any extra cols as a result of maing cols large 
    var RemoveRow = LoadedRowsCols.rowTo_scroll;
    while (ip_GridProps[GridID].dimensions.accumulativeScrollHeight - ip_RowHeight(GridID, RemoveRow, true) > ip_GridProps[GridID].dimensions.scrollHeight && RemoveRow > ip_GridProps[GridID].frozenRows) {

        ip_RemoveRow(GridID, { row: RemoveRow });
        ip_GridProps[GridID].loadedRows--;
        RemoveRow--;
    }

}

function ip_OptimizeLoadedCols(GridID) {
//Adds or removes cols so that they fit perfectly in the grid visible area so that we do not have excess hidden cols

    var LoadedRowsCols = ip_LoadedRowsCols(GridID, false, false, true);

    //Full in any extra cols as a result of maing cols small            
    var AddCol = LoadedRowsCols.colTo_scroll + 1;
    while (ip_GridProps[GridID].dimensions.accumulativeScrollWidth < ip_GridProps[GridID].dimensions.scrollWidth && AddCol < ip_GridProps[GridID].cols) {
        ip_GridProps[GridID].loadedCols++;
        ip_AddCol(GridID, { col: AddCol, appendTo: 'end' });
        AddCol++;
    }

    //Remove any extra cols as a result of maing cols large 
    var RemoveCol = LoadedRowsCols.colTo_scroll;
    while (ip_GridProps[GridID].dimensions.accumulativeScrollWidth - ip_ColWidth(GridID, RemoveCol, true) > ip_GridProps[GridID].dimensions.scrollWidth && RemoveCol > ip_GridProps[GridID].frozenCols) {

        ip_RemoveCol(GridID, { col: RemoveCol });
        ip_GridProps[GridID].loadedCols--;
        RemoveCol--;
    }

}


//----- CELL FUNCTIONS  ------------------------------------------------------------------------------------------------------------------------------------

function ip_CellDataType(GridID, row, col, adviseDefault, value, oldMask) {

    // Tests cell value and returns the cell dataType.
    // adviseDefault: returns the columns data type if the cells datatype is null
    // value: lets you override the test for cells current value with a potential future value
    var OldMask = oldMask;
    var Mask = ip_GetMaskObj(GridID, row, col, adviseDefault);
    var Decimals = ip_GetEnabledDecimals(GridID, null, row, col, true);
    var DataType = { output: function () { return this.value.toString() }, mask: null, dataType: ip_dataTypeObject('default'), display:null, value: null, valid: false, expectedDataType: ip_dataTypeObject('default'), decimals:Decimals } //{ dataType: 'default', defaultAlign:'center', value:Cell.value }
    
    if (row < 0) {

        DataType.expectedDataType.dataType = ip_GridProps[GridID].colData[col].dataType.dataType;
        DataType.expectedDataType.dataTypeName = ip_GridProps[GridID].colData[col].dataType.dataTypeName;
        DataType.dataType.dataType = ip_GridProps[GridID].colData[col].dataType.dataType;
        DataType.dataType.dataTypeName = ip_GridProps[GridID].colData[col].dataType.dataTypeName;
        return DataType;

    }
    else {

        var Cell = ip_CellData(GridID, row, col);

        DataType.display = Cell.display;
        DataType.value = (value != null ? value : Cell.value);

        if (Mask != null) {

            DataType.mask = Mask;
            try {
                var mValue = Mask.input(DataType.value);
                DataType.value = (isNaN(mValue) ? DataType.value : mValue);
                DataType.output = function () { return this.mask.output(value, OldMask, Decimals); }
            }
            catch (ex) { DataType.value = (value ? value : Cell.value); }

        } //Converts a masked value into native value

        if (Cell != null) {

            //Check cell data type
            if (Cell.dataType.dataType != null) {
                DataType.expectedDataType.dataType = (!Cell.dataType.dataType ? 'default' : Cell.dataType.dataType.toLowerCase());
                DataType.expectedDataType.dataTypeName = (!Cell.dataType.dataTypeName ? null : Cell.dataType.dataTypeName.toLowerCase());
                DataType.dataType.dataTypeName = DataType.expectedDataType.dataTypeName;
            }

            //Check column data type
            if (adviseDefault) {
                if (DataType.expectedDataType.dataType.toLowerCase() == 'default' && ip_GridProps[GridID].colData[col].dataType.dataType != null) {
                    DataType.expectedDataType.dataType = (!ip_GridProps[GridID].colData[col].dataType.dataType ? 'default' : ip_GridProps[GridID].colData[col].dataType.dataType.toLowerCase());
                    DataType.expectedDataType.dataTypeName = (!ip_GridProps[GridID].colData[col].dataType.dataTypeName ? null : ip_GridProps[GridID].colData[col].dataType.dataTypeName.toLowerCase());
                }
            }

            
            if (((DataType.value != null && DataType.expectedDataType.dataType == 'default') || DataType.expectedDataType.dataType == 'number') && !isNaN(ip_parseNumber(DataType.value))) {
                DataType.dataType.dataType = 'number';
                DataType.value = ip_parseNumber(DataType.value);
                if (!Mask) { DataType.output = function () { return this.decimals == null ? this.value : this.value.toFixed(this.decimals); }; }
                DataType.valid = true;
            }
            else if (((DataType.value != null && DataType.expectedDataType.dataType == 'default') || DataType.expectedDataType.dataType == 'date') && !isNaN(ip_parseDate(DataType.value))) {
                DataType.dataType.dataType = 'date';
                DataType.value = ip_parseDate(DataType.value);
                DataType.valid = true;
            }
            else if (((DataType.value != null && DataType.expectedDataType.dataType == 'default') || DataType.expectedDataType.dataType == 'currency') && !isNaN(ip_parseCurrency(DataType.value))) {
                DataType.dataType.dataType = 'currency';
                DataType.value = ip_parseCurrency(DataType.value);
                if (!Mask) { DataType.output = function () { return this.decimals == null ? this.value : this.value.toFixed(this.decimals); }; }
                DataType.valid = true;
            }
            else if ((DataType.expectedDataType.dataType == 'text')) {
                DataType.dataType.dataType = 'text';
                DataType.value = ip_parseString(DataType.value);
                DataType.valid = true;
            }
            else if ((DataType.expectedDataType.dataType == 'default')) {

                //If we made it this far and the value is null we cannot parse it so we dont know what it is, so assume text
                if (DataType.value == null) { DataType.dataType.dataType = 'text'; DataType.value = ''; DataType.valid = true; }
                else {
                    //Determine the UKNOWN datatype of the cell                
                    DataType.value = ip_parseAny(GridID, DataType.value);
                    DataType.dataType.dataType = typeof (DataType.value);

                    if (DataType.dataType.dataType == 'string') { DataType.dataType.dataType = 'text' }
                    else if (DataType.dataType.dataType == 'object' && typeof DataType.value.getMonth === 'function') { DataType.dataType.dataType = 'date' }

                    DataType.valid = true;
                }
            }
            else {
                //Data type undetermined
                DataType.dataType.dataTypeName = null;
            }
            
            //Default option
            if (DataType.dataType.dataType == 'default') { DataType.dataType.dataType = ''; }
            if (DataType.dataType.dataTypeName == null) { DataType.dataType.dataTypeName = DataType.dataType.dataType; }
            if (DataType.expectedDataType.dataType == 'default') { DataType.expectedDataType.dataType = ''; }
            if (DataType.expectedDataType.dataTypeName == null) { DataType.expectedDataType.dataTypeName = DataType.expectedDataType.dataType; }
            
            return DataType;
        }
    }

    return null;

}

function ip_CellStyle(GridID, row, col) {
    
    return (ip_GridProps[GridID].colData[col].style == null ? '' : ip_GridProps[GridID].colData[col].style) + (ip_GridProps[GridID].rowData[row].cells[col].style == null ? '' : ip_GridProps[GridID].rowData[row].cells[col].style);
    //return (ip_GridProps[GridID].rowData[row].cells[col].style == null ? ip_GridProps[GridID].colData[col].style : ip_GridProps[GridID].rowData[row].cells[col].style);

}

function ip_CellData(GridID, row, col, withControls) {

    //if (row == 3 && col == 0) { debugger; }
    //This is the cell object,
    //WithControls will return the cell with all its appropriate properties

    if (ip_GridProps[GridID].rowData.length > row && row >= 0) {
        if (ip_GridProps[GridID].rowData[row].cells.length > col && col >= 0) {
            
            var cell = ip_GridProps[GridID].rowData[row].cells[col];
            var column = ip_GridProps[GridID].colData[col];

            if (withControls) {

                var value = '';
                var groupCount = ip_GridProps[GridID].rowData[row].groupCount;
                var isLoading = ip_GridProps[GridID].rowData[row].loading;

                var errorCss = (ip_GridProps[GridID].rowData[row].cells[col].error && ip_GridProps[GridID].rowData[row].cells[col].error.errorCode ? ' ip_grid_cell_error' : '');
                var rowHideCss = (ip_GridProps[GridID].rowData[row].hide ? ' ip_grid_cell_hideRow' : '');
                var colHideCss = (ip_GridProps[GridID].colData[col].hide ? ' ip_grid_cell_hideCol' : '');
                var groupCss = (groupCount > 0 ? ' ip_grid_rowGroup' : '');
                var loadingCss = (isLoading ? ' ip_grid_cell_loading' : '');
                var mergeCss = '';
                var border = (column.border == null ? '' : column.border) + (cell.border == null ? '' : cell.border);

                var rowSpan = '';
                var colSpan = '';
                var MergedCell = null;
                var mergedHeight = '';

                //Get merge properties
                if (cell.merge != null) {

                    mergeCss = ' ip_grid_MergedCell';
                    
                    if (cell.merge.mergedWithRow != row || cell.merge.mergedWithCol != col) { MergedCell = ip_CellData(GridID, cell.merge.mergedWithRow, cell.merge.mergedWithCol); }


                    if (cell.merge.mergedWithRow == row && cell.merge.mergedWithCol == col) { mergeCss = ''; }
                    else if (col == ip_GridProps[GridID].scrollX && row == cell.merge.mergedWithRow) { mergeCss = ''; }
                    else if (row == ip_GridProps[GridID].scrollY && col == cell.merge.mergedWithCol) { mergeCss = ''; }
                    else if (col == ip_NextNonHiddenCol(GridID, cell.merge.mergedWithCol, col + cell.merge.colSpan - 1, col, 'right', true) && row == ip_NextNonHiddenRow(GridID, cell.merge.mergedWithRow, row + cell.merge.rowSpan - 1, row, 'down', true)) { mergeCss = ''; }



                    //COUNT THE AMOUNT OF HIDES IN MERGE            
                    if (mergeCss == '') { rowSpan = ip_CalculateHiddenRowSpan(GridID, row, row + cell.merge.rowSpan - 1); }
                    else { rowSpan = cell.merge.rowSpan; }
                    if (mergeCss == '') { colSpan = ip_CalculateHiddenColSpan(GridID, col, col + cell.merge.colSpan - 1); }
                    else { colSpan = cell.merge.colSpan; }

                    rowSpan = 'rowspan="' + rowSpan + '"';
                    colSpan = 'colspan="' + colSpan + '"';

                    if (MergedCell != null) { cell = MergedCell; }

                    if (border != '') {
                        if (ip_GridProps['index'].browser.name == 'ie') { mergedHeight = 'height:' + ip_MergedHeight(GridID, row, col) + 'px;'; }
                        else { mergedHeight = 'height:100%;'; }
                    }
                    
                }

                
                return {
                    
                    cell: cell,
                    isLoading: isLoading,
                    iconDD: (ip_GetEnabledControlType(GridID, null, row, col, true) == 'dropdown' ? '<div class="ip_grid_cell_dropdownIconContainer"><div class="ip_grid_cell_dropdownIcon"></div></div>' : ''),
                    iconGP: (groupCount > 0 && ip_GridProps[GridID].rowData[row].groupColumn == col ? '<div class="ip_grid_cell_groupIcon ' + (ip_GridProps[GridID].rowData[row + 1].hide ? 'collapsed' : 'expanded') + '">' + (ip_GridProps[GridID].rowData[row + 1].hide ? '+' : '-') + '</div>' : ''),
                    css: mergeCss + loadingCss + rowHideCss + groupCss + colHideCss + errorCss,
                    style: ip_CellStyle(GridID, row, col) + mergedHeight,
                    rowSpan: rowSpan,
                    colSpan: colSpan,
                    border: border,

                }

            }
            else { return cell; }
            
        }
    }

}

function ip_CellInput(GridID, options) {
    //Sets cellvalue for indevidual cells for any array of ranges, syncing indevidual cells with the server

    var options = $.extend({
          
        range: null,  //[{ startRow:null, startCol: null, endRow: null, endCol: null }]
        row: null,
        col: null,
        valueRAW: null, //Is specified, auto calculates the formula and sets the value, overriding value and formula specified in options
        value: null, //Set this value
        formula: null, //Set this formula
        formulaObj: null,
        validation: null,
        controlType: null,
        hashTags:null,
        style: null,
        dataType: null,
        raiseEvent: true,
        render: true,

    }, options);
    

    var TransactionID = ip_GenerateTransactionID();
    var Effected = { cell: null, rowData: [],  PropertyAppendModes:'' } 
    var EffectedRecalculated = null;
    var row = null;
    var col = null;
    var rIndex = 0;

    if (options.range == null) { options.range = [ip_rangeObject(options.row, options.col, options.row, options.col)] }
    else if (!options.range.length) { options.range = [options.range]; }

    for (var rng = 0; rng < options.range.length; rng++) {
                
        var CellUndoData = ip_AddUndo(GridID, 'ip_CellInput', TransactionID, 'CellData', options.range[rng], options.range[rng], (rng == 0 ? { row: options.range[rng].startRow, col: options.range[rng].startCol } : null));

        for (var r = options.range[rng].startRow; r <= options.range[rng].endRow; r++) {

            for (var c = options.range[rng].startCol; c <= options.range[rng].endCol; c++) {

                row = r;
                col = c;

                var fxResult = ip_ValidateCellValue(GridID, (options.valueRAW != null ? ip_fxValue(GridID, options.valueRAW, row, col) : { error: null, value: options.value, formula: options.formula }), row, col, true);
                if (!fxResult.reject) {

                    //Record undo
                    ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, r, c));

                    ip_GridProps[GridID].rowData[row].cells[col].error = (fxResult.error && fxResult.validation.validationAction != 'allow' ? fxResult.error : ip_errorObject('', ''));
                    ip_SetValue(GridID, row, col, fxResult.value, undefined, undefined, fxResult.dataType.decimals); 
                    ip_SetCellFormula(GridID, { row: row, col: col, formula: fxResult.formula, formulaObj: options.formulaObj });
                    ip_SetCellControlType(GridID, row, col);
                    ip_SetCellFormat(GridID, { row: row, col: col, validation: options.validation, dataType: options.dataType, style: options.style, hashTags: options.hashTags, raiseEvent: false, render: false, transactionID: TransactionID })
                    
                    if (Effected.rowData[rIndex] == null) { Effected.rowData[rIndex] = { cells: [] } }
                    Effected.rowData[rIndex].cells[Effected.rowData[rIndex].cells.length] = ip_CloneCell(GridID, row, col);
                    if (Effected.rowData[rIndex].cells.length == ip_GridProps[GridID].cols) { rIndex++; }

                }
                else { ip_TextEditbleCancel(GridID); }

            }

        }
    }

    EffectedRecalculated = ip_ReCalculateFormulas(GridID, { recalcSource: false, render: false, raiseEvent: false, transactionID: TransactionID, range: options.range }).Effected;
    Effected.rowData = Effected.rowData.concat(EffectedRecalculated.rowData);

    if (options.raiseEvent || options.render) {
        if (Effected.rowData.length == 1 && Effected.rowData[0].cells.length == 1 && (EffectedRecalculated == null || EffectedRecalculated.rowData.length == 0)) {

            //Raise event for just one cell

            //Setup event options    
            Effected.cell = ip_GridProps[GridID].rowData[row].cells[col];
            Effected.cell.row = row;
            Effected.cell.col = col;

            if (options.render) {  ip_ReRenderCell(GridID, row, col);  } //$('#' + GridID + '_cell_' + row + '_' + col).html(ip_CellOuter(GridID, row, col)); }

            //Raise cell change event
            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_CellInput', TransactionID, { SetCellValue: { Inputs: options, Effected: Effected } }); }

        }
        else if(Effected.rowData.length > 0) {
            //Raise event for multple one cells
 
            if (options.render) { ip_ReRenderCells(GridID, Effected.rowData); }

            if (options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_SetCellValues', TransactionID, { SetCellValues: { Inputs: { range: Effected.range }, Effected: Effected } }); }

        }

    }


}

function ip_SetValue(GridID, row, col, value, oldMask) {
    //This is the central point for setting a cell value, handles the mask, decimal places
    //All instances of 'ip_GridProps[GridID].rowData[row].cells[col].value =' must do this.
    //Returns true if the value actually changed
    //oldMask should typically only be used if the mask has changed - it is rare that we need to populate this
    var oldValue = ip_GridProps[GridID].rowData[row].cells[col].value;    

    if (oldValue != value) {

        if (typeof value == 'undefined') { value = ip_GridProps[GridID].rowData[row].cells[col].value; }

        if (value == null) {

            ip_GridProps[GridID].rowData[row].cells[col].value = null;
            ip_GridProps[GridID].rowData[row].cells[col].display = '';

        }
        else {
                       
            var error = ip_GridProps[GridID].rowData[row].cells[col].error;
            var dataType = {};
            var formatted = '';

            if (error == null || (error && error.errorCode == '')) {
                dataType = ip_CellDataType(GridID, row, col, true, value, oldMask);
                if (dataType.valid) { formatted = dataType.output(); }
                else { formatted = value; }
            }
            else {
                dataType.value = value;
                formatted = value;
            }

            ip_GridProps[GridID].rowData[row].cells[col].value = dataType.value;
            ip_GridProps[GridID].rowData[row].cells[col].display = formatted;
        }

        return true;
    }

    return false;
}

function ip_SetCellFormat(GridID, options) {
    //This is physically different to SetCellValue because it accepts formatting options and applies it to the entire range, this single value for the entire range then syncs
    //this offers more performance, but has the limitation of not being able to update indevidual rowdata/celldata with the server
    //If indevidual row/cell data is required, SetCellValue should be used rather.
    var options = $.extend({

        transactionID: ip_GenerateTransactionID(),
        range: null, // [{ startRow:null, startCol: null, endRow: null, endCol: null }],     
        row: null,
        col: null,
        dataType: null, //ip_dataTypeObject()
        style: null,
        border: null,
        validation: null, //ip_validationObject()
        PropertyAppendModes: 'style:append;cell:auto', //supporting 'style:append;cell:notnull;cell:auto'
        controlType: null, //dropdown, gantt, checkbox
        hashTags: null,
        mask: null,
        decimals: null,
        decimalsInc: null,        
        fxCallerName: null,
        recalculate: false, //Recalculate formulas
        raiseEvent: true,
        render: true,
        createUndo: true

    }, options);

    //Validate if we have formatting 
    if (options.dataType == null && options.style == null && options.validation == null && options.controlType == null && options.hashTags == null
        && options.mask == null && options.decimals == null && options.decimalsInc == null) { return; }
    
    var TransactionID = options.transactionID;
    var Error = '';    
    var Effected = null;
    var formatObject = ip_formatObject();
    if (options.range == null && options.row != null && options.col != null) { options.range = [ip_rangeObject(options.row, options.col, (options.row == -1 ? ip_GridProps[GridID].rows-1 : options.row), options.col)] }
    
    //Validate range
    if (options.range == null && ip_GridProps[GridID].selectedRange.length == 0) { Error = 'Please specify a range'; }    
    else if (options.range == null) {

        //Copied ranges
        options.range = new Array();
        for (var i = 0; i < ip_GridProps[GridID].selectedRange.length; i++) { options.range[i] = ip_rangeObject(ip_GridProps[GridID].selectedRange[i][0][0], ip_GridProps[GridID].selectedRange[i][0][1], ip_GridProps[GridID].selectedRange[i][1][0], ip_GridProps[GridID].selectedRange[i][1][1], null, options.col); }

    }
    
    if (options.range != null && Error == '')
    {
        //Format cells
        Effected = { range: [], rowData: [], cellFormat: { style: options.style, dataType: options.dataType, validation: options.validation, controlType: options.controlType, hashTags: options.hashTags, mask: options.mask, decimals: options.decimals, decimalsInc: options.decimalsInc } }

        for (var i = 0; i < options.range.length; i++) {

            options.range[i].PropertyAppendModes = options.PropertyAppendModes;

            var origonalRow = options.range[i].startRow;
            var startRow = options.range[i].startRow;
            var startCol = options.range[i].startCol;
            var endRow = options.range[i].endRow;
            var endCol = options.range[i].endCol;            

            if (startRow < 0) { startRow = 0; }
            if (startCol < 0) { startCol = 0; }
            if (endRow >= ip_GridProps[GridID].rows) { startRow = ip_GridProps[GridID].rows - 1; }
            if (startCol >= ip_GridProps[GridID].cols) { startCol = ip_GridProps[GridID].cols - 1; }

            var FormatRange = { startRow: startRow, startCol: startCol, endRow: endRow, endCol: endCol }
            if (options.createUndo) { var CellUndoData = ip_AddUndo(GridID, 'ip_FormatCell', TransactionID, 'CellData', FormatRange, FormatRange, { row: FormatRange.startRow, col: FormatRange.startCol }); }
            if (options.range[i].PropertyAppendModes.indexOf('cell:auto') != -1 && origonalRow <= 0 && endRow == ip_GridProps[GridID].rows - 1) { options.range[i].PropertyAppendModes = options.range[i].PropertyAppendModes.replace('cell:auto', 'cell:colnotnull'); }

            var PropertyAppendModes =
                {
                    colnotnull: (options.range[i].PropertyAppendModes.indexOf('cell:colnotnull') != -1 ? true : false), //Sets the col object + cells with existing styles, leaves nulls cells alone
                    cellnotnull: (options.range[i].PropertyAppendModes.indexOf('cell:notnull') != -1 ? true : false), //Sets cells with existing styles, leaves non cells alone
                    stylappend: (options.range[i].PropertyAppendModes.indexOf('style:append') != -1 ? true : false)
                }


            //FORMAT COLUMN - this improves performance
            if (PropertyAppendModes.colnotnull && Error == '') {

                var columns = [];
                var oldMasks = {};
                for (var c = options.range[i].startCol; c <= options.range[i].endCol; c++) {
                    columns[columns.length] = c;
                    oldMasks[c] = ip_GridProps[GridID].colData[c].mask;
                }

                //Format columns 
                var cOptions = $.extend(cOptions, options);
                cOptions.columns = columns;
                cOptions.raiseEvent = false;
                cOptions.render = false;
                cOptions.createUndo = options.createUndo;
                cOptions.transactionID = TransactionID;
                formatObject = ip_SetColumnFormat(GridID, cOptions);

            }

            //FORMAT CELL
            var defaultMaskInit = false;
            for (var r = startRow; r <= endRow; r++) {
                for (var c = startCol; c <= endCol; c++) {

                    if (((PropertyAppendModes.cellnotnull || PropertyAppendModes.colnotnull) && (!ip_isCellEmpty(GridID, r, c))) || (!PropertyAppendModes.cellnotnull && !PropertyAppendModes.colnotnull)) {

                        var cell = ip_GridProps[GridID].rowData[r].cells[c];
                        var oldMask = undefined;
                        var setvalue = false;
                        var validate = false;

                        if (options.createUndo) { ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, r, c)); }
                        
                        if (options.style != null) { cell.style = (options.style == '' ? null : (PropertyAppendModes.stylappend ? ip_AppendCssStyle(GridID, ip_GridProps[GridID].rowData[r].cells[c].style, options.style) : options.style)); }

                        if (options.controlType != null) { cell.controlType = (options.controlType == '' || PropertyAppendModes.colnotnull ? null : options.controlType); }

                        if (options.hashTags != null) { cell.hashTags = (options.hashTags == '' || PropertyAppendModes.colnotnull ? null : options.hashTags); }

                        if (options.border != null) { cell.border = (options.border == '' || PropertyAppendModes.colnotnull ? null : options.border); }
                        
                        if (options.validation != null) {                            
                            if (options.validation.validationCriteria != null) { cell.validation.validationCriteria = (options.validation.validationCriteria == '' || PropertyAppendModes.colnotnull ? null : options.validation.validationCriteria); validate = true; }
                            if (options.validation.validationAction != null) { cell.validation.validationAction = (options.validation.validationAction == '' || PropertyAppendModes.colnotnull ? null : options.validation.validationAction); validate = true; }                                                       
                        }

                        if (options.dataType != null) {
                            
                            cell.dataType.dataType = (options.dataType.dataType == '' || PropertyAppendModes.colnotnull ? null : options.dataType.dataType);
                            cell.dataType.dataTypeName = (options.dataType.dataTypeName == '' || PropertyAppendModes.colnotnull ? null : options.dataType.dataTypeName);

                            //Switch to a default mask for the datatype  
                            if (!defaultMaskInit) {
                                options.mask = ip_DefaultMask(GridID, options.dataType, (!options.mask ? cell.mask : options.mask));
                                options.mask = (options.mask == null ? '' : options.mask);                                
                                Effected.cellFormat.mask = options.mask;
                                defaultMaskInit = true;
                            }
                            
                        }
                                              
                        if (options.mask != null) {
                            oldMask = (cell.mask == null && (oldMasks != undefined && oldMasks[c] != null) ? oldMasks[c] : ip_GetEnabledMask(GridID, null, r, c, true));
                            ip_GridProps[GridID].rowData[r].cells[c].mask = (PropertyAppendModes.colnotnull ? null : options.mask);
                            setvalue = true;                            
                        }

                        if (options.decimalsInc != null) {

                            var decimals = cell.decimals;
                            if (decimals == null && !PropertyAppendModes.colnotnull) {
                                decimals = ip_GetEnabledDecimals(GridID, null, r, c, true);
                                if (decimals == null) { decimals = 0; }
                            }
                            if (decimals != null) { cell.decimals = (decimals + options.decimalsInc); }

                        }

                        if (options.decimals != null) { cell.decimals = (PropertyAppendModes.colnotnull ? null : options.decimals); }
                        if (cell.decimals < 0) { cell.decimals = null; }
                        if (options.decimals || options.decimalsInc) { setvalue = true; }
                                                
                        if (validate) {

                            var error = ip_ValidateCellValue(GridID, null, r, c, false).error;
                            
                            if (!ip_CompareError(GridID, error, cell.error)) {
                                cell.error = error;
                                if (error == null) { error = ip_errorObject("", ""); }
                                ip_AppendEffectedRowData(GridID, Effected, { row: r, col: c, error: error });
                            }

                        }

                        if (setvalue) {
                            ip_SetValue(GridID, r, c, undefined, oldMask);
                        }


                        formatObject = ip_GetEnabledFormats(GridID, formatObject, r, c);
                    }
                }
            }                

            ip_AddDataTypeToIndex(GridID, options.dataType);

            FormatRange.PropertyAppendModes = options.range[i].PropertyAppendModes;
            Effected.range[Effected.range.length] = FormatRange;

        }


        if (options.recalculate) { Effected.rowData = Effected.rowData.concat( ip_ReCalculateFormulas(GridID, { range: options.range, transactionID: TransactionID, render: true, raiseEvent: false }).Effected.rowData ); }

    }


   
    if (Error == '')
    {

        //ReRender
        if (options.render) { ip_ReRenderRanges(GridID, options.range, null, PropertyAppendModes.colnotnull); }

        //Raise cell change event - server module should automatically handle the queing of the data types
        if (options.raiseEvent && Effected.range.length > 0) { ip_RaiseEvent(GridID, 'ip_FormatCell', TransactionID, { CellFormat: { Inputs: options, Effected: Effected } }); }

        return formatObject;

    }
    else { ip_RaiseEvent(GridID, 'warning', TransactionID, Error); return null; }



}

function ip_SetCellFormula(GridID, options) {

    var options = $.extend({

        row: null,
        col: null,
        formula: null,
        isRowColShuffel: false, //Used when moving rows/cols only

    }, options);

    
    if (options.row != null && options.col != null) {

        var row = options.row;
        var col = options.col;
        var fx = options.fx;
        var TransactionID = options.transactionID;

        //Remove any existing formulas
        ip_RemoveCellFormulaIndex(GridID, { row: row, col: col, clearFormula: true });
        
        if (fx == null && typeof (options.formula) == 'string') { fx = ip_fxObject(GridID, options.formula, row, col); }
        else if (typeof (options.formula == 'object')) { fx = options.formula }

        //Link formula range cells
        if (fx) {

            //Build local index
            var cell = ip_GridProps[GridID].rowData[options.row].cells[options.col];
            var fxIndex = ip_GenerateTransactionID(); 

            //Create formula index
            ip_GridProps[GridID].indexedData.formulaData[fxIndex] = fx;
            ip_GridProps[GridID].rowData[options.row].cells[options.col].fxIndex = fxIndex;
            ip_GridProps[GridID].rowData[options.row].cells[options.col].formula = fx.formula;
            

            for(key in fx.ranges)
            {
                var range = fx.ranges[key].range;

                for (var r = range.startRow; r <= range.endRow; r++) {

                    //Create formula index
                    for (var c = range.startCol; c <= range.endCol; c++) {
                        if (r != row || c != col || options.isRowColShuffel) {

                            //put in an index on the cell that links to a formula
                            if (ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c] == null) { ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c] = {} }
                            if (ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c].fxIndexLnk == null) { ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c].fxIndexLnk = {} }
                            ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c].fxIndexLnk[fxIndex] = true;

                        }
                    }

                }
            }

           
        }

        return fx;
    }

    return null;
}

function ip_SetCellControlType(GridID, row, col) {

    if (GridID != null && row != null && col != null) {
        var fxIndex = ip_GridProps[GridID].rowData[row].cells[col].fxIndex;
        var ControlTypeArgs = ip_GetCellControlType(GridID, null, fxIndex, row, col);
        if (ControlTypeArgs) {
            ip_GridProps[GridID].rowData[row].cells[col].validation.validationCriteria = ControlTypeArgs.validation.validationCriteria;
            ip_GridProps[GridID].rowData[row].cells[col].formula = ControlTypeArgs.formula;
            ip_GridProps[GridID].rowData[row].cells[col].controlType = ControlTypeArgs.controlType;
        }

    }
}

function ip_GetCellControlType(GridID, formula, fxIndex, row, col) {

    //Looks at the formula and deciceds what the the cells control type
    //Rules for control type: the formula must be explicit and singular, so e.g. 'dropdown(a1:a10) + a1' will not be converted, but 'dropdown(a1:a10)' will be
    
    if (GridID != null && (formula != null || fxIndex != null)) {
        
        var fx = (fxIndex ? ip_GridProps[GridID].indexedData.formulaData[fxIndex] : ip_fxObject(GridID, formula, row, col))

        if (fx ) {

            if (!fx.functions) { fx.functions = ['']; }
            if (fx.formula == '=()') { fx.formula = fx.formula.replace('=()', '') }
        
            var returnObj = {
                controlType: null,
                formula: formula,
                validation: ip_validationObject()
            }

            switch (fx.functions[0].toLowerCase()) {
                case 'gantt':
                    returnObj.controlType = 'gantt';
                    //returnObj.formula = null;
                    returnObj.validation.validationCriteria = (fx.inputs ? fx.formula : null);
                    return returnObj;
                case 'dropdown':
                    returnObj.validation.validationCriteria = (fx.inputs ? fx.formula : null);
                    returnObj.formula = null;
                    returnObj.controlType = 'dropdown';
                    return returnObj;
                case 'text':
                    returnObj.controlType = 'text';
                    //returnObj.formula = null;
                    returnObj.validation.validationCriteria = (fx.inputs ? fx.formula : null);
                    return returnObj;
                default:
                    returnObj.controlType = '';
                    returnObj.formula = fx.formula;
                    returnObj.validation.validationCriteria = '';
                    return returnObj;
            }
        }

    }
    else if (GridID) {

        var returnObj = {
            controlType: null,
            formula: formula,
            validation: ip_validationObject()
        }

    }

}

function ip_ValidateCellValue(GridID, fxResult, row, col, showWarning) {

    var validateCell = ip_GridProps[GridID].rowData[row].cells[col];

    if (fxResult == null) { fxResult = { reject: false, Error: null, value: validateCell.value, formula: validateCell.formula }  }

    fxResult.reject = false;

    if (!fxResult.Error) {

        var Error = null;

        fxResult.validation = ip_GetEnabledValidation(GridID, null, row, col, true);
        fxResult.dataType = ip_CellDataType(GridID, row, col, true, fxResult.value);

        if (!fxResult.dataType.valid) {
            //1. Validate the data type
            fxResult.error = ip_errorObject("02", "Value should be a \"" + fxResult.dataType.expectedDataType.dataTypeName + "\"" + (fxResult.dataType.mask ? ", in the format " + fxResult.dataType.mask.mask : ""));
        }
        else if (fxResult.validation.validationAction && fxResult.validation.validationAction != 'allow') {
            //2. Validate that the contents are correct
            if (fxResult.validation.inputs) {
                var validated = false;
                for (var i = 0; i < fxResult.validation.inputs.length; i++) {

                    //Validate range
                    var range = ip_fxRangeObject(GridID, null, null, fxResult.validation.inputs[i]);
                    if (range) {
                        for (var r = range.startRow; r <= range.endRow; r++) {
                            for (var c = range.startCol; c <= range.endCol; c++) {
                                var cell = ip_CellData(GridID, r, c);
                                if (cell && cell.value && cell.value.match("\\b" + fxResult.value + "\\b", 'gi')) { validated = true; break; }
                            }
                        }
                    }
                    else {

                        //validate normal value
                        if (fxResult.validation.inputs[i].match("\\b" + fxResult.value + "\\b", 'gi')) { validated = true; break; }

                    }

                }
                if (!validated) {

                    var validationText = fxResult.validation.inputs.join();
                    fxResult.error = ip_errorObject("03", '"' + fxResult.value + '" is an invalid cell value, try one of these: ' + (validationText.length > 40 ? validationText.substr(0, 40) + ' ...' : validationText))

                }
            }
        }


    }

    if (fxResult.error && fxResult.error.errorDescription && showWarning) {
        ip_ShowFooterAlert("You have entered an invalid cell value", fxResult.error.errorDescription + ".", null, (fxResult.validation.validationAction == 'allow' ? 'green' : 'red'));
        if (fxResult.validation.validationAction == 'prevent') { fxResult.reject = true; }
    }

    return fxResult;

}

function ip_RemoveCellFormulaIndex(GridID, options) {

    var options = $.extend({
        
        row: null,
        col: null,
        fxIndex: null,
        clearFormula: false,

    }, options);

    var fxIndex = (!options.fxIndex ? ip_GridProps[GridID].rowData[options.row].cells[options.col].fxIndex : options.fxIndex);
    var fx = ip_GridProps[GridID].indexedData.formulaData[fxIndex];
           
    //Link formula range cells
    if (fx != null) {

        for (key in fx.ranges) {

            var range = fx.ranges[key].range;

            for (var r = range.startRow; r <= range.endRow; r++) {
                
                for (var c = range.startCol; c <= range.endCol; c++) {

                    //delete formula index
                    if (ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c] != null && ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c].fxIndexLnk[fxIndex] != null) { delete ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c].fxIndexLnk[fxIndex]; }

                }

            }

        }

        if (options.clearFormula) { ip_GridProps[GridID].rowData[fx.row].cells[fx.col].formula = ''; }
        ip_GridProps[GridID].rowData[fx.row].cells[fx.col].fxIndex = null;
        delete ip_GridProps[GridID].indexedData.formulaData[fxIndex];           
        
    }       

}

function ip_ReCalculateFormulas(GridID, options) {

    //Triggers the formulas linked to this range to be recaculated

    var options = $.extend({

        range: null,// [] array of range object
        level: 0,
        transactionID: null,
        render: true,
        raiseEvent: true,
        createUndo: true,
        undoMode: null,
        recalcSource: true,
        cellUndoData: null //used for recursive recalculations

    }, options);

    var Effected = { rowData: [], range: options.range }
    var fxDone = {};

    
    if (options.range != null) {
        if (options.range.length > 0) {

            var TransactionID = options.transactionID;
            var changeRIndex = false;
            var CellUndoData = options.cellUndoData;

            if (TransactionID == null) { TransactionID = ip_GenerateTransactionID(); }

            for (var i = 0; i < options.range.length; i++) {

                //xxxx-
                var range = options.range[i];
                if (range.endRow == null) { range.endRow = range.startRow; }
                if (range.endCol == null) { range.endCol = range.startCol; }

                var fxList = ip_FormulasInRange(GridID, range, options.recalcSource, true);

                for (var key in fxList) {

                    if (TransactionID == null) { TransactionID = ip_GenerateTransactionID(); }

                    var fx = ip_GridProps[GridID].indexedData.formulaData[key];



                    var row = fx.row;
                    var col = fx.col;
                    var fxIndex = key; //row + '-' + col;
                    var recursiveResult = null;

                    if (options.level == 100) {
                        ip_SetValue(GridID, row, col, "#ERROR");
                        ip_GridProps[GridID].rowData[row].cells[col].error = ip_errorObject('2', 'Circular dependancy detected');                        
                        return { Effected: Effected, fxDone: fxDone }
                    }

                    if (fxDone[fxIndex] == null) {

                        fxDone[fxIndex] = { row: row, col: col, valChange: false };

                        if (changeRIndex) { rIndex++; changeRIndex = false; }

                        if (options.createUndo) {
                            if (options.undoMode != 'server') {
                                if (CellUndoData == null) { CellUndoData = ip_AddUndo(GridID, 'ip_ReCalculateFormulas', TransactionID, 'CellData', { startRow: row, startCol: col, endRow: row, endCol: col }); }
                                ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, row, col));
                            }
                            else {
                                if (CellUndoData == null) { CellUndoData = ip_AddUndo(GridID, 'ip_ReCalculateFormulas', TransactionID, 'undoServer', options.range[i], options.range[i], { row: options.range[i].startRow, col: options.range[i].startCol }); }
                            }
                        }
                             
                        var oldCell = ip_CloneCell(GridID, row, col);
                        var fxResult = ip_fxValue(GridID, fx.formula, row, col);
                        ip_SetValue(GridID, row, col, fxResult.value);
                        ip_GridProps[GridID].rowData[row].cells[col].error = (fxResult.error ? fxResult.error :  ip_errorObject('','') );

                        //Check if the value of the cell has changed at all
                        if (!ip_CompareCell(GridID, oldCell, ip_GridProps[GridID].rowData[row].cells[col])) { fxDone[fxIndex].valChange = true; }

                        //Check if the value errored
                        if (!fxResult.error) {

                            //Check if we need to execute a recursive calculation
                            if (ip_GridProps[GridID].indexedData.cellIndex[row + '-' + col] != null && ip_GridProps[GridID].indexedData.cellIndex[row + '-' + col].fxIndexLnk != null) {
                                var testlnk = ip_GridProps[GridID].indexedData.cellIndex[row + '-' + col];
                                var formulatemp = ip_GridProps[GridID].indexedData.formulaData[testlnk];
                                recursiveResult = ip_ReCalculateFormulas(GridID, { cellUndoData: CellUndoData, recalcSource: false, createUndo: options.createUndo, level: options.level + 1, transactionID: TransactionID, range: [{ startRow: row, startCol: col }] });
                                fxDone = $.extend(fxDone, recursiveResult.fxDone);

                            }
                        }

                    }
      
                }
                
            }

            if (options.level == 0) {

                //This operation is only required in level 0 (the root calculation and not subsequent recursion)
                for (var keyR in fxDone) {

                    if (fxDone[keyR].valChange) {

                        ip_AppendEffectedRowData(GridID, Effected, { row: fxDone[keyR].row, col: fxDone[keyR].col, error: ip_GridProps[GridID].rowData[fxDone[keyR].row].cells[fxDone[keyR].col].error, value: ip_GridProps[GridID].rowData[fxDone[keyR].row].cells[fxDone[keyR].col].value, formula: ip_GridProps[GridID].rowData[fxDone[keyR].row].cells[fxDone[keyR].col].formula });

                    }

                }

                if (Effected.rowData.length > 0 && options.render) { ip_ReRenderCells(GridID, Effected.rowData); }
                if (Effected.rowData.length > 0 && options.raiseEvent) { ip_RaiseEvent(GridID, 'ip_SetCellValues', TransactionID, { SetCellValues: { Inputs: null, Effected: { rowData: Effected.rowData } } }); }
            }

        }
    }
    
    return { Effected: Effected, fxDone: fxDone }
}

function ip_ChangeFormulaOrigin(GridID, options) {

    //Moves formula indexes from a range to a new range, usefull for move row, col -
    //NB unless changeFormula is set to true, does not change the formula - just its origon and points its indexes

    var options = $.extend({

        fxIndex: null,

        range: null, //range object
        rowDiff: 0,  //Use this only when range object is used
        colDiff: 0,  //Use this only when range object is used

        fromRow: null,
        fromCol: null,
        toRow: null,
        toCol: null,

        changeFormula: false,

    }, options);

    var fxIndex = null;
    var fxList = null;
    var row, col = null;

    if (options.fromRow == 2) { debugger; }

    if (options.fxIndex != null && ip_GridProps[GridID].indexedData.formulaData[options.fxIndex] != null) { fxIndex = options.fxIndex; }
    else if (options.fromRow != null && options.fromCol != null && ip_GridProps[GridID].rowData[options.fromRow].cells[options.fromCol].fxIndex != null) { fxIndex = ip_GridProps[GridID].rowData[options.fromRow].cells[options.fromCol].fxIndex; }
    else if (options.toRow != null && options.toCol != null && ip_GridProps[GridID].rowData[options.toRow].cells[options.toCol].fxIndex != null) { fxIndex = ip_GridProps[GridID].rowData[options.toRow].cells[options.toCol].fxIndex; }
    else if (options.range != null && (options.rowDiff != 0 || options.colDiff != 0)) { fxList = ip_FormulasInRange(GridID, options.range, true, false); }

    if (fxIndex != null) {

        var fx = ip_GridProps[GridID].indexedData.formulaData[fxIndex];

        if (options.toRow == null) { options.toRow = fx.row; }
        if (options.toCol == null) { options.toCol = fx.col; }
        row = fx.row;
        col = fx.col;
        fx.row = options.toRow;
        fx.col = options.toCol;

        if (options.changeFormula) { ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: ip_MoveFormulaOrigon(GridID, ip_GridProps[GridID].rowData[row].cells[col].formula, row, col, fx.row, fx.col) }); }

        fxList = {};
        fxList[fxIndex] = true;
        
    }
    else if (fxList != null) {

        for (var key in fxList) {

            fx = ip_GridProps[GridID].indexedData.formulaData[key];
            row = fx.row;
            col = fx.col;
            fx.row += options.rowDiff;
            fx.col += options.colDiff;

            if (options.changeFormula) { ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: ip_MoveFormulaOrigon(GridID, ip_GridProps[GridID].rowData[row].cells[col].formula, row, col, fx.row, fx.col) }); }

        }
    }

    return fxList;
}

function ip_MoveFormulaOrigon(GridID, formula, fromRow, fromCol, toRow, toCol, exlcudeIfInRange, includeIfInRange, rowIncrement, colIncrement, incrementRange, selectedRange) {

    //Changes a formula by working out the difference in change and moving it to a new location
    //wont change the indevidual range if it overlaps exlcudeIfInRange
    if (formula != null && fromRow != null && fromCol != null && toRow != null && toCol != null) {

        var rowInc = toRow - fromRow;
        var colInc = toCol - fromCol;
        
        formula = formula.replace(ip_GridProps['index'].regEx.range, function (val) {

            var range = ip_fxRangeObject(GridID,null, null, val);
            if (range != null) {


                if (exlcudeIfInRange && ip_IsRangeOverLap(GridID, exlcudeIfInRange, range)) { return val; }
                
                var includeIfInRangeOverlap = (includeIfInRange ? ip_IsRangeOverLap(GridID, includeIfInRange, range) : false);
                if (includeIfInRange != null && !includeIfInRangeOverlap) { return val; } // || overlap == "inside"

                var selectedRangeOverlap = (selectedRange ? ip_IsRangeOverLap(GridID, range, selectedRange) : false);
                if (selectedRange != null && selectedRangeOverlap != 'exact' && selectedRangeOverlap != 'inside') { return val; }
                
                if (incrementRange != null && ip_IsRangeOverLap(GridID, incrementRange, range)) {

                    if (!range.startRowLock) { range.startRow += rowIncrement; }
                    if (!range.startColLock) { range.startCol += colIncrement; }
                    if (!range.endRowLock) { range.endRow += rowIncrement; }
                    if (!range.endColLock) { range.endCol += colIncrement; }
                }
                else {
                                                           
                    if (!range.startRowLock) { range.startRow += rowInc; }
                    if (!range.startColLock) { range.startCol += colInc; }
                    if (!range.endRowLock) { range.endRow += rowInc; }
                    if (!range.endColLock) { range.endCol += colInc; }
                }

                return ip_fxRangeToString(GridID, range);
            }

        });

    }

    return formula;
}

function ip_ChangeFormulasForCellMove(GridID, CellUndoData, fromRow, fromCol, toRow, toCol, selectedRange) {

    var Effected = { rowData: [] }

    if (GridID == null || fromRow == null && fromCol == null && toRow == null && toCol == null) { return }
    
    //Update linked cell formuals
    var fxInRangeXX = ip_FormulasInRange(GridID, ip_rangeObject(fromRow, fromCol, fromRow, fromCol), true, true);
    for (var key in fxInRangeXX) {
        
        var fx = ip_GridProps[GridID].indexedData.formulaData[key];
        //if (!fx) { continue; }
        var oldFormula = fx.formula;
        var formula = ip_MoveFormulaOrigon(GridID, fx.formula, fromRow, fromCol, toRow, toCol, null, null, null, null, null, selectedRange);
        
        if (oldFormula != formula) {
            //Create an undo
            ip_AddUndoTransactionData(GridID, CellUndoData, ip_CloneCell(GridID, fx.row, fx.col));
            ip_SetCellFormula(GridID, { row: fx.row, col: fx.col, formula: formula });
            ip_AppendEffectedRowData(GridID, Effected, { row: fx.row, col: fx.col, formula: formula });
        }

    }
    

    return Effected;
}

function ip_FormulasInRange(GridID, range, sources, linkIndexes) {

    //Fetches formulas within the specified range.
    //If sources, then looks for formulas which origonate in that range
    //If linkIndex, then looks for cells that are linked to the formula

    var FXList = {};

    if (range != null) {

        range = ip_ValidateRangeObject(GridID, range);
        var rangeSize = (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);

        //Determine which iteration will be fastest
        if (rangeSize < Object.keys(ip_GridProps[GridID].indexedData.formulaData).length) { // ip_GridProps[GridID].indexedData.formulaData.length

            for (var r = range.startRow; r <= range.endRow; r++) {

                for (var c = range.startCol; c <= range.endCol; c++) {

                    if (sources && ip_GridProps[GridID].rowData[r].cells[c].fxIndex != null ) { FXList[ip_GridProps[GridID].rowData[r].cells[c].fxIndex] = true; }
                    else if (linkIndexes && ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c] != null && ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c].fxIndexLnk != null) {  FXList = $.extend(FXList, ip_GridProps[GridID].indexedData.cellIndex[r + '-' + c].fxIndexLnk);   }

                }

            }

        }
        else
        {
            //Iterate through all the indexex formulas we have
            for (var key in ip_GridProps[GridID].indexedData.formulaData) {

                var fx = ip_GridProps[GridID].indexedData.formulaData[key];

                //Test the source method
                if (sources && fx.row >= range.startRow && fx.row <= range.endRow && fx.col >= range.startCol && fx.col <= range.endCol) { FXList[key] = true; }
                else if (linkIndexes)
                {

                    for (rangeKey in fx.ranges) {
                        if (ip_IsRangeOverLap(GridID, fx.ranges[rangeKey].range, range)) { FXList[key] = true; }
                    }

                }
            }
        }
    }

    return FXList;
}

function ip_SetColumnFormat(GridID, options) {

    var options = $.extend({

        transactionID: ip_GenerateTransactionID(),
        render:true,
        raiseEvent: true,
        columns: null, // [],  array of columns
        style: null,
        border: null,
        validation: null, //ip_validationObject()
        dataType: null, //ip_dataTypeObject()
        controlType: null, //dropdown, gantt, checkbox
        hashTags: null,
        mask: null,
        decimals: null,
        decimalsInc: null,
        PropertyAppendModes: 'style:append',
        createUndo: true,

    }, options);

    var TransactionID = options.transactionID;
    var Error = '';
    var Effected = null;
    var formatObject = ip_formatObject();

    //Validate columns

    Effected = { colData: new Array() }


    if (options.columns != null && Error == '') {

        options.columns.sort(function sortNumber(a, b) { return a - b; });

        //Record undo
        var defaultMaskInit = false;
        var FormatRange = { startRow: -1, startCol: options.columns[0], endRow: ip_GridProps[GridID].rows - 1, endCol: options.columns[options.columns.length -1 ] }
        if (options.createUndo) { var ColUndoData = ip_AddUndo(GridID, 'ip_SetColDataType', TransactionID, 'ColData', FormatRange, FormatRange, { row: FormatRange.startRow, col: FormatRange.startCol }); }

        for (var i = 0; i < options.columns.length; i++) {

            //FORMAT COLUMN
            var c = options.columns[i];

            if (options.createUndo) { ip_AddUndoTransactionData(GridID, ColUndoData, ip_CloneCol(GridID, c)); }

            Effected.colData[Effected.colData.length] = { col: c, style: null, dataType: null, controlType:null, hashTags:null, validation:null, decimalsInc: null, decimals:null, mask:null  };

            //Set column style
            if (options.style != null) {

                ip_GridProps[GridID].colData[c].style = (options.style == '' ? null : (options.PropertyAppendModes.indexOf('style:append') != -1 ? ip_AppendCssStyle(GridID, ip_GridProps[GridID].colData[c].style, options.style) : options.style));
                Effected.colData[Effected.colData.length - 1].style = (ip_GridProps[GridID].colData[c].style == null ? '' : ip_GridProps[GridID].colData[c].style);

            }

            //Set controlType
            if (options.controlType != null) {

                ip_GridProps[GridID].colData[c].controlType = (options.controlType == '' ? null : options.controlType);
                Effected.colData[Effected.colData.length - 1].controlType = options.controlType;

            }

            //Set hashTags
            if (options.hashTags != null) {

                ip_GridProps[GridID].colData[c].hashTags = (options.hashTags == '' ? null : options.hashTags);
                Effected.colData[Effected.colData.length - 1].hashTags = options.hashTags;

            }

            //Set border
            if (options.border != null) {

                ip_GridProps[GridID].colData[c].border = (options.border == '' ? null : options.border);
                Effected.colData[Effected.colData.length - 1].border = options.border;

            }
            
            //Set validation criteria
            if (options.validation != null) {

                ip_GridProps[GridID].colData[c].validation.validationCriteria = (options.validation.validationCriteria == '' ? null : options.validation.validationCriteria);
                ip_GridProps[GridID].colData[c].validation.validationAction = (options.validation.validationAction == '' ? null : options.validation.validationAction);

                Effected.colData[Effected.colData.length - 1].validation = ip_validationObject(options.validation.validationCriteria, options.validation.validationAction);
            }

            //Set column datatype
            if (options.dataType != null) {

                ip_GridProps[GridID].colData[c].dataType.dataType = (options.dataType.dataType == '' ? null : options.dataType.dataType);
                ip_GridProps[GridID].colData[c].dataType.dataTypeName = (options.dataType.dataTypeName == '' ? null : options.dataType.dataTypeName);

                Effected.colData[Effected.colData.length - 1].dataType = ip_dataTypeObject(options.dataType.dataType, options.dataType.dataTypeName);
                
                //Switch to a default mask for the datatype  
                if (!defaultMaskInit) {
                    options.mask = ip_DefaultMask(GridID, ip_GridProps[GridID].colData[c].dataType, (!options.mask ? ip_GridProps[GridID].colData[c].mask : options.mask));
                    options.mask = (options.mask == null ? '' : options.mask);                    
                    defaultMaskInit = true;
                }
            }


            //Set mask
            if (options.mask != null) {

                ip_GridProps[GridID].colData[c].mask = (options.mask == '' ? null : options.mask);
                Effected.colData[Effected.colData.length - 1].mask = options.mask;

            }

            //Set decimal places
            if (options.decimalsInc != null) {

                var decimals = ip_GridProps[GridID].colData[c].decimals;
                if (decimals == null) { decimals = 0; }
                ip_GridProps[GridID].colData[c].decimals = decimals + options.decimalsInc;

            }

            //Set decimal places
            if (options.decimals != null) { ip_GridProps[GridID].colData[c].decimals = options.decimals; }
            if (ip_GridProps[GridID].colData[c].decimals < 0) { ip_GridProps[GridID].colData[c].decimals = null; }

            formatObject = ip_GetEnabledFormats(GridID, formatObject, -1, c);

        }

        ip_AddDataTypeToIndex(GridID, options.dataType);

        //Rerender
        if (options.render) { ip_ReRenderCols(GridID); }
        
        if (options.raiseEvent && Effected.colData != null && Effected.colData.length > 0) { ip_RaiseEvent(GridID, 'ip_FormatColumn', TransactionID, { ColDataType: { Inputs: options, Effected: Effected } }); }

    }
    else if (Error == '') { Error = 'Please specify columns'; }

    if (Error != '') { ip_RaiseEvent(GridID, 'warning', TransactionID, Error); return null; }
    else { return formatObject; }


}

function ip_AddDataTypeToIndex(GridID, dataType) {

    if (GridID && dataType) {

        var found = false;
        var dataTypeIndex = ip_GridProps[GridID].dataTypes;

        for(var i =0; i < dataTypeIndex.length; i++){

            if (dataTypeIndex[i].dataTypeName == dataType.dataTypeName && dataTypeIndex[i].dataType == dataType.dataType) { found = true; break; }

        }

        if (!found) { dataTypeIndex.push(dataType); }
        
    }

}

function ip_ResetGridValues(GridID, loading) {


    //Reset grid 
    for (var r = 0; r < ip_GridProps[GridID].rowData.length; r++) {

        ip_GridProps[GridID].rowData[r].loading = loading;

        for (var c = 0; c < ip_GridProps[GridID].rowData[r].cells.length; c++) {

            ip_SetValue(GridID, r, c, null);

        }

    }

}

function ip_ResetCell(GridID, row, col, preserveMerge, preserveFormulaIndex, preserverCellFormula) {
    
    if (row >= 0 && col >= 0) {
        if (ip_GridProps[GridID].rowData.length > row) {
            if (ip_GridProps[GridID].rowData[row].cells.length > col) {

                var cell = ip_cellObject();

                if (preserveMerge == true) { cell.merge = ip_GridProps[GridID].rowData[row].cells[col].merge; }
                else { ip_ResetCellMerge(GridID, row, col); }

                if (preserveFormulaIndex == true) { cell.fxIndexLnk = ip_GridProps[GridID].rowData[row].cells[col].fxIndexLnk; }

                if (!preserverCellFormula) { ip_RemoveCellFormulaIndex(GridID, { row: row, col: col }); } //Reset the cells formula           
                ip_GridProps[GridID].rowData[row].cells[col] = cell;                

            }
        }
    }

}

function ip_GetColumnValues(GridID, col) {

    var valuesList = {};
    var values = []; //{ displayField: 'dataTypeName', style: 'font-weight:bold;' }
    
    for (var r = 0; r < ip_GridProps[GridID].rowData.length; r++) {

        var val = ip_GridProps[GridID].rowData[r].cells[col].value;
        if (valuesList[val] == null) {
            valuesList[val] = true;
            values[values.length] = { displayField: val };
        }

    }

    return values;
}

function ip_AppendCssStyle(GridID, currentStyle, newStyle) {
    //This method replaces a css style with a new one if it already exists

    if (newStyle != null && newStyle != '') {

        if (newStyle.indexOf(';', newStyle.length - 1) == -1) { newStyle += ';'; };
        
        if (currentStyle != '' && currentStyle != null) {

            var newProperties = newStyle.split(';');

            for (var np = 0; np < newProperties.length; np++) {

                var property = newProperties[np].split(':');
                if (property.length >= 2) { currentStyle = ip_ReplaceCssProperty(GridID, currentStyle, property[0], property[1]); }

            }
        }
        else
        {
            currentStyle = newStyle;
        }
   
    }

    return currentStyle;
}

function ip_GetCssProperty(GridID, style, property) {
    
    var result = '';

    if (style != '' && property != '') {

        var startI = style.indexOf(property + ':');        

        if (startI != -1) {
                        
            startI += property.length + 1;

            var endI = style.indexOf(';', startI);
            if (endI == -1) { endI = style.length; }
            result = style.substring(startI, endI);

        }
    }

    return result;
}

function ip_EnabledFormats(GridID, options)
{
    var options = $.extend({
                
        range: null, // [{ startRow:null, startCol: null, endRow: null, endCol: null }],    
        maxCells: 1000, //Because this is resource intensive, this wont get properties for ranges bigger than this
        getStyle: true,
        getDataType: false,
        getControlType: false,
        getValidation: false,
        getHashTags: false,
        getMask: false,
        getDecimals: false,
        adviseDefault: null // if this is set to true, it provides the column default if the cell properties are null

    }, options);

    var Error = '';
    var TransactionID = ip_GenerateTransactionID();
    var formatObject = ip_formatObject();

    //Validate range
    if (options.range == null && ip_GridProps[GridID].selectedRange.length == 0 && ip_GridProps[GridID].selectedColumn.length == 0 && ip_GridProps[GridID].selectedRow.length == 0) { Error = 'Please specify a range'; }
    else if (options.range == null) {
              
        //Copied ranges
        options.range = new Array();
        if (ip_GridProps[GridID].selectedRange.length > 0) {
            for (var i = 0; i < ip_GridProps[GridID].selectedRange.length; i++) { options.range[i] = ip_rangeObject(ip_GridProps[GridID].selectedRange[i][0][0], ip_GridProps[GridID].selectedRange[i][0][1], ip_GridProps[GridID].selectedRange[i][1][0], ip_GridProps[GridID].selectedRange[i][1][1], null, options.col); }
        }
        else if (ip_GridProps[GridID].selectedColumn.length > 0) {
            for (var i = 0; i < ip_GridProps[GridID].selectedColumn.length; i++) { options.range[i] = ip_rangeObject(0, ip_GridProps[GridID].selectedColumn[i], ip_GridProps[GridID].rows - 1, ip_GridProps[GridID].selectedColumn[i], null, null); }
        }
        else if (ip_GridProps[GridID].selectedRow.length > 0) {
            for (var i = 0; i < ip_GridProps[GridID].selectedRow.length; i++) { options.range[i] = ip_rangeObject(ip_GridProps[GridID].selectedRow[i], 0, ip_GridProps[GridID].selectedRow[i], ip_GridProps[GridID].cols - 1, null); }
        }

    }

    
    if (Error == '') {

        var cellCount = 0;

        for (var i = 0; i < options.range.length; i++) {

            if (options.adviseDefault == null && options.range[i].startRow <= 0 && options.range[i].endRow >= ip_GridProps[GridID].rows - 1) { options.adviseDefault = true; }
            if (options.range[i].startRow < -1) { options.range[i].startRow = -1; }
            if (options.range[i].startCol < 0) { options.range[i].startCol = 0; }
            

            for (var r = options.range[i].startRow; r <= options.range[i].endRow; r++)
            {
                for (var c = options.range[i].startCol; c <= options.range[i].endCol; c++) {

                    cellCount++;

                    //Reteive formatting relating to style
                    if (options.getStyle) {

                        formatObject = ip_GetEnabledFormats(GridID, formatObject, r, c);
                        if (formatObject.fontFamily == '') { formatObject.fontFamily = 'default' }
                        if (formatObject.fontSize == '') { formatObject.fontSize = 'default' }

                    }
                    
                    //Retreive formatting relating to datatype
                    if(options.getDataType){ ip_GetEnabledDataType(GridID, formatObject, r, c, options.adviseDefault); }
                    if (options.getControlType) { ip_GetEnabledControlType(GridID, formatObject, r, c, options.adviseDefault); }
                    if (options.getValidation) { ip_GetEnabledValidation(GridID, formatObject, r, c, options.adviseDefault); }
                    if (options.getHashTags) { ip_GetEnabledHashTags(GridID, formatObject, r, c, options.adviseDefault); }
                    if (options.getMask) { ip_GetEnabledMask(GridID, formatObject, r, c, options.adviseDefault); }
                    if (options.getDecimals) { ip_GetEnabledDecimals(GridID, formatObject, r, c, options.adviseDefault); }
                    
                    if (options.maxCells != null && cellCount >= options.maxCells) {
                    
                        //Break out if we reach the maximum amount of cells we are allowed to test for
                        r = options.range[i].endRow + 1;
                        c = options.range[i].endCol + 1;                        

                    }

                }
                
            }

            if (options.maxCells != null && cellCount >= options.maxCells) { i = options.range.length; }

        }

        

    }

    if (formatObject.fontFamily == 'default') { formatObject.fontFamily = '' }
    if (formatObject.fontSize == 'default') { formatObject.fontSize = '' }
    if (!formatObject.validation.validationAction) { formatObject.validation.validationAction = '' }
    if (!formatObject.validation.validationCriteria) { formatObject.validation.validationCriteria = '' }
    if (!formatObject.dataType.dataType) { formatObject.dataType.dataType = '' }
    if (!formatObject.dataType.dataTypeName) { formatObject.dataType.dataTypeName = '' }

    return formatObject;
    //if (Error != '') { ip_RaiseEvent(GridID, 'warning', TransactionID, Error); return Error; }
}

function ip_GetEnabledHashTags(GridID, formatObject, row, col, adviseDefault) {

    if (formatObject) {

        if (formatObject.hashTags == null) {

            if (row >= 0) { formatObject.hashTags = (ip_GridProps[GridID].rowData[row].cells[col].hashTags != null || ip_GridProps[GridID].colData[col].hashTags != null ? (ip_GridProps[GridID].rowData[row].cells[col].hashTags || !adviseDefault ? ip_GridProps[GridID].rowData[row].cells[col].hashTags : ip_GridProps[GridID].colData[col].hashTags) : null); }
            else { formatObject.hashTags = ip_GridProps[GridID].colData[col].hashTags; }

        }

        return formatObject;
    }
    else {
        if (row < 0) { return ip_GridProps[GridID].colData[col].hashTags; }
        else { return (ip_GridProps[GridID].rowData[row].cells[col].hashTags != null || ip_GridProps[GridID].colData[col].hashTags != null ? (ip_GridProps[GridID].rowData[row].cells[col].hashTags || !adviseDefault ? ip_GridProps[GridID].rowData[row].cells[col].hashTags : ip_GridProps[GridID].colData[col].hashTags) : ''); }
    }

}

function ip_GetEnabledDataType(GridID, formatObject, row, col, adviseDefault) {

    if (formatObject.dataType.dataType == null && formatObject.dataType.dataTypeName == null) {
        formatObject.dataType = ip_CellDataType(GridID, row, col, adviseDefault).expectedDataType;
        if (formatObject.dataType.dataType == '' && formatObject.dataType.dataTypeName == '') {

            formatObject.dataType.dataType = null;
            formatObject.dataType.dataTypeName = null;

        }
    }

    return formatObject;
}

function ip_GetEnabledControlType(GridID, formatObject, row, col, adviseDefault) {
       

    if (formatObject) {

        if (formatObject.controlType == null) {
                        
            if (row >= 0) { formatObject.controlType = (ip_GridProps[GridID].rowData[row].cells[col].controlType != null || ip_GridProps[GridID].colData[col].controlType != null ? (ip_GridProps[GridID].rowData[row].cells[col].controlType || !adviseDefault ? ip_GridProps[GridID].rowData[row].cells[col].controlType : ip_GridProps[GridID].colData[col].controlType) : null); }
            else { formatObject.controlType = ip_GridProps[GridID].colData[col].controlType; }

        }

        return formatObject;
    }
    else { return (ip_GridProps[GridID].rowData[row].cells[col].controlType != null || ip_GridProps[GridID].colData[col].controlType != null ? (ip_GridProps[GridID].rowData[row].cells[col].controlType ||  !adviseDefault ? ip_GridProps[GridID].rowData[row].cells[col].controlType : ip_GridProps[GridID].colData[col].controlType) : ''); }
    
}

function ip_GetEnabledValidation(GridID, formatObject, row, col, adviseDefault) {

    if (formatObject) {

        if (formatObject.validation.validationCriteria == null && formatObject.validation.validationAction == null) {

            if (row >= 0) {

                formatObject.validation.validationCriteria = (ip_GridProps[GridID].rowData[row].cells[col].validation.validationCriteria != null || !adviseDefault ? ip_GridProps[GridID].rowData[row].cells[col].validation.validationCriteria : ip_GridProps[GridID].colData[col].validation.validationCriteria);
                formatObject.validation.validationAction = (ip_GridProps[GridID].rowData[row].cells[col].validation.validationAction != null || !adviseDefault ? ip_GridProps[GridID].rowData[row].cells[col].validation.validationAction : ip_GridProps[GridID].colData[col].validation.validationAction);

            }
            else { formatObject.validation = ip_GridProps[GridID].colData[col].validation; }

            if (formatObject.validation.validationCriteria) { formatObject.validation.inputs = formatObject.validation.validationCriteria.match(/[^()]+(?=\))/g); }

        }

        return formatObject;

    }
    else
    {
        var validation = ip_validationObject();
        
        if (row >= 0) {

            validation.validationCriteria = (ip_GridProps[GridID].rowData[row].cells[col].validation.validationCriteria != null || !adviseDefault ? ip_GridProps[GridID].rowData[row].cells[col].validation.validationCriteria : ip_GridProps[GridID].colData[col].validation.validationCriteria);
            validation.validationAction = (ip_GridProps[GridID].rowData[row].cells[col].validation.validationAction != null || !adviseDefault ? ip_GridProps[GridID].rowData[row].cells[col].validation.validationAction : ip_GridProps[GridID].colData[col].validation.validationAction);
            
        }
        else { validation = ip_GridProps[GridID].colData[col].validation; }

        if (validation.validationCriteria) { validation.inputs = validation.validationCriteria.match(/[^()]+(?=\))/g); }

        return validation;
    }
}

function ip_GetEnabledFormats(GridID, formatObject, row, col) {
//Return the formatting for a range
    if (formatObject == null) { formatObject = ip_formatObject();  }

    var style = (row == -1 ? ip_GridProps[GridID].colData[col].style : ip_GridProps[GridID].rowData[row].cells[col].style);
    var styleC = ip_GridProps[GridID].colData[col].style;
        

    if (style != null || styleC != null) {

        if (ip_HasFormatting(GridID, style, styleC, 'font-weight', 'bold')) { formatObject.bold = true; }
        if (ip_HasFormatting(GridID, style, styleC, 'font-style', 'italic')) { formatObject.italic = true; }
        if (ip_HasFormatting(GridID, style, styleC, 'text-decoration', 'line-through')) { formatObject.linethrough = true; }
        if (ip_HasFormatting(GridID, style, styleC, 'text-decoration', 'underline')) { formatObject.underline = true; }
        if (ip_HasFormatting(GridID, style, styleC, 'text-decoration', 'overline')) { formatObject.overline = true; }

        if (ip_HasFormatting(GridID, style, styleC, 'text-align','left')) { formatObject.alignleft = true; }
        if (ip_HasFormatting(GridID, style, styleC, 'text-align','center', true)) { formatObject.aligncenter = true; }
        if (ip_HasFormatting(GridID, style, styleC, 'text-align','right')) { formatObject.alignright = true; }

        if (ip_HasFormatting(GridID, style, styleC, 'vertical-align','top')) { formatObject.aligntop = true; }
        if (ip_HasFormatting(GridID, style, styleC, 'vertical-align','middle', true)) { formatObject.alignmiddle = true; }
        if (ip_HasFormatting(GridID, style, styleC, 'vertical-align','bottom')) { formatObject.alignbottom = true; }


        if (formatObject.fontSize == '') { formatObject.fontSize = ip_GetCssProperty(GridID, style + styleC, 'font-size'); }
        if (formatObject.fontFamily == '') { formatObject.fontFamily = ip_GetCssProperty(GridID, style + styleC, 'font-family'); }

        if (formatObject.fill == '') { formatObject.fill = ip_GetCssProperty(GridID, style + styleC, 'background-color'); }
        if (formatObject.color == '') { formatObject.color = ip_GetCssProperty(GridID, style + styleC, '/**/color'); }
        
    }
    else
    {
        //SET DEFAULTS
        formatObject.alignmiddle = true;
        formatObject.aligncenter = true;
    }

    if (row == -1 && ip_GridProps[GridID].colData[col].containsMerges != null) { formatObject.merge = true; } else if (row == -1) { formatObject.hasUnmergedCells = true; }
    else if (row > -1 && ip_GridProps[GridID].rowData[row].cells[col].merge != null) { formatObject.merge = true; } else if (row > -1) { formatObject.hasUnmergedCells = true; }

    return formatObject;
}

function ip_GetEnabledDecimals(GridID, formatObject, row, col, adviseDefault) {

    if (formatObject && formatObject.decimals != null) { return formatObject; }

    var decimals = null;

    if (row < 0) { Mask = ip_GridProps[GridID].colData[col].decimals; }
    else {

        var cellDecimals = ip_GridProps[GridID].rowData[row].cells[col].decimals;
        var colDecimals = ip_GridProps[GridID].colData[col].decimals;
        var cellColDecimals = (cellDecimals != null || colDecimals != null ? (cellDecimals || !adviseDefault ? cellDecimals : colDecimals) : null);

        if (cellDecimals == null) { decimals = cellColDecimals; }
        else { decimals = cellDecimals; }//Use cells mask because it has a datatype

    }

    if (formatObject && formatObject.decimals == null) { formatObject.decimals = decimals; return formatObject; }

    return decimals;

}

function ip_GetEnabledMask(GridID, formatObject, row, col, adviseDefault) {

    if (formatObject && formatObject.mask != null) { return formatObject; }

    var Mask = null;

    if (row < 0) { Mask = ip_GridProps[GridID].colData[col].mask; }
    else {

        var cellMask = ip_GridProps[GridID].rowData[row].cells[col].mask;
        var colMask = ip_GridProps[GridID].colData[col].mask;
        var cellColMask = (cellMask != null || colMask != null ? (cellMask || !adviseDefault ? cellMask : colMask) : '');
        
        if (!cellMask) { Mask = cellColMask; }
        else { Mask = cellMask; }//Use cells mask because it has a datatype        
                    
    }

    if (formatObject && formatObject.mask == null) { formatObject.mask = Mask; return formatObject; }
    
    return Mask;
    
}

function ip_GetMaskObj(GridID, row, col, adviseDefault, mask) {
    //typeof mask == 'undefined' || mask == ''
    var mask = (!mask ? ip_GetEnabledMask(GridID, null, row, col, true) : mask);
    if (mask != null && mask != '' & ip_GridProps[GridID].mask.input[mask] != null) {
        return { mask: mask, input: ip_GridProps[GridID].mask.input[mask], output: ip_GridProps[GridID].mask.output[mask] }
    }
    return null;
}

function ip_HasFormatting(GridID, cellStyle, colStyle, styleName, styleValue, isDefaultStyle) {

    var combinedStyle = (cellStyle == null ? '' : cellStyle) + (colStyle == null ? '' : colStyle);
    var testedStyles = {};

    if (combinedStyle != '') {

        var styles = combinedStyle.split(';');        

        for (var i = 0; i < styles.length; i++) {

            var style = styles[i].split(':');

            if (testedStyles[style[0]] == null) {

                testedStyles[style[0]] = style[1];
                if (style[0] == styleName) {

                    if (style[1] == styleValue) { return true; }
                    else { return false; }
                    
                }

            }
            else if (style[0] == styleName) { return false; }

        }

    }

    if (isDefaultStyle && testedStyles[styleName] == null) { return true; }

    return false;

}

function ip_ShowCellToolTips(GridID, cell) {

    if (cell == null) { cell = ip_GridProps[GridID].selectedCell; }
    


}

function ip_isCellEmpty(gridID, row, col) {

    var cell = ip_GridProps[gridID].rowData[row].cells[col];

    if (cell.merge != null) { return false }
    else if (cell.value != '' && cell.value != null) { return false }
    else if (cell.style != '' && cell.style != null) { return false }
    else if (cell.dataType.dataType != '' && cell.dataType.dataType != null) { return false }
    else if (cell.dataType.dataTypeName != '' && cell.dataType.dataTypeName != null) { return false }
    else if (cell.validation.validationCriteria != '' && cell.validation.validationCriteria != null) { return false }
    else if (cell.validation.validationAction != '' && cell.validation.validationAction != null) { return false }
    ///else if (ip_GridProps[gridID].rowData[row].loading == true) { return false; }

    return true;
}

function ip_cellBorder(GridID, row, col, borderStyle, borderSize, borderColor, borderPlacement) {

    //borderStyle: 'solid',
    //borderSize: 1,
    //borderColor: 'black',
    //borderPlacement: '', //all, 'top', 'right', bottom, left, inner, outer, horizontal, vertical, none
    if (!borderPlacement || borderPlacement.length == 0 || row < 0 || col < 0 || row >= ip_GridProps[GridID].rows || col >= ip_GridProps[GridID].cols) { return; }
    var cell = ip_GridProps[GridID].rowData[row].cells[col];
    var top = '';
    var left = '';

    if (borderStyle == 'none' && borderPlacement.length == 4) {
        cell.border = '';
        return;
    }
    else if (borderStyle == 'none') {
        borderStyle = 'solid';
        borderSize = 1;
        borderColor = 'transparent';
    }
       
    
    for (var i = 0; i < borderPlacement.length; i++) {
            
        if (borderPlacement[i] == 'top') { top = 'top:' + (0 - parseInt(borderSize)) + 'px;'; }
        if (borderPlacement[i] == 'left') { left = 'left:' + (0 - parseInt(borderSize)) + 'px;'; }
        var newBorder = border = 'border-' + borderPlacement[i] + ':' + (borderStyle == 'none' ? ';' : borderStyle + ' ' + borderSize + 'px ' + borderColor + ';');
        cell.border = ip_AppendCssStyle(GridID, cell.border, newBorder);        

    }

    if (top != '') { cell.border = ip_AppendCssStyle(GridID, cell.border, top); }
    if (left != '') { cell.border = ip_AppendCssStyle(GridID, cell.border, left); }

    if (cell.border.indexOf('border-top') == -1) { cell.border += 'border-top:solid 1px transparent;' }
    if (cell.border.indexOf('border-left') == -1) { cell.border += 'border-left:solid 1px transparent;' }
     
}


//----- EDIT TOOL FUNCTIONS  ------------------------------------------------------------------------------------------------------------------------------------

function ip_TextEditbleBlur(GridID, defaultValue, editTool, editToolInput, elementToEdit, contentType, row, col, selectCell, cancel) {

    if (!ip_GridProps[GridID].editing.editing) { return; }

    ip_GridProps[GridID].editing.editing = false;
        
    editTool = (editTool == null ? ip_GridProps[GridID].editing.editTool : editTool);
    editToolInput = (editToolInput == null ? $(editTool).children('.ip_grid_EditTool_Input')[0] : editToolInput);
    elementToEdit = (elementToEdit == null ? ip_GridProps[GridID].editing.element : elementToEdit);
    contentType = (contentType == null ? ip_GridProps[GridID].editing.contentType : contentType);
    defaultValue = (defaultValue == null ? editTool.text() : defaultValue);
    row = (row == null ? parseInt($(elementToEdit).attr('row')) : row);
    col = (col == null ? parseInt($(elementToEdit).attr('col')) : col);

    var dropDown = $(editTool).children('.ip_grid_EditTool_DropDown');
    var Error = '';                

    //validate dropdown
    if (!cancel && $(dropDown).attr('validate') == 'true') {

        var valid = false;
        var NewValue = ip_fxValue(GridID, ip_GridProps[GridID].editing.editTool.text().toLowerCase()).value;
        var dropDownItems = $(dropDown).find('.ip_grid_EditTool_DropDownItem');

        if ($(dropDown).attr('allowEmpty') == 'false' || NewValue != '') {
            for (var i = 0; i < dropDownItems.length; i++) {

                if ($(dropDownItems[i]).attr('val').toLowerCase() == NewValue) {

                    ip_EditToolDropDownSelect(GridID, dropDownItems[i], false, true);
                    i = dropDownItems.length;
                    valid = true;
                }

            }
        }
        else { valid = true; }

        if (!valid) { Error = 'Please make sure you choose one of the dropdown items.' }
    }

    if (Error == '') {

        row = (row == null ? -1 : row);
        col = (col == null ? -1 : col);


        ip_GridProps[GridID].events.textEditTool_Blur = ip_UnBindEvent(editTool, 'focusout', ip_GridProps[GridID].events.textEditTool_FocusOut);

        if (!cancel) {

            if ((contentType == 'ip_grid_CellEdit' || contentType == 'ip_grid_fxBar') && row != -1 && col != -1) {

                //Normal Cell   
                var input = ip_GridProps[GridID].editing.editTool.text();

                ip_CellInput(GridID, { row: row, col: col, valueRAW: input });
                ip_SetFxBarValue(GridID, { row: row, col: col, cell: elementToEdit });


            }
            else if (contentType == 'ip_grid_ColumnType' && col != -1) {

                //Column Header
                var NewDataType = $(editToolInput).attr('key');
                var NewDataTypeName = ip_GridProps[GridID].editing.editTool.text();

                ip_SetCellFormat(GridID, { row:-1, col:col, dataType: { dataType: NewDataType, dataTypeName: NewDataTypeName } });
                //ip_SetColumnFormat(GridID, { columns: [col], dataType: { dataType: NewDataType, dataTypeName: NewDataTypeName }, mask: '' });
                $(elementToEdit).find('.ip_grid_cell_innerContent').html(NewDataTypeName);

            }
        }

        ip_DisableSelection(GridID, true);
        ip_EditToolHelpClear(GridID, editTool, editToolInput);      
        $(editTool).hide();
        

        ip_GridProps[GridID].editing.row = -1;
        ip_GridProps[GridID].editing.col = -1;
        ip_GridProps[GridID].editing.element = null;
        ip_GridProps[GridID].editing.contentType = '';

        if (selectCell) { $('#' + GridID).ip_SelectCell({ raiseEvent: false, unselect: false }); }
                        
        return true;
    }
    else {

        $(editToolInput).focus();
        ip_RaiseEvent(GridID, 'warning', null, Error);

    }

    return false;
}

function ip_TextEditbleCancel(GridID, editTool, editToolInput, elementToEdit, contentType, row, col, selectCell) {

    elementToEdit = (elementToEdit == null ? ip_GridProps[GridID].editing.element : elementToEdit);
    row = (row == null ? parseInt($(elementToEdit).attr('row')) : row);
    col = (col == null ? parseInt($(elementToEdit).attr('col')) : col);

    ip_TextEditbleBlur(GridID, null, editTool, editToolInput, elementToEdit, contentType, row, col, false, true);

    ip_ReRenderCell(GridID, row, col, selectCell);
}

function ip_EditToolShowDropDown(GridID, editTool, dropDownData) {

    //dropDownData: { allowEmpty: true, validate: false, autoComplete: true, keyField: null, displayField: null, data: [], noMatchData: null }

    if (editTool != null && dropDownData != null) {
        

        var dropDown = $(editTool).children('.ip_grid_EditTool_DropDown')[0];

        //fetch the dropdown data from function
        if (typeof (dropDownData.data) == 'string') {
            
            var fx = ip_fxValue(GridID, dropDownData.data, null, null).data;
            if (fx != null) {

                dropDownData.displayField = fx.displayField;
                dropDownData.data = fx.data;
            }
            else { dropDownData.data = []; }
        }
        

        //Setup dropdown values
        if (dropDownData.data != null && dropDownData.data.length > 0) {

            ip_GridProps[GridID].editing.editing = true;

            var dropDownContent = '<table border="0" cellpadding="0" cellspacing="0" class="ip_grid_EditTool_DropDownTable">';

            //Add non match data
            if (dropDownData.noMatchData != null) {

                var replaceWithVal = false;
                var aval = '';
                var val = '';
                var acontent = '';
                var akey = dropDownData.noMatchData[dropDownData.keyField];

                if (typeof dropDownData.displayField === 'string') {

                    replaceWithVal = dropDownData.noMatchData[dropDownData.displayField[b].displayField] == null ? true : false;
                    aval = dropDownData.noMatchData[dropDownData.displayField];
                    val = aval;
                    if (aval == null) { aval = ''; }
                    if (val == null) { val = '..'; }
                    acontent = '<td replaceWithInput="' + replaceWithVal + '">' + aval + '</td>';

                }
                else if (dropDownData.displayField.length > 0) {
                    for (var b = 0; b < dropDownData.displayField.length; b++) {

                        val = dropDownData.noMatchData[dropDownData.displayField[b].displayField];
                        replaceWithVal = dropDownData.noMatchData[dropDownData.displayField[b].displayField] == null ? true : false;
                        if (b == 0) { aval = dropDownData.noMatchData[dropDownData.displayField[b].displayField] }
                        if (val == null) { val = '..'; }
                        if (aval == null) { aval = ''; }
                        acontent += '<td replaceWithInput="' + replaceWithVal + '" style="' + dropDownData.displayField[b].style + ';" >' + val + '</td>';

                    }
                }

                dropDownContent += '<tr key="' + akey + '" val="' + aval + '" class="ip_grid_EditTool_DropDownItem noMatch">' + acontent + '</tr>';
            }

            //for (t = 0; t < 30; t++) { //<-- this is debug
            //Add regular dropdown items
            for (var a = 0; a < dropDownData.data.length; a++) {

                var aval = '';
                var acontent = '';
                var akey = dropDownData.data[a][dropDownData.keyField];

                if (typeof dropDownData.displayField === 'string') {
                    
                    aval = dropDownData.data[a][dropDownData.displayField]
                    acontent = '<td>' + aval + '</td>';

                }
                else if (dropDownData.displayField.length > 0) {
                    for (var b = 0; b < dropDownData.displayField.length; b++) {

                        if (b == 0) { aval = dropDownData.data[a][dropDownData.displayField[b].displayField] }
                        acontent += '<td style="' + dropDownData.displayField[b].style + ';" >' + dropDownData.data[a][dropDownData.displayField[b].displayField] + '</td>';

                    }
                }

                dropDownContent += '<tr key="' + akey + '" val="' + aval + '" class="ip_grid_EditTool_DropDownItem">' + acontent + '</tr>';

            }
            //}

            dropDownContent += '</table>';

            $(dropDown).html(dropDownContent);


            var dropDownItems = $(dropDown).find('.ip_grid_EditTool_DropDownItem');
            var MaxDropDownHeight = (ip_GridProps[GridID].dimensions.gridHeight - ip_GridProps[GridID].dimensions.defaultBorderHeight - parseInt($('#' + GridID + '_q4_scrollbar_container_x').height()) - ip_GridProps[GridID].dimensions.columnSelectorHeight);
            var minWidth = $(editTool).width();

            $(dropDown).attr('validate', dropDownData.validate);
            $(dropDown).attr('allowEmpty', dropDownData.allowEmpty);
            $(dropDown).css('max-height', MaxDropDownHeight + 'px');       
            $(dropDown).show();

            //Calculate dropdown width
            var ddTableWidth = $(dropDown).children('.ip_grid_EditTool_DropDownTable').width() - 8;            
            if (minWidth > ddTableWidth) { ddTableWidth = minWidth; $(dropDown).children('.ip_grid_EditTool_DropDownTable').css('width', '100%'); }
            $(dropDown).width(ddTableWidth);
            
            //Setup DropDownEvents
            ip_UnBindEvent(dropDown, 'focusin', ip_GridProps[GridID].events.textEditToolDropdown_FocusIn);
            $(dropDown).focusin(ip_GridProps[GridID].events.textEditToolDropdown_FocusIn = function (e) {

                clearTimeout(ip_GridProps[GridID].timeouts.textEditToolBlurTimeout);

            });

            //Select a drop down item
            ip_UnBindEvent(dropDownItems, 'click', ip_GridProps[GridID].events.textEditToolDropdown_Click);
            $(dropDownItems).click(ip_GridProps[GridID].events.textEditToolDropdown_Click = function (e) {

                ip_EditToolDropDownSelect(GridID, this, true, true);

            });

            //Highlight a dropdown item
            ip_UnBindEvent(dropDownItems, 'mouseenter', ip_GridProps[GridID].events.textEditToolDropdown_Mouseenter);
            $(dropDownItems).mouseenter(ip_GridProps[GridID].events.textEditToolDropdown_Mouseenter = function (e) {

                ip_EditToolDropDownSelect(GridID, this, false, false);

            });

            

        }
        else {

            $(dropDown).attr('validate', false);
            $(dropDown).html('');
            $(dropDown).hide();
        }
    }
}

function ip_EditToolDropDownHighlightNext(GridID, Direction) {


    clearTimeout(ip_GridProps[GridID].textEditToolDropdownScrollTimeout);

    var editTool = ip_GridProps[GridID].editing.editTool;
    var dropDown = $(editTool).children('.ip_grid_EditTool_DropDown')[0];
    var dropDownItems = $(dropDown).find('.ip_grid_EditTool_DropDownItem:not(.filtered)');
    var dropDownItemSelected = $(dropDown).find('.Selected');
       


    if (dropDownItemSelected.length == 0) { ip_EditToolDropDownSelect(GridID, dropDownItems[0], false, true); }
    else {
        var Index = $(dropDownItems).index(dropDownItemSelected);

        if (Direction == 'down' && Index < dropDownItems.length - 1) { 
            
            ip_EditToolDropDownSelect(GridID, dropDownItems[Index + 1], false, true);


            //Scroll to selected element
            var dropDownHeight = $(dropDown).height();
            var dropDownScrollTop = $(dropDown).scrollTop();
            var elementHeight = $(dropDownItems[Index + 1]).outerHeight();
            var elementTop = ($(dropDownItems[Index + 1]).position().top + dropDownScrollTop - elementHeight);
            var scrollTo = elementTop - (dropDownHeight / 2);

            if ((dropDownHeight + dropDownScrollTop) < elementTop) {
                                
                $(dropDown).animate({
                    scrollTop: scrollTo
                }, 200);

            }


        }
        else if (Direction == 'up' && Index > 0) {

           
            ip_EditToolDropDownSelect(GridID, dropDownItems[Index - 1], false, true);

            //Scroll to selected element
            var dropDownHeight = $(dropDown).height();
            var dropDownScrollTop = $(dropDown).scrollTop();
            var elementHeight = $(dropDownItems[Index - 1]).outerHeight();
            var elementTop = ($(dropDownItems[Index - 1]).position().top + dropDownScrollTop - elementHeight);
            var scrollTo = elementTop - (dropDownHeight / 2);

            if ((dropDownScrollTop) > elementTop) {
                                
                $(dropDown).animate({
                    scrollTop: scrollTo
                }, 200);

            }

        }
    }

    

}

function ip_EditToolDropDownSelect(GridID, DropDownItem, commit, select) {

    var editTool = ip_GridProps[GridID].editing.editTool;
    var dropDown = $(editTool).children('.ip_grid_EditTool_DropDown')[0];

    $(dropDown).find('.ip_grid_EditTool_DropDownItem').removeClass('Selected');
    $(DropDownItem).addClass('Selected');

    if (select) {
        ip_GridProps[GridID].editing.editTool.text($(DropDownItem).attr('val'), $(DropDownItem).attr('key'));           
    }

    if (commit) { ip_TextEditbleBlur(GridID, '', editTool, null, ip_GridProps[GridID].editing.element, ip_GridProps[GridID].editing.contentType, ip_GridProps[GridID].editing.row, ip_GridProps[GridID].editing.col); }
}

function ip_EditToolDropDownFilter(GridID, text) {


    var editTool = ip_GridProps[GridID].editing.editTool;
    var dropDown = $(editTool).children('.ip_grid_EditTool_DropDown')[0];
    var dropDownItems = $(dropDown).find('.ip_grid_EditTool_DropDownItem');
    var dropDownItemNoData = $(dropDown).find('.ip_grid_EditTool_DropDownItem.noMatch');
    var matchFound = false;
    var testText = text.toLowerCase().trim();
    var itemCount = 0;
    
    if (testText == '') {
        $(dropDownItems).removeClass('filtered');
    }
    else {

        
        for (var i = 0; i < dropDownItems.length; i++) {

            var value = $(dropDownItems[i]).attr('val').toLowerCase().trim();

            if (value.indexOf(testText) > -1) { $(dropDownItems[i]).removeClass('filtered'); itemCount++; }
            else { $(dropDownItems[i]).addClass('filtered'); }

            if (testText == value && !$(dropDownItems[i]).hasClass('noMatch')) { matchFound = true; }
        }


    }

    //Check if an exact match is found for typed in value and if so put in the dropdown
    if (dropDownItemNoData.length > 0) {

        if (!matchFound) {

            $(dropDownItemNoData).removeClass('filtered');
            $(dropDownItemNoData).find('td').each(function () {

                if ($(this).attr('replaceWithInput') == 'true') {
                    
                    var replaceText = text + '..';
                    $(this).html(replaceText);
                    $(this).parent().attr('val', text);
                }

            });
        }
        else { $(dropDownItemNoData).addClass('filtered'); }

        $(dropDown).animate({ width: $(dropDown).children('.ip_grid_EditTool_DropDownTable').width() },200);
    }

    //if (itemCount == 0) { $(dropDown).css('overflow-y', 'hidden'); }
    //else { $(dropDown).css('overflow-y', 'auto'); }

}

function ip_EditToolHelp(GridID, text, carret, carretIncr, showToolTip, setCursor, setEditTool) {

    //Shows typing assistance for the edit tool
    //Only GridID is required
    if (text == null || text == '') { if (setEditTool) { ip_GridProps[GridID].editing.editTool.text(''); } return false; }

    editTool = ip_GridProps[GridID].editing.editTool;       
    carret = (carret == null ? { x: text.lengh } : carret);

    var fx = false;
    var cStart = carret.x + carretIncr;
    var cEnd = (carret.length == null || carret.length == 0 ? null : carret.x + carret.length);

    //Show formula help
    if (text[0] == '=') {

        carret = (carret == null ? carret = { x: text.length, length: 0 } : carret); //ip_GetCursorPos(GridID, editToolInput)
        text = (text == null ? ip_GridProps[GridID].editing.editTool.text() : text);

        fx = ip_fxObject(GridID, text, null, null, cStart - carretIncr);

        //fx.formatted = fx.formatted.replace(/\s$/, "&nbsp;");
        //fx.formatted = fx.formatted.replace(/(?![^<]*[>])\s/, "&nbsp;");

        $('#' + GridID).ip_RemoveRangeHighlight({ highlightType: 'ip_grid_cell_rangeHighlight_fx' });
        for (var key in fx.ranges) { $('#' + GridID).ip_RangeHighlight({ highlightType: 'ip_grid_cell_rangeHighlight_fx', multiselect: true, opacity: 0.2, fillColor: fx.ranges[key].color, borderColor: fx.ranges[key].color, range: fx.ranges[key].range }); }

        text = fx.formatted;
    }


    if (showToolTip && fx) { ip_ShowEditToolHelpToolTip(GridID, ip_fxInfo(GridID, fx.focused.value)); };// else { $(helpTool).hide(); }
    if (setEditTool) { ip_GridProps[GridID].editing.editTool.text(text, null, cStart, cEnd); }

    return fx;      
    
                 
     
}

function ip_EditToolHelpClear(GridID, editTool, editToolInput) {
    
    editTool = (editTool == null ? ip_GridProps[GridID].editing.editTool : editTool);
    editToolInput = (editToolInput == null ? $(editTool).children('.ip_grid_EditTool_Input')[0] : editToolInput);

    $('#' + GridID).ip_RemoveRangeHighlight({ fadeOut:true, highlightType: 'ip_grid_cell_rangeHighlight_fx' });
    $(editToolInput).removeClass('ip_grid_EditTool_Input_FX');

    $().ip_Modal({ show: false});
}

function ip_ShowEditToolHelpToolTip(GridID, fxInfo) {

    var fxToolTip = '';

    if (fxInfo != null) {

        var modalSessionID = fxInfo.name;
        var fxToolTip = '<div class="ip_grid_EditTool_HelpToolTipItem Title"><b>' + fxInfo.name.toUpperCase() + ' ' + fxInfo.inputs + '</b></div>';
        fxToolTip += '<div class="ip_grid_EditTool_HelpToolTipItem Text">' + fxInfo.tip + '</div>';
        fxToolTip += '<div class="ip_grid_EditTool_HelpToolTipItem Example"><span>Example: </span>' + fxInfo.example + '</div>';
        

    }

    if (fxToolTip != '') {

        $().ip_Modal({
            modalSessionID: modalSessionID, 
            greyOut: false,
            axisContainment:true,
            cssClass: 'ip_grid_EditTool_Help',
            animate: 0,
            relativeTo: '.ip_grid_EditTool_Input',
            message: fxToolTip,
            position: null, 
            positionNAN: ['right', 'left'], //['right', 'left']
            fade:100,
            buttons: {
                //anchor: { text: 'ANCHOR', style: "background-color:#4545a1;", onClick: function () { $().ip_Modal({ show: false }) } },
                hide: { text: 'HIDE', style: "background-color:#9c3b3b;", onClick: function () { $().ip_Modal({ show: false }) } },                
                ok: { },
                cancel: { }
            }
        });

    }
    else {
        $().ip_Modal({ show:false, fade:100 });
    }

    return fxToolTip;
}


//----- FX BAR FUNCTION  ------------------------------------------------------------------------------------------------------------------------------------

function ip_SetFxBarValue(GridID, options) {

    var options = $.extend({

        value: null,
        cell: ip_GridProps[GridID].selectedCell, //cell td object
        row: null,
        col: null,

    }, options);

    var value = options.value;
    var cell = options.cell;
    var row = options.row;
    var col = options.col;
    
    if (row == null && cell != null) { row = parseInt($(cell).attr('row')); }
    if (col == null && cell != null) { col = parseInt($(cell).attr('col')); }

    if ((row != null && col != null) || cell != null) {

        if (row == null) { row = parseInt($(cell).attr('row')); }
        if (col == null) { col = parseInt($(cell).attr('col')); }

        var cellData = ip_CellData(GridID, row, col);
        

        if (cellData != null) {

            var fx = ip_fxObject(GridID, cellData.formula, row, col, 0);
            var value = (!fx ? cellData.value : fx.formatted);

            if (value == null) { value = '&nbsp;'; }
        }

    }
    
    if (value != null) {
        var fBar = ip_GridProps[GridID].fxBar;
        $(fBar).children('.ip_grid_fbar_text').html(value);
    }

    
}


//----- MASK FUNCTIONS  ------------------------------------------------------------------------------------------------------------------------------------

function ip_DefaultMask(GridID, dataType, currentMask) {
//Gets the default mask for a datatype, however you can specify the current mask, if the current mask is of that datatype, it will use it
    if (dataType == null) { return null; }

    var masks = ip_GetMasksForDataType(GridID, dataType);

    if (masks != null && masks.length > 0) {
        if (currentMask) {
            var index = masks.map(function (e) { return e.mask; }).indexOf(currentMask);
            if (index != -1) { return masks[index].mask; }
        }
        return masks[0].mask;
    }


    return null;
}

function ip_GetMasksForDataType(GridID, dataType) {

    if (!dataType) { return null; }

    var Masks = [];
    var dataTypeName = dataType.dataTypeName == null ? '' : dataType.dataTypeName.toLowerCase();
    var dataTypeType = dataType.dataType == null ? '' : dataType.dataType.toLowerCase();

    if (dataTypeName && dataTypeType != dataTypeName && ip_GridProps[GridID].mask[dataTypeName] != null) { Masks = Masks.concat(ip_GridProps[GridID].mask[dataTypeName]) }
    if (dataTypeType != null && ip_GridProps[GridID].mask[dataTypeType] != null) { Masks = Masks.concat(ip_GridProps[GridID].mask[dataTypeType]) }

    return Masks;

}


//----- FORMULA FUNCTIONS  ------------------------------------------------------------------------------------------------------------------------------------

function ip_fxInfo(GridID, fxName) {
    
    if (!ip_GridProps[GridID].fxList[fxName]) { return null; }
    return ip_GridProps[GridID].fxList[fxName].fxInfo;

}

function ip_fxValue(GridID, formula, row, col) {

    //Returns the calculated value of a formula, 
    //note: formulas MUST start with =
    var returnObj = { data: null, value: null, error: null, formula: formula }

    if (formula == null || formula == '' || formula[0] != '=') {
        var dataType = (!row || !col ? { value: formula  } :  ip_CellDataType(GridID, row, col, true, formula));
        returnObj.value = dataType.valid == true ? dataType.value : formula;
        returnObj.formula = null;
    }
    else {

        var result = ip_fxCalculate(GridID, formula.substring(1, formula.length), row, col);
             
        if (result != null && !result.errorCode) {

            if (typeof (result) == 'object') {
                returnObj.data = result;
                returnObj.value = result.toString(); //result.toString();
            }
            else { returnObj.value = result }

        }
        else if (result != null) {
            returnObj.data = null;
            returnObj.value = '#ERROR';
            returnObj.error = ip_errorObject(result.errorCode, result.errorDescription);
        }

    }

    return returnObj;
}

function ip_fxObject(GridID, formula, row, col, cursorI) {
    
    //Accepts a forumla string and returns it in a structured fx object
    //e.g. '=sum(A1,A2,A3) + 25 * A3'
    //Will break the formula for analysis up to cursorI -> this resource intensive leave out cursor I if analysis is not needed

    if (formula == null || formula == '' || formula[0] != '=') { return false; }

    var fxRanges = formula.match(ip_GridProps['index'].regEx.range);
    var fxFunctions = formula.match(ip_GridProps['index'].regEx.fx);
    var fxInputs = formula.match(/[^()]+(?=\))/g)
    
    var fx = {

        formatted: formula,
        formula: formula,
        inputs: fxInputs,
        ranges: {},
        functions: fxFunctions,
        row: row,
        col: col,
        focused: { value: '', type: '', partIndex:0, cursorI:0 },
        parts: []

    }

    

    //Find the method or range in focus
    if (cursorI != null) {

        var index = 0;
        fx.parts = formula.match(new RegExp(ip_GridProps['index'].regEx.range.source + '|' + ip_GridProps['index'].regEx.words.source + '|' + ip_GridProps['index'].regEx.nonWords.source, 'gi'));
        fx.focused.partIndex = fx.parts.length - 1;

        for (var f = 0; f < fx.parts.length; f++) {

            index += fx.parts[f].length;

            fx.focused.value = fx.parts[f];
            fx.focused.partIndex = f;
            fx.focused.cursorI = index - fx.parts[f].length; 

            if (index > cursorI) { fx.focused.value = fx.focused.value.trim(); break;  }            

        }
        
    }

    
    //Structure all the ranges in the formula
    if (fxRanges != null) {
        var hueInc = 255 / fxRanges.length;
        var hue = 0;
        for (var r = 0; r < fxRanges.length; r++) {

            if (fx.ranges[fxRanges[r]] == null) {

                var range = ip_fxRangeObject(GridID, row, col, fxRanges[r]);

                if (range != null) {

                    hue += hueInc;

                    fx.ranges[fxRanges[r]] = {
                        color: ip_ChangeHue('#ff0000', hue),
                        range: range,
                        rangeText: fxRanges[r],
                    }

                    //Set formatting
                    if (cursorI) {
                        //var replaced = expr.replace(/(^|[^:])a1\b(?!:)/gi, '$1foo');
				
                        var rangeval = fxRanges[r].replace(/\$/gi, '\\$');
                        var regExpColor = new RegExp(ip_GridProps['index'].regEx.notInQuotes.source + '(?!<[^>]*?>)' + rangeval + '(?![^<]*?</)(?![0-9:]+)', 'gi'); // + '(?![^<]*>|[^<>]*</)(?!:|</|[0-9])', 'gi');
                        
                        var matchdata = fx.formatted.match(regExpColor);
                        var replaceRange = fxRanges[r].replace(/\$/gi, '$$$');
                        fx.formatted = fx.formatted.replace(regExpColor, '<span style="color:' + fx.ranges[fxRanges[r]].color + ';">' + replaceRange + '</span>');
                  
                    }
                }
            }

        }
    }
    
    return fx;
}

function ip_fxException(errorCode, errorDescription, fxName, row, col) {
    return {
        errorCode: errorCode,
        errorDescription: errorDescription,
        fxName: fxName,
        row: row,
        col: col,
    }
}

function ip_fxRangeObject(GridID, row, col, fxRange) {

    //Accepts a range string e.g. A1:B3 and returns a validated range object
    if (fxRange != '' && fxRange != null) {

        var returnRanges = [];

        fxRange = fxRange.toUpperCase().trim();

        fxRange = fxRange.split(',');

        for (var x = 0; x < fxRange.length; x++) {

            var range = ip_rangeObject(null, null, null, null, null, null, fxRange[x].match(ip_GridProps['index'].regEx.hashtag));
            //fxRange[x] = fxRange[x].replace(ip_GridProps['index'].regEx.hashtag, '').split(':');
            var fxRanges = fxRange[x].replace(ip_GridProps['index'].regEx.hashtag, '').split(':'); //fxRange[x].split(':');

            for (var r = 0; r < fxRanges.length; r++) {

                var rangeStr = fxRanges[r].trim();
                var ColLetter = '';
                var RowNumber = '';
                var rLock = false;
                var cLock = false;

                for (var i = 0; i < rangeStr.length; i++) {

                    if (rangeStr[i] == '$') {
                        if (ColLetter == '') { cLock = true; }
                        else { rLock = true; }
                    }
                    else if (isNaN(parseInt(rangeStr[i]))) { ColLetter += rangeStr[i]; }
                    else { RowNumber += rangeStr[i]; }

                }

                if (ColLetter == '' || RowNumber == '') { return null; }

                if (r == 0) {

                    range.startRowLock = rLock;
                    range.startColLock = cLock;
                    range.endRowLock = rLock;
                    range.endColLock = cLock;
                    range.startRow = (RowNumber == '' ? 0 : parseInt(RowNumber));
                    range.startCol = ip_GridProps[GridID].indexedData.colSymbols.symbolCols[ColLetter];
                    range.endRow = (RowNumber == '' ? ip_GridProps[GridID].rows - 1 : range.startRow);
                    range.endCol = range.startCol;
                    if (range.startRow == null || range.startCol == null) { return null; }
                    if (fxRanges.length > 1) { r = fxRanges.length - 2; }

                }
                else {

                    range.endRowLock = rLock;
                    range.endColLock = cLock;
                    range.endRow = (RowNumber == '' ? ip_GridProps[GridID].rows - 1 : parseInt(RowNumber));
                    range.endCol = ip_GridProps[GridID].indexedData.colSymbols.symbolCols[ColLetter];
                    if (range.endRow == null || range.endCol == null) { return null; }

                }

            }

            if (range.startRow < 0) { range.startRow = 0; }
            if (range.startRow >= ip_GridProps[GridID].rows) { range.startRow = ip_GridProps[GridID].rows - 1; }
            if (range.endRow < 0) { range.endRow = 0; }
            if (range.endRow >= ip_GridProps[GridID].rows) { range.endRow = ip_GridProps[GridID].rows - 1; }

            if (range.startCol < 0) { range.startCol = 0; }
            if (range.startCol >= ip_GridProps[GridID].cols) { range.startCol = ip_GridProps[GridID].cols - 1; }
            if (range.endCol < 0) { range.endCol = 0; }
            if (range.endCol >= ip_GridProps[GridID].cols) { range.endCol = ip_GridProps[GridID].cols - 1; }

            returnRanges[returnRanges.length] = range;
        }

        return (returnRanges.length > 1? returnRanges : returnRanges[0]);
    }

    return null;

}

function ip_fxRangeToString(GridID, range, splitter) {

    //Accepts a range object or an ARRAY of range objects and returns the range as a string e.g. a1 or a1, b1, c1
    

    if (range != null) {

        var newRange = '';

        if (!range.length) { range = [range]; }
        if (!splitter) { splitter = ', '; }

        for (var i = 0; i < range.length; i++) {
                        
            var vrange = ip_ValidateRangeObject(GridID, range[i]);

            newRange += (vrange.startColLock ? '$' : '') + ip_GridProps[GridID].indexedData.colSymbols.colSymbols[vrange.startCol] + (vrange.startRowLock ? '$' : '') + vrange.startRow +
                            (vrange.startRow != vrange.endRow || vrange.startCol != vrange.endCol ? ':' + (vrange.endColLock ? '$' : '') + ip_GridProps[GridID].indexedData.colSymbols.colSymbols[vrange.endCol] + (vrange.endRowLock ? '$' : '') + vrange.endRow : '') +
                            (i < range.length - 1 ? splitter : '');
            if (vrange.hashtags) { newRange += vrange.hashtags.join(); }
        }
        return newRange;
    }

}

function ip_fxCalculate(GridID, fxString, row, col) {

    try {

        //
        var rxRootRanges = new RegExp(ip_GridProps['index'].regEx.notInBrackets.source + ip_GridProps['index'].regEx.range.source, 'gi');  // /(?=[^"]*(?:"[^"]*"[^"]*)*$)(?![^(]*[,)])[a-z]\d+(:\w+)?/gi; ///(?![^("]*[)"])(([a-z]+[0-9]+[:][a-z]+[0-9]+)|([a-z]+[0-9]+))/gi;
        
        ip_fxValidate(GridID, fxString, row, col);

        fxString = fxString.replace(rxRootRanges, function (arg) { return 'ip_fxRange("' + GridID + '",' + row + ',' + col + ',"' + arg + '")'; }); //regular expression to replace ranges with quotes
        fxString = fxString.replace(ip_GridProps['index'].regEx.range, function (arg) { return 'ip_fxRangeObject("' + GridID + '",' + row + ',' + col + ',"' + arg + '")'; });

        for (var key in ip_GridProps[GridID].fxList) {
                        
            fxString = fxString.replace(new RegExp('\\b' + key + '\\(\\)', 'gi'), ip_GridProps[GridID].fxList[key].fxName + '("' + GridID + '",' + row + ',' + col + ')');
            fxString = fxString.replace(new RegExp('\\b' + key + '\\(', 'gi'), ip_GridProps[GridID].fxList[key].fxName + '("' + GridID + '",' + row + ',' + col + ',');

        }
                
        return eval(fxString);
    }
    catch (ex) {
        if (ex.fxName) { return ex; }
        else { return ip_fxException(1, ex.message, 'eval', row, col); }
    }

}

function ip_fxValidate(GridID, fxString, row, col) {

    if (fxString == "") { return true; }

    var rxParts = new RegExp(ip_GridProps['index'].regEx.range.source + '|' + ip_GridProps['index'].regEx.words.source + '|' + ip_GridProps['index'].regEx.nonWords.source, 'gi')
    var fxParts = fxString.match(rxParts);

    for (var i = 0; i < fxParts.length; i++) {

        var part = fxParts[i].toLowerCase().trim();

        if (ip_GridProps[GridID].fxList[part]) { } //formula
        else if (part == "") { } //space
        else if (part.match(ip_GridProps['index'].regEx.operator)) { } //operator
        else if (part[0] == "\"" && part[part.length - 1] == "\"") { } //is a string
        else if (ip_parseNumber(part)) { } //number
        else if (part.match(ip_GridProps['index'].regEx.range)) { } // range
        else { throw ip_fxException('1', 'Cant calculate "' + part + '" in ' + fxString + ", please fix formula", 'formula', row, col); }

    }

    
    return true;
}

function ip_fxValidateCellHashTags(GridID, row, col, hashTags) {

    if (!hashTags) { return true; }
    
    for (var i = 0; i < hashTags.length; i++) {

        var cellHashTags = ip_GetEnabledHashTags(GridID, null, row, col, true);
        if (cellHashTags.match(new RegExp(hashTags[i] + '\\b', 'gi'))) { return true; }
    }

    return false;
}

function ip_fxRange(GridID, row, col, fxRanges) {

    //returns the first values in a range
    //fxRanges is an array of ranges e.g. ip_rangeObject and determines the collective value

    if (arguments.length < 4) { throw ip_fxException('1', "Missing input paramiters", 'range', row, col); }

    var value = '';
    var type = '';

    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    fxRanges = Array.prototype.slice.call(arguments).splice(3);

    for (var i = 0; i < fxRanges.length; i++) {

        var range = fxRanges[i];

        if (typeof (range) == 'string') { range = ip_fxRangeObject(GridID, row, col, range); }
        if (ip_GridProps[GridID].rowData[range.startRow].loading) { throw ip_fxException('1','Row data not loaded', 'range', row, col); }

        if (ip_fxValidateCellHashTags(GridID, range.startRow, range.startCol, range.hashtags)) {
            return ip_CellDataType(GridID, range.startRow, range.startCol, true).value;
        }

    }

    return value;

}

function ip_fxCount(GridID, row, col, fxRanges) {


    //fxRanges is an array of ranges e.g. ip_rangeObject or simply a number, and counts the number of cells in range containing numeric values

    if (arguments.length < 4) { throw ip_fxException('1', "Missing input paramiters", 'count', row, col); }

    var value = 0;
    var type = '';

    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    fxRanges = Array.prototype.slice.call(arguments).splice(3);

    for (var i = 0; i < fxRanges.length; i++) {

        if (typeof (fxRanges[i]) == 'object') {

            var range = fxRanges[i];

            if (typeof (range) == 'string') { range = ip_fxRangeObject(GridID, row, col, range); }

            for (var r = range.startRow; r <= range.endRow; r++) {

                if (ip_GridProps[GridID].rowData[r].loading) { throw ip_fxException('1', 'Row data not loaded', 'range', row, col); }

                for (var c = range.startCol; c <= range.endCol; c++) {

                    if (r == row && c == col) { throw ip_fxException('1', "Circular dependency detected, your formula range may not include the cell that contains the formula", 'count', row, col); }
                    
                    var val = ip_CellDataType(GridID, r, c, true);
                    if (val.value != null && ip_fxValidateCellHashTags(GridID, r, c, range.hashtags)) {
                        if (typeof (val.value) == 'number') { value++; }
                    }
                }

            }

        }
        else if (!isNaN(parseFloat(fxRanges[i]))) { value += parseFloat(fxRanges[i]); }
        else { throw ip_fxException('1', "Inputs are incorrect, they must be numbers or ranges", 'count', row, col); }

    }

    return value;

}

function ip_fxSum(GridID, row, col, fxRanges) {
    
    //fxRanges is an array of ranges e.g. ip_rangeObject or simply a number, and sums up the numebers in that range, ignoring other datatypes
    if (arguments.length < 4) { throw ip_fxException('1', "Missing input paramiters", 'sum', row, col); }

    var tmp = null;
    var value = 0;
    var type = '';
    
    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    fxRanges = Array.prototype.slice.call(arguments).splice(3);

    for (var i = 0; i  < fxRanges.length; i ++) {        

        if (typeof(fxRanges[i]) == 'object') {

            var range = fxRanges[i];

            if (typeof (range) == 'string') { range = ip_fxRangeObject(GridID, row, col, range); }

            for (var r = range.startRow; r <= range.endRow; r++) {
                                
                if (ip_GridProps[GridID].rowData[r].loading) { throw ip_fxException('1', 'Row data not loaded', 'sum', row, col); }

                for (var c = range.startCol; c <= range.endCol; c++) {

                    if (r == row && c == col) { throw ip_fxException('1', "Circular dependency detected, your formula range may not include the cell that contains the formula", 'sum', row, col); }
                    
                    var val = ip_CellDataType(GridID, r, c, true);
                    
                    if (val.value != null && ip_fxValidateCellHashTags(GridID, r, c, range.hashtags)) {
                        if (!isNaN(parseFloat(val.value))) { value += parseFloat(val.value); }
                    }
                    
                }

            }

        }
        else if (typeof (tmp = ip_fxCalculate(GridID, fxRanges[i])) == 'number') { value += tmp; }
        else if (!isNaN(parseFloat(fxRanges[i]))) { value += parseFloat(fxRanges[i]); }        
        else { throw ip_fxException('1', "Inputs are incorrect, they must be numbers or ranges", 'sum', row, col); }

    }

    return value;

}

function ip_fxConcat(GridID, row, col,  fxRanges) {


    //fxRanges is an array of ranges e.g. ip_rangeObject and joins the values, max 500 chars
    if (arguments.length < 4) { throw ip_fxException('1', "Missing input paramiters", 'concat', row, col); }

    var value = '';
    var type = '';

    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    fxRanges = Array.prototype.slice.call(arguments).splice(3);
    
    for (var i = 0; i < fxRanges.length; i++) {

        if (typeof (fxRanges[i]) == 'object') {

            var range = fxRanges[i];
            if (typeof (range) == 'string') { range = ip_fxRangeObject(GridID, row, col, range); }

            for (var r = range.startRow; r <= range.endRow; r++) {

                if (ip_GridProps[GridID].rowData[r].loading) { throw ip_fxException('1', 'Row data not loaded', 'range', row, col); }

                for (var c = range.startCol; c <= range.endCol; c++) {

                    if (r == row && c == col) { throw ip_fxException('1', "Circular dependency detected, your formula range may not include the cell that contains the formula", 'concat', row, col); }
                    
                    var val = ip_CellDataType(GridID, r, c, true);
                    if (val.display != null && ip_fxValidateCellHashTags(GridID, r, c, range.hashtags)) {
                        if (val.display != null) { value += String(val.display); }
                        if (value.length == 500) { break; }
                    }

                    

                }

            }

        }
        else if (fxRanges[i] != null) { value += fxRanges[i]; }

    }

    if (value.length > 500) { value = value.substr(0, 500); }

    return value;

}

function ip_fxDropDown(GridID, row, col, fxRanges) {

    //fxRanges is an array of ranges e.g. ip_rangeObject and joins the values, max 500 chars
    if (arguments.length < 4) { throw ip_fxException('1', "Missing input paramiters", 'dropdown', row, col); }


    var data = { displayField: 'displayField', data: [], toString: function () { return null; }  }
    var dataIndex = {};

    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    fxRanges = Array.prototype.slice.call(arguments).splice(3);

    for (var i = 0; i < fxRanges.length; i++) {

        var range = fxRanges[i];

        if (typeof (range) == 'object') {
            for (var r = range.startRow; r <= range.endRow; r++) {

                if (ip_GridProps[GridID].rowData[r].loading) { throw ip_fxException('1', 'Row data not loaded', 'range', row, col); }

                for (var c = range.startCol; c <= range.endCol; c++) {

                    var val = ip_GridProps[GridID].rowData[r].cells[c].value;
                    var display = ip_GridProps[GridID].rowData[r].cells[c].display;
                    if (display != null && display != '' && dataIndex[display] == null && ip_fxValidateCellHashTags(GridID, r, c, range.hashtags)) {

                        dataIndex[display] = true;
                        data.data[data.data.length] = { displayField: display, valueField: val }

                    }

                }

            }
        }
        else {

            var val = range;
            if (val != null && val != '' && dataIndex[val] == null) {

                dataIndex[val] = true;
                data.data[data.data.length] = { displayField: val, valueField: val }

            }

        }
        

    }
    

    return data;

}

function ip_fxGantt(GridID, row, col, fxInputs) {

    //fxInputs is an array with 5 inputs: BaseDate, StartDate, EndDate, TaskName, ProjectName (optional)
    if (arguments.length < 6) { throw ip_fxException('1', "Missing input paramiters", 'gantt', row, col); }

    var data = false;

    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    fxInputs = Array.prototype.slice.call(arguments).splice(3);

    if (fxInputs.length >= 3) {

        var baseDateFxObj = false;
        var startDateFxObj = false;
        var endDateFxObj = false;
                

        var baseDate = fxInputs[0];
        var startDate = fxInputs[1];
        var endDate = fxInputs[2];
        var taskName = (fxInputs.length >= 3 ? fxInputs[3] : '');
        var projectName = (fxInputs.length >= 4 ? fxInputs[4] : '');

        var test = typeof (baseDate);
        var test2 = baseDate.toString();

        //Fetch date if it is a range
        if (!ip_parseDate(baseDate) && typeof (baseDate) == 'object') { baseDate = ip_GridProps[GridID].rowData[baseDate.startRow].cells[baseDate.startCol].value; }
        if (!ip_parseDate(startDate) && typeof (startDate) == 'object') { startDate = ip_GridProps[GridID].rowData[startDate.startRow].cells[startDate.startCol].value; }
        if (!ip_parseDate(endDate) && typeof (endDate) == 'object') { endDate = ip_GridProps[GridID].rowData[endDate.startRow].cells[endDate.startCol].value; }
              
        
        //Convert to date format allowing for javascript eval
        if (Date.parse(baseDateFxObj = new Date(baseDate))) { baseDate = baseDateFxObj; }
        else if (baseDateFxObj = ip_fxCalculate(GridID, baseDate)) { if (baseDateFxObj.error) { return baseDateFxObj; } baseDate = baseDateFxObj; }
        else { throw ip_fxException('1', "The BaseDate must be a valid date format", 'gantt', row, col); }

        if (Date.parse(startDateFxObj = new Date(startDate))) { startDate = startDateFxObj; }
        else if (startDateFxObj = ip_fxCalculate(GridID, startDate)) { if (startDateFxObj.error) { return startDateFxObj; } startDate = startDateFxObj; }
        else { throw ip_fxException('1', "The StartDate must be a valid date format", 'gantt', row, col); }

        if (Date.parse(endDateFxObj = new Date(endDate))) { endDate = endDateFxObj; }
        else if (endDateFxObj = ip_fxCalculate(GridID, endDate)) { if (endDateFxObj.error) { return endDateFxObj; } endDate = endDateFxObj; }
        else { throw ip_fxException('1', "The EndDate must be a valid date format", 'gantt', row, col); }
        
        startDate.setHours(0, 0, 0, 0);
        endDate.setHours(23, 59, 59, 999);

        if (startDate <= baseDate && endDate >= baseDate) { data = true; }
        
    }
    else { throw ip_fxException('1', "Missing inputs, must have the following: BaseDate, StartDate, EndDate, TaskName", 'gantt', row, col); }

    return data;

}

function ip_fxMax(GridID, row, col, fxRanges) {


    //fxRanges is an array of ranges e.g. ip_rangeObject or simply a number, and counts the number of cells in range containing numeric values

    if (arguments.length < 4) { throw ip_fxException('1', "Missing input paramiters", 'max', row, col); }

    var value = null;
    var type = '';

    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    increment = arguments[3];

    if (typeof (fxRanges) == 'string') { fxRanges = [ip_fxRangeObject(GridID, row, col, fxRanges)]; }
    if (fxRanges.length == undefined) { fxRanges = [fxRanges]; }

    for (var i = 0; i < fxRanges.length; i++) {

        if (typeof (fxRanges[i]) == 'object') {

            var range = fxRanges[i];

            for (var r = range.startRow; r <= range.endRow; r++) {

                if (ip_GridProps[GridID].rowData[r].loading) { throw ip_fxException('1', 'Row data not loaded', 'range', row, col); }

                for (var c = range.startCol; c <= range.endCol; c++) {

                    if (r == row && c == col) { throw ip_fxException('1', "Circular dependency detected, your formula range may not include the cell that contains the formula", 'max', row, col); }

                    var val = ip_CellDataType(GridID, r, c, true);
                    if (val.value != null && ip_fxValidateCellHashTags(GridID, r, c, range.hashtags)) {
                        
                        if (value == null || val.value > value) { value = val.value; }
                        
                    }
                }

            }

        }
        else if (tmp == null || fxRanges[i] > tmp) { tmp = fxRanges[i]; }
        else { throw ip_fxException('1', "Inputs are incorrect, they must be numbers or ranges", 'max', row, col); }

    }

    return value;

}

function ip_fxToday(GridID,  row, col, increment) {

    if (arguments.length < 4) { throw ip_fxException('1', "Missing input paramiters", 'today', row, col); }

    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    increment = arguments[3];

    if (increment == null) { increment = 0; }
    else {

        if (typeof (increment) == 'object' && increment.length > 0 && !isNaN(parseInt(increment[0]))) { increment = parseInt(increment[0]); }
        else if (typeof (increment) == 'object' && increment.length > 0) { increment = increment[0]; }

        if (typeof (increment) == 'string') { increment = ip_fxRangeObject(GridID, row,col, increment); }
        if (typeof (increment) == 'object') { increment = ip_fxRange(GridID, row, col, increment); }

    }

    return new Date(new Date().valueOf() + 86400000 * increment);

}

function ip_fxDate(GridID,  row, col, increment) {

    if (arguments.length < 4) { throw ip_fxException('1', "Missing input paramiters", 'date', row, col); }

    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    increment = arguments[3];

    if (increment == null) { increment = 0; }
    else {

        if (typeof (increment) == 'object' && increment.length > 0 && !isNaN(parseInt(increment[0]))) { increment = parseInt(increment[0]); }
        else if (typeof (increment) == 'object' && increment.length > 0) { increment = increment[0]; }

        if (typeof (increment) == 'string') { increment = ip_fxRangeObject(GridID, row, col, increment); }
        if (typeof (increment) == 'object') { increment = ip_fxRange(GridID, row, col, increment); }

    }

    return new Date(new Date().valueOf() + 86400000 * increment);

}

function ip_fxDay(GridID, row, col, increment) {

    if (arguments.length < 4) { throw ip_fxException('1', "Missing input paramiters", 'day', row, col); }

    GridID = arguments[0];
    row = arguments[1];
    col = arguments[2];
    increment = arguments[3];

    if (increment == null) { increment = 0; }
    else {

        if (typeof (increment) == 'object' && increment.length > 0 && !isNaN(parseInt(increment[0]))) { increment = parseInt(increment[0]); }
        else if (typeof (increment) == 'object' && increment.length > 0) { increment = increment[0]; }

        if (typeof (increment) == 'string') { increment = ip_fxRangeObject(GridID, row, col, increment); }
        if (typeof (increment) == 'object') { increment = ip_fxRange(GridID, row, col, increment); }

    }

    return new Date(new Date().valueOf() + 86400000 * increment).getUTCDate();

}



//----- IP GRID CALCULATIONS ------------------------------------------------------------------------------------------------------------------------------------

function ip_TableWidth(TableID) {

    
    var RealWidth = 0;
    var CellsWidth = 0;
    var TableWidth = $('#' + TableID).outerWidth();
    var objTable = document.getElementById(TableID);
    var dvCells = objTable.querySelectorAll("thead th");

    for (var col = 0; col < dvCells.length; col++) {

        CellsWidth += $(dvCells[col]).outerWidth();

    }

    //Adjust correct width based on browser (ie returns stlye width, crome etc return real width
    if (TableWidth > CellsWidth) { RealWidth = TableWidth; }
    else { RealWidth = CellsWidth + TableWidth; }
    
    return RealWidth;

}

function ip_SetQuadContainerWidth(GridID, Quad) {

    

    var ScrolBarWidth = $('#' + GridID + '_q4_scrollbar_container_y').outerWidth();

    if (Quad == 1 || !Quad) { $('#' + GridID + '_q1_container').width(ip_TableWidth(GridID + '_q1_table')); }
    if (Quad == 2 || !Quad) { $('#' + GridID + '_q2_container').width(ip_TableWidth(GridID + '_q2_table') + ScrolBarWidth); }
    if (Quad == 3 || !Quad) { $('#' + GridID + '_q3_container').width(ip_TableWidth(GridID + '_q3_table')); }
    if (Quad == 4 || !Quad) { $('#' + GridID + '_q4_container').width(ip_TableWidth(GridID + '_q4_table') + ScrolBarWidth); }

}

function ip_GetQuad(GridID, row, col) {

    var Quad = 1;

    if (row != -1 && col != -1) {

        Quad = (col >= ip_GridProps[GridID].frozenCols ? 2 : 1);

        if (Quad == 1) {

            Quad = (row >= ip_GridProps[GridID].frozenRows ? 3 : 1);
        }
        else {

            Quad = (row >= ip_GridProps[GridID].frozenRows ? 4 : 2);
        }

    }
    else if (row == -1) {

        Quad = (col >= ip_GridProps[GridID].frozenCols ? 2 : 1);

    }
    else if (col == -1) {

        Quad = (row >= ip_GridProps[GridID].frozenRows ? 3 : 1);

    }
    
    return Quad;
}

function ip_CellRelativePosition(GridID, row, col, returnLastIfNull) {

    var pos = {
        relativeTop: 0,
        relativeLeft: 0,
        localTop: 0,
        localLeft: 0,
        localBottom: 0,
        localRight: 0,
        width: 0,
        height: 0
    };

    if (row == null || isNaN(parseFloat(row)) || col == null || isNaN(parseFloat(col))) { return pos; }

    //Validate merge
    var rowSpan = 1;
    var colSpan = 1;
    if (row != -1 && col != -1) {
        var cellMerge = ip_CellData(GridID, row, col).merge;
        if (cellMerge != null) {
            //row = cellMerge.mergedWithRow;
            //col = cellMerge.mergedWithCol;
            //cellMerge = ip_CellData(GridID, row, col).merge;
            rowSpan = cellMerge.rowSpan;
            colSpan = cellMerge.colSpan;

        }
    }

    var Quad = ip_GetQuad(GridID, row, col);
    var QuadElement = $('#' + GridID + '_q' + Quad + '_table');

    if (returnLastIfNull) {

        var ScrollY = ip_GridProps[GridID].scrollY;
        var ScrollX = ip_GridProps[GridID].scrollX;
        var FrozenRows = ip_GridProps[GridID].frozenRows;
        var FrozenCols = ip_GridProps[GridID].frozenCols;
        var LoadedRows = ScrollY + ip_GridProps[GridID].loadedRows - FrozenRows - 1;
        var LoadedCols = ScrollX + ip_GridProps[GridID].loadedCols - FrozenCols - 1;
        

        if (row != -1 && row < ScrollY && row >= FrozenRows) { row = ScrollY; }
        else if (row != -1 && row > LoadedRows) { row = LoadedRows; }
        
        if (col != -1 && col < ScrollX && col >= FrozenCols) { col = ScrollX; }        
        else if (col != -1 && col > LoadedCols) { col = LoadedCols; }


        if (row != -1 && row < 0) { row = 0; }
        else if (row >= ip_GridProps[GridID].rows) { row = ip_GridProps[GridID].rows - 1; }
        if (col != -1 && col < 0) { col = 0; }
        else if (col >= ip_GridProps[GridID].cols) { col = ip_GridProps[GridID].cols - 1; }

    }

    if (row == -1) {
        //var Quad = ip_GetQuad(GridID, row, col);
        var Cell = $('#' + GridID + '_q' + Quad + '_columnSelectorCell_' + col);
        if (Cell.length > 0) {

            pos.localTop = $(Cell).offset().top - $(QuadElement).offset().top;
            pos.localLeft = $(Cell).offset().left - $(QuadElement).offset().left;

            pos.localBottom = pos.localTop; // + ip_RowHeight(GridID, row);
            for (var b = 0; b < rowSpan; b++) { pos.localBottom += ip_RowHeight(GridID, row + b, true) }

            pos.localRight = pos.localLeft;// + ip_ColWidth(GridID, col);
            for (var l = 0; l < colSpan; l++) { pos.localRight += ip_ColWidth(GridID, col + l, true) }
        }

    }
    else if (col == -1) {
        //var Quad = ip_GetQuad(GridID, row, col);
        var Cell = $('#' + GridID + '_q' + Quad + '_rowSelecterCell_' + row);
        if (Cell.length > 0) {

            pos.localTop = $(Cell).offset().top - $(QuadElement).offset().top;
            pos.localLeft = $(Cell).offset().left - $(QuadElement).offset().left;

            pos.localBottom = pos.localTop; // + ip_RowHeight(GridID, row);
            for (var b = 0; b < rowSpan; b++) { pos.localBottom += ip_RowHeight(GridID, row + b, true) }

            pos.localRight = pos.localLeft; // + ip_ColWidth(GridID, col);
            for (var l = 0; l < colSpan; l++) { pos.localRight += ip_ColWidth(GridID, col + l, true) }
        }

    }
    else {

        var Cell = $('#' + GridID + '_cell_' + row + '_' + col);
        if (Cell.length > 0) {

            pos.localTop = $(Cell).offset().top - $(QuadElement).offset().top;
            pos.localLeft = $(Cell).offset().left - $(QuadElement).offset().left;

            pos.localBottom = pos.localTop; // + ip_RowHeight(GridID, row);
            for (var b = 0; b < rowSpan; b++) { pos.localBottom += ip_RowHeight(GridID, row + b, true) }

            pos.localRight = pos.localLeft; // + ip_ColWidth(GridID, col);
            for (var l = 0; l < colSpan; l++) { pos.localRight += ip_ColWidth(GridID, col + l, true) }

        }

    }

    return pos;


}

function ip_LoadedRowsCols(GridID, loadVisualRepresentation, rows, cols) {

    var loaded = {

        rowFrom_scroll: 0,
        rowTo_scroll: 0,
        rowCount_scroll: 0,
        rowLoaded_scroll: 0,

        colFrom_scroll: 0,
        colTo_scroll: 0,
        colCount_scroll: 0,
        colLoaded_scroll: 0,

        rowFrom_frozen: 0,
        rowTo_frozen: 0,
        rowCount_frozen: 0,
        rowLoaded_frozen: 0,

        colFrom_frozen: 0,
        colTo_frozen: 0,
        colCount_frozen: 0,
        colLoaded_frozen: 0
    };


    var el1 = null;
    var el2 = null;
    var el3 = null;
    var el4 = null;

    if(rows == null) { rows = true;}
    if(cols == null) { cols = true;}

    //FROZEN    
    if(rows){
        loaded.rowFrom_frozen = 0;
        loaded.rowCount_frozen = ip_GridProps[GridID].frozenRows;
        if (loaded.rowCount_frozen >= ip_GridProps[GridID].rows) { loaded.rowCount_frozen = ip_GridProps[GridID].rows; }
        loaded.rowTo_frozen = loaded.rowFrom_frozen + loaded.rowCount_frozen - 1;
    }

    if(cols){
        loaded.colFrom_frozen = 0;
        loaded.colCount_frozen = ip_GridProps[GridID].frozenCols;
        if (loaded.colCount_frozen >= ip_GridProps[GridID].cols) { loaded.colCount_frozen = ip_GridProps[GridID].cols; }
        loaded.colTo_frozen = loaded.colFrom_frozen + loaded.colCount_frozen - 1;
    }


    //SCROLLABLE    
    if(rows){
        loaded.rowFrom_scroll = (ip_GridProps[GridID].scrollY == -1 ? ip_GridProps[GridID].frozenRows : ip_GridProps[GridID].scrollY);
        loaded.rowCount_scroll = ip_GridProps[GridID].loadedRows - ip_GridProps[GridID].frozenRows;
        if (loaded.rowFrom_scroll + loaded.rowCount_scroll >= ip_GridProps[GridID].rows) { loaded.rowCount_scroll = ip_GridProps[GridID].rows - loaded.rowFrom_scroll; }
        loaded.rowTo_scroll = loaded.rowFrom_scroll + loaded.rowCount_scroll - 1;
    }

    if(cols){
        loaded.colFrom_scroll = (ip_GridProps[GridID].scrollX == -1 ? ip_GridProps[GridID].frozenCols : ip_GridProps[GridID].scrollX);
        loaded.colCount_scroll = ip_GridProps[GridID].loadedCols - ip_GridProps[GridID].frozenCols;
        if (loaded.colFrom_scroll + loaded.colCount_scroll >= ip_GridProps[GridID].cols) {
            loaded.colCount_scroll = ip_GridProps[GridID].cols - loaded.colFrom_scroll;
        }
        loaded.colTo_scroll = loaded.colFrom_scroll + loaded.colCount_scroll - 1;
    }

    //Turned off by default for performance reasons only used in special situations when absolutely nessisary
    if (loadVisualRepresentation) {

        var table = null;

        table = document.getElementById(GridID + '_q1_table');
        el1 = (table == null ? null : table.querySelectorAll('tbody tr.ip_grid_row'));
        el2 = (table == null ? null : table.querySelectorAll('thead th.ip_grid_columnSelectorCell'));

        table = document.getElementById(GridID + '_q3_table');
        el3 = (table == null ? null : table.querySelectorAll('tbody tr.ip_grid_row'));

        table = document.getElementById(GridID + '_q2_table');//
        el4 = (table == null ? null : table.querySelectorAll('thead th.ip_grid_columnSelectorCell'));
    }

    if(rows){
        loaded.rowLoaded_frozen = (el1 != null ? el1.length : loaded.rowCount_frozen);
        loaded.rowLoaded_scroll = (el3 != null ? el3.length : loaded.rowCount_scroll);
    }

    if(cols){
        loaded.colLoaded_frozen = (el2 != null ? el2.length : loaded.colCount_frozen);    
        loaded.colLoaded_scroll = (el4 != null ? el4.length : loaded.colCount_scroll);
    }

    return loaded;
}

function ip_IsCellLoaded(GridID, row, col, LoadedRowsCols) {

    if (LoadedRowsCols == null) { LoadedRowsCols = ip_LoadedRowsCols(GridID, false, true, false); }

    if (((row >= LoadedRowsCols.rowFrom_frozen && row <= LoadedRowsCols.rowTo_frozen)
        || (row >= LoadedRowsCols.rowFrom_scroll && row <= LoadedRowsCols.rowTo_scroll))
        &&
        ((col >= LoadedRowsCols.colFrom_frozen && col <= LoadedRowsCols.colTo_frozen)
        || (col >= LoadedRowsCols.colFrom_frozen && col <= LoadedRowsCols.colTo_scroll))) { return true; }
}

function ip_IsRowLoaded(GridID, row, LoadedRowsCols, checkScrollQuadOnly) {

    if (LoadedRowsCols == null) { LoadedRowsCols = ip_LoadedRowsCols(GridID, false, true, false); }
    if (checkScrollQuadOnly == null) { checkScrollQuadOnly = false; }

    if (!checkScrollQuadOnly && row >= LoadedRowsCols.rowFrom_frozen && row <= LoadedRowsCols.rowTo_frozen) { return true; }
    else if (row >= LoadedRowsCols.rowFrom_scroll && row <= LoadedRowsCols.rowTo_scroll) { return true; }

    return false;
}

function ip_IsColLoaded(GridID, col, LoadedRowsCols, checkScrollQuadOnly) {

    if (LoadedRowsCols == null) { LoadedRowsCols = ip_LoadedRowsCols(GridID, false, false, true); }
    if (checkScrollQuadOnly == null) { checkScrollQuadOnly = false; }

    if (!checkScrollQuadOnly && col >= LoadedRowsCols.colFrom_frozen && col <= LoadedRowsCols.colTo_frozen) { return true; }
    else if (col >= LoadedRowsCols.colFrom_scroll && col <= LoadedRowsCols.colTo_scroll) { return true; }

    return false;
}

function ip_RecalculateLoadedRowsCols(GridID, render, doRows, doCols, scrollY, scrollX) {

    var Loaded = ip_LoadedRowsCols(GridID);

    if (doRows) {
        
        if (!render) {

            //if (ip_GridProps[GridID].dimensions.scrollHeight <= 0) {

            //    //Calculate scroll height                
            //    var ScrollBarX_height = parseInt($('#' + GridID + '_q4_scrollbar_container_x').height());
            //    var FirstQuadTableHeight = parseInt($('#' + GridID + '_q1_table').outerHeight());

            //    if (!isNaN(FirstQuadTableHeight)) { ip_GridProps[GridID].dimensions.scrollHeight = ip_GridProps[GridID].dimensions.gridHeight - FirstQuadTableHeight - ScrollBarX_height }
            //    else { ip_GridProps[GridID].dimensions.scrollHeight = ip_GridProps[GridID].dimensions.gridHeight - ip_GridProps[GridID].dimensions.columnSelectorHeight - ip_GridProps[GridID].dimensions.defaultBorderHeight; }

            //}
            //else
            //{
            //    ip_GridProps[GridID].dimensions.scrollHeight = $('#' + GridID + '_q4_scrollbar_container_y').height();
            //}

            ////var GridTop = $('#' + GridID + '_table').position().top;
            //var scrollHeight = ip_GridProps[GridID].dimensions.scrollHeight;// - GridTop;
            //var ScrollY = (scrollY == null ? ip_GridProps[GridID].scrollY : scrollY);
            //var TotalHeight = 0;
            //var rows = 0;
            //var totalRows = ip_GridProps[GridID].rows;

            

            //for (var r = ScrollY; r < totalRows; r++) {
            
            //    TotalHeight += ip_RowHeight(GridID, r, true);
            //    if (TotalHeight > scrollHeight) { r = ip_GridProps[GridID].rows; } //Break loop
            //    else { rows++; }

            //}

            //defaultRowHeight: 30, //Default height for a row
            //defaultColWidth: 130, //Default width for a column
            //gridHeight: 0, //The overall height of the entire grid
            //gridWidth: 0, //The overall with of the entire grid
            //scrollHeight: 0,
            //scrollWidth: 0,
            //columnSelectorHeight: 30, //Column selector height
            //rowSelectorWidth: 50, //Row selector width
            //defaultBorderHeight: 1,
            //accumulativeScrollHeight: 0,
            //accumulativeScrollWidth: 0

            var ScrollBarX_height = 10; //parseInt($('#' + GridID + '_q4_scrollbar_container_x').height());
            var GridHeight = ip_GridProps[GridID].dimensions.gridHeight - ip_GridProps[GridID].dimensions.columnSelectorHeight - ip_GridProps[GridID].dimensions.defaultBorderHeight - ScrollBarX_height;
            var TotalHeight = 0;
            var rows = 0;

            //Calculate total frozen row height
            for (var r = 0; r < ip_GridProps[GridID].frozenRows; r++) {

                rows++;
                TotalHeight += ip_RowHeight(GridID, r, true);
                if (TotalHeight > GridHeight) { rows--; r = ip_GridProps[GridID].rows; } //Break loop

            }

            //Calculate total scroll row height
            for (var r = ip_GridProps[GridID].scrollY; r < ip_GridProps[GridID].rows; r++) {

                rows++;
                TotalHeight += ip_RowHeight(GridID, r, true);
                if (TotalHeight >= GridHeight) { r = ip_GridProps[GridID].rows; } //Break loop

            }

            ip_GridProps[GridID].loadedRows = rows;
        }
        else if (render) { ip_ReRenderRows(GridID); }
    }

    if (doCols) {

        //var ScrollY = (scrollY == null ? ip_GridProps[GridID].scrollY : scrollY);
        //var ScrollX = (scrollX == null ? ip_GridProps[GridID].scrollX : scrollX);
        //var GridWidth = ip_GridProps[GridID].dimensions.gridWidth;
        //var TotalWidth = 0;
        //var cols = 0;
        
        //for (var c = ScrollY; c < ip_GridProps[GridID].cols; c++) {
         
        //    TotalWidth += ip_GridProps[GridID].colData[c].width;
        //    if (TotalWidth > GridWidth) { c = ip_GridProps[GridID].cols; } //Break loop
        //    else { cols++; }

        //}

        var ScrollBarY_width = 10; //parseInt($('#' + GridID + '_q4_scrollbar_container_y').width());
        var GridWidth = ip_GridProps[GridID].dimensions.gridWidth - ip_GridProps[GridID].dimensions.rowSelectorWidth - ip_GridProps[GridID].dimensions.defaultBorderHeight - ScrollBarY_width;
        var TotalWidth = 0;
        var cols = 0;
        
        //Calculate total frozen row height
        for (var c = 0; c < ip_GridProps[GridID].frozenCols; c++) {

            cols++;
            TotalWidth += ip_ColWidth(GridID, c, true);
            if (TotalWidth > GridWidth) { cols--; c = ip_GridProps[GridID].cols; } //Break loop

        }

        //Calculate total scroll row height
        for (var c = ip_GridProps[GridID].scrollX; c < ip_GridProps[GridID].cols; c++) {

            cols++;
            TotalWidth += ip_ColWidth(GridID, c, true);
            if (TotalWidth >= GridWidth) { c = ip_GridProps[GridID].cols; } //Break loop

        }

        ip_GridProps[GridID].loadedCols = cols;

    }
    
}

function ip_BorderSize(GridID, style) {

    $('#' + GridID + ' .ip_grid').css('border-right', style);
    $('#' + GridID + ' .ip_grid_columnSelectorCell').css('border-bottom', style);
    $('#' + GridID + ' .ip_grid_columnSelectorCell').css('border-left', style);
    $('#' + GridID + ' .ip_grid_columnSelectorCell').css('border-top', style);
    $('#' + GridID + ' .ip_grid_columnSelectorCellCorner').css('border-bottom', style);
    $('#' + GridID + ' .ip_grid_columnSelectorCellCorner').css('border-left', style);
    $('#' + GridID + ' .ip_grid_columnSelectorCellCorner').css('border-top', style);
    $('#' + GridID + ' .ip_grid_rowSelecterCell').css('border-bottom', style);
    $('#' + GridID + ' .ip_grid_rowSelecterCell').css('border-left', style);
    $('#' + GridID + ' .ip_grid_cell').css('border-bottom', style);
    $('#' + GridID + ' .ip_grid_cell').css('border-left', style);

    $('#' + GridID).ip_Scrollable();
}

function ip_RowHeight(GridID, row, addBorderHeight) {

    if (row < 0) { return 0; }

    var rowData = ip_GridProps[GridID].rowData[row];
    var height = (rowData.height == '' || rowData.height == null ? ip_GridProps[options.GridID].dimensions.defaultRowHeight : rowData.height);

    if (rowData.hide) { height = 0; addBorderHeight = false; }
    if (row >= 0 && row < ip_GridProps[GridID].rows && addBorderHeight) { height = height + ip_GridProps[GridID].dimensions.defaultBorderHeight; }
    
    

    return height;
}

function ip_ColWidth(GridID, col, addBorderHeight) {

    if (col < 0) { return 0; }

    var colData = ip_GridProps[GridID].colData[col];
    var width = (colData.width == '' || colData.width == null ? ip_GridProps[GridID].dimensions.defaultColWidth : colData.width);

    if (colData.hide) { width = 0; addBorderHeight = false; }
    if (col >= 0 && col < ip_GridProps[GridID].cols && addBorderHeight) { width = width + ip_GridProps[GridID].dimensions.defaultBorderHeight; }

    return width;
}


//----- SORTING ------------------------------------------------------------------------------------------------------------------------------------

function ip_CompareRow(GridID, col, order, rowSort) {

    if (col == null) { col = 0; }
    if (order == null) { order = 'az'; }

    if (order == 'az') {
        return function (a, b) {

            var aVal = (rowSort ? a.cells[col].groupValue : a.cells[col].value);
            var bVal = (rowSort ? b.cells[col].groupValue : b.cells[col].value);
            var aRowA = a.cells[col].row;
            var aRowB = b.cells[col].row;
            
            if (a.cells[col].merge != null && !rowSort) { aVal = ip_GridProps[GridID].rowData[a.cells[col].merge.mergedWithRow].cells[a.cells[col].merge.mergedWithCol].value; }
            if (b.cells[col].merge != null && !rowSort) { bVal = ip_GridProps[GridID].rowData[b.cells[col].merge.mergedWithRow].cells[b.cells[col].merge.mergedWithCol].value; }
            
            if (aVal == null || aVal == '') { return 1; }
            if (bVal == null || bVal == '') { return -1; }

            if (!isNaN(parseFloat(aVal))) { aVal = parseFloat(aVal); }
            if (!isNaN(parseFloat(bVal))) { bVal = parseFloat(bVal); }

            var aType = typeof (aVal);
            var bType = typeof (bVal);

            if (aType == 'object') { aType = 'date'; }
            if (bType == 'object') { bType = 'date'; }

            if (aType == 'string') { aVal = aVal.toLowerCase(); }
            if (bType == 'string') { bVal = bVal.toLowerCase(); }
            
            //if (typeof (aVal) == 'object') { aVal = JSON.stringify(aVal).replace(/"/g,"") }
            //if (typeof (bVal) == 'object') { bVal = JSON.stringify(bVal).replace(/"/g, "") }

            if (aType == 'string' && bType == 'date') { return -1; } 
            if (aType == 'date' && bType == 'string') { return 1; } 

            if (aType == 'number' && bType == 'date') { bVal = bVal.getFullYear(); }
            if (aType == 'date' && bType == 'number') { aVal = aVal.getFullYear(); }

            if (typeof (aVal) == 'number' && typeof (bVal) == 'object') { return 1; } 
            if (typeof (aVal) == 'object' && typeof (bVal) == 'number') { return -1; }

            if (aType == 'string' && bType == 'number') { return -1; } //make string smaller than numbers (this is because MySQL does it this way)
            if (aType == 'number' && bType == 'string') { return 1; } //make string smaller than numbers (this is because MySQL does it this way)

            if (aVal < bVal) { return -1; }
            if (aVal > bVal) { return 1; }

            // values match exactly maintain row sequence
            if (aRowA < aRowB) { return -1; }
            if (aRowA > aRowB) { return 1; }
                       
            return 0;
        }
    }
    else {
        return function (a, b) {

            var aVal = (rowSort ? a.cells[col].groupValue : a.cells[col].value);
            var bVal = (rowSort ? b.cells[col].groupValue : b.cells[col].value);
            var aRowA = a.cells[col].row;
            var aRowB = b.cells[col].row;
            
            if (a.cells[col].merge != null && !rowSort) { aVal = ip_GridProps[GridID].rowData[a.cells[col].merge.mergedWithRow].cells[a.cells[col].merge.mergedWithCol].value; }
            if (b.cells[col].merge != null && !rowSort) { bVal = ip_GridProps[GridID].rowData[b.cells[col].merge.mergedWithRow].cells[b.cells[col].merge.mergedWithCol].value; }

            if (aVal == null || aVal == '') { return 1; }
            if (bVal == null || bVal == '') { return -1; }

            if (!isNaN(parseFloat(aVal))) { aVal = parseFloat(aVal); }
            if (!isNaN(parseFloat(bVal))) { bVal = parseFloat(bVal); }

            var aType = typeof (aVal);
            var bType = typeof (bVal);

            if (aType == 'object') { aType = 'date'; }
            if (bType == 'object') { bType = 'date'; }

            if (aType == 'string') { aVal = aVal.toLowerCase(); }
            if (bType == 'string') { bVal = bVal.toLowerCase(); }

            //if (typeof (aVal) == 'object') { aVal = JSON.stringify(aVal).replace(/"/g, "") }
            //if (typeof (bVal) == 'object') { bVal = JSON.stringify(bVal).replace(/"/g, "") }

            if (aType == 'string' && bType == 'date') { return 1; }
            if (aType == 'date' && bType == 'string') { return -1; }

            if (aType == 'number' && bType == 'date') { bVal = bVal.getFullYear(); }
            if (aType == 'date' && bType == 'number') { aVal = aVal.getFullYear(); }

            if (aType == 'string' && bType == 'number') { return 1; } //make string smaller than numbers (this is because MySQL does it this way)
            if (aType == 'number' && bType == 'string') { return -1; } //make string smaller than numbers (this is because MySQL does it this way)

            if (aVal > bVal) { return -1; }
            if (aVal < bVal) { return 1; }

            // values match exactly maintain row sequence
            if (aRowA < aRowB) { return -1; }
            if (aRowA > aRowB) { return 1; }

            return 0;
        }
    }
}


//----- UNDO ------------------------------------------------------------------------------------------------------------------------------------

function ip_ValidateUndo(GridID, TransactionID, Ranges, ResetUndoStack) {

    var totalUndoSize = 0;

    if (Ranges != null) {        

        for (var r = 0; r < Ranges.length; r++) {

            totalUndoSize += (Ranges[r].endRow - Ranges[r].startRow + 1) * (Ranges[r].endCol - Ranges[r].startCol + 1);

        }

    }

    if (ip_GridProps[GridID].undo.maxUndoRangeSize > totalUndoSize) {
        if (ResetUndoStack) { ip_ClearUndoStack(GridID); }
        return false;
    }
    else { return true; }
}

function ip_AddUndo(GridID, Method, TransactionID, TransactionType, Range, RevertSelectRange, RevertSelectCell, TransactionIndex, Data) {


    if (Range == null) { Range = ip_rangeObject(); }
    if (Range.endRow == null) { Range.endRow = Range.startRow; }
    if (Range.endCol == null) { Range.endCol = Range.startCol; }
    
    //undoStack
    if (ip_GridProps[GridID].undo.undoStack[TransactionID] == null) {

        var undoStackSize = Object.keys(ip_GridProps[GridID].undo.undoStack).length;
        var MaxUndoTransaction = undoStackSize;

        //Remove older transactions from the undo stack tranactions
        if (undoStackSize >= ip_GridProps[GridID].undo.maxTransactions) {  MaxUndoTransaction = ip_TrimUndoStack(GridID) + 1;  }

        ip_GridProps[GridID].undo.undoStack[TransactionID] = {

            transactionID: TransactionID,
            transactionSeq: MaxUndoTransaction,
            method: Method,
            transactions: new Array()

        }

    }
    
    TransactionIndex = (TransactionIndex == null ? ip_GridProps[GridID].undo.undoStack[TransactionID].transactions.length : TransactionIndex);

    ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex] = {};    

    if (TransactionType != null) { ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].transactionType = TransactionType; }
    if (Range.startRow != null) { ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].startRow = Range.startRow; }
    if (Range.startCol != null) { ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].startCol = Range.startCol; }
    if (Range.endRow != null) { ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].endRow = Range.endRow; }
    if (Range.endCol != null) { ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].endCol = Range.endCol; }
    if (RevertSelectRange != null) { ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].RevertSelectRange = RevertSelectRange; }
    if (RevertSelectCell != null) { ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].RevertSelectCell = RevertSelectCell; }

    ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].dataIndex = {};
    ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].data = new Array();
    ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex].index = TransactionIndex;
    
    if (Data != null) { ip_AddUndoTransactionData(GridID, ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex], Data); }

    return ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[TransactionIndex];
}
                           
function ip_AddUndoTransactionData(GridID, UndoTransaction, Data) {

    if (UndoTransaction != null) {
        
        //PREVENT DUPLICATE, specifically used for recursive methods such as 'recalculate'
        if (UndoTransaction.transactionType == "CellData" && Data.row != null && Data.col != null) {
            if (UndoTransaction.dataIndex[Data.row + '-' + Data.col] == null) { UndoTransaction.dataIndex[Data.row + '-' + Data.col] = UndoTransaction.data.length; }
            else { return; }
        }

        if (typeof Data == "function") { UndoTransaction.data = Data; }
        else { UndoTransaction.data[UndoTransaction.data.length] = Data; }

    }

}

function ip_UndoTransaction(GridID, TransactionID, SelectRanges, ReRender) {
    //Performs the undo transction, the following is supported: CellData, MergeData
    //TransactionID: ID of undo transaction to execute
    //ReRender: redraw the grid after the undo

    //IMPORTAINT TO NOTE:
    //CellData: does not effect merges (use Merge data for that)
    //ColData: does not effect merges (use Merge data for that)
    //RowData: does not effect cells or merges (use CellData / MergeData for that)
    //MergeData: will change the marges within a range, however will not affect merges that overlap the range 
    //undoServer: undo can only take plase on the server, this will cause resync to be set to true

    var result = { success: false, rowDataLoading: false, method: '' }

    if (ip_GridProps[GridID].undo.undoStack[TransactionID] != null) {

        var clearRanges = true;
        var SelectRanges = [];
        var SelectCell = null;
        var renderType = 'rows';

        result.method = ip_GridProps[GridID].undo.undoStack[TransactionID].method;

        switch (ip_GridProps[GridID].undo.undoStack[TransactionID].method) {
            case 'ip_InsertCol': renderType = 'cols'; break;
            case 'ip_AddCol': renderType = 'cols'; break;
            case 'ip_RemoveCol': renderType = 'cols'; break;
            case 'ip_MoveCol': renderType = 'cols'; break;
            case 'ip_ResizeColumn': renderType = 'cols'; break;
            case 'ip_HideShowColumns': renderType = 'cols'; break;
            case 'ip_SetColDataType': renderType = 'cols'; break;
        }

        for (var t = ip_GridProps[GridID].undo.undoStack[TransactionID].transactions.length -1; t >= 0; t--) {

            var currentTransaction = ip_GridProps[GridID].undo.undoStack[TransactionID].transactions[t];

            if (currentTransaction.transactionType == 'undoServer') {
                result.rowDataLoading = true;
            }
            else if (currentTransaction.transactionType == 'function')
            {
                if (typeof currentTransaction.data == "function") { currentTransaction.data();  }
            }
            else if (currentTransaction.transactionType == 'CellData') {

                //Clear range but preserve merges and indexes
                for (var r = currentTransaction.startRow; r <= currentTransaction.endRow; r++) {
                    for (var c = currentTransaction.startCol; c <= currentTransaction.endCol; c++) {
                        ip_ResetCell(GridID, r, c, true, true, false);
                    }
                }

                //Set new values still preserving merges
                for (var c = 0; c < currentTransaction.data.length; c++) {

                    var row = currentTransaction.data[c].row;
                    var col = currentTransaction.data[c].col;                    

                    var merge = ip_GridProps[GridID].rowData[row].cells[col].merge;                    

                    ip_ResetCell(GridID, row, col, true, true, false);
                    ip_GridProps[GridID].rowData[row].cells[col] = currentTransaction.data[c];
                    ip_GridProps[GridID].rowData[row].cells[col].merge = merge;                    
                    ip_SetCellFormula(GridID, {  row:row, col:col, formula: ip_GridProps[GridID].rowData[row].cells[col].formula });

                }


            }
            else if (currentTransaction.transactionType == 'MergeData') {

                //Clear existing merges
                var mergeData = ip_ValidateRangeMergedCells(GridID, currentTransaction.startRow, currentTransaction.startCol, currentTransaction.endRow, currentTransaction.endCol);
                for (var m = 0; m < mergeData.merges.length; m++) {
                    if (!mergeData.merges[m].containsOverlap) { ip_ResetCellMerge(GridID, mergeData.merges[m].mergedWithRow, mergeData.merges[m].mergedWithCol); }
                }

                //Add undo merges
                for (var m = 0; m < currentTransaction.data.length; m++) {

                    var startRow = currentTransaction.data[m].mergedWithRow;
                    var startCol = currentTransaction.data[m].mergedWithCol;
                    var endRow = startRow + currentTransaction.data[m].rowSpan - 1;
                    var endCol = startCol + currentTransaction.data[m].colSpan - 1;

                    ip_ResetCellMerge(GridID, startRow, startCol);
                    ip_SetCellMerge(GridID, startRow, startCol, endRow, endCol);

                }

            }
            else if (currentTransaction.transactionType == 'ColData') {

                //Set new column values 
                for (var c = 0; c < currentTransaction.data.length; c++) {

                    var col = currentTransaction.data[c].col;
                    var containsMerges = ip_GridProps[GridID].colData[col].containsMerges;
                    ip_GridProps[GridID].colData[col] = currentTransaction.data[c];
                    ip_GridProps[GridID].colData[col].containsMerges = containsMerges;
                    

                }

            }
            else if (currentTransaction.transactionType == 'RowData') {

                //Set new row values - preserving the existing cells - use cell data to change cell values
                for (var r = 0; r < currentTransaction.data.length; r++) {
                                       
                    var row = currentTransaction.data[r].row;
                    var containsMerges = ip_GridProps[GridID].rowData[row].containsMerges;
                    var cells = ip_GridProps[GridID].rowData[row].cells;
                    ip_GridProps[GridID].rowData[row] = currentTransaction.data[r];
                    ip_GridProps[GridID].rowData[row].cells = cells;
                    ip_GridProps[GridID].rowData[row].containsMerges = containsMerges;

                }

            }
            

            if (SelectRanges && currentTransaction.RevertSelectCell != null) { SelectCell = { row: currentTransaction.RevertSelectCell.row, col: currentTransaction.RevertSelectCell.col }  }
            if (SelectRanges && currentTransaction.RevertSelectRange != null) { SelectRanges[SelectRanges.length] = currentTransaction.RevertSelectRange;  }
            
        }
        
        delete ip_GridProps[GridID].undo.undoStack[TransactionID];

        if (ReRender) {

        
            
            if (renderType == 'cols') { ip_ReRenderCols(GridID); } else { ip_ReRenderRows(GridID); }
            $('#' + GridID).ip_Scrollable();
            ip_RePoistionRanges(GridID, 'all', false, true);

        }

        //Select cell
        if (SelectCell != null) { $('#' + GridID).ip_SelectCell({ row: SelectCell.row, col: SelectCell.col, multiselect: false, scrollTo: true, selectAsRange:false }); }

        //Select range
        for (var r = 0; r < SelectRanges.length; r++) {

            if (SelectRanges[r].startRow == -1 && SelectRanges[r].startCol == -1) {

                $('#' + GridID).ip_SelectColumn({ multiselect: false, col: -1 });
                $('#' + GridID).ip_SelectRow({ multiselect: true, row: -1, fetchRange: false });

            }
            else if (SelectRanges[r].startRow == -1) { for (var col = SelectRanges[r].startCol; col <= SelectRanges[r].endCol; col++) { $('#' + GridID).ip_SelectColumn({ col: col, multiselect: true }); } }
            else if (SelectRanges[r].startCol == -1) { for (var row = SelectRanges[r].startRow; row <= SelectRanges[r].endRow; row++) { $('#' + GridID).ip_SelectRow({ row: row, multiselect: true }); } }
            else { $('#' + GridID).ip_SelectRange({ range: SelectRanges[r], multiselect: true }); }

        }

        result.success = true;
    }

    return result;
}

function ip_TrimUndoStack(GridID) {

    //Taking a chance with this code - because though its faster, it may not work properly with all javascript implementations
    //var MinTransactionID = Object.keys(ip_GridProps[GridID].undo.undoStack)[0];
    //var MaxTransactionID = Object.keys(ip_GridProps[GridID].undo.undoStack)[Object.keys(ip_GridProps[GridID].undo.undoStack).length - 1];
    //var MaxTransactionSeq = ip_GridProps[GridID].undo.undoStack[MaxTransactionID].transactionSeq;

    //This option will work with all javascript implementations but is slower
    var sortedStack = [];
    var MinTransactionSeq = -1;
    var MaxTransactionSeq = -1;
    var MinTransactionID = '';
    var MaxTransactionID = '';


    //Buld an array of the undo stack so we can sort it
    for (var key in ip_GridProps[GridID].undo.undoStack) {

        var TransactionSeq = ip_GridProps[GridID].undo.undoStack[key].transactionSeq;
        if (MinTransactionSeq == -1 || MinTransactionSeq > TransactionSeq) { MinTransactionSeq = TransactionSeq; MinTransactionID = key; }
        if (MaxTransactionSeq == -1 || MaxTransactionSeq < TransactionSeq) { MaxTransactionSeq = TransactionSeq; MaxTransactionID = key; }

    }


    delete ip_GridProps[GridID].undo.undoStack[MinTransactionID];

    return MaxTransactionSeq;
}

function ip_ClearUndoStack(GridID) {

    if (ip_GridProps[GridID] != null) {
        for (var key in ip_GridProps[GridID].undo.undoStack) { delete ip_GridProps[GridID].undo.undoStack[key]; }
    }
}


//----- GROUP ROWS ------------------------------------------------------------------------------------------------------------------------------------

function ip_GetGroupedRows(GridID) {

    var groupedRows = { rowData: {} };
    var inGroupIndex = [];

    for (var r = 0; r < ip_GridProps[GridID].rows; r++) {

        if (inGroupIndex.length > 0) {

            //These are grouped rows, add rows to the return object
            groupedRows.rowData[r] = ip_CloneRow(GridID, r);
            groupedRows.rowData[r].groupedWithRow = inGroupIndex[inGroupIndex.length - 1].row;
            groupedRows.rowData[r].groupCountDown = inGroupIndex[inGroupIndex.length - 1].groupCountDown;
            groupedRows.rowData[r].groupCount = inGroupIndex[inGroupIndex.length - 1].groupCount;
            groupedRows.rowData[r].groupHeader = false;

            inGroupIndex[inGroupIndex.length - 1].groupCountDown--;
            if (inGroupIndex[inGroupIndex.length - 1].groupCountDown == 0) { inGroupIndex.splice(inGroupIndex.length - 1, 1); }

        }

        if (ip_GridProps[GridID].rowData[r].groupCount != null && ip_GridProps[GridID].rowData[r].groupCount > 0) {
               
            //Set the grouped row index counter
            inGroupIndex[inGroupIndex.length] = {
                row: r,
                groupCountDown: ip_GridProps[GridID].rowData[r].groupCount,
                groupCount: ip_GridProps[GridID].rowData[r].groupCount
            }
            
            groupedRows.rowData[r] = ip_CloneRow(GridID, r);
            groupedRows.rowData[r].groupedWithRow = r;
            groupedRows.rowData[r].groupCountDown = ip_GridProps[GridID].rowData[r].groupCount + 1;
            groupedRows.rowData[r].groupCount = ip_GridProps[GridID].rowData[r].groupCount + 1;
            groupedRows.rowData[r].groupHeader = true;
        }

    }

    return groupedRows;
}


//----- COLOR FUNCTIONS ------------------------------------------------------------------------------------------------------------------------------------

function ip_ChangeHue(rgb, hueDegree, saturationDegree, darkenLighten) {

    if (rgb == '') { return ''; }

    var hsl = ip_rgbToHSL(rgb);

    if (hueDegree != null) { hsl.h += hueDegree; }
    if (saturationDegree != null) { hsl.l = saturationDegree; }
    if (darkenLighten != null) {
        
        if (hsl.l - darkenLighten > 0) { hsl.l -= darkenLighten; }
        else { hsl.l += darkenLighten; }        

        if (hsl.l < 0) { hsl.l = 0; }
        if (hsl.l > 1) { hsl.l = 1; }

    }


    if (hsl.h > 360) {
        hsl.h -= 360;
    }
    else if (hsl.h < 0) {
        hsl.h += 360;
    }

    if (hsl.l > 1) {
        hsl.l -= 1;
    }
    else if (hsl.l < 0) {
        hsl.l += 1;
    }

    return ip_hslToRGB(hsl);
}

function ip_rgbToHSL(rgb) {
    // strip the leading # if it's there
    rgb = rgb.replace(/^\s*#|\s*$/g, '');

    // convert 3 char codes --> 6, e.g. `E0F` --> `EE00FF`
    if (rgb.length == 3) {
        rgb = rgb.replace(/(.)/g, '$1$1');
    }

    var r = parseInt(rgb.substr(0, 2), 16) / 255,
        g = parseInt(rgb.substr(2, 2), 16) / 255,
        b = parseInt(rgb.substr(4, 2), 16) / 255,
        cMax = Math.max(r, g, b),
        cMin = Math.min(r, g, b),
        delta = cMax - cMin,
        l = (cMax + cMin) / 2,
        h = 0,
        s = 0;

    if (delta == 0) {
        h = 0;
    }
    else if (cMax == r) {
        h = 60 * (((g - b) / delta) % 6);
    }
    else if (cMax == g) {
        h = 60 * (((b - r) / delta) + 2);
    }
    else {
        h = 60 * (((r - g) / delta) + 4);
    }

    if (delta == 0) {
        s = 0;
    }
    else {
        s = (delta / (1 - Math.abs(2 * l - 1)))
    }

    return {
        h: h,
        s: s,
        l: l
    }
}

function ip_hslToRGB(hsl) {
    var h = hsl.h,
        s = hsl.s,
        l = hsl.l,
        c = (1 - Math.abs(2 * l - 1)) * s,
        x = c * (1 - Math.abs((h / 60) % 2 - 1)),
        m = l - c / 2,
        r, g, b;

    if (h < 60) {
        r = c;
        g = x;
        b = 0;
    }
    else if (h < 120) {
        r = x;
        g = c;
        b = 0;
    }
    else if (h < 180) {
        r = 0;
        g = c;
        b = x;
    }
    else if (h < 240) {
        r = 0;
        g = x;
        b = c;
    }
    else if (h < 300) {
        r = x;
        g = 0;
        b = c;
    }
    else {
        r = c;
        g = 0;
        b = x;
    }

    r = ip_Normalize_rgb_value(r, m);
    g = ip_Normalize_rgb_value(g, m);
    b = ip_Normalize_rgb_value(b, m);

    return ip_rgbToHex(r, g, b);
}

function ip_rgbToHex(r, g, b) {
    return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

function ip_Normalize_rgb_value(color, m) {
    color = Math.floor((color + m) * 255);
    if (color < 0) {
        color = 0;
    }
    return color;
}

//----- PARSE AND FORMAT ------------------------------------------------------------------------------------------------------------------------------------

function ip_parseBool(value) {

    if (value == null) { return false; }
    if (value == true) { return true; }

    if (typeof (value) == 'string') { value = value.toLowerCase(); }

    if (value == 'true') { return true; }
    if (value == '1') { return true; }
    if (value == 1) { return true; }

    return false;
}

function ip_parseDate(value, mask) {



    if (Object.prototype.toString.call(value) === "[object Date]") {
        // it is a date
        if (isNaN(value.getTime())) {
            return Number.NaN;
        }
        else {
            return new Date(value);
        }
    }
    else {
        if (value == null) { return new Date(value); } //return Number.NaN; }
        if (typeof (value) == 'string') {

            if (mask != null) { mask = mask.trim().toLowerCase(); }
            value = value.replace(/\//gi, '-');
            value = value.toLowerCase().trim();

            if (mask == null || mask == 'full') {

                if (!value.toString().match(/^\d+(\.\d{1,2})?$/)) {
                    var processed = new Date(value);
                    if (processed != 'Invalid Date') { return processed; }
                }

            }
            
            var date = [];
            var ShortMonths = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11, ocx: 9, january:0, february:1, march:2, april:3, may:4, june:5, july:6, august:7, september:8, october:9, november:10, december:11  }
            var RegEx;
            var universal = (value[value.length - 1] == 'z' ? true : false);
            var timeZoneOffset = (universal ? -(new Date()).getTimezoneOffset() / 60 : 0);

            //YYYY-MM-DD (THH:MM:SS OPTIONAL)
            if (mask == null || mask == 'yyyy-mm-dd') {
                RegEx = /20\d{2}(-|\/)((0[1-9])|(1[0-2]))(-|\/)((0[1-9])|([1-2][0-9])|(3[0-1]))((T|\s)(([0-1][0-9])|(2[0-3])):([0-5][0-9])(:([0-5][0-9]))?)?/gi;
                if (value.match(RegEx)) {


                    date = value.split(/[-t\s:.z]/);
                    if (date.length <= 3) { return new Date(parseInt(date[0]), parseInt(date[1]) - 1, parseInt(date[2])); }
                    else if (date.length <= 5) { return new Date(parseInt(date[0]), parseInt(date[1]) - 1, parseInt(date[2]), parseInt(date[3]) + timeZoneOffset, parseInt(date[4])); }
                    else if (date.length >= 6) { return new Date(parseInt(date[0]), parseInt(date[1]) - 1, parseInt(date[2]), parseInt(date[3]) + timeZoneOffset, parseInt(date[4]), parseInt(date[5])); }

                }
            }

            //DD-MM-YYYY (THH:MM:SS OPTIONAL)
            if (mask == null || mask == 'dd-mm-yyyy') {
            }

            //MM-DD-YYYY (THH:MM:SS OPTIONAL)
            if (mask == null || mask == 'mm-dd-yyyy') {
                RegEx = /^(((((((0?[13578])|(1[02]))[\.\-/]?((0?[1-9])|([12]\d)|(3[01])))|(((0?[469])|(11))[\.\-/]?((0?[1-9])|([12]\d)|(30)))|((0?2)[\.\-/]?((0?[1-9])|(1\d)|(2[0-8]))))[\.\-/]?(((19)|(20))?([\d][\d]))))|((0?2)[\.\-/]?(29)[\.\-/]?(((19)|(20))?(([02468][048])|([13579][26])))))((T|\s)(([0-1][0-9])|(2[0-3])):([0-5][0-9])(:([0-5][0-9]))?)?$/gi
                if (value.match(RegEx)) {

                    date = value.split(/[-t\s:.z]/);
                    if (date.length <= 3) { return new Date(parseInt(date[2]), parseInt(date[0]) - 1, parseInt(date[1])); }
                    else if (date.length <= 5) { return new Date(parseInt(date[2]), parseInt(date[0]) - 1, parseInt(date[1]), parseInt(date[3]) + timeZoneOffset, parseInt(date[4])); }
                    else if (date.length >= 6) { return new Date(parseInt(date[2]), parseInt(date[0]) - 1, parseInt(date[1]), parseInt(date[3]) + timeZoneOffset, parseInt(date[4]), parseInt(date[5])); }

                }
            }

            //DD-MON-YYYY || DD MON YYYY (THH:MM:SS OPTIONAL)
            if (mask == null || mask == 'dd-mon-yyyy') {
                RegEx = /^((31(?! (FEB|APR|JUN|SEP|NOV)))|((30|29)(?! FEB))|(29(?= FEB (((1[6-9]|[2-9]\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00)))))|(0?[1-9])|1\d|2[0-8])[-\s](JAN|FEB|MAR|MAY|APR|JUL|JUN|AUG|OCT|SEP|NOV|DEC)[-\s]((1[6-9]|[2-9]\d)\d{2})((T|\s)(([0-1][0-9])|(2[0-3])):([0-5][0-9])(:([0-5][0-9]))?)?$/gi; //((T|\s)(([0-1][0-9])|(2[0-3])):([0-5][0-9])(:([0-5][0-9]))?)?
                if (value.match(RegEx)) {

                    value = value.replace(/oct/, 'ocx');
                    date = value.split(/[-t\s:]/);
                    if (date.length <= 3) { return new Date(parseInt(date[2]), ShortMonths[date[1]], parseInt(date[0])); }
                    else if (date.length <= 5) { return new Date(parseInt(date[2]), ShortMonths[date[1]], parseInt(date[0]), parseInt(date[3]) + timeZoneOffset, parseInt(date[4])); }
                    else if (date.length >= 6) { return new Date(parseInt(date[2]), ShortMonths[date[1]], parseInt(date[0]), parseInt(date[3]) + timeZoneOffset, parseInt(date[4]), parseInt(date[5])); }

                }
            }




        }
        return Number.NaN;
    }
}

function ip_parseNumber(value, decimals) {
    if (value == null) { return NaN; }
    if (typeof (value) == 'string' && value.match(/[^.0-9%]/)) { return NaN }

    value = parseFloat(value);

    if (decimals != null) { return value.toFixed(decimals); }

    return value;

}

function ip_parseCurrency(value, decimals) {

    if (value == null) { return null; }

    var processedVal = value;

    if (typeof (value) == 'string') {
        var processedVal = value.replace(/[$R]/gi, ''); //all symboles for currency    
        if (processedVal.match(/[^.0-9]/)) { return NaN }
    }

    var numberVal = ip_parseNumber(processedVal, decimals);

    return numberVal;
}

function ip_parseString(value) {
    if (value == null) { return ''; }
    return value.toString();
}

function ip_parseRange(GridID, value) {

    if (value == null || value == '' || isNaN(value)) { return null; }
    return ip_fxRangeObject(GridID, null, null, value.trim());

}

function ip_parseAny(GridID, value) {

    var val = null;

    if (value == null) { return value; }
    else if (val = ip_parseDate(value)) { return val; }
    else if (val = ip_parseBool(value)) { return val; }
    else if (val = ip_parseNumber(value)) { return val; }
    else if (val = ip_parseRange(GridID, value)) { return val; }

    return ip_parseString(value);

}

function ip_formatDate(GridID, value, oldMask, newMask) {

    var date = ip_parseDate(value, oldMask);
    if (!date && oldMask && GridID) { date = ip_GridProps[GridID].mask.input[oldMask](value); }
    if (!date) { date = ip_parseDate(value); }
    if (!date) { return false; }
    else {

        try
        {
            var ShortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug','Sep', 'Oct', 'Nov', 'Dec'];
            var dd = date.getDate();
            var mm = date.getMonth() + 1;
            var yy = date.getFullYear();
            var hh = date.getHours();
            var min = date.getMinutes();
            var ss = date.getSeconds();

            var fdd = dd.toString();
            var fmm = mm.toString();
            var fmnth = ShortMonths[mm - 1];
            var fyy = yy.toString();
            var fhh = hh.toString();
            var fmin = min.toString();
            var fss = ss.toString();

            if (dd < 10) { fdd = '0' + dd; }
            if (mm < 10) { fmm = '0' + mm; }
            if (hh < 10) { fhh = '0' + hh; }
            if (min < 10) { fmin = '0' + min; }
            if (ss < 10) { fss = '0' + ss; }

            if (newMask == null) { mask = 'yyyy-mm-dd'; }
            if (newMask == 'yyyy-mm-dd') { return fyy + '-' + fmm + '-' + fdd; }
            else if (newMask == 'yyyy-mm-dd hh:mm') { return fyy + '-' + fmm + '-' + fdd + ' ' + fhh + ':' + fmin; }
            else if (newMask == 'yyyy-mm-dd hh:mm:ss') { return fyy + '-' + fmm + '-' + fdd + ' ' + fhh + ':' + fmin + ':' + fss; }
            else if (newMask == 'dd-mon-yyyy') { return fdd + '-' + fmnth + '-' + fyy; }
            else if (newMask == 'full') { return date.toString(); }
        }
        catch (ex) { return false; }
    }

    return value;
}

function ip_formatNumber(GridID, value, oldMask, newMask, decimals) {

    if (value == null || value == '') { return value; }

    var number = value.toString().replace(/[\s,R$]/gi, '');    

    if (oldMask && GridID) {
        number = ip_GridProps[GridID].mask.input[oldMask](number);
        if (isNaN(number)) { return false; }
        if (number == null) { number = ''; }
        else { number = number.toString(); }
    }

    number = number.split('.');
    var decimal = (number.length > 1 ? '.' + number[1] : decimals != undefined ? '.' : '');

    if (decimals != undefined && decimals <= 0) { decimal = ''; }
    else if (decimals != undefined && decimal.length - 1 > decimals) {
        //Decimals long
        decimal = decimal.substring(0, decimals+1);
    }
    else if (decimals != undefined && decimal.length - 1 < decimals) {
        //decimals too short
        for (var i = decimal.length - 1; i < decimals; i++) { decimal += '0'; }
    }
    
    if (newMask == '123') { return number[0] + decimal; }
    else if (newMask == '1 000 000.00') { return number[0].replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1 ") + decimal; }
    else if (newMask == '1,000,000.00') { return number[0].replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1,") + decimal; }

    return value;


}

function ip_formatCurrency(GridID, value, oldMask, newMask, decimals) {

    var number = value;

    if (newMask == '$1,000,000.00') {
        number = ip_formatNumber(GridID, value, oldMask, '1,000,000.00', decimals);
        if (number == false) { return false; } else { number = '$' + number; }
    }
    if (newMask == 'R1,000,000.00') {
        number = ip_formatNumber(GridID, value, oldMask, '1,000,000.00', decimals);
        if (number == false) { return false; } else { number = 'R' + number; }
    }

    return number; 

}


//----- SPECIFIC FUNCTIONS FOR IP GRID ------------------------------------------------------------------------------------------------------------------------------------

function ip_FocusGrid(GridID, raiseEvent) {

    var TransactionID = ip_GenerateTransactionID();
    var oldGridID = ip_GridProps['index'].focusedGrid;

    if (oldGridID != GridID || raiseEvent) {
        ip_GridProps['index'].focusedGrid = GridID;
        ip_RaiseEvent(GridID, 'ip_FocusGrid', TransactionID, { FocusGrid: { Inputs: null, Effected: { GridID: GridID } } });
    }

}

function ip_GetCursorPos(GridID, element) {

    var caretOffset = 0;
    var rangelength = 0;
    var text = '';

    if (element != null)
    {
        if (!ip_GridProps[GridID].editing.selectionState) { ip_EnableSelection(GridID);  }

        if (typeof window.getSelection != "undefined") {
                  
            try
            {
                var sel = window.getSelection && window.getSelection();
                var range = sel.getRangeAt(0);
        
                text = range.toString();
                rangelength = text.length;

                var preCaretRange = range.cloneRange();
                preCaretRange.selectNodeContents(element);
                preCaretRange.setEnd(range.endContainer, range.endOffset);
                caretOffset = preCaretRange.toString().length;
            }
            catch (ex) {
            }

        } else if (typeof document.selection != "undefined" && document.selection.type != "Control") {

            var textRange = document.selection.createRange();
            text = textRange.toString();
            rangelength = text.length;

            var preCaretTextRange = document.body.createTextRange();
            preCaretTextRange.moveToElementText(element);
            preCaretTextRange.setEndPoint("EndToEnd", textRange);
            caretOffset = preCaretTextRange.text.length;

        }

    }

    return { x: (caretOffset - rangelength), length: rangelength, text: text };
}

function ip_SetCursorPos(GridID, element, start, end) {

    var carret = { x: start, length: 0, text:'' }
       
    if (element != null) {

        if (ip_GridProps[GridID] != undefined && !ip_GridProps[GridID].editing.selectionState) { ip_EnableSelection(GridID); }

        carret.text = element.innerText || element.textContent;

        if (carret.text.length > start && (start != null || end != null)) {
                
            var rng = document.createRange(),
                    sel = getSelection(),
                    n,
                    o = 0,
                    x2 = 0,
                    tw = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null, null);

            while (n = tw.nextNode()) {
                o += n.nodeValue.length;
                if (o > start) {
                    carret.x = n.nodeValue.length + start - o;
                    rng.setStart(n, carret.x);
                    start = Infinity;
                }
                if (end != null && o >= end) {
                    x2 = n.nodeValue.length + end - o;
                    rng.setEnd(n, x2);
                    break;
                }
            }

            sel.removeAllRanges();
            sel.addRange(rng);

        }
        else {


            var range = document.createRange();//Create a range (a range is a like the selection but invisible)
            range.selectNodeContents(element);//Select the entire contents of the element with the range
            range.collapse(false);//collapse the range to the end point. false means collapse to end rather than the start        
            var selection = window.getSelection();//get the selection object (allows you to change selection)
            selection.removeAllRanges();//remove any selections already made
            selection.addRange(range);//make the range you have just created the visible selection
                
            carret.x = carret.text.length;
        }

        carret.length = x2 - carret.x;
        if (carret.length < 0) { carret.length = 0; }

        return carret;

    }
       
       
    return null;
}

function ip_EnableSelection(GridID, removeRanges) {

    //clearTimeout(ip_GridProps[GridID].timeouts.disableSelectionTimout);
    $(document).enableSelection();
    if (removeRanges) { document.getSelection().removeAllRanges(); }
    ip_GridProps[GridID].editing.selectionState = true;
    //debugtext('enabled selection', true);
}

function ip_DisableSelection(GridID, removeRanges) {

    try{
        if (removeRanges) { document.getSelection().removeAllRanges(); }
        $(document).disableSelection();
        ip_GridProps[GridID].editing.selectionState = false;
    }
    catch (ex) {
        console.warn(ex);
    }


}

function ip_AppendEffectedRowData(GridID, Effected, cell) {

    var rIndex = 0;
    if (Effected.rowData == null)
    {
        Effected.rowData = []; rIndex = 0;        
    }
    else
    {        
        if (Effected.rowData.length > 0 && Effected.rowData[rIndex].cells != null && Effected.rowData[rIndex].cells.length == ip_GridProps[GridID].cols) { rIndex++; }
        else if (Effected.rowData.length > 0 && Effected.rowData[rIndex].cells != null) { rIndex = Effected.rowData.length - 1; }
    }

    if (Effected.rowData[rIndex] == null) { Effected.rowData[rIndex] = { cells: [] } }
    Effected.rowData[rIndex].cells[Effected.rowData[rIndex].cells.length] = cell;
    
    return Effected;
}

function ip_PurgeEffectedRowData(GridID, Effected) {
    //Removes duplicate cells, keeping the ones added to the array last 

    if (Effected.rowData == null) { return Effected; }
    
    var EffectedCells = {};
    for (var r = Effected.rowData.length - 1; r >= 0; r--) {
        
        var rowData = Effected.rowData[r];
        for (var c = rowData.cells.length - 1; c >= 0; c--) {

            var index = rowData.cells[c].row + '-' + rowData.cells[c].col;
            if (EffectedCells[index] == null) { EffectedCells[index] = true; }
            else { rowData.cells.splice(c, 1); }
            
        }

        if (rowData.cells.length == 0) { Effected.rowData.splice(r, 1); }
    }

    return Effected;
}

function replaceAt(GridID, source, index, matchText, replaceText) {
    return source.substr(0, index) + replaceText + source.substr(index + matchText.length);
}

function ip_Browser() {

    var thisBrowser = { name:'', version:'' }
    
    //Normalize name
    if (navigator.appName == 'Microsoft Internet Explorer' || !!window.MSStream) { thisBrowser.name = 'ie' }
    else { thisBrowser.name = navigator.appName }

    //Normalize version
    if (navigator.appVersion.indexOf('MSIE 10') != -1) { thisBrowser.version = 10; }
    else if (navigator.appVersion.indexOf('MSIE 9') != -1) { thisBrowser.version = 9; }
    else if (navigator.appVersion.indexOf('MSIE 8') != -1) { thisBrowser.version = 8; }
    else if (navigator.appVersion.indexOf('MSIE 7') != -1) { thisBrowser.version = 7; }
    else if (navigator.appVersion.indexOf('MSIE 6') != -1) { thisBrowser.version = 6; }
    else { thisBrowser.version = navigator.appVersion }
        
    return thisBrowser;
}

function ip_RaiseEvent(GridID, eType, transactionID, args) {
    
    switch (eType)
    {
        case 'warning': ip_ShowFooterAlert('hmmm', args, '', 'red', 3000); break;
        case 'message': ip_ShowFooterAlert('hmmm', args, '', 'lightblue', 3000); break;
        case 'error': ip_ShowFooterAlert('hmmm', args, '', 'red', 3000); break;
        case 'ip_CellInput': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_SetCellValues': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_ResizeColumn': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_ResizeRow': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_ColumnSelector': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_RowSelector': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_RemoveRow': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_AddRow': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_InsertRow': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_RemoveCol': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_AddCol': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_InsertCol': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_FrozenRowsCols': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_MoveCol': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_MoveRow': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_MergeRange': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_UnMergeRange': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_Paste': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_ResetRange': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_ScrollComplete': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_Sort': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_Undo': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_HideShowRows': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_HideShowColumns': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_GroupRows': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_UngroupRows': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_FormatColumn': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_FormatCell': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_DragRange': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_SelectCell': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_SelectColumn': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_SelectRow': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_SelectRange': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_ResizeGrid': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_DisposeGrid': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_Border': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_GridMeta': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
        case 'ip_FocusGrid': $('#' + GridID).trigger(eType, ip_eventObject(GridID, eType, transactionID, args)); break;
            
    }
}

function ip_UnbindAllEvents(GridID) {

    $('#' + GridID).find("*").andSelf().unbind();

    if (ip_GridProps[GridID] != null) {

        //Unbind non grid events (events that are not bound on the grid dom)
        ip_GridProps[GridID].events.SetupEvents_document_mouseup = ip_UnBindEvent(document, 'mouseup', ip_GridProps[GridID].events.SetupEvents_document_mouseup);
        ip_GridProps[GridID].events.SetupEvents_document_keydown = ip_UnBindEvent(document, 'keydown', ip_GridProps[GridID].events.SetupEvents_document_keydown);
        ip_GridProps[GridID].events.SetupEvents_document_keyup = ip_UnBindEvent(document, 'keyup', ip_GridProps[GridID].events.SetupEvents_document_keyup);

    }
}

function ip_UnBindEvent(element, eventName, eventFn) {

    if (eventFn != null) {
        $(element).unbind(eventName, eventFn);
        eventFn = null;
    }

    return null;

}

function ip_SetHoverCell(GridID, topElement, e, allBrowsers, callBack) {

    allBrowsers = (allBrowsers == null ? false : allBrowsers);

    if (ip_GridProps['index'].browser.name == 'ie' || allBrowsers) {

        var loopCount = 0;
        var elementsHidden = new Array();
        var x = e.clientX;
        var y = e.clientY;
        var hoverElement = null;

        //Seeks out the cell for up to 50 elements deep
        while (loopCount < 50) {

            elementsHidden[loopCount] = topElement;
            $(topElement).hide();
            var e = document.elementFromPoint(x, y);

            if (e.id == '') { hoverElement = e.parentNode } else { hoverElement = e; }

            if (hoverElement.id.indexOf('_columnSelectorCell') > -1 || hoverElement.id.indexOf('_rowSelecterCell') > -1 || hoverElement.id.indexOf('_cell') > -1) {

                for (var i = 0; i < elementsHidden.length; i++) { $(elementsHidden[i]).show(); }
                ip_GridProps[GridID].hoverCell = hoverElement;
                return hoverElement;

            }
            else { topElement = hoverElement; }

            loopCount++;
        }

    }

    $(elementsHidden).show();

    return null;
}

function ip_ReplaceInTags(GridID, SourceString, OpenTag, CloseTag, NewValue, removeTagIfEmpty) {

    var Result = SourceString;
    var StartI = SourceString.indexOf(OpenTag);

    if (StartI >= 0) {

        EndI = SourceString.indexOf(CloseTag, StartI);
        if (EndI == -1) { EndI = SourceString.length; }

        if (!removeTagIfEmpty || NewValue != '') {
            Result = SourceString.substring(StartI, EndI);
            Result = SourceString.replace(Result, OpenTag + NewValue);
        }
        else
        {
            Result = SourceString.substring(StartI, EndI + 1);
            Result = SourceString.replace(Result, '');
        }
    }
    else if(NewValue!='') { Result += OpenTag + NewValue + CloseTag; }

    return Result;

}

function ip_ReplaceCssProperty(GridID, SourceString, PropertyName, ProperyValue) {

    var Added = false;
    var css = '';

    if (SourceString == null) { SourceString = ''; }

    var Properties = SourceString.split(';');

    
    //if (PropertyName == 'color') { PropertyName = '/*x*/' + PropertyName; }

    for (var p = 0; p < Properties.length; p++) {

        var attributes = Properties[p].split(':');

        if (attributes.length > 1) {
            
            if (attributes[0] == PropertyName) {
                
                css += PropertyName + ':' + ProperyValue + ';';
                Added = true;

            }
            else {
                css += attributes[0] + ':' + attributes[1] + ';';
            }
        }

    }

    if (!Added) { css += PropertyName + ':' + ProperyValue + ';'; }

    return css;
}

function debugtext(text, append) {

    if (append) { text = text + '<br />' + $('#output').html(); }

    
    $('#output').html(text);
    

}

function ouputmerges(GridID, what) {

    debugtext('');


    if (what == null || what == 'col') {

        debugtext('COL-INDEX ', true);
        debugtext('', true);

        for (var c = 0; c < ip_GridProps[GridID].cols; c++) {

            if (ip_GridProps[GridID].colData[c].containsMerges != null) {

                for (var m = 0; m < ip_GridProps[GridID].colData[c].containsMerges.length; m++) {

                    var merge = ip_GridProps[GridID].colData[c].containsMerges[m];
                    debugtext('COLUMN ' + c, true);
                    debugtext('merge : ' + merge.mergedWithRow + ' x ' + merge.mergedWithCol, true);
                    debugtext('rspan : ' + merge.rowSpan, true);
                    debugtext('cspan : ' + merge.colSpan, true);
                    debugtext('--------------------------------', true);

                }

            }

        }
    }


    if (what == null || what == 'row') {

        debugtext('', true);
        debugtext('', true);
        debugtext('ROW-INDEX ', true);
        debugtext('', true);

        for (var r = 0; r < ip_GridProps[GridID].rows; r++) {

            if (ip_GridProps[GridID].rowData[r].containsMerges != null) {

                for (var m = 0; m < ip_GridProps[GridID].rowData[r].containsMerges.length; m++) {

                    var merge = ip_GridProps[GridID].rowData[r].containsMerges[m];
                    debugtext('ROW ' + r, true);
                    debugtext('merge : ' + merge.mergedWithRow + ' x ' + merge.mergedWithCol, true);
                    debugtext('rspan : ' + merge.rowSpan, true);
                    debugtext('cspan : ' + merge.colSpan, true);
                    debugtext('--------------------------------', true);

                }

            }

        }
    }



    if (what == null || what == 'cell') {


        debugtext('', true);
        debugtext('', true);
        debugtext('CELL-INDEX ', true);
        debugtext('', true);


        for (var r = 0; r < ip_GridProps[GridID].rows; r++) {

            for (var c = 0; c < ip_GridProps[GridID].cols; c++) {

                if (ip_GridProps[GridID].rowData[r].cells[c].merge != null) {

                    var merge = ip_GridProps[GridID].rowData[r].cells[c].merge;
                    debugtext('CELL ' + r + 'x' + c, true);
                    debugtext('merge : ' + merge.mergedWithRow + ' x ' + merge.mergedWithCol, true);
                    debugtext('rspan : ' + merge.rowSpan, true);
                    debugtext('cspan : ' + merge.colSpan, true);
                    debugtext('--------------------------------', true);

                }

         

            }

        }
    }
}

function ip_GeneratePublicKey() {
    'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });

    var GUID = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) { var r = Math.random() * 16 | 0, v = c == 'x' ? r : r & 0x3 | 0x8; return v.toString(16); });
    return GUID;
}

function ip_GenerateTransactionID()
{
    return ip_GeneratePublicKey();
}

//var ip_ShowFooterAlert_Timeout = null;
function ip_ShowFooterAlert(Title, Message, Icon, BackColor, Duration) {

    if (BackColor == null) { BackColor = ''; }
    if (Duration == null) { Duration = 5000; }

    $(document.body).ip_FooterDialog({
        
        Message:Message,
        BackColor: BackColor,
        Duration: Duration

    });

    //clearTimeout(ip_ShowFooterAlert_Timeout);

    //if (BackColor == null) { BackColor = ''; }
    //if (Duration == null) { Duration = 5000; }

    //var FooterObject = $('.ip_FooterDialog');
    //if (FooterObject.length == 0) {

    //    Footer = '<div class="ip_FooterDialog"></div>';
    //    $(document.body).append(Footer);
    //    FooterObject = $('.ip_FooterDialog')[0];

    //}
    //else { FooterObject = FooterObject[0]; }
    
    //$(FooterObject).html(Message.replace(/([a-z])([A-Z])/g, '$1 $2'));
    //$(FooterObject).css('background-color', BackColor);

    //$(FooterObject).slideDown(200);

    //ip_ShowFooterAlert_Timeout = setTimeout(function () { $(FooterObject).slideUp(200); }, Duration);
}







//----- 3rd party dependancies ------------------------------------------------------------------------------------------------------------------------------------

//----- MOUSEWHEEL ------------------------------------------------------------------------------------------------------------------------------------

/*! Copyright (c) 2013 Brandon Aaron (http://brandonaaron.net)
 * Licensed under the MIT License (LICENSE.txt).
 *
 * Thanks to: http://adomas.org/javascript-mouse-wheel/ for some pointers.
 * Thanks to: Mathias Bank(http://www.mathias-bank.de) for a scope bug fix.
 * Thanks to: Seamus Leahy for adding deltaX and deltaY
 *
 * Version: 3.1.3
 *
 * Requires: 1.2.2+
 */

(function (factory) {
    if (typeof define === 'function' && define.amd) {
        // AMD. Register as an anonymous module.
        define(['jquery'], factory);
    } else if (typeof exports === 'object') {
        // Node/CommonJS style for Browserify
        module.exports = factory;
    } else {
        // Browser globals
        factory(jQuery);
    }
}(function ($) {

    var toFix = ['wheel', 'mousewheel', 'DOMMouseScroll', 'MozMousePixelScroll'];
    var toBind = 'onwheel' in document || document.documentMode >= 9 ? ['wheel'] : ['mousewheel', 'DomMouseScroll', 'MozMousePixelScroll'];
    var lowestDelta, lowestDeltaXY;

    if ($.event.fixHooks) {
        for (var i = toFix.length; i;) {
            $.event.fixHooks[toFix[--i]] = $.event.mouseHooks;
        }
    }

    $.event.special.mousewheel = {
        setup: function () {
            if (this.addEventListener) {
                for (var i = toBind.length; i;) {
                    this.addEventListener(toBind[--i], handler, false);
                }
            } else {
                this.onmousewheel = handler;
            }
        },

        teardown: function () {
            if (this.removeEventListener) {
                for (var i = toBind.length; i;) {
                    this.removeEventListener(toBind[--i], handler, false);
                }
            } else {
                this.onmousewheel = null;
            }
        }
    };

    $.fn.extend({
        mousewheel: function (fn) {
            return fn ? this.bind("mousewheel", fn) : this.trigger("mousewheel");
        },

        unmousewheel: function (fn) {
            return this.unbind("mousewheel", fn);
        }
    });


    function handler(event) {
        var orgEvent = event || window.event,
            args = [].slice.call(arguments, 1),
            delta = 0,
            deltaX = 0,
            deltaY = 0,
            absDelta = 0,
            absDeltaXY = 0,
            fn;
        event = $.event.fix(orgEvent);
        event.type = "mousewheel";

        // Old school scrollwheel delta
        if (orgEvent.wheelDelta) { delta = orgEvent.wheelDelta; }
        if (orgEvent.detail) { delta = orgEvent.detail * -1; }

        // New school wheel delta (wheel event)
        if (orgEvent.deltaY) {
            deltaY = orgEvent.deltaY * -1;
            delta = deltaY;
        }
        if (orgEvent.deltaX) {
            deltaX = orgEvent.deltaX;
            delta = deltaX * -1;
        }

        // Webkit
        if (orgEvent.wheelDeltaY !== undefined) { deltaY = orgEvent.wheelDeltaY; }
        if (orgEvent.wheelDeltaX !== undefined) { deltaX = orgEvent.wheelDeltaX * -1; }

        // Look for lowest delta to normalize the delta values
        absDelta = Math.abs(delta);
        if (!lowestDelta || absDelta < lowestDelta) { lowestDelta = absDelta; }
        absDeltaXY = Math.max(Math.abs(deltaY), Math.abs(deltaX));
        if (!lowestDeltaXY || absDeltaXY < lowestDeltaXY) { lowestDeltaXY = absDeltaXY; }

        // Get a whole value for the deltas
        fn = delta > 0 ? 'floor' : 'ceil';
        delta = Math[fn](delta / lowestDelta);
        deltaX = Math[fn](deltaX / lowestDeltaXY);
        deltaY = Math[fn](deltaY / lowestDeltaXY);

        // Add event and delta to the front of the arguments
        args.unshift(event, delta, deltaX, deltaY);

        return ($.event.dispatch || $.event.handle).apply(this, args);
    }

}));
