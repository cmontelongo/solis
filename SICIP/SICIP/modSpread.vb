Option Strict Off
Option Explicit On
Module modSpread
	'----------------------------------------------------------
	'
	' File: SSOCX.BAS
	'
	' Copyright (C) 2002 FarPoint Technologies.
	' All rights reserved.
	'
	'----------------------------------------------------------
	
	' ********** SPREADSHEET PROPERTY SETTINGS **********
	
	' Action property settings
	Public Const SS_ACTION_ACTIVE_CELL As Short = 0
	Public Const SS_ACTION_GOTO_CELL As Short = 1
	Public Const SS_ACTION_SELECT_BLOCK As Short = 2
	Public Const SS_ACTION_CLEAR As Short = 3
	Public Const SS_ACTION_DELETE_COL As Short = 4
	Public Const SS_ACTION_DELETE_ROW As Short = 5
	Public Const SS_ACTION_INSERT_COL As Short = 6
	Public Const SS_ACTION_INSERT_ROW As Short = 7
	Public Const SS_ACTION_RECALC As Short = 11
	Public Const SS_ACTION_CLEAR_TEXT As Short = 12
	Public Const SS_ACTION_PRINT As Short = 13
	Public Const SS_ACTION_DESELECT_BLOCK As Short = 14
	Public Const SS_ACTION_DSAVE As Short = 15
	Public Const SS_ACTION_SET_CELL_BORDER As Short = 16
	Public Const SS_ACTION_ADD_MULTISELBLOCK As Short = 17
	Public Const SS_ACTION_GET_MULTI_SELECTION As Short = 18
	Public Const SS_ACTION_COPY_RANGE As Short = 19
	Public Const SS_ACTION_MOVE_RANGE As Short = 20
	Public Const SS_ACTION_SWAP_RANGE As Short = 21
	Public Const SS_ACTION_CLIPBOARD_COPY As Short = 22
	Public Const SS_ACTION_CLIPBOARD_CUT As Short = 23
	Public Const SS_ACTION_CLIPBOARD_PASTE As Short = 24
	Public Const SS_ACTION_SORT As Short = 25
	Public Const SS_ACTION_COMBO_CLEAR As Short = 26
	Public Const SS_ACTION_COMBO_REMOVE As Short = 27
	Public Const SS_ACTION_RESET As Short = 28
	Public Const SS_ACTION_SEL_MODE_CLEAR As Short = 29
	Public Const SS_ACTION_VMODE_REFRESH As Short = 30
	Public Const SS_ACTION_SMARTPRINT As Short = 32
	
	' Appearance property settings
	Public Const SS_APPEARANCE_FLAT As Short = 0
	Public Const SS_APPEARANCE_3D As Short = 1
	Public Const SS_APPEARANCE_3DWITHBORDER As Short = 2
	
	' BackColorStyle property settings
	Public Const SS_BACKCOLORSTYLE_OVERGRID As Short = 0
	Public Const SS_BACKCOLORSTYLE_UNDERGRID As Short = 1
	Public Const SS_BACKCOLORSTYLE_OVERHORZGRIDONLY As Short = 2
	Public Const SS_BACKCOLORSTYLE_OVERVERTGRIDONLY As Short = 3
	
	' ButtonDrawMode property settings
	Public Const SS_BDM_ALWAYS As Short = 0
	Public Const SS_BDM_CURRENT_CELL As Short = 1
	Public Const SS_BDM_CURRENT_COLUMN As Short = 2
	Public Const SS_BDM_CURRENT_ROW As Short = 4
	Public Const SS_BDM_ALWAYS_BUTTON As Short = 8
	Public Const SS_BDM_ALWAYS_COMBO As Short = 16
	
	' CellBorderStyle property settings
	Public Const SS_BORDER_STYLE_DEFAULT As Short = 0
	Public Const SS_BORDER_STYLE_SOLID As Short = 1
	Public Const SS_BORDER_STYLE_DASH As Short = 2
	Public Const SS_BORDER_STYLE_DOT As Short = 3
	Public Const SS_BORDER_STYLE_DASH_DOT As Short = 4
	Public Const SS_BORDER_STYLE_DASH_DOT_DOT As Short = 5
	Public Const SS_BORDER_STYLE_BLANK As Short = 6
	Public Const SS_BORDER_STYLE_FINE_SOLID As Short = 11
	Public Const SS_BORDER_STYLE_FINE_DASH As Short = 12
	Public Const SS_BORDER_STYLE_FINE_DOT As Short = 13
	Public Const SS_BORDER_STYLE_FINE_DASH_DOT As Short = 14
	Public Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT As Short = 15
	
	' CellBorderType property settings
	Public Const SS_BORDER_TYPE_NONE As Short = 0
	Public Const SS_BORDER_TYPE_LEFT As Short = 1
	Public Const SS_BORDER_TYPE_RIGHT As Short = 2
	Public Const SS_BORDER_TYPE_TOP As Short = 4
	Public Const SS_BORDER_TYPE_BOTTOM As Short = 8
	Public Const SS_BORDER_TYPE_OUTLINE As Short = 16
	
	' CellNoteIndicator property settings
	Public Const SS_CELLNOTEINDICATOR_SHOWANDFIREEVENT As Short = 0
	Public Const SS_CELLNOTEINDICATOR_SHOWANDDONOTFIREEVENT As Short = 1
	Public Const SS_CELLNOTEINDICATOR_DONOTSHOWANDFIREEVENT As Short = 2
	Public Const SS_CELLNOTEINDICATOR_DONOTSHOWANDDONOTFIREEVENT As Short = 3
	
	' CellType property settings
	Public Const SS_CELL_TYPE_DATE As Short = 0
	Public Const SS_CELL_TYPE_EDIT As Short = 1
	Public Const SS_CELL_TYPE_FLOAT As Short = 2
	Public Const SS_CELL_TYPE_INTEGER As Short = 3
	Public Const SS_CELL_TYPE_PIC As Short = 4
	Public Const SS_CELL_TYPE_STATIC_TEXT As Short = 5
	Public Const SS_CELL_TYPE_TIME As Short = 6
	Public Const SS_CELL_TYPE_BUTTON As Short = 7
	Public Const SS_CELL_TYPE_COMBOBOX As Short = 8
	Public Const SS_CELL_TYPE_PICTURE As Short = 9
	Public Const SS_CELL_TYPE_CHECKBOX As Short = 10
	Public Const SS_CELL_TYPE_OWNER_DRAWN As Short = 11
	Public Const SS_CELL_TYPE_CURRENCY As Short = 12
	Public Const SS_CELL_TYPE_NUMBER As Short = 13
	Public Const SS_CELL_TYPE_PERCENT As Short = 14
	
	' ClipboardOptions property settings
	Public Const SS_CLIP_NOHEADERS As Short = 0
	Public Const SS_CLIP_COPYROWHEADERS As Short = 1
	Public Const SS_CLIP_PASTEROWHEADERS As Short = 2
	Public Const SS_CLIP_COPYCOLHEADERS As Short = 4
	Public Const SS_CLIP_PASTECOLHEADERS As Short = 8
	Public Const SS_CLIP_COPYPASTEALLHEADERS As Short = 15
	
	' ColHeadersAutoText and RowHeadersAutoText property settings
	Public Const SS_HEADER_BLANK As Short = 0
	Public Const SS_HEADER_NUMBERS As Short = 1
	Public Const SS_HEADER_LETTERS As Short = 2
	
	' ColMerge and RowMerge property settings
	Public Const SS_MERGE_NONE As Short = 0
	Public Const SS_MERGE_ALWAYS As Short = 1
	Public Const SS_MERGE_RESTRICTED As Short = 2
	
	' ColUserSortIndicator property settings
	Public Const SS_COLUSERSORTINDICATOR_NONE As Short = 0
	Public Const SS_COLUSERSORTINDICATOR_ASCENDING As Short = 1
	Public Const SS_COLUSERSORTINDICATOR_DESCENDING As Short = 2
	Public Const SS_COLUSERSORTINDICATOR_DISABLED As Short = 3
	
	' CursorStyle property settings
	Public Const SS_CURSOR_STYLE_USER_DEFINED As Short = 0
	Public Const SS_CURSOR_STYLE_DEFAULT As Short = 1
	Public Const SS_CURSOR_STYLE_ARROW As Short = 2
	Public Const SS_CURSOR_STYLE_DEFCOLRESIZE As Short = 3
	Public Const SS_CURSOR_STYLE_DEFROWRESIZE As Short = 4
	
	' CursorType property settings
	Public Const SS_CURSOR_TYPE_DEFAULT As Short = 0
	Public Const SS_CURSOR_TYPE_COLRESIZE As Short = 1
	Public Const SS_CURSOR_TYPE_ROWRESIZE As Short = 2
	Public Const SS_CURSOR_TYPE_BUTTON As Short = 3
	Public Const SS_CURSOR_TYPE_GRAYAREA As Short = 4
	Public Const SS_CURSOR_TYPE_LOCKEDCELL As Short = 5
	Public Const SS_CURSOR_TYPE_COLHEADER As Short = 6
	Public Const SS_CURSOR_TYPE_ROWHEADER As Short = 7
	Public Const SS_CURSOR_TYPE_DRAGDROPAREA As Short = 8
	Public Const SS_CURSOR_TYPE_DRAGDROP As Short = 9
	
	' DAutoSizeCols property settings
	Public Const SS_AUTOSIZE_NO As Short = 0
	Public Const SS_AUTOSIZE_MAX_COL_WIDTH As Short = 1
	Public Const SS_AUTOSIZE_BEST_GUESS As Short = 2
	
	' EditEnterAction property settings
	Public Const SS_CELL_EDITMODE_EXIT_NONE As Short = 0
	Public Const SS_CELL_EDITMODE_EXIT_UP As Short = 1
	Public Const SS_CELL_EDITMODE_EXIT_DOWN As Short = 2
	Public Const SS_CELL_EDITMODE_EXIT_LEFT As Short = 3
	Public Const SS_CELL_EDITMODE_EXIT_RIGHT As Short = 4
	Public Const SS_CELL_EDITMODE_EXIT_NEXT As Short = 5
	Public Const SS_CELL_EDITMODE_EXIT_PREVIOUS As Short = 6
	Public Const SS_CELL_EDITMODE_EXIT_SAME As Short = 7
	Public Const SS_CELL_EDITMODE_EXIT_NEXTROW As Short = 8
	
	' OperationMode property settings
	Public Const SS_OP_MODE_NORMAL As Short = 0
	Public Const SS_OP_MODE_READONLY As Short = 1
	Public Const SS_OP_MODE_ROWMODE As Short = 2
	Public Const SS_OP_MODE_SINGLE_SELECT As Short = 3
	Public Const SS_OP_MODE_MULTI_SELECT As Short = 4
	Public Const SS_OP_MODE_EXT_SELECT As Short = 5
	
	' Position property settings
	Public Const SS_POSITION_UPPER_LEFT As Short = 0
	Public Const SS_POSITION_UPPER_CENTER As Short = 1
	Public Const SS_POSITION_UPPER_RIGHT As Short = 2
	Public Const SS_POSITION_CENTER_LEFT As Short = 3
	Public Const SS_POSITION_CENTER_CENTER As Short = 4
	Public Const SS_POSITION_CENTER_RIGHT As Short = 5
	Public Const SS_POSITION_BOTTOM_LEFT As Short = 6
	Public Const SS_POSITION_BOTTOM_CENTER As Short = 7
	Public Const SS_POSITION_BOTTOM_RIGHT As Short = 8
	
	' PrintOrientation property settings
	Public Const SS_PRINTORIENT_DEFAULT As Short = 0
	Public Const SS_PRINTORIENT_PORTRAIT As Short = 1
	Public Const SS_PRINTORIENT_LANDSCAPE As Short = 2
	
	' PrintPageOrder property settings
	Public Const SS_PAGEORDER_AUTO As Short = 0
	Public Const SS_PAGEORDER_DOWNTHENOVER As Short = 1
	Public Const SS_PAGEORDER_OVERTHENDOWN As Short = 2
	
	' PrintType property settings
	Public Const SS_PRINT_ALL As Short = 0
	Public Const SS_PRINT_CELL_RANGE As Short = 1
	Public Const SS_PRINT_CURRENT_PAGE As Short = 2
	Public Const SS_PRINT_PAGE_RANGE As Short = 3
	
	' ScrollBars property settings
	Public Const SS_SCROLLBAR_NONE As Short = 0
	Public Const SS_SCROLLBAR_H_ONLY As Short = 1
	Public Const SS_SCROLLBAR_V_ONLY As Short = 2
	Public Const SS_SCROLLBAR_BOTH As Short = 3
	
	' ScrollBarTrack property settings
	Public Const SS_SCROLLBARTRACK_OFF As Short = 0
	Public Const SS_SCROLLBARTRACK_VERTICAL As Short = 1
	Public Const SS_SCROLLBARTRACK_HORIZONTAL As Short = 2
	Public Const SS_SCROLLBARTRACK_BOTH As Short = 3
	
	' SelBackColor property settings
	Public Const SPREAD_COLOR_NONE As Integer = &H8000000B
	
	' SelectBlockOptions property settings
	Public Const SS_SELBLOCKOPT_COLS As Short = 1
	Public Const SS_SELBLOCKOPT_ROWS As Short = 2
	Public Const SS_SELBLOCKOPT_BLOCKS As Short = 4
	Public Const SS_SELBLOCKOPT_ALL As Short = 8
	
	' ShowScrollTips property settings
	Public Const SS_SHOWSCROLLTIPS_OFF As Short = 0
	Public Const SS_SHOWSCROLLTIPS_VERT As Short = 1
	Public Const SS_SHOWSCROLLTIPS_HORZ As Short = 2
	Public Const SS_SHOWSCROLLTIPS_BOTH As Short = 3
	
	' SortKeyOrder property settings
	Public Const SS_SORT_ORDER_NONE As Short = 0
	Public Const SS_SORT_ORDER_ASCENDING As Short = 1
	Public Const SS_SORT_ORDER_DESCENDING As Short = 2
	
	' TextTip property settings
	Public Const SS_TEXTTIP_OFF As Short = 0
	Public Const SS_TEXTTIP_FIXED As Short = 1
	Public Const SS_TEXTTIP_FLOATING As Short = 2
	Public Const SS_TEXTTIP_FIXEDFOCUSONLY As Short = 3
	Public Const SS_TEXTTIP_FLOATINGFOCUSONLY As Short = 4
	
	' TypeButtonAlign property settings
	Public Const SS_CELL_BUTTON_ALIGN_BOTTOM As Short = 0
	Public Const SS_CELL_BUTTON_ALIGN_TOP As Short = 1
	Public Const SS_CELL_BUTTON_ALIGN_LEFT As Short = 2
	Public Const SS_CELL_BUTTON_ALIGN_RIGHT As Short = 3
	
	' TypeButtonType property settings
	Public Const SS_CELL_BUTTON_NORMAL As Short = 0
	Public Const SS_CELL_BUTTON_TWO_STATE As Short = 1
	
	' TypeCheckTextAlign property settings
	Public Const SS_CHECKBOX_TEXT_LEFT As Short = 0
	Public Const SS_CHECKBOX_TEXT_RIGHT As Short = 1
	
	' TypeCheckType property settings
	Public Const SS_CHECKBOX_NORMAL As Short = 0
	Public Const SS_CHECKBOX_THREE_STATE As Short = 1
	
	' TypeComboBoxAutoSearch property settings
	Public Const SS_COMBOBOX_AUTOSEARCH_NONE As Short = 0
	Public Const SS_COMBOBOX_AUTOSEARCH_SINGLECHAR As Short = 1
	Public Const SS_COMBOBOX_AUTOSEARCH_MULTIPLECHAR As Short = 2
	Public Const SS_COMBOBOX_AUTOSEARCH_SINGLECHARGREATER As Short = 3
	
	'TypeComboBoxWidth property settings
	Public Const SS_COMBOWIDTH_CELLWIDTH As Short = 0
	Public Const SS_COMBOWIDTH_AUTORIGHT As Short = 1
	Public Const SS_COMBOWIDTH_AUTOLEFT As Short = -1
	
	' TypeCurrencyLeadingZero, TypeNumberLeadingZero,
	' TypePercentLeadingZero property settings
	Public Const SS_LEADINGZERO_INTL As Short = 0
	Public Const SS_LEADINGZERO_NO As Short = 1
	Public Const SS_LEADINGZERO_YES As Short = 2
	
	' TypeCurrencyNegStyle property settings
	Public Const SS_CELL_CURRENCY_NEGSTYLE_INTL As Short = 0
	Public Const SS_CELL_CURRENCY_NEGSTYLE_1 As Short = 1
	Public Const SS_CELL_CURRENCY_NEGSTYLE_2 As Short = 2
	Public Const SS_CELL_CURRENCY_NEGSTYLE_3 As Short = 3
	Public Const SS_CELL_CURRENCY_NEGSTYLE_4 As Short = 4
	Public Const SS_CELL_CURRENCY_NEGSTYLE_5 As Short = 5
	Public Const SS_CELL_CURRENCY_NEGSTYLE_6 As Short = 6
	Public Const SS_CELL_CURRENCY_NEGSTYLE_7 As Short = 7
	Public Const SS_CELL_CURRENCY_NEGSTYLE_8 As Short = 8
	Public Const SS_CELL_CURRENCY_NEGSTYLE_9 As Short = 9
	Public Const SS_CELL_CURRENCY_NEGSTYLE_10 As Short = 10
	Public Const SS_CELL_CURRENCY_NEGSTYLE_11 As Short = 11
	Public Const SS_CELL_CURRENCY_NEGSTYLE_12 As Short = 12
	Public Const SS_CELL_CURRENCY_NEGSTYLE_13 As Short = 13
	Public Const SS_CELL_CURRENCY_NEGSTYLE_14 As Short = 14
	Public Const SS_CELL_CURRENCY_NEGSTYLE_15 As Short = 15
	Public Const SS_CELL_CURRENCY_NEGSTYLE_16 As Short = 16
	
	' TypeCurrencyPosStyle property settings
	Public Const SS_CELL_CURRENCY_POSSTYLE_INTL As Short = 0
	Public Const SS_CELL_CURRENCY_POSSTYLE_1 As Short = 1
	Public Const SS_CELL_CURRENCY_POSSTYLE_2 As Short = 2
	Public Const SS_CELL_CURRENCY_POSSTYLE_3 As Short = 3
	Public Const SS_CELL_CURRENCY_POSSTYLE_4 As Short = 4
	
	' TypeDateFormat property settings
	Public Const SS_CELL_DATE_FORMAT_DDMONYY As Short = 0
	Public Const SS_CELL_DATE_FORMAT_DDMMYY As Short = 1
	Public Const SS_CELL_DATE_FORMAT_MMDDYY As Short = 2
	Public Const SS_CELL_DATE_FORMAT_YYMMDD As Short = 3
	Public Const SS_CELL_DATE_FORMAT_YYMM As Short = 4
	Public Const SS_CELL_DATE_FORMAT_MMDD As Short = 5
	Public Const SS_CELL_DATE_FORMAT_DEFAULT As Short = 99
	
	' TypeEditCharCase property settings
	Public Const SS_CELL_EDIT_CASE_LOWER_CASE As Short = 0
	Public Const SS_CELL_EDIT_CASE_NO_CASE As Short = 1
	Public Const SS_CELL_EDIT_CASE_UPPER_CASE As Short = 2
	
	' TypeEditCharSet property settings
	Public Const SS_CELL_EDIT_CHAR_SET_ASCII As Short = 0
	Public Const SS_CELL_EDIT_CHAR_SET_ALPHA As Short = 1
	Public Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC As Short = 2
	Public Const SS_CELL_EDIT_CHAR_SET_NUMERIC As Short = 3
	
	' TypeHAlign property settings
	Public Const SS_CELL_H_ALIGN_LEFT As Short = 0
	Public Const SS_CELL_H_ALIGN_RIGHT As Short = 1
	Public Const SS_CELL_H_ALIGN_CENTER As Short = 2
	
	' TypeNumberNegStyle property settings
	Public Const SS_CELL_NUMBER_NEGSTYLE_INTL As Short = 0
	Public Const SS_CELL_NUMBER_NEGSTYLE_1 As Short = 1
	Public Const SS_CELL_NUMBER_NEGSTYLE_2 As Short = 2
	Public Const SS_CELL_NUMBER_NEGSTYLE_3 As Short = 3
	Public Const SS_CELL_NUMBER_NEGSTYLE_4 As Short = 4
	Public Const SS_CELL_NUMBER_NEGSTYLE_5 As Short = 5
	
	' TypePercentNegStyle property settings
	Public Const SS_CELL_PERCENT_NEGSTYLE_INTL As Short = 0
	Public Const SS_CELL_PERCENT_NEGSTYLE_1 As Short = 1
	Public Const SS_CELL_PERCENT_NEGSTYLE_2 As Short = 2
	Public Const SS_CELL_PERCENT_NEGSTYLE_3 As Short = 3
	Public Const SS_CELL_PERCENT_NEGSTYLE_4 As Short = 4
	Public Const SS_CELL_PERCENT_NEGSTYLE_5 As Short = 5
	Public Const SS_CELL_PERCENT_NEGSTYLE_6 As Short = 6
	Public Const SS_CELL_PERCENT_NEGSTYLE_7 As Short = 7
	Public Const SS_CELL_PERCENT_NEGSTYLE_8 As Short = 8
	
	' TypeTextAlignVert property settings
	Public Const SS_CELL_STATIC_V_ALIGN_BOTTOM As Short = 0
	Public Const SS_CELL_STATIC_V_ALIGN_CENTER As Short = 1
	Public Const SS_CELL_STATIC_V_ALIGN_TOP As Short = 2
	
	' TypeTextOrient property settings
	Public Const SS_CELL_TEXTORIENT_HORIZONTAL As Short = 0
	Public Const SS_CELL_TEXTORIENT_VERTICAL_LTR As Short = 1
	Public Const SS_CELL_TEXTORIENT_DOWN As Short = 2
	Public Const SS_CELL_TEXTORIENT_UP As Short = 3
	Public Const SS_CELL_TEXTORIENT_INVERT As Short = 4
	Public Const SS_CELL_TEXTORIENT_VERTICAL_RTL As Short = 5
	
	' TypeTime24Hour property settings
	Public Const SS_CELL_TIME_12_HOUR_CLOCK As Short = 0
	Public Const SS_CELL_TIME_24_HOUR_CLOCK As Short = 1
	Public Const SS_CELL_TIME_24_HOUR_DEFAULT As Short = 2
	
	' TypeVAlign property settings
	Public Const SS_CELL_V_ALIGN_TOP As Short = 0
	Public Const SS_CELL_V_ALIGN_BOTTOM As Short = 1
	Public Const SS_CELL_V_ALIGN_VCENTER As Short = 2
	
	' UnitType property settings
	Public Const SS_CELL_UNIT_NORMAL As Short = 0
	Public Const SS_CELL_UNIT_VGA As Short = 1
	Public Const SS_CELL_UNIT_TWIPS As Short = 2
	
	' UserColAction property settings
	Public Const SS_USERCOLACTION_DEFAULT As Short = 0
	Public Const SS_USERCOLACTION_SORT As Short = 1
	Public Const SS_USERCOLACTION_SORTNOINDICATOR As Short = 2
	
	' UserResize property settings
	Public Const SS_USER_RESIZE_NONE As Short = 0
	Public Const SS_USER_RESIZE_COL As Short = 1
	Public Const SS_USER_RESIZE_ROW As Short = 2
	Public Const SS_USER_RESIZE_BOTH As Short = 3
	
	' UserResizeCol and UserResizeRow property settings
	Public Const SS_USER_RESIZE_DEFAULT As Short = 0
	Public Const SS_USER_RESIZE_ON As Short = 1
	Public Const SS_USER_RESIZE_OFF As Short = 2
	
	' VScrollSpecialType property settings
	Public Const SS_VSCROLLSPECIAL_NO_HOME_END As Short = 1
	Public Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN As Short = 2
	Public Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN As Short = 4
	
	
	
	' ********** SPREADSHEET METHOD SETTINGS ***********
	
	' ActionKey method settings
	Public Const SS_KBA_CLEAR As Short = 0
	Public Const SS_KBA_CURRENT As Short = 1
	Public Const SS_KBA_POPUP As Short = 2
	
	' AddCustomFunctionExt, GetCustomFunction method Flags parameter settings
	Public Const SS_CUSTFUNC_WANTCELLREF As Short = 1
	Public Const SS_CUSTFUNC_WANTRANGEREF As Short = 2
	
	' CFGetParamInfo method Type parameter settings
	Public Const SS_VALUE_TYPE_LONG As Short = 0
	Public Const SS_VALUE_TYPE_DOUBLE As Short = 1
	Public Const SS_VALUE_TYPE_STR As Short = 2
	Public Const SS_VALUE_TYPE_CELL As Short = 3
	Public Const SS_VALUE_TYPE_RANGE As Short = 4
	
	' CFGetParamInfo method Status parameter settings
	Public Const SS_VALUE_STATUS_OK As Short = 0
	Public Const SS_VALUE_STATUS_ERROR As Short = 1
	Public Const SS_VALUE_STATUS_EMPTY As Short = 2
	
	' GetCellSpan method return values
	Public Const SS_SPAN_NO As Short = 0
	Public Const SS_SPAN_YES As Short = 1
	Public Const SS_SPAN_ANCHOR As Short = 2
	
	' ExportTextFile, ExportRangeToTextFile, ExportToXML and  LoadTextFile
	Public Const SS_EXPORTTEXT_CREATE As Short = 1
	Public Const SS_EXPORTTEXT_APPEND As Short = 2
	Public Const SS_EXPORTTEXT_UNFORMATTED As Short = 4
	Public Const SS_EXPORTTEXT_COLHEADERS As Short = 8
	Public Const SS_EXPORTTEXT_ROWHEADERS As Short = 16
	
	Public Const SS_EXPORTXML_FORMATTED As Short = 0
	Public Const SS_EXPORTXML_UNFORMATTED As Short = 1
	
	Public Const SS_LOADTEXT_NOHEADERS As Short = 0
	Public Const SS_LOADTEXT_COLHEADERS As Short = 1
	Public Const SS_LOADTEXT_ROWHEADERS As Short = 2
	Public Const SS_LOADTEXT_CLEARDATAONLY As Short = 4
	
	' GetRefStyle/SetRefStyle methods return values/parameter settings
	Public Const SS_REFSTYLE_DEFAULT As Short = 0
	Public Const SS_REFSTYLE_A1 As Short = 1
	Public Const SS_REFSTYLE_R1C1 As Short = 2
	
	' PrintSheet flags
	Public Const SS_PRINTFLAGS_NONE As Short = 0
	Public Const SS_PRINTFLAGS_SHOWCOMMONDIALOG As Short = 1
	
	' SearchCol and SearchRow method's SearchFlags values
	Public Const SS_SEARCHFLAGS_NONE As Short = 0
	Public Const SS_SEARCHFLAGS_GREATEROREQUAL As Short = 1
	Public Const SS_SEARCHFLAGS_PARTIALMATCH As Short = 2
	Public Const SS_SEARCHFLAGS_VALUE As Short = 4
	Public Const SS_SEARCHFLAGS_CASESENSITIVE As Short = 8
	Public Const SS_SEARCHFLAGS_SORTEDASCENDING As Short = 16
	Public Const SS_SEARCHFLAGS_SORTEDDESCENDING As Short = 32
	
	' Sort method's SortBy parameter settings
	Public Const SS_SORT_BY_ROW As Short = 0
	Public Const SS_SORT_BY_COL As Short = 1
	
	
	
	' ********** SPREADSHEET EVENT SETTINGS **********
	
	Public Const SS_BEFOREUSERSORT_DEFAULTACTION_CANCEL As Short = 0
	Public Const SS_BEFOREUSERSORT_DEFAULTACTION_AUTOSORT As Short = 1
	Public Const SS_BEFOREUSERSORT_DEFAULTACTION_MANUALSORT As Short = 2
	
	Public Const SS_BEFOREUSERSORT_STATE_NONE As Short = 0
	Public Const SS_BEFOREUSERSORT_STATE_ASCENDING As Short = 1
	Public Const SS_BEFOREUSERSORT_STATE_DESCENDING As Short = 2
	
	' TextTipFetch event MultiLine parameter settings
	Public Const SS_TT_MULTILINE_SINGLE As Short = 0
	Public Const SS_TT_MULTILINE_MULTI As Short = 1
	Public Const SS_TT_MULTILINE_AUTO As Short = 2
	
	
	' ********** PRINT PREVIEW PROPERTY SETTINGS **********
	
	' GrayAreaMarginType property values
	Public Const SPV_GRAYAREAMARGINTYPE_SCALED As Short = 0
	Public Const SPV_GRAYAREAMARGINTYPE_ACTUAL As Short = 1
	
	' MousePointer property values
	Public Const SPV_MOUSEPOINTER_DEFAULT As Short = 0
	Public Const SPV_MOUSEPOINTER_ARROW As Short = 1
	Public Const SPV_MOUSEPOINTER_CROSS As Short = 2
	Public Const SPV_MOUSEPOINTER_I_BEAM As Short = 3
	Public Const SPV_MOUSEPOINTER_ICON As Short = 4
	Public Const SPV_MOUSEPOINTER_SIZE As Short = 5
	Public Const SPV_MOUSEPOINTER_SIZE_NE_SW As Short = 6
	Public Const SPV_MOUSEPOINTER_SIZE_N_S As Short = 7
	Public Const SPV_MOUSEPOINTER_SIZE_NW_SE As Short = 8
	Public Const SPV_MOUSEPOINTER_SIZE_W_E As Short = 9
	Public Const SPV_MOUSEPOINTER_UP_ARROW As Short = 10
	Public Const SPV_MOUSEPOINTER_HOURGLASS As Short = 11
	Public Const SPV_MOUSEPOINTER_NO_DROP As Short = 12
	
	' PageViewType property values
	Public Const SPV_PAGEVIEWTYPE_WHOLE_PAGE As Short = 0
	Public Const SPV_PAGEVIEWTYPE_NORMAL_SIZE As Short = 1
	Public Const SPV_PAGEVIEWTYPE_PERCENTAGE As Short = 2
	Public Const SPV_PAGEVIEWTYPE_PAGE_WIDTH As Short = 3
	Public Const SPV_PAGEVIEWTYPE_PAGE_HEIGHT As Short = 4
	Public Const SPV_PAGEVIEWTYPE_MULTIPLE_PAGES As Short = 5
	
	' ScrollBarH property values
	Public Const SPV_SCROLLBARH_SHOW As Short = 0
	Public Const SPV_SCROLLBARH_AUTO As Short = 1
	Public Const SPV_SCROLLBARH_HIDE As Short = 2
	
	' ScrollBarV property values
	Public Const SPV_SCROLLBARV_SHOW As Short = 0
	Public Const SPV_SCROLLBARV_AUTO As Short = 1
	Public Const SPV_SCROLLBARV_HIDE As Short = 2
	
	' ZoomState property values
	Public Const SPV_ZOOMSTATE_INDETERMINATE As Short = 0
	Public Const SPV_ZOOMSTATE_IN As Short = 1
	Public Const SPV_ZOOMSTATE_OUT As Short = 2
	Public Const SPV_ZOOMSTATE_SWITCH As Short = 3
	
	
	Public gintZoom As Short
	
	Public Sub ProtegeSpread(ByRef Spread As Object)
		'*************************************************************************
		'   Procedimiento que protege todo al spread de ser alterado en sus celdas
		'*************************************************************************
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Row. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.Row = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.col. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.col = -1
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Lock. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.Lock = True
		'UPGRADE_WARNING: Couldn't resolve default property of object Spread.Protect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Spread.Protect = True
	End Sub
End Module