'--------------------------------------------------
'■履歴
'◇ ver 2014/11/03
'・ 作成
'・ 文字列処理First/Last/Delimiter
'・ グラフ処理
'・ DataLastRow/Col
'・ ArrayCount
'・ Assert/Check/OrValue
'・ IncludeLastPathDelim
'・ IniFile_GetString/SetString
'・ GetAbsolutePath
'・ MaxValue/MinValue
'・ LongToStrDigitZero
'・ PixelToPoint/PointToPixel
'・ ADOStream
'◇ ver 2014/11/06
'・ CommandExecuteReturn
'・ IncludeBothEndsStr/ExcludeBothEndsStr
'・ GetFirstStr---/GetLastStr---
'・ TrimLast/TrimFirst
'・ IsLong
'◇ ver 2014/11/07
'・ ClearLast
'・ CommandExecuteReturn
'◇ ver 2014/11/08
'・ ChartObjectExists/ShapeExists
'◇ ver 2014/11/19
'・ ExcludeLastPathDelim追加
'・ UBound/LBound
'・ ArrayStr/StringArrayCombine/StringCombine/PathCombine
'・ GetExtensionIncludePeriod/ChangeFileExtension
'・ Get/SetWindowLong
'・ SetWindowStyle/SetWindowExStyle/SetWindowTopMost
'◇ ver 2014/11/20
'・ GetAsyncKeyState
'・ BooleanToString
'・ FormatYYYY_MM_DD/FormatHH_MM_SS
'・ GetFolderPathListTopFolder
'・ ClearLineColumn
'・ SetTaskbarButtonAppID
'◇ ver 2014/11/21
'・ SetIcon/ResetIcon
'◇ ver 2014/11/24
'・ BooleanToString>>BoolToStr
'・ RectToStr/StrToRect
'・ NewRect/NewRectSize/NewPoint/NewRect_PositionSize
'   /GetRectSize/RectEqual
'   /GetRectInsideDesktopRect
'・ PopupMenu
'・ Form_GetRectPixel/Form_SetRectPixel
'・ GetDesktopWindow/GetWindowRect/SystemParametersInfo
'   GetRectDesktop/GetRectWorkArea
'・ GetSpecialFolderPath
'・ Form_IniWritePosition/Form_IniReadPosition
'・ TaskDialog
'◇ ver 2014/11/26
'・ IsWindowsOffice64/32bit
'   WindowsMajor/MinorVersion
'   IsTaskbarPinWindows
'・ ForceCreateFolder
'・ CreateShortcutFile
'・ GetWindowState
'・ GetRectInsideDesktopRect修正
'◇ ver 2014/12/01
'・ TaskDialog系の修正
'・ SetWindowIcon/ResetWindowIcon
'・ GetBitmapDrawIcon/Image_Picture_SetBitmap
'・ GetDC/FillRect/DrawIcon
'   /CreateCompatibleDC/CreateCompatibleBitmap
'   /SelectObject/DeleteObject/GetStockObject
'・ GetWindowCloseButton/GetWindowStyle/GetWindowExStyle
'   /GetWindowIcon
'◇ ver 2014/12/02
'・ MouseMove/MouseClick
'◇ ver 2014/12/04
'・ SetShortcutIcon/SetTaskbarPinShortcutIcon/SetTaskbarPin
'◇ ver 2014/12/06
'・ StrToLongDefault
'・ ArrayAdd
'・ ApplicationMode/SetExcelWindowTitle
'◇ ver 2015/02/02
'・ Microsoft Forms 2.0 Object Libraryの参照設定追加
'・ FirstStrFirstDelim/FirstStrLastDelim
'   /LastStrFirstDelim/LastStrLastDelim
'◇ ver 2015/02/06
'・ ReCreateFolder作成
'◇ ver 2015/02/13
'・ DataLastCol修正
'   DataLastCell作成
'◇ ver 2015/03/05
'・ 参照設定ReferenceAdd系処理追加
'・ 配列関連処理追加
'   ArrayInsert/ArrayDelete
'   /ArrayIndexOf/ArrayDeleteSameItem
'   /ArrayDimension/ArrayToString
'・ ListView関連処理追加
'   ListView_SelectedItemCount/ListView_CheckedItemCount
'   /ListView_SelectAll/ListView_CheckSelectedItem
'   /ListView_IsCheckSelectedItem/ListView_MultiSelectChecked
'   /ListView_IndexOfKey
'・ ファイル日時関連処理追加
'   DateToApiFILETIME/GetFileFolderTime/SetFileFolderTime
'・ FormatDateTimeNormal追加
'・ ファイルフォルダ一覧処理追加
'   FolderPathListTopFolder/FolderPathListSubFolder
'   /FilePathListTopFolder/FilePathListSubFolder
'・ ComboBox関連処理追加
'   ComboBox_GetStrings/ComboBox_SetStrings
'   /Combobox_ClearList
'・ 名前変更 GetAbsolutePath>>AbsolutePath
'・ StringCombine/StringCombineArray
'   /PathCombine修正
'◇ ver 2015/03/11
'・ ArraySetValueObjectを追加
'   ArrayAdd/ArrayInsertを修正
'◇ ver 2015/03/19
'・ ArrayAdd/ArrayInsert/ArrayDeleteを修正
'・ コメントの修正
'◇ ver 2015/07/23
'・ StarndardSoftwareLibraryからst_vbaに名前変更
'◇ ver 2015/07/29
'・ 64bit版Excelへの暫定対応(既存は32bit版Excelのみの対応)
'   TaskDialogAPIを削除
'・ GetDPIの正しい実装を行った。
'◇ ver 2015/08/07
'・ FileExists(Win/Mac両対応版)を追加
'・ GetClipboardText/SetClipboardText(Win/Mac両対応版)を追加
'◇ ver 2015/08/23
'・ CommandExecuteを追加
'・ PopupMenu_PopupReturn_NoPositionを追加
'・ IsShortcutLinkFile追加
'・ IsJpegImageFile/IsJpegExifFile追加
'・ GetJpegExifDateTime追加
'◇ ver 2015/12/12
'・ Excel64bit定数追加
'・ SleepAPI追加
'・ IE_NewObject/IE_GetObject/IE_Navigate
'   /IE_NavigateWait/IE_RunJavaScript追加
'・ IsIncludeStr追加
'◇ ver 2015/12/16
'・ ClearLastRange/ClearLastColumn/ClearLastRow
'   /ClearLastRangeContents
'   /ClearLastColumnContents/ClearLastRowContentsを修正追加
'◇ ver 2016/01/08
'・ ClearLastRange/ClearLastColumn/ClearLastRowを修正
'   ClearContents機能を追加
'・ TrimFirstChar/TrimLastChar/TrimBothEndsCharを廃止
'   TrimFirstStrs/TrimLastStrs/TrimBothEndsStrs
'   /TrimFirstSpace/TrimLastSpace/TrimBothEndsSpaceを追加
'・ DataLastRow/DataLastColがデータがないときにエラー発生するので
'   OnErrorResumeするように修正
'◇ ver 2016/02/06
'・ Enum AlineHorizontal/AlineVertical の定義
'・ URLDownloadToFile APIとURLDownloadFileの追加
'・ 日付時刻書式指定関数の追加
'   FormatYYYYMMDD/FormatYYYY_MM
'   /FormatHHMMSS/FormatHH_MM
'   /FormatYYYYMMDDHHMMSS/FormatYYYYMMDDHHMMSS_Hyphen
'・ クリア形処理の名前変更
'   ClearLastRange→ClearRangeLast
'   ClearLastColumn→ClearColumnLast
'   ClearLastRow→ClearRowLast
'・ Shape処理の追加
'   GetShapeFromImageFile/ShapeCompressUseClipboard
'・ IE処理の修正 IE_NewObject/IE_Refresh
'   /IE_Navigate/IE_NavigateWait
'◇ ver 2016/02/20
'・ GetWorkbook追加
'・ GetWorksheet/WorksheetExists追加
'・ DeleteSheet/DeleteDefaultSheet追加
'・ SetTextSheet追加
'・ TagInnerText/TagOuterText追加
'・ IfEmptyStr追加
'・ セルクリア系処理の名前変更
'   ClearRangeLast→ClearRangeLastData
'   ClearColumnLast→ClearColumnLastRow
'   ClearRowLast→ClearRowLastColumn
'・ URLDownloadFileの戻り値をBooleanに変更
'◇ ver 2016/02/21
'・ IsNothing/IsNotNothing追加
'・ CastExcludeComma追加
'・ IE_GetElementByTagNameClassName/IE_GetElementByTagNameInnerHTMLの追加
'・ FormulaDeleteRange追加
'・ ColumnNumberByTitle追加
'・ ColumnNumber追加
'・ CopyFile追加
'◇ ver 2016/02/23
'・ ThisWeekDay/LastWeekDay/NextWeekDay追加
'◇ ver 2016/02/24
'・ IsDrivePath/IsNetworkPath追加
'・ SettingFullPath追加
'   AbsolutePath修正
'◇ ver 2016/02/28
'・ ThisWeekDay/LastWeekDay/NextWeekDay修正
'・ ColumnNumberByTitle修正
'・ RangeClear機能追加MergeCellOption対応
'・ RangeCopyNumberFormat/RangeCopyFormat/RangeCopyAll追加
'・ FormulaDeleteRange→RangeDeleteFormula名前変更
'・ FirstStrFirstDelim/FirstStrLastDelim
'   /LastStrFirstDelim/LastStrLastDelim の修正
'・ DeleteSheetの修正
'・ SetTextSheetの修正
'・ IE_GetElementByTagNameId追加
'◇ ver 2016/02/29
'・ ClearRangeLastData/ClearColumnLastRow/ClearRowLastColumn修正
'◇ ver 2016/03/04
'・ TagOuterTextの修正
'・ TagOuterTextList追加
'・ ReplaceHTMLTag追加
'◇ ver 2016/03/10
'・ Wingdings_Checkbox_Checked/UnChecked追加
'・ urlEncode追加
'・ ArrayAddNotDuplicate/ArrayExists追加
'・ ArraySortQuick追加
'・ RangeUpRow/RangeDownRow追加
'・ RangeMoveUpRowOne/RangeMoveDownRowOne追加
'・ LengthSjisByte
'   /LeftSjisByte/RightSjisByte
'   /MidSjisByte追加
'◇ ver 2016/03/13
'・ urlEncode修正
'・ TopLeftCell追加
'・ StrCount追加
'・ StrToBool追加
'・ st_vba_Baseから、st_vba_Coreに名称変更
'・ ListView処理を、st_vba_ListViewに移行
'・ InternetExplorer処理を、st_vba_IEに移行
'◇ ver 2016/03/20
'・ IE_GetElementByTagNameを追加
'・ ReplaceContinuousSpace追加
'・ RangeCopyValue追加
'・ MatchRegExp追加
'・ ArrayIndexOfに完全一致/部分一致/ワイルドカード/正規表現
'   の機能を追加。ArrayExistsも追加。
'◇ ver 2016/03/23
'・ ArrayIndexOfを改良して
'   ワイルドカード配列/正規表現配列の機能を追加
'・ ReplaceArrayValue/DeleteArrayValueを追加
'・ ArraySortOrderを追加
'◇ ver 2016/03/26
'・ ArraySortOrderを修正
'   ArraySortCustomOrderに名称変更
'・ ReplaceRegExpを追加
'・ ReplaceArrayRegExpを追加
'・ DeleteArrayRegExpを追加
'・ ArraySortQuickにSortOrder機能追加
'・ ArraySortStrLength追加
'・ ArrayReverse追加
'・ ShapeCompressUseClipboard修正
'・ RowNumberByTitle追加
'◇ ver 2016/03/27
'・ ArrayIsUnique追加
'・ 2次元配列系の処理を追加
'   Array2dSetColumn
'   /Array2dSetRowValues/Array2dGetRowValues
'   /Array2dAdd/Array2dInsert/Array2dDelete
'   /Array2dSortQuick/Array2dIsUnique
'◇ ver 2016/03/28
'・ Array2dAddを修正
'◇ ver 2016/03/29
'・ DeleteRegExp追加
'・ ReplaceHTMLTag>>DeleteHTMLTag名前変更と修正
'・ st_vba_IE.IE_GetElementの処理を修正
'   引数をieからElement=ie.Documentに変更
'   IE_GetElementByTagNameName追加
'◇ ver 2016/03/30
'・ Array2dSetRowValues/Array2dGetRowValues 追加
'・ Array2dRowsCount/Array2dColumnsCount 追加
'◇ ver 2016/03/31
'・ Array2dColumnsCount/Array2dRowsCount 追加
'・ Array2dColumnsCount/Array2dRowsCount 追加
'・ Array2dSetColumnValues/Array2dGetColumnValues 追加
'・ Array2dSortStrLength/Array2dSortStrLengthSetKeyValue 追加
'・ Array2dSortCustomOrder/Array2dSortCustomOrderSetKeyValue 追加
'・ ArraySort系処理のAssertとメッセージ修正
'◇ ver 2016/04/02
'・ Array2dSort系の処理修正
'◇ ver 2017/02/05
'・ FileCreateWaitをFileExistWaitに変更し
'   ファイルの存在の有無を待つように機能追加
'◇ ver 2017/02/11
'・ Array2dRowsStartIndex/Array2dRowsEndIndex
'   Array2dColumnsStartIndex/Array2dColumnsEndIndex 追加
'・ Array2dRowsCount/Array2dColumnsCount 修正
'・ Array2dSortCustomOrder
'   Array2dSetRowValues/Array2dGetRowValues
'   Array2dSetColumnValues/Array2dGetColumnValues
'   1Originの配列(Indexの最小値が1の配列)に対応
'◇ ver 2017/02/13
'・ Application_StatusBar_Progress 修正
'   ProgressText 追加
'・ DeleteSheets 追加
'・ SheetRangeSortCustomOrder 追加
'◇ ver 2017/02/17
'・ WorksheetExistsのBook指定の不具合修正
'◇ ver 2017/02/20
'・ CellValueIncrement 追加
'・ WorkbookFullPath 追加
'◇ ver 2017/02/26
'・ IPアドレスを処理するために
'   InRangeCurrency/IPAddressToCurrency
'   /InRangeIPAddress/IsIPAddress 追加
'・ ProgressText 修正
'・ st_vba_WaitForm の組み込み
'・ Standard Software URL Facebookページに変更
'◇ ver 2017/03/06
'・ SheetRangeSortCustomOrder に WorksheetFunction.Transpose の
'   限界値がある不具合があり、Transpose関数と同じものを
'   自作の Array2dTranspose 関数に置き換えた
'◇ ver 2017/03/09
'・ テキストファイル読み書き関数のエンコードのEnumでの指定版を作成
'   GetEncodingTypeJpCharCode/GetEncodingTypeName
'   /String_LoadFromFile/String_SaveToFile 追加
'◇ ver 2017/03/12
'・ TagInnerText 修正
'◇ ver 2017/03/14
'・ 参照設定追加コードを st_vba_SetReference に分離
'◇ ver 2017/03/19
'・ st_vba_SetReference の処理順序入れ替え
'・ 各モジュールやクラスの説明がないものは先頭に説明を記載
'・ st_vba_CSheetData_Sample を追加
'◇ ver 2017/03/21
'・ AbsolutePath のテストを追加
'・ ForceCreateFolderを修正
'・ 全体的に関数名をリファクタリング
'   App_ / Book_ / Sheet_ / Range_ を関数先頭に追加
'・ Sheet_ColumnNumberByTitle / Sheet_RowNumberByTitle を
'   ワイルドカード対応
'・ App_GetOpenedBookOrOpenBook 追加
'・ Folder_DeleteIfNoFile / Folder_DeleteIfNoFileToUpFolder 追加
'・ Format_Date_UseOnlyYMDHNS 追加
'◇ ver 2017/03/23
'・ RandomValue を追加
'◇ ver 2017/03/25
'・ IsDrivePath / IsNetworkPath の不具合を修正
'・ AbsolutePathのネットワークパス対応
'   テストの確立
'・ Book_SaveAs の追加
'・ FileDialog_FilePicker / FileDialog_Open
'   / FileDialog_SaveAs / FileDialog_FolderPicker 追加
'◇ ver 2017/04/01
'・ st_vba_CSetting 追加
'・ AbsolutePathを修正
'・ String_GetOutShiftJIS
'   /String_GetMachineDependentCharacter 追加
'◇ ver 2017/04/02
'・ TagInnerTextLast 追加
'・ Folder_HasSubItem / Folder_DeleteSubItem 追加
'・ Folder_DeleteIfNoFile 修正
'◇ ver 2017/04/03
'・ ChangeFileExtension 修正
'◇ ver 2017/04/05
'・ st_vba_SetReference ReferenceAdd_VBAExtensibility を
'   64bit版Windowsでも動くように対応した
'・ BookFullPath の修正
'◇ ver 2017/04/06
'・ FilePath_IsIncludeFileNameOutString / FilePath_ReplaceFileNameOutString 追加
'◇ ver 2017/04/16
'・ SetArrayCount / ArrayAddArray 追加
'・ Application_StatusBar_Progress / ProgressText 修正
'・ String_DeleteSpaceLine / String_LineTrim
'   / String_TagDelete / String_HTMLtoText 追加
'・ Sheet_RowNumberByTitle 追加
'・ Range_DeleteShape 追加
'・ Sheet_CheckBoxColumn 追加
'◇ ver 2017/05/04
'・ Col__A→Col_A 等、修正
'◇ ver 2017/06/11
'・ keybd_event / GetKeyboardState API 追加
'・ NumLockOn 追加
'◇ ver 2017/07/01
'・ Long最大値の定義
'・ IsLong不具合修正
'◇ ver 2017/07/02
'・ ADODB.StreamがExcel2016 64bit環境で不具合があり
'   環境依存かもしれないが問題を解決できなかったためにコードを分離
'・ ShiftJISのみ対応の String_LoadFromFile/String_SaveToFile 追加
'・ CommandExecuteReturnからADOStream_LoadTextFile削除
'・ st_vba_WaitForm.Update_ProgressInfo を .Updateに処理分離
'・ Book_SaveAsの対応をxlsのみから、xlsx/xlsmを追加した
'◇ ver 2017/07/03
'・ ファイルフォルダ処理の分類を整頓
'・ CopyFileに加えて、MoveFile/CopyFolder/MoveFolderを追加
'◇ ver 2017/08/08
'・ Sheet_OpenCSV/Sheet_SaveCSV を追加
'◇ ver 2017/08/14
'・ GetShapeFromImageFileでExifの回転画像対応
'・ IE_Navigate 修正
'   IE_Navigate_AuthBasic 作成
'   IE_Navigate_AuthBasicInput 作成
'・ IE_GetElementByTagNameClassName に除外条件指定可能にした
'◇ ver 2017/09/19
'・ GetJpegExifRotate 作成
'・ ImageSize 作成
'・ Sheet_CellRange 作成
'・ IE_GetElementByTagNameSearch 作成
'・ NowMilliSec 作成
'・ Format_Date_UseOnlyYMDHNS を FormatOnlyYMDHNS に名前変更
'・ GetShapeFromImageFile の内部を小数点以下サイズに対応
'   Exif での回転画像に対応
'・ Range_DeleteShape 軽微な修正
'◇ ver 2017/11/06
'・ ColorToStr/StrToColor 作成
'・ MonthLastDay/MonthDayCount の引数修正
'・ App_/Book_/Sheet_などの名前を修正した
'◇ ver 2017/11/13
'・ Sheet_などの名前をさらに修正した
'◇ ver 2017/12/03
'・ StrToBoolDef 追加
'・ CellText 追加
'・ IsReadOnlyFile/IsUseFile 追加
'・ NewSheetNameNumbering 追加
'・ SheetCopyBookAdd/SheetTextCopyBookAdd 追加
'・ GetFirstSheet/GetLastSheet 追加
'--------------------------------------------------

