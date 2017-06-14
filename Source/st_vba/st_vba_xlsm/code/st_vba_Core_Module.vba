'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    Core Module
'ObjectName:    st_vba_Core
'--------------------------------------------------
'Discription:   Standard Software Library For Windows Excel VBA
'--------------------------------------------------
'OpenSource:    https://github.com/standard-software/st_vba/
'License:       MIT License
'   URL:        https://github.com/standard-software/st_vba/blob/master/Document/Readme_jp.txt
'All Right Reserved:
'   Name:       Standard Software
'   URL:        https://www.facebook.com/stndardsoftware/
'--------------------------------------------------
'Version:       2017/06/11
'--------------------------------------------------

'--------------------------------------------------
'■マーク
'--------------------------------------------------

    '--------------------------------------------------
    '■
    '--------------------------------------------------

    '----------------------------------------
    '◆
    '----------------------------------------
    '◇
    '----------------------------------------
    '・
    '----------------------------------------

'--------------------------------------------------
'■参照設定
'・ Microsoft Scripting Runtime
'       FileSystemObject
'・ Windows Script Host Object Model
'       WshShell
'・ Microsoft AxtiveX Data Objects 6.1 Library
'       ADODB.Stream
'・ Microsoft Forms 2.0 Object Library
'       Image
'       ComboBox
'・ Microsoft Internet Controls
'       InternetExplorer
'・ Microsoft Windows Common Controls 6.0 (SP6)
'       ListView
'       32bit Windows / 32bit Excel
'           C:\Windows\system32\MSCOMCTL.OCX
'       64bit Windows / 32bit Excel
'           C:\Windows\SysWOW64\mscomctl.ocx
'       64bit Windows / 64bit Excel
'           使用不可
'--------------------------------------------------
'・ Microsoft Windows Common Controls 6.0 (SP6)
'       64bit Windows / 32bit Excel
'           C:\Windows\SysWOW64\mscomctl.ocx
'   ・  http://www.microsoft.com/ja-jp/download/details.aspx?id=10019
'       Download
'           VisualBasic6-KB896559-v1-JPN.exe
'       Unzip
'           mscomctl.ocx / comctl32.ocx / etc..
'       FileMove
'           C:\Windows\SysWOW64\mscomctl.ocx
'           C:\Windows\SysWOW64\comctl32.ocx (?)
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'■定数/型宣言
'--------------------------------------------------

'----------------------------------------
'◆位置・サイズ
'----------------------------------------
Public Type Point
    X As Long
    Y As Long
End Type

Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RectSize
    Width As Long
    Height As Long
End Type

Public Enum AlineHorizontal
    alLeft
    alCenter
    alRight
End Enum

Public Enum AlineVertical
    alTop
    alCenter
    alBottom
End Enum

'----------------------------------------
'◆FileSystemObject
'----------------------------------------
Public fso As New FileSystemObject

'----------------------------------------
'◆Shell
'----------------------------------------
Public Shell As New WshShell

'----------------------------------------
'◆文字列比較
'----------------------------------------
Public Enum MatchType
    FullMatch = 0   '完全一致
    PartMatch = 1   '部分一致
    WildCardValue = 2
    WildCardArray = 3
    RegExpValue = 4
    RegExpArray = 5
End Enum

Public Enum CaseCompare
    CaseSensitive
    IgnoreCase
End Enum

Public Enum StrAddType
    FirstAdd
    LastAdd
End Enum

'----------------------------------------
'◆配列
'----------------------------------------
Public Enum SortOrder
    Ascending
    Descending
End Enum

'----------------------------------------
'◆Excel
'----------------------------------------

'----------------------------------------
'◇列指定
'----------------------------------------

Public Const Col_A = 1, Col_B = 2, Col_C = 3, Col_D = 4, Col_E = 5, Col_F = 6
Public Const Col_G = 7, Col_H = 8, Col_I = 9, Col_J = 10, Col_K = 11, Col_L = 12
Public Const Col_M = 13, Col_N = 14, Col_O = 15, Col_P = 16, Col_Q = 17, Col_R = 18
Public Const Col_S = 19, Col_T = 20, Col_U = 21, Col_V = 22, Col_W = 23, Col_X = 24
Public Const Col_Y = 25, Col_Z = 26
Public Const Col_AA = 27, Col_AB = 28, Col_AC = 29, Col_AD = 30, Col_AE = 31, Col_AF = 32
Public Const Col_AG = 33, Col_AH = 34, Col_AI = 35, Col_AJ = 36, Col_AK = 37, Col_AL = 38
Public Const Col_AM = 39, Col_AN = 40, Col_AO = 41, Col_AP = 42, Col_AQ = 43, Col_AR = 44
Public Const Col_AS = 45, Col_AT = 46, Col_AU = 47, Col_AV = 48, Col_AW = 49, Col_AX = 50
Public Const Col_AY = 51, Col_AZ = 52
Public Const Col_CA = 53, Col_CB = 54, Col_CC = 55, Col_CD = 56, Col_CE = 57, Col_CF = 58
Public Const Col_CG = 59, Col_CH = 60, Col_CI = 61, Col_CJ = 62, Col_CK = 63, Col_CL = 64
Public Const Col_CM = 65, Col_CN = 66, Col_CO = 67, Col_CP = 68, Col_CQ = 69, Col_CR = 70
Public Const Col_CS = 71, Col_CT = 72, Col_CU = 73, Col_CV = 74, Col_CW = 75, Col_CX = 76
Public Const Col_CY = 77, Col_CZ = 78
Public Const Col_DA = 105, Col_DB = 106, Col_DC = 107, Col_DD = 108, Col_DE = 109, Col_DF = 110
Public Const Col_DG = 111, Col_DH = 112, Col_DI = 113, Col_DJ = 114, Col_DK = 115, Col_DL = 116
Public Const Col_DM = 117, Col_DN = 118, Col_DO = 119, Col_DP = 120, Col_DQ = 121, Col_DR = 122
Public Const Col_DS = 123, Col_DT = 124, Col_DU = 125, Col_DV = 126, Col_DW = 127, Col_DX = 128
Public Const Col_DY = 129, Col_DZ = 130
Public Const Col_EA = 131, Col_EB = 132, Col_EC = 133, Col_ED = 134, Col_EE = 135, Col_EF = 136
Public Const Col_EG = 137, Col_EH = 138, Col_EI = 139, Col_EJ = 140, Col_EK = 141, Col_EL = 142
Public Const Col_EM = 143, Col_EN = 144, Col_EO = 145, Col_EP = 146, Col_EQ = 147, Col_ER = 148
Public Const Col_ES = 149, Col_ET = 150, Col_EU = 151, Col_EV = 152, Col_EW = 153, Col_EX = 154
Public Const Col_EY = 155, Col_EZ = 156
Public Const Col_FA = 157, Col_FB = 158, Col_FC = 159, Col_FD = 160, Col_FE = 161, Col_FF = 162
Public Const Col_FG = 163, Col_FH = 164, Col_FI = 165, Col_FJ = 166, Col_FK = 167, Col_FL = 168
Public Const Col_FM = 169, Col_FN = 170, Col_FO = 171, Col_FP = 172, Col_FQ = 173, Col_FR = 174
Public Const Col_FS = 175, Col_FT = 176, Col_FU = 177, Col_FV = 178, Col_FW = 179, Col_FX = 180
Public Const Col_FY = 181, Col_FZ = 182
Public Const Col_GA = 183, Col_GB = 184, Col_GC = 185, Col_GD = 186, Col_GE = 187, Col_GF = 188
Public Const Col_GG = 189, Col_GH = 190, Col_GI = 191, Col_GJ = 192, Col_GK = 193, Col_GL = 194
Public Const Col_GM = 195, Col_GN = 196, Col_GO = 197, Col_GP = 198, Col_GQ = 199, Col_GR = 200
Public Const Col_GS = 201, Col_GT = 202, Col_GU = 203, Col_GV = 204, Col_GW = 205, Col_GX = 206
Public Const Col_GY = 207, Col_GZ = 208
Public Const Col_HA = 209, Col_HB = 210, Col_HC = 211, Col_HD = 212, Col_HE = 213, Col_HF = 214
Public Const Col_HG = 215, Col_HH = 216, Col_HI = 217, Col_HJ = 218, Col_HK = 219, Col_HL = 220
Public Const Col_HM = 221, Col_HN = 222, Col_HO = 223, Col_HP = 224, Col_HQ = 225, Col_HR = 226
Public Const Col_HS = 227, Col_HT = 228, Col_HU = 229, Col_HV = 230, Col_HW = 231, Col_HX = 232
Public Const Col_HY = 233, Col_HZ = 234
Public Const Col_IA = 235, Col_IB = 236, Col_IC = 237, Col_ID = 238, Col_IE = 239, Col_IF = 240
Public Const Col_IG = 241, Col_IH = 242, Col_II = 243, Col_IJ = 244, Col_IK = 245, Col_IL = 246
Public Const Col_IM = 247, Col_IN = 248, Col_IO = 249, Col_IP = 250, Col_IQ = 251, Col_IR = 252
Public Const Col_IS = 253, Col_IT = 254, Col_IU = 255, Col_IV = 256, Col_IW = 257, Col_IX = 258
Public Const Col_IY = 259, Col_IZ = 260

'----------------------------------------
'◇Cell削除処理
'----------------------------------------
'   ・  ClearComments/ClearOutlineは
'       特に用途がなさそうなので実装しなかった
'   ・  rcClear:            全てクリア
'       rcClearContents:    数式・文字列のクリア
'       rcClearFormats:     書式のクリア
'----------------------------------------
Enum RangeClearType
    rcClear
    rcClearContents
    rcClearFormats

End Enum

'----------------------------------------
'◆グラフ処理
'----------------------------------------
Public Type GraphFormulaData
    SeriesName As String
    ItemXAxisRangeStr As String
    DataRangeStr As String
    SeriesNumber As Long
End Type

'----------------------------------------
'◆メニュー処理
'----------------------------------------
Private PopupMenu_Return As String

'----------------------------------------
'◆ファイルフォルダパス取得
'----------------------------------------
Enum SpecialFolderType
    Desktop
    MyDocument
    StartMenu
    StartMenuProgram
    StartMenuStartup
    SendTo
    AppData
    AllUsersDesktop
    AllUsersStartMenu
    AllUsersStartMenuProgram
    AllUsersStartMenuStartup
    TaskbarPin
    Windows
    System
    Temporary
End Enum

'----------------------------------------
'◆システム
'----------------------------------------
#If VBA7 And Win64 Then
    Const Excel64bit As Boolean = True
#Else
    Const Excel64bit As Boolean = False
#End If


'----------------------------------------
'◆テキストファイルエンコード
'----------------------------------------
Public Enum EncodingTypeJpCharCode
    NONE
    ASCII
    JIS
    EUC_JP
    UTF_7
    Shift_JIS
    UTF8_BOM
    UTF8_BOM_NO
    UTF16_LE_BOM
    UTF16_LE_BOM_NO
    UTF16_BE_BOM
    UTF16_BE_BOM_NO
End Enum

'--------------------------------------------------
'■API
'--------------------------------------------------

'----------------------------------------
'◆ファイル日時
'----------------------------------------

'ファイルを作成またはオープン
Private Declare PtrSafe Function CreateFile Lib "kernel32.dll" _
    Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwdesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare PtrSafe Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long) As Long

'システムタイムをファイルタイムに変換する
Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32.dll" ( _
    ByRef lpSystemTime As SYSTEMTIME, _
    ByRef lpFileTime As FILETIME) As Long

'ローカルファイルタイムをUTCファイルタイム形式で取得する
Private Declare PtrSafe Function LocalFileTimeToFileTime Lib "kernel32.dll" ( _
    ByRef lpLocalFileTime As FILETIME, _
    ByRef lpFileTime As FILETIME) As Long

'ファイルのファイルタイムを設定する
Private Declare PtrSafe Function SetFileTime Lib "kernel32.dll" ( _
    ByVal hFile As Long, _
    ByRef lpCreationTime As FILETIME, _
    ByRef lpLastAccessTime As FILETIME, _
    ByRef lpLastWriteTime As FILETIME) As Long

'FILETIME 構造体
Private Type FILETIME
    dwLowDateTime As Long    '下位32ビット値
    dwHighDateTime As Long   '上位32ビット値
End Type

'SECURITY_ATTRIBUTES 構造体
Private Type SECURITY_ATTRIBUTES
    nLength              As LongPtr '構造体のバイト数
    lpSecurityDescriptor As LongPtr 'セキュリティデスクリプタ(Win95,98では無効)
    bInheritHandle       As LongPtr '1のとき属性を継承する
End Type

'CreateFileで使用する定数
Private Const FILE_FLAG_BACKUP_SEMANTICS As Long = &H2000000  'NT系OSのみ
Private Const GENERIC_READ               As Long = &H80000000
Private Const GENERIC_WRITE              As Long = &H40000000
Private Const FILE_SHARE_READ            As Long = &H1
Private Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
Private Const OPEN_EXISTING              As Long = 3
Private Const OPEN_ALWAYS                As Long = 4
Private Const INVALID_HANDLE_VALUE       As Long = &HFFFFFFFF

'SYSTEMTIME 構造体
Private Type SYSTEMTIME
    wYear         As Integer '年
    wMonth        As Integer '月
    wDayOfWeek    As Integer '曜日(日=0, 月=1 ...)
    wDay          As Integer '日
    wHour         As Integer '時
    wMinute       As Integer '分
    wSecond       As Integer '秒
    wMilliseconds As Integer 'ミリ秒
End Type

Public Type FileFolderTime
    CreataionTime As Date
    LastWriteTime As Date
    LastAccessTime As Date
End Type

'----------------------------------------
'◆Iniファイル
'----------------------------------------
Public Declare PtrSafe Function GetPrivateProfileString _
    Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpAppName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Declare PtrSafe Function WritePrivateProfileString _
    Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

'----------------------------------------
'◆キーボード入力
'----------------------------------------
Public Declare PtrSafe Function GetAsyncKeyState _
    Lib "User32.dll" (ByVal vKey As Long) As Long

Private Declare PtrSafe Sub keybd_event _
    Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long)

Private Declare PtrSafe Function GetKeyboardState _
    Lib "user32" ( _
    pbKeyState As Byte) As Long

Public Const VK_NUMLOCK = &H90   '「NumLock」キー
Public Const KEYEVENTF_EXTENDEDKEY = &H1 'キーを押す
Public Const KEYEVENTF_KEYUP = &H2   'キーを放す

'----------------------------------------
'◆マウス
'----------------------------------------
Public Declare PtrSafe Function GetCursorPos _
    Lib "user32" (lpPoint As Point) As Long

Public Declare PtrSafe Sub mouse_event _
    Lib "user32" ( _
    ByVal dwFlags As Long, _
    ByVal dx As Long, ByVal dy As Long, _
    ByVal cButtons As Long, _
    ByVal dwExtraInfo As Long)

Public Const MOUSE_MOVED = &H1              'マウスを移動する(相対座標)
Public Const MOUSEEVENTF_ABSOLUTE = &H8000& 'MOUSE_MOVED or で絶対座標を指定
Public Const MOUSEEVENTF_LEFTUP = &H4       '左ボタンUP
Public Const MOUSEEVENTF_LEFTDOWN = &H2     '左ボタンDown
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20  '中央ボタンDown
Public Const MOUSEEVENTF_MIDDLEUP = &H40    '中央ボタンUP
Public Const MOUSEEVENTF_RIGHTDOWN = &H8    '右ボタンDown
Public Const MOUSEEVENTF_RIGHTUP = &H10     '右ボタンUP

'----------------------------------------
'◇マウスボタンイベント
'----------------------------------------
Enum MouseButton
    fmButtonLeft = 1       'レフトボタンクリック
    fmButtonRight = 2      'ライトボタンクリック
    fmButtonLeftRight = 3  'レフト+ライトボタンを同時クリック
    fmButtonMiddle = 4     '中ボタンクリック
End Enum


'----------------------------------------
'◆Form
'----------------------------------------

'----------------------------------------
'◇Windowハンドル
'----------------------------------------
Public Declare PtrSafe Function WindowFromAccessibleObject _
    Lib "oleacc.dll" ( _
    ByVal IAcessible As Object, _
    ByRef hWnd As Long) As Long

'----------------------------------------
'◇Windowスタイル
'----------------------------------------
Public Const GWL_HINSTANCE = (-6) 'インスタンスハンドルを取得
Public Const GWL_HWNDPARENT = (-8) '親ウインドウのハンドルを取得
Public Const GWL_ID = (-12) 'ウインドウのIDを取得
Public Const GWL_EXSTYLE = (-20) '拡張ウインドウスタイルを取得
Public Const GWL_STYLE = (-16) 'ウインドウスタイルを取得
Public Const GWL_WNDPROC = (-16) 'ウインドウ関数のアドレスを取得
Public Const GWL_USERDATA = (-21) 'ユーザー定義の32ビット値を取得

Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000

Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_TOPMOST = &H8

Public Declare PtrSafe Function GetWindowLong _
    Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare PtrSafe Function SetWindowLong _
    Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

'----------------------------------------
'◇SystemMenu/Closeボタン
'----------------------------------------
Public Declare PtrSafe Function GetSystemMenu _
    Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Declare PtrSafe Function GetMenuItemID _
    Lib "user32" (ByVal hMenu As Long, _
    ByVal nPos As Long) As Long

Public Declare PtrSafe Function DeleteMenu _
    Lib "user32" (ByVal hMenu As Long, _
    ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Declare PtrSafe Function EnableMenuItem _
    Lib "user32" (ByVal hMenu As Long, _
    ByVal uItem As Long, ByVal fuFlags As Long) As Long

Public Declare PtrSafe Function DrawMenuBar _
    Lib "user32" (ByVal hWnd As Long) As Long

Public Const SC_CLOSE As Long = &HF060
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_ENABLED As Long = &H0&
Public Const MF_GRAYED As Long = &H1&
Public Const MF_DISABLED As Long = &H2&

'----------------------------------------
'◇FormIcon
'----------------------------------------
Public Declare PtrSafe Function ExtractIcon _
    Lib "shell32" Alias "ExtractIconA" ( _
    ByVal hInst As Long, _
    ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long) As Long

Public Declare PtrSafe Function DestroyIcon _
    Lib "user32" (ByVal hIcon As Long) As Long

Public Const WM_GETICON As Long = &H7F
Public Const WM_SETICON = &H80
Public Const ICON_SMALL = 0&
Public Const ICON_BIG = 1&

Public Declare PtrSafe Function SendMessage _
    Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

'----------------------------------------
'◇Window位置/TopMost
'----------------------------------------
Public Declare PtrSafe Function SetWindowPos _
    Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST = -&H2
Public Const SWP_NOSIZE As Long = &H1&
Public Const SWP_NOMOVE As Long = &H2&
Public Const SWP_SHOWWINDOW = &H40

'----------------------------------------
'◇Window位置/TopMost
'----------------------------------------
Public Declare PtrSafe Function GetWindowPlacement _
    Lib "user32" ( _
    ByVal hWnd As Long, _
    lpwndpl As WINDOWPLACEMENT) As Long

Public Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    showCmd As Long
    ptMinPosition As Point
    ptMaxPosition As Point
    rcNormalPosition As Rect
End Type

Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMINIMIZED As Long = 2
Public Const SW_SHOWMAXIMIZED  As Long = 3

'----------------------------------------
'◆システム
'----------------------------------------

'----------------------------------------
'・OSバージョン
'----------------------------------------
Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

'----------------------------------------
'・Sleep
'----------------------------------------
'#If VBA7 And Win64 Then
'    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
'#Else
'    Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
'#End If
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

'----------------------------------------
'◆タスクバーボタン登録
'----------------------------------------
Public Declare PtrSafe Function SetCurrentProcessExplicitAppUserModelID _
    Lib "shell32.dll" ( _
    ByVal lpString As LongPtr) As Long

'----------------------------------------
'◆Window
'----------------------------------------
Public Declare PtrSafe Function GetDesktopWindow _
    Lib "user32" () As Long

'----------------------------------------
'◆描画
'----------------------------------------

Private Const LOGPIXELSX As Long = &H58&
Private Const LOGPIXELSY As Long = &H5A&
Private Declare PtrSafe Function GetDeviceCaps _
    Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal nIndex As Long _
    ) As Long

Public Declare PtrSafe Function GetDC _
    Lib "user32" (ByVal hHnd As Long) As Long

Private Declare PtrSafe Sub ReleaseDC _
    Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hDC As Long _
    )

Public Declare PtrSafe Function FillRect _
    Lib "user32" ( _
    ByVal hDC As Long, _
    ByRef lpRect As Rect, _
    ByVal hBrush As Long) As Long

Public Declare PtrSafe Function DrawIcon _
    Lib "user32" ( _
    ByVal hDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal hIcon As Long) As Long

Public Declare PtrSafe Function SelectObject _
    Lib "gdi32" ( _
    ByVal hDC As Long, ByVal hGdiObject As Long) As Long

Public Declare PtrSafe Function CreateCompatibleDC _
    Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare PtrSafe Function CreateCompatibleBitmap _
    Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal pWidth As Long, _
    ByVal pHeight As Long) As Long

Public Declare PtrSafe Function DeleteObject _
    Lib "gdi32" ( _
    ByVal hObj As Long) As Long

Public Declare PtrSafe Function GetStockObject _
    Lib "gdi32" (pIx As Long) As Long


'----------------------------------------
'◆アイコン
'----------------------------------------
Public Type PictDesc
    cbSizeOfStruct As Long
    picType As Long
    hImage As Long
    Option1 As Long
    Option2 As Long
End Type

Public Type guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Declare PtrSafe Function OleCreatePictureIndirect _
    Lib "oleaut32.dll" ( _
    ByRef lpPictDesc As PictDesc, _
    ByRef RefIID As guid, _
    ByVal fPictureOwnsHandle As Long, _
    ByRef IPic As IPicture) As Long

Public Type IconFilePathIndex
    Path As String
    Index As Long
End Type

Public Enum ImageresDllIcon
    ID_ICON_WINDOWS_SHIELD = 1
    ID_ICON_SECURITY_SHIELD = 73
    ID_ICON_INFORMATION = 76
    ID_ICON_WARNING = 79
    ID_ICON_ERROR = 93
    ID_ICON_QUESTION = 94
    ID_ICON_SECURITY_QUESTION = 99
    ID_ICON_SECURITY_ERROR = 100
    ID_ICON_SECURITY_SUCCESS = 101
    ID_ICON_SECURITY_WARNING = 102
End Enum

'----------------------------------------
'◆Rect
'----------------------------------------
Public Declare PtrSafe Function GetWindowRect _
    Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long

Private Declare PtrSafe Function GetSystemMetrics _
    Lib "user32" (ByVal nIndex As Long) As Long
    Private Const SM_CXSCREEN As Long = 0
    Private Const SM_CYSCREEN As Long = 1

Public Declare PtrSafe Function SystemParametersInfo _
    Lib "user32" Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, ByVal uParam As Long, _
    ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_GETWORKAREA As Long = 48

'----------------------------------------
'◆インターネット
'----------------------------------------
Public Declare PtrSafe Function URLDownloadToFile _
    Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long


'--------------------------------------------------
'■実装
'--------------------------------------------------

'----------------------------------------
'◆条件判断
'----------------------------------------
Public Sub Assert(ByVal Value As Boolean, Optional ByVal Message As String)
    If Value = False Then
        Call Err.Raise(9999, , Message)
    End If
End Sub

Private Sub testAssert()
    Call Assert(False, "テスト")
End Sub

Public Function Check(ByVal A As Variant, ByVal B As Variant) As Boolean
    Check = (A = B)
    If Check = False Then
        Call MsgBox("A != B" + vbCrLf + _
            "A = " + CStr(A) + vbCrLf + _
            "B = " + CStr(B))
    End If
End Function

'----------------------------------------
'・OrValue
'----------------------------------------
Public Function OrValue(ByVal Value As Variant, ParamArray Values() As Variant) As Boolean
    OrValue = False
    Dim I As Long
    For I = LBound(Values) To UBound(Values)
        If Value = Values(I) Then
            OrValue = True
            Exit For
        End If
    Next
End Function

Private Sub testOrValue()
    Call Check(True, OrValue(10, 20, 30, 40, 10))
    Call Check(False, OrValue(50, 20, 30, 40, 10))
End Sub

'----------------------------------------
'・IsNothing/IsNotNothing
'----------------------------------------
Public Function IsNothing(ByRef Value As Object) As Boolean
    IsNothing = (Value Is Nothing)
End Function

Public Function IsNotNothing(ByRef Value As Object) As Boolean
    IsNotNothing = Not (Value Is Nothing)
End Function

'----------------------------------------
'・値が空文字の場合だけ別の値を返す関数
'----------------------------------------
Function IfEmptyStr(ByVal Value As String, ByVal EmptyStrCaseValue) As String
    Dim Result As String: Result = ""
    If Value = "" Then
        Result = EmptyStrCaseValue
    Else
        Result = Value
    End If
    IfEmptyStr = Result
End Function

'----------------------------------------
'◆型、型変換
'----------------------------------------

'----------------------------------------
'・変数に値やオブジェクトをセットする
'----------------------------------------
Public Sub SetValue(ByRef Variable, ByVal Value)
    If IsObject(Value) Then
        Set Variable = Value
    Else
        Variable = Value
    End If
End Sub

Private Sub testSetValue()
    Dim A As Long
    A = 1
    Call SetValue(A, 2)
    Call Check(2, A)

    Dim B As Object
    'Set B = fso
    Call SetValue(B, fso)
    Call Check("test.txt", B.GetFileName("C:\temp\test.txt"))
End Sub


'----------------------------------------
'◇Long
'----------------------------------------
Public Function IsLong(Value As String) As Boolean
    Dim Result As Boolean: Result = False
    If IsNumeric(Value) Then
        If CInt(Value) = CDbl(Value) Then
            Result = True
        End If
    End If
    IsLong = Result
End Function

Private Sub testIsLong()
    Call Check(True, IsLong("123"))
    Call Check(False, IsLong("12a"))
    Call Check(False, IsLong("123.4"))

End Sub

Public Function LongToStrDigitZero(ByVal Value As Long, ByVal Digit As Long) As String
    Dim Result As String: Result = ""
    If 0 <= Value Then
        Result = String(MaxValue(0, Digit - Len(CStr(Value))), "0") + CStr(Value)
    Else
        Result = "-" + String(Digit - Len(CStr(Abs(Value))), "0") + CStr(Abs(Value))
    End If
    LongToStrDigitZero = Result
End Function

Private Sub testLongToStrDigitZero()
    Call Check("003", LongToStrDigitZero(3, 3))
    Call Check("000", LongToStrDigitZero(0, 3))
    Call Check("1000", LongToStrDigitZero(1000, 3))
    Call Check("-050", LongToStrDigitZero(-50, 3))
End Sub

Public Function StrToLongDefault(ByVal S As String, ByVal Default As Long) As Long
On Error Resume Next
    Dim Result As Long
    Result = Default
    Result = CLng(S)
    StrToLongDefault = Result
End Function

Private Sub testStrToLongDefault()
    Call Check(123, StrToLongDefault("123", 0))
    Call Check(123, StrToLongDefault(" 123 ", 0))
    Call Check(0, StrToLongDefault(" A123 ", 0))
    Call Check(123, StrToLongDefault("BBB", 123))
End Sub

'----------------------------------------
'◇カンマ付き文字の変換
'----------------------------------------
Public Function CastExcludeComma(ByVal CommaNumber As String) As Double
    Dim Result As Double: Result = 0
    If CommaNumber <> "" Then
        Result = CDbl( _
        Replace(CommaNumber, ",", ""))
    End If
    CastExcludeComma = Result
End Function

Public Sub testCastExcludeComma()
    Call Check(1000, CastExcludeComma("1,000"))
    Call Check(1000000, CastExcludeComma("1,000,000"))
End Sub

'----------------------------------------
'◇Boolean
'----------------------------------------
Public Function BoolToStr(ByVal Value As Boolean) As String
    Dim Result As String: Result = ""
    If Value Then
        Result = "True"
    Else
        Result = "False"
    End If
    BoolToStr = Result
End Function

Function StrToBool(ByVal Value As String) As Boolean
    Dim Result As Boolean: Result = False
    Select Case UCase(Value)
        Case "TRUE"
            Result = True
    End Select
    StrToBool = Result
End Function

'----------------------------------------
'◇Point
'----------------------------------------
Public Function NewPoint( _
ByVal Left As Long, _
ByVal Top As Long) As Point
    Dim Result As Point
    Result.X = Left
    Result.Y = Top
    NewPoint = Result
End Function

Public Function PointEqual( _
ByRef A As Point, ByRef B As Point) As Boolean
    PointEqual = False
    If (A.X = B.X) _
    And (A.Y = B.Y) Then
        PointEqual = True
    End If
End Function

Public Function GetPointRectCenter(ByRef RectValue As Rect) As Point
    Dim Result As Point
    Result.X = _
        RectValue.Left + ((RectValue.Right - RectValue.Left) / 2)
    Result.Y = _
        RectValue.Top + ((RectValue.Bottom - RectValue.Top) / 2)
    GetPointRectCenter = Result
End Function

'----------------------------------------
'◇Rect
'----------------------------------------

'----------------------------------------
'・Rect文字列変換
'----------------------------------------
Public Function RectToStr(ByRef RectValue As Rect) As String
    Dim Result As String: Result = ""
    Result = _
        StringCombine(",", _
            CStr(RectValue.Left), _
            CStr(RectValue.Top), _
            CStr(RectValue.Right), _
            CStr(RectValue.Bottom))
    RectToStr = Result
End Function

Private Sub testRectToStr()
    Call Check("5,10,15,25", RectToStr(NewRect(5, 10, 15, 25)))
End Sub

Public Function StrToRect(ByVal S As String) As Rect
    Dim Result As Rect

    Dim Strs() As String
    Strs = Split(S, ",")
    Call Assert(ArrayCount(Strs) = 4, _
        "Error:StrToRect")

    Result.Left = CLng(Strs(0))
    Result.Top = CLng(Strs(1))
    Result.Right = CLng(Strs(2))
    Result.Bottom = CLng(Strs(3))
    StrToRect = Result
End Function

Private Sub testStrToRect()
    Call Check(True, RectEqual(NewRect(5, 10, 15, 25), StrToRect("5,10,15,25")))
End Sub

Public Function CanStrToRect(ByVal S As String) As Boolean
On Error GoTo Err:
    Call StrToRect(S)
    CanStrToRect = True
    Exit Function
Err:
    CanStrToRect = False
End Function

'----------------------------------------
'・Rect生成
'----------------------------------------
Public Function NewRect( _
ByVal Left As Long, _
ByVal Top As Long, _
ByVal Right As Long, _
ByVal Bottom As Long) As Rect
    Dim Result As Rect
    Result.Left = Left
    Result.Top = Top
    Result.Right = Right
    Result.Bottom = Bottom
    NewRect = Result
End Function

Public Function NewRect_PositionSize( _
ByRef Position As Point, _
ByRef RectSize As RectSize) As Rect
    Dim Result As Rect
    Result.Left = Position.X
    Result.Top = Position.Y
    Result.Right = Position.X + RectSize.Width
    Result.Bottom = Position.Y + RectSize.Height
    NewRect_PositionSize = Result
End Function

'----------------------------------------
'・Rect比較
'----------------------------------------
Public Function RectEqual( _
ByRef A As Rect, ByRef B As Rect) As Boolean
    RectEqual = False
    If (A.Left = B.Left) _
    And (A.Top = B.Top) _
    And (A.Right = B.Right) _
    And (A.Bottom = B.Bottom) Then
        RectEqual = True
    End If
End Function

'----------------------------------------
'・Rect Width/Height値取得
'----------------------------------------
Public Function GetRectWidth( _
ByRef r As Rect) As Long
    GetRectWidth = _
        GetRectSize(r).Width
End Function

Public Sub SetRectWidth( _
ByRef r As Rect, ByVal Width As Long)
    r.Right = r.Left + Width
End Sub

Public Function GetRectHeight( _
ByRef r As Rect) As Long
    GetRectHeight = _
        GetRectSize(r).Height
End Function

Public Sub SetRectHeight( _
ByRef r As Rect, ByVal Height As Long)
    r.Bottom = r.Top + Height
End Sub

'----------------------------------------
'◇Rect Get系
'----------------------------------------
Public Function GetRectMoveCenter( _
ByRef r As Rect, ByRef Center As Point) As Rect
    Dim OriginalCenter As Point
    OriginalCenter = GetPointRectCenter(r)
    Dim Move As Point
    Move.X = Center.X - OriginalCenter.X
    Move.Y = Center.Y - OriginalCenter.Y
    GetRectMoveCenter = GetRectMove(r, Move)
End Function

Public Function GetRectMove( _
ByRef r As Rect, ByRef Move As Point) As Rect
    Dim Result As Rect
    Result.Left = r.Left + Move.X
    Result.Top = r.Top + Move.Y
    Result.Right = r.Right + Move.X
    Result.Bottom = r.Bottom + Move.Y
    GetRectMove = Result
End Function

Public Function GetRectMovePosition( _
ByRef r As Rect, ByRef Position As Point) As Rect
    Dim Result As Rect
    Dim RectSize As RectSize
    RectSize = GetRectSize(r)
    Result = NewRect_PositionSize(Position, RectSize)
    GetRectMovePosition = Result
End Function

'はみ出していたら中にいれる
Public Function GetRectInsideDesktopRect( _
ByRef r As Rect, ByRef DesktopRect As Rect) As Rect
    Dim Result As Rect: Result = r
    Dim RectSizeDesktop As RectSize
    RectSizeDesktop = GetRectSize(DesktopRect)
    If RectSizeDesktop.Width < GetRectWidth(r) Then
        Call SetRectWidth(Result, RectSizeDesktop.Width)
    End If
    If RectSizeDesktop.Height <= GetRectHeight(r) Then
        Call SetRectHeight(Result, RectSizeDesktop.Height)
    End If

    If Result.Left < DesktopRect.Left Then
        Result = GetRectMove(Result, NewPoint(DesktopRect.Left - Result.Left, 0))
    End If
    If DesktopRect.Right < Result.Right Then
        Result = GetRectMove(Result, NewPoint(DesktopRect.Right - Result.Right, 0))
    End If
    If Result.Top < DesktopRect.Top Then
        Result = GetRectMove(Result, NewPoint(0, DesktopRect.Top - Result.Top))
    End If
    If DesktopRect.Bottom < Result.Bottom Then
        Result = GetRectMove(Result, NewPoint(0, DesktopRect.Bottom - Result.Bottom))
    End If


    GetRectInsideDesktopRect = Result
End Function

'----------------------------------------
'◇RectSize
'----------------------------------------
Public Function NewRectSize( _
ByVal Width As Long, _
ByVal Height As Long) As RectSize
    Dim Result As RectSize
    Result.Width = Width
    Result.Height = Height
    NewRectSize = Result
End Function

Public Function GetRectSize(ByRef RectValue As Rect) As RectSize
    Dim Result As RectSize
    Result.Width = Abs(RectValue.Right - RectValue.Left)
    Result.Height = Abs(RectValue.Bottom - RectValue.Top)
    GetRectSize = Result
End Function

'----------------------------------------
'◇Pixel Point 相互変換
'----------------------------------------

Public Function GetDPI() As Long
    Dim Result As Long: Result = 0

    Dim hWnd As Long
    Dim hDC As Long
    hWnd = Excel.Application.hWnd
    hDC = GetDC(hWnd)
    '水平方向DPI
    Result = GetDeviceCaps(hDC, LOGPIXELSX)
    '垂直方向DPI
    Result = GetDeviceCaps(hDC, LOGPIXELSY)
    Call ReleaseDC(hWnd, hDC)

    GetDPI = Result
 End Function

'96しか取得できない！
Public Function GetDPI1() As Long
    GetDPI1 = ActiveWorkbook.WebOptions.PixelsPerInch
End Function

'120しか取得できない！
Public Function GetDPI2() As Long
    Dim Result As Long: Result = 0
    Dim Locator As Object
    Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Dim Service As Object
    Set Service = Locator.ConnectServer
    Dim ClassItems As Object
    Set ClassItems = Service.ExecQuery("Select * From Win32_DisplayConfiguration")
    Dim ClassItem As Object
    For Each ClassItem In ClassItems
        Result = ClassItem.LogPixels
    Next
    GetDPI2 = Result
End Function

Private Sub testGetDPI()
    Call MsgBox(GetDPI)
End Sub


Public Function PointToPixel(ByVal PointValue As Double) As Double
    PointToPixel = PointValue * GetDPI / 72
End Function

Public Function PixelToPoint(ByVal PixelValue As Double) As Double
    PixelToPoint = PixelValue * 72 / GetDPI
End Function

Private Sub testPixelToPoint()
    Call Check(7.5, PixelToPoint(10))
    Call Check(24, PixelToPoint(32))
End Sub


'----------------------------------------
'◆数値処理
'----------------------------------------

'----------------------------------------
'・最大値最小値
'----------------------------------------
Public Function MaxValue(ParamArray Values() As Variant) As Variant
    MaxValue = Empty
    Dim Value As Variant
    For Each Value In Values
        If IsEmpty(MaxValue) Then
            MaxValue = Value
        ElseIf MaxValue < Value Then
            MaxValue = Value
        End If
    Next
End Function

Private Sub testMaxValue()
    Call Check(100, MaxValue(50, 20, 30, 100, 9))
End Sub

Public Function MinValue(ParamArray Values() As Variant) As Variant
    MinValue = Empty
    Dim Value As Variant
    For Each Value In Values
        If MinValue = Empty Then
            MinValue = Value
        ElseIf MinValue > Value Then
            MinValue = Value
        End If
    Next
End Function

Private Sub testMinValue()
    Call Check(9, MinValue(50, 20, 30, 100, 9))
End Sub

'----------------------------------------
'・値範囲
'----------------------------------------
Public Function InRange( _
ByVal MinValue As Long, _
ByVal Value As Long, _
ByVal MaxValue As Long) As Boolean

    InRange = ((MinValue <= Value) And (Value <= MaxValue))

End Function

Public Function InRangeCurrency( _
ByVal MinValue As Currency, _
ByVal Value As Currency, _
ByVal MaxValue As Currency) As Boolean

    InRangeCurrency = ((MinValue <= Value) And (Value <= MaxValue))
    
End Function

'----------------------------------------
'・乱数
'----------------------------------------
'   ・  指定範囲の乱数を生成する
'   ・  実行前に乱数列を初期化するには
'           Call Randomize
'       を行う
'----------------------------------------
Function RandomValue( _
ByVal MinValue As Long, ByVal MaxValue As Long) As Long

    RandomValue = Int((MaxValue - MinValue + 1) * Rnd + MinValue)

End Function

'----------------------------------------
'◆文字列処理
'----------------------------------------

'----------------------------------------
'・StrCount
'----------------------------------------
'   ・  文字列の数を数える関数
'       AAAからAAを数えると2を返す
'----------------------------------------
Public Function StrCount(str As String, SubStr As String) As Long
    Dim Result As Long
    Result = 0
    Dim Index As Long
    Index = 0
    Do
        Index = InStr(Index + 1, str, SubStr)
        If Index = 0 Then
            Exit Do
        Else
            Result = Result + 1
        End If
    Loop
    StrCount = Result
End Function

Sub testStrCount()
    Call Check(2, StrCount("AAA", "AA"))
End Sub

'----------------------------------------
'・IsInclude
'----------------------------------------
Public Function IsIncludeStr(ByVal str As String, ByVal SubStr As String)
    IsIncludeStr = _
        (1 <= InStr(str, SubStr))
End Function


'----------------------------------------
'◇First / Last
'----------------------------------------

'----------------------------------------
'・First
'----------------------------------------
Public Function IsFirstStr(ByVal str As String, ByVal SubStr As String) As Boolean
    Dim Result As Boolean: Result = False
    Do
        If SubStr = "" Then Exit Do
        If str = "" Then Exit Do
        If Len(str) < Len(SubStr) Then Exit Do

        If InStr(1, str, SubStr) = 1 Then
            Result = True
        End If
    Loop While False
    IsFirstStr = Result
End Function

Private Sub testIsFirstStr()
    Call Check(True, IsFirstStr("12345", "1"))
    Call Check(True, IsFirstStr("12345", "12"))
    Call Check(True, IsFirstStr("12345", "123"))
    Call Check(False, IsFirstStr("12345", "23"))
    Call Check(False, IsFirstStr("", "34"))
    Call Check(False, IsFirstStr("12345", ""))
    Call Check(False, IsFirstStr("123", "1234"))
End Sub

Public Function IncludeFirstStr(ByVal str As String, ByVal SubStr As String) As String
    If IsFirstStr(str, SubStr) Then
        IncludeFirstStr = str
    Else
        IncludeFirstStr = SubStr + str
    End If
End Function

Private Sub testIncludeFirstStr()
    Call Check("12345", IncludeFirstStr("12345", "1"))
    Call Check("12345", IncludeFirstStr("12345", "12"))
    Call Check("12345", IncludeFirstStr("12345", "123"))
    Call Check("2312345", IncludeFirstStr("12345", "23"))
End Sub

Public Function ExcludeFirstStr(ByVal str As String, ByVal SubStr As String) As String
    If IsFirstStr(str, SubStr) Then
        ExcludeFirstStr = Mid$(str, Len(SubStr) + 1)
    Else
        ExcludeFirstStr = str
    End If
End Function

Private Sub testExcludeFirstStr()
    Call Check("2345", ExcludeFirstStr("12345", "1"))
    Call Check("345", ExcludeFirstStr("12345", "12"))
    Call Check("45", ExcludeFirstStr("12345", "123"))
    Call Check("12345", ExcludeFirstStr("12345", "23"))
End Sub

'----------------------------------------
'・Last
'----------------------------------------
Public Function IsLastStr(ByVal str As String, ByVal SubStr As String) As Boolean
    Dim Result As Boolean: Result = False
    Do
        If SubStr = "" Then Exit Do
        If str = "" Then Exit Do
        If Len(str) < Len(SubStr) Then Exit Do

        If Mid$(str, Len(str) - Len(SubStr) + 1) = SubStr Then
            Result = True
        End If
    Loop While False
    IsLastStr = Result
End Function

Private Sub testIsLastStr()
    Call Check(True, IsLastStr("12345", "5"))
    Call Check(True, IsLastStr("12345", "45"))
    Call Check(True, IsLastStr("12345", "345"))
    Call Check(False, IsLastStr("12345", "34"))
    Call Check(False, IsLastStr("", "34"))
    Call Check(False, IsLastStr("12345", ""))
    Call Check(False, IsLastStr("123", "1234"))
End Sub

Public Function IncludeLastStr(ByVal str As String, ByVal SubStr As String) As String
    If IsLastStr(str, SubStr) Then
        IncludeLastStr = str
    Else
        IncludeLastStr = str + SubStr
    End If
End Function

Private Sub testIncludeLastStr()
    Call Check("12345", IncludeLastStr("12345", "5"))
    Call Check("12345", IncludeLastStr("12345", "45"))
    Call Check("12345", IncludeLastStr("12345", "345"))
    Call Check("1234534", IncludeLastStr("12345", "34"))
End Sub

Public Function ExcludeLastStr(ByVal str As String, ByVal SubStr As String) As String
    If IsLastStr(str, SubStr) Then
        ExcludeLastStr = Mid$(str, 1, Len(str) - Len(SubStr))
    Else
        ExcludeLastStr = str
    End If
End Function

Private Sub testExcludeLastStr()
    Call Check("1234", ExcludeLastStr("12345", "5"))
    Call Check("123", ExcludeLastStr("12345", "45"))
    Call Check("12", ExcludeLastStr("12345", "345"))
    Call Check("12345", ExcludeLastStr("12345", "34"))
End Sub

'----------------------------------------
'・Both
'----------------------------------------
Public Function IncludeBothEndsStr(ByVal str As String, ByVal SubStr As String) As String
    IncludeBothEndsStr = _
        IncludeFirstStr(IncludeLastStr(str, SubStr), SubStr)
End Function

Public Function ExcludeBothEndsStr(ByVal str As String, ByVal SubStr As String) As String
    ExcludeBothEndsStr = _
        ExcludeFirstStr(ExcludeLastStr(str, SubStr), SubStr)
End Function


'----------------------------------------
'◇First / Last Delim
'----------------------------------------

'----------------------------------------
'・FirstStrFirstDelim
'----------------------------------------
'   ・  先頭で見つかれば空文字を返す
'   ・  見つからなければ文字をそのまま返す
'----------------------------------------
Public Function FirstStrFirstDelim( _
ByVal Value As String, ByVal Delimiter As String) As String
    Dim Result As String: Result = ""
    Dim Index As Long: Index = InStr(Value, Delimiter)
    If 1 <= Index Then
        Result = Left$(Value, Index - 1)
    Else
        Result = Value
    End If
    FirstStrFirstDelim = Result
End Function

Public Sub testFirstStrFirstDelim()
    Call Check("123", FirstStrFirstDelim("123,456", ","))
    Call Check("123", FirstStrFirstDelim("123,456,789", ","))
    Call Check("123", FirstStrFirstDelim("123ttt456", "ttt"))
    Call Check("123", FirstStrFirstDelim("123ttt456", "tt"))
    Call Check("123", FirstStrFirstDelim("123ttt456", "t"))
    Call Check("123ttt456", FirstStrFirstDelim("123ttt456", ","))
    Call Check("", FirstStrFirstDelim(",123,", ","))
End Sub

'----------------------------------------
'・FirstStrLastDelim
'----------------------------------------

Public Function FirstStrLastDelim( _
ByVal Value As String, ByVal Delimiter As String) As String
    Dim Result As String: Result = ""
    Dim Index: Index = InStrRev(Value, Delimiter)
    If 1 <= Index Then
        Result = Left$(Value, Index - 1)
    Else
        Result = Value
    End If
    FirstStrLastDelim = Result
End Function

Public Sub testFirstStrLastDelim()
    Call Check("123", FirstStrLastDelim("123,456", ","))
    Call Check("123,456", FirstStrLastDelim("123,456,789", ","))
    Call Check("123", FirstStrLastDelim("123ttt456", "ttt"))
    Call Check("123t", FirstStrLastDelim("123ttt456", "tt"))
    Call Check("123tt", FirstStrLastDelim("123ttt456", "t"))
    Call Check("123ttt456", FirstStrLastDelim("123ttt456", ","))
    Call Check(",123", FirstStrLastDelim(",123,", ","))
End Sub


'----------------------------------------
'・LastStrFirstDelim
'----------------------------------------
Public Function LastStrFirstDelim( _
ByVal Value As String, ByVal Delimiter As String) As String
    Dim Result As String: Result = ""
    Dim Index: Index = InStr(Value, Delimiter)
    If 1 <= Index Then
        Result = Mid$(Value, Index + Len(Delimiter))
    Else
        Result = Value
    End If
    LastStrFirstDelim = Result
End Function

Public Sub testLastStrFirstDelim()
    Call Check("456", LastStrFirstDelim("123,456", ","))
    Call Check("456,789", LastStrFirstDelim("123,456,789", ","))
    Call Check("456", LastStrFirstDelim("123ttt456", "ttt"))
    Call Check("t456", LastStrFirstDelim("123ttt456", "tt"))
    Call Check("tt456", LastStrFirstDelim("123ttt456", "t"))
    Call Check("123ttt456", LastStrFirstDelim("123ttt456", ","))
    Call Check("123,", LastStrFirstDelim(",123,", ","))
End Sub

'----------------------------------------
'・LastStrLastDelim
'----------------------------------------
Public Function LastStrLastDelim( _
ByVal Value As String, ByVal Delimiter As String) As String
    Dim Result: Result = ""
    Dim Index As Long: Index = InStrRev(Value, Delimiter)
    If 1 <= Index Then
        Result = Mid$(Value, Index + Len(Delimiter))
    Else
        Result = Value
    End If
    LastStrLastDelim = Result
End Function

Public Sub testLastStrLastDelim()
    Call Check("456", LastStrLastDelim("123,456", ","))
    Call Check("789", LastStrLastDelim("123,456,789", ","))
    Call Check("456", LastStrLastDelim("123ttt456", "ttt"))
    Call Check("456", LastStrLastDelim("123ttt456", "tt"))
    Call Check("456", LastStrLastDelim("123ttt456", "t"))
    Call Check("123ttt456", LastStrLastDelim("123ttt456", ","))
    Call Check("", LastStrLastDelim(",123,", ","))
End Sub


'----------------------------------------
'◇Trim
'----------------------------------------
'Public Function TrimFirstChar(ByVal Str As String, ByVal TrimChar As String) As String
'    Do While IsFirstStr(Str, TrimChar)
'        Str = ExcludeFirstStr(Str, TrimChar)
'    Loop
'    TrimFirstChar = Str
'End Function
'
'Public Function TrimLastChar(ByVal Str As String, ByVal TrimChar As String) As String
'    Do While IsLastStr(Str, TrimChar)
'        Str = ExcludeLastStr(Str, TrimChar)
'    Loop
'    TrimLastChar = Str
'End Function
'
'Public Function TrimBothEndsChar(ByVal Str As String, ByVal TrimChar As String) As String
'    TrimBothEndsChar = _
'        TrimFirstChar(TrimLastChar(Str, TrimChar), TrimChar)
'End Function


Public Function TrimFirstStrs(ByVal str As String, ByRef TrimStrs() As String) As String
    Call Assert(IsArray(TrimStrs), "Error:TrimFirstStrs:TrimStrs is not Array.")
    Dim Result As String: Result = str
    Do
        str = Result
        Dim I As Long
        For I = LBound(TrimStrs) To UBound(TrimStrs)
            Result = ExcludeFirstStr(Result, TrimStrs(I))
        Next
    Loop While Result <> str
    TrimFirstStrs = Result
End Function

Private Sub testTrimFirstStrs()
    Call Check("123 ", TrimFirstStrs("   123 ", ArrayStr(" ")))
    Call Check(vbTab + "  123 ", TrimFirstStrs("   " + vbTab + "  123 ", ArrayStr(" ")))
    Call Check("123 ", TrimFirstStrs("   " + vbTab + "  123 ", ArrayStr(" ", vbTab)))
End Sub

Public Function TrimLastStrs(ByVal str As String, ByRef TrimStrs() As String) As String
    Call Assert(IsArray(TrimStrs), "Error:TrimLastStrs:TrimStrs is not Array.")
    Dim Result As String: Result = str
    Do
        str = Result
        Dim I As Long
        For I = LBound(TrimStrs) To UBound(TrimStrs)
            Result = ExcludeLastStr(Result, TrimStrs(I))
        Next
    Loop While Result <> str
    TrimLastStrs = Result
End Function

Private Sub testTrimLastStrs()
    Call Check(" 123", TrimLastStrs(" 123   ", ArrayStr(" ")))
    Call Check(" 123  " + vbTab, TrimLastStrs(" 123  " + vbTab + "   ", ArrayStr(" ")))
    Call Check(" 123", TrimLastStrs(" 123  " + vbTab + "   ", ArrayStr(" ", vbTab)))
End Sub

Public Function TrimBothEndsStrs(ByVal str As String, ByRef TrimStrs() As String) As String
    TrimBothEndsStrs = _
        TrimFirstStrs(TrimLastStrs(str, TrimStrs), TrimStrs)
End Function

Public Function TrimFirstSpace(ByVal str As String) As String
    TrimFirstSpace = TrimFirstStrs(str, ArrayStr("　", " ", vbCr, vbLf, vbTab))
End Function

Public Function TrimLastSpace(ByVal str As String) As String
    TrimLastSpace = TrimLastStrs(str, ArrayStr("　", " ", vbCr, vbLf, vbTab))
End Function

Public Function TrimBothEndsSpace(ByVal str As String) As String
    TrimBothEndsSpace = _
        TrimFirstSpace(TrimLastSpace(str))
End Function

'----------------------------------------
'◇置き換え処理
'----------------------------------------
'----------------------------------------
'・連続スペースを単独スペースに変換
'----------------------------------------
Public Function ReplaceContinuousSpace(ByVal Value As String, _
Optional Space As String = " ") As String
    Call Assert(Space <> "", "Error:ReplaceContinuousSpace:Space is Empty.")

    Dim Result As String
    Result = Value
    Do While IsIncludeStr(Result, Space + Space)
        Result = Replace(Result, Space + Space, Space)
    Loop
    ReplaceContinuousSpace = Result
End Function

Public Sub testReplaceContinuousSpace()
    Call Check(" A B C ", ReplaceContinuousSpace("  A  B   C "))

End Sub

'----------------------------------------
'・ HTML特殊文字の変換
'----------------------------------------
'   ・  HTML特殊文字は非常に多くあるみたいなので
'       全て対応はしないけど、主なものを変換できる関数
'----------------------------------------
Public Function String_HTMLtoText(ByVal Value As String) As String
    String_HTMLtoText = ReplaceArrayValue(Value, _
        ArrayStr("&amp;", "&gt;", "&lt;", "&nbsp;"), _
        ArrayStr("&", ">", "<", " "))
End Function


'----------------------------------------
'◇タグ処理
'----------------------------------------

'----------------------------------------
'・タグの内部文字列
'----------------------------------------
Public Function TagInnerText(ByVal Text As String, _
    ByVal StartTag As String, ByVal EndTag As String) As String

    Dim Result As String
    Result = LastStrFirstDelim(Text, StartTag)
    Result = FirstStrFirstDelim(Result, EndTag)
    TagInnerText = Result
End Function

Public Sub testTagInnerText()
    Call Check("456", TagInnerText("000<123>456<789>000", "<123>", "<789>"))
    Call Check("456", TagInnerText("<123>456<789>", "<123>", "<789>"))
    Call Check("456", TagInnerText("000<123>456", "<123>", "<789>"))
    Call Check("456", TagInnerText("456<789>000", "<123>", "<789>"))
    Call Check("456", TagInnerText("456", "<123>", "<789>"))
    Call Check("", TagInnerText("000<123><789>000", "<123>", "<789>"))
    
    Dim Text As String
    Text = "<123>123<789> <123>456<789> <123>789<789>"
    Call Check("123", TagInnerText(Text, "<123>", "<789>"))
    Call Check("<123>123", TagInnerText(Text, "<456>", "<789>"))
    Call Check("", TagInnerText(Text, "<456>", "<123>"))
    Call Check(Text, TagInnerText(Text, "<321>", "<456>"))
    
End Sub

'----------------------------------------
'・タグの内部文字列、後方検索版
'----------------------------------------
Public Function TagInnerTextLast(ByVal Text As String, _
    ByVal StartTag As String, ByVal EndTag As String) As String

    Dim Result As String
    Result = LastStrLastDelim(Text, StartTag)
    Result = FirstStrLastDelim(Result, EndTag)
    TagInnerTextLast = Result
End Function


'----------------------------------------
'・タグを含んだ内部文字列
'----------------------------------------
Public Function TagOuterText(ByVal Text As String, _
    ByVal StartTag As String, ByVal EndTag As String) As String

    Dim Result1 As String
    Dim Result2 As String
    Result1 = LastStrFirstDelim(Text, StartTag)
    If Result1 <> Text Then
        Result1 = StartTag + Result1
    End If

    Result2 = FirstStrFirstDelim(Result1, EndTag)
    If Result2 <> Result1 Then
        Result2 = Result2 + EndTag
    End If
    TagOuterText = Result2
End Function

Public Sub testTagOuterText()
    Call Check("<123>456<789>", TagOuterText("000<123>456<789>000", "<123>", "<789>"))
    Call Check("<123>456<789>", TagOuterText("<123>456<789>", "<123>", "<789>"))
    Call Check("<123>456", TagOuterText("000<123>456", "<123>", "<789>"))
    Call Check("456<789>", TagOuterText("456<789>000", "<123>", "<789>"))
    Call Check("456", TagOuterText("456", "<123>", "<789>"))
End Sub


'----------------------------------------
'・指定のタグではさまれた文字列のリストを出力する
'----------------------------------------
'   ・ 結果は改行コードで区切られて出力される
'----------------------------------------
Public Function TagOuterTextList(ByVal Text As String, _
    ByVal StartTag As String, ByVal EndTag As String) As String

    Dim Result As String: Result = ""
    Dim StartTagToEnd As String
    Dim InnerText As String
    Do
        StartTagToEnd = LastStrFirstDelim(Text, StartTag)
        If StartTagToEnd = Text Then Exit Do
        InnerText = FirstStrFirstDelim(StartTagToEnd, EndTag)
        If InnerText = StartTagToEnd Then Exit Do
        Result = StringCombine(vbCrLf, Result, _
            StartTag + InnerText + EndTag)
        Text = LastStrFirstDelim(StartTagToEnd, EndTag)
    Loop While True
    TagOuterTextList = Result
End Function

Public Sub testTagOuterTextList()

    Call Check("http://a.jpg" + vbCrLf + "http://b.jpg", _
        TagOuterTextList("abc http://a.jpg def http://b.jpg ghi", _
            "http://", ".jpg"))
End Sub

'----------------------------------------
'・ 任意のタグを削除する関数
'---------------------------------------
'   ・  HTMLタグ削除や
'       特定のタグだけ削除、という処理が行える
'---------------------------------------
Public Function String_TagDelete(ByVal Text As String, _
ByVal StartTag As String, ByVal EndTag As String) As String
    Dim Result As String
    
    Result = Text
    Do
        If IsIncludeStr(Text, StartTag) = False Then Exit Do
        Result = FirstStrFirstDelim(Text, StartTag)
        Result = Result + LastStrFirstDelim(Text, EndTag)
        Text = Result
    Loop While True
    
    String_TagDelete = Result
End Function

Public Sub test_String_TagDelete()
    Call Check("123", String_TagDelete("123<a>456</a>", "<a>", "</a>"))
    Call Check("456", String_TagDelete("<a>123</a>456", "<a>", "</a>"))
    Call Check("", String_TagDelete("<a>123</a><a>456</a>", "<a>", "</a>"))
    Call Check("123456", String_TagDelete("123456", "<a>", "</a>"))

    Call Check("123456", String_TagDelete("<a>123</a><a>456</a>", "<", ">"))

End Sub


'----------------------------------------
'◇文字列リスト処理
'----------------------------------------

'----------------------------------------
'・ 文字列から空白行だけの行を削除する
'----------------------------------------
Public Function String_DeleteSpaceLine(ByVal Value As String) As String
    Dim Lines() As String
    Lines = Split(Replace(Replace(Value, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim Line As String
    
    Dim I As Long
    For I = ArrayCount(Lines) - 1 To 0 Step -1
        If TrimBothEndsSpace(Lines(I)) = "" Then
            Call ArrayDelete(Lines, I)
        End If
    Next
    String_DeleteSpaceLine = ArrayToString(Lines, vbCrLf)
End Function

'----------------------------------------
'・ 文字列から行ごとにTrimを行う
'----------------------------------------
Public Function String_LineTrim(ByVal Value As String) As String
    Dim Lines() As String
    Lines = Split(Replace(Replace(Value, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim Line As String
    
    Dim I As Long
    For I = ArrayCount(Lines) - 1 To 0 Step -1
        Lines(I) = TrimBothEndsSpace(Lines(I))
    Next
    String_LineTrim = ArrayToString(Lines, vbCrLf)
End Function




'----------------------------------------
'◇文字列結合
'----------------------------------------

'----------------------------------------
'・文字列結合
'----------------------------------------
'   ・  少なくとも1つのDelimiterが間に入って接続される。
'   ・  Delimiterが結合の両端に付属する場合も1つになる。
'   ・  2連続で結合の両端にある場合は1つが削除される
'       (テストでの動作参照)
'----------------------------------------

Public Function StringCombine(ByVal Delimiter As String, _
ParamArray Values()) As String

    Dim Parameter() As String
    ReDim Parameter(UBound(Values))
    Dim I As Long
    For I = 0 To UBound(Values)
        Parameter(I) = Values(I)
    Next

    StringCombine = StringCombineArray(Delimiter, Parameter)
End Function

Private Sub testStringCombine()
    Call Check("1.2.3.4", StringCombine(".", "1", "2", "3", "4"))

    Call Check("1.2", StringCombine(".", "1.", "2"))
    Call Check("1.2", StringCombine(".", "1.", ".2"))
    Call Check("1..2", StringCombine(".", "1..", "2"))
    Call Check("1..2", StringCombine(".", "1", "..2"))

    Call Check("1..2", StringCombine(".", "1..", ".2"))
    Call Check("1..2", StringCombine(".", "1.", "..2"))
    Call Check("1...2", StringCombine(".", "1..", "..2"))

    Call Check("1.2.3", StringCombine(".", "1.", ".2", "3"))
    Call Check("1.2.3", StringCombine(".", "1.", ".2.", "3"))
    Call Check("1.2.3", StringCombine(".", "1.", ".2", ".3"))
    Call Check("1.2.3", StringCombine(".", "1.", ".2.", ".3"))

    Call Check("1.2.3..4", StringCombine(".", "1.", ".2", "3", "..4"))
    Call Check("1..2.3..4", StringCombine(".", "1..", ".2", "3.", "..4"))
    Call Check(".1..2.3..4..", StringCombine(".", ".1..", ".2", "3.", "..4.."))

    Call Check("1.2.3.4", StringCombine(".", "1", "2", "3", "4"))
    Call Check("1..2.3..4", StringCombine(".", "1..", ".2", "3.", "..4"))
    Call Check(".1..2.3..4..", StringCombine(".", ".1..", ".2", "3.", "..4.."))

    Call Check("", StringCombine(vbCrLf, "", ""))
    Call Check("", StringCombine(vbCrLf, "", "", ""))
    Call Check("A", StringCombine(vbCrLf, "A", ""))
    Call Check("A", StringCombine(vbCrLf, "A", "", ""))
    Call Check("A", StringCombine(vbCrLf, "", "A"))
    Call Check("A", StringCombine(vbCrLf, "", "", "A"))
    Call Check("A", StringCombine(vbCrLf, "", "A", ""))

    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "A", "B", ""))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "A", "B", "", ""))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "", "A", "B"))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "", "", "A", "B"))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "", "A", "B", ""))

    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "A", "", "B"))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "", "A", "", "B"))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "A", "", "B", ""))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "", "A", "", "B", ""))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "A", "", "", "B"))
    Call Check("A" + vbCrLf + "B", StringCombine(vbCrLf, "", "", "A", "", "", "B", "", ""))

    Call Check("\\test\temp\temp\temp\", StringCombine("\", "\\test\", "\temp\", "temp", "\temp\"))
End Sub


Public Function StringCombineArray(ByVal Delimiter As String, _
ByRef Values() As String) As String

    Call Assert(IsArray(Values), _
        "Error:StringCombineArray:ArrayValue is not Array.")

    Dim Result: Result = ""
    Dim Count: Count = ArrayCount(Values)
    If Count = 0 Then

    ElseIf Count = 1 Then
        Result = Values(0)
    Else
        Dim I
        For I = 0 To Count - 1
        Do
            If Values(I) = "" Then Exit Do
            If Result = "" Then
                Result = Values(I)
            Else
                Result = _
                    ExcludeLastStr(Result, Delimiter) + _
                    Delimiter + _
                    ExcludeFirstStr(Values(I), Delimiter)
            End If

        Loop While False
        Next
    End If
    StringCombineArray = Result
End Function

Private Sub testStringArrayCombine()
    Call Check("1.2.3.4", StringCombineArray(".", ArrayStr("1", "2", "3", "4")))
    Call Check("1..2.3..4", StringCombineArray(".", ArrayStr("1..", ".2", "3.", "..4")))
    Call Check(".1..2.3..4..", StringCombineArray(".", ArrayStr(".1..", ".2", "3.", "..4..")))

    Dim Values() As String
    Values = Split("A,B,C", ",")
    Call Check("A,B,C", StringCombineArray(",", Values))

    Values = Split("A,B,,C", ",")
    Call Check("A,B,C", StringCombineArray(",", Values))

    Values = Split("A,B,,,C", ",")
    Call Check("A,B,C", StringCombineArray(",", Values))

    Values = Split("A,B,,,,C", ",")
    Call Check("A,B,C", StringCombineArray(",", Values))

    Values = Split(",,A,B,,,,C,,", ",")
    Call Check("A,B,C", StringCombineArray(",", Values))
End Sub


'----------------------------------------
'◇Byte指定文字列処理
'----------------------------------------
'   ・  SJISに変換しているためUnicode固有文字は非対応
'       標準のLenB関数はUnicode文字扱いなのでSJIS対応ではない
'----------------------------------------

'----------------------------------------
'・Byte数取得
'----------------------------------------
Public Function LengthSjisByte(ByVal S As String) As Long
    LengthSjisByte = LenB(StrConv(S, vbFromUnicode))
End Function

'----------------------------------------
'・Byte数で切り出すLeft関数
'----------------------------------------
Public Function LeftSjisByte(ByVal S As String, _
ByVal ByteLength As Long) As String
    Dim Result As String: Result = ""
    Do
        If S = "" Then Exit Do
        Result = StrConv( _
            LeftB(StrConv(S, vbFromUnicode), ByteLength) _
            , vbUnicode)
    Loop While False
    LeftSjisByte = Result
End Function

'----------------------------------------
'・Byte数で切り出すRight関数
'----------------------------------------
Function RightSjisByte(S As String, ByteLength As Long) As String
    Dim Result As String
    Result = ""
    Do
        If S = "" Then Exit Do
        Result = StrConv( _
            RightB(StrConv(S, vbFromUnicode), ByteLength) _
            , vbUnicode)
    Loop While False
    RightSjisByte = Result
End Function

'----------------------------------------
'・Byte数で切り出すMid関数
'----------------------------------------
Function MidSjisByte(S As String, Start As Long, Optional ByteLength As Long)
    Dim Result As String
    Result = ""
    Do
        If IsMissing(ByteLength) Then
            Result = StrConv( _
                MidB(StrConv(S, vbFromUnicode), Start) _
                , vbUnicode)
        Else
            Result = StrConv( _
                MidB(StrConv(S, vbFromUnicode), Start, ByteLength) _
                , vbUnicode)
        End If
    Loop While False
    MidSjisByte = Result
End Function


'----------------------------------------
'◇文字列正規表現
'----------------------------------------

'----------------------------------------
'・正規表現の一致を確認する
'----------------------------------------
'   ・  動作対象は1行テキスト
'   ・  RegExpオブジェクトは外部から指定可能
'   ・  Matchオブジェクトは戻り値として利用可能
'----------------------------------------
Function MatchRegExp(ByVal Text As String, ByVal MatchWord As String, _
Optional CaseCompare As CaseCompare = IgnoreCase, _
Optional RegExp As Object = Nothing, _
Optional Match As Object = Nothing) As Boolean

On Error GoTo Err:
    Dim Result As Boolean
    Result = False
    Do
        If (MatchWord = "") Or (Text = "") Then Exit Do

        '正規表現用オブジェクト用意
        Dim RegCreateFlag As Boolean
        RegCreateFlag = False
        If RegExp Is Nothing Then
            RegCreateFlag = True
            Set RegExp = CreateObject("VBScript.RegExp")
        End If

        '正規表現マッチ調査
        RegExp.Pattern = MatchWord
        RegExp.Global = True
        RegExp.IgnoreCase = CaseCompare = IgnoreCase
        Set Match = RegExp.Execute(Text)
        If 1 <= Match.Count Then
            Result = True
        End If

        If RegCreateFlag Then
            Set RegExp = Nothing
        End If

    Loop While False
    MatchRegExp = Result
    Exit Function
Err:
    MatchRegExp = False
End Function


'----------------------------------------
'・正規表現での置き換え
'----------------------------------------
'   ・  動作対象は1行テキスト
'   ・  RegExpオブジェクトは外部から指定可能
'----------------------------------------
Public Function ReplaceRegExp(ByVal Value As String, ByVal Pattern As String, _
ByVal NewValue As String, _
Optional ByVal CaseCompare As CaseCompare = CaseSensitive, _
Optional RegExp As Object = Nothing) As String

On Error GoTo Err:
    Dim Result As String: Result = Value
    Do
        If (Pattern = "") Or (Value = "") Then Exit Do

        '正規表現用オブジェクト用意
        Dim RegCreateFlag As Boolean
        RegCreateFlag = False
        If RegExp Is Nothing Then
            RegCreateFlag = True
            Set RegExp = CreateObject("VBScript.RegExp")
        End If

        '正規表現マッチ調査
        RegExp.Pattern = Pattern
        RegExp.IgnoreCase = (CaseCompare = IgnoreCase)
        RegExp.Global = True

        Result = RegExp.Replace(Value, NewValue)

        If RegCreateFlag Then
            Set RegExp = Nothing
        End If

    Loop While False
Err:
    ReplaceRegExp = Result
End Function

'----------------------------------------
'・正規表現での削除
'----------------------------------------
'   ・  動作対象は1行テキスト
'   ・  RegExpオブジェクトは外部から指定可能
'----------------------------------------
Public Function DeleteRegExp(ByVal Value As String, ByVal Pattern As String, _
Optional ByVal CaseCompare As CaseCompare = CaseSensitive, _
Optional RegExp As Object = Nothing) As String

    DeleteRegExp = _
        ReplaceRegExp(Value, Pattern, "", CaseCompare, RegExp)

End Function

'----------------------------------------
'・HTMLタグを削除する関数
'----------------------------------------
Public Function DeleteHTMLTag(ByVal Value As String) As String
    DeleteHTMLTag = _
        DeleteRegExp(Value, "<[^>]*>", IgnoreCase)
End Function

Public Sub testDeleteHTMLTag()
    Call Check("abc", DeleteHTMLTag("<b>abc</b>"))
End Sub

'----------------------------------------
'◇配列指定処理
'----------------------------------------

'----------------------------------------
'・文字列の連続置き換え、配列指定
'----------------------------------------
Public Function ReplaceArrayValue(ByVal Value As String, _
ByRef OldTableArray() As String, NewTableArray() As String) As String
    Call Assert(ArrayCount(OldTableArray) = ArrayCount(NewTableArray), _
        "Error:ReplaceArrayValue:OldTableArray's Count is not same NewTableArray's Count")

    Dim Result As String
    Result = Value
    Dim I As Long
    For I = 0 To ArrayCount(OldTableArray) - 1
        Result = Replace(Result, OldTableArray(I), NewTableArray(I))
    Next
    ReplaceArrayValue = Result
End Function

'----------------------------------------
'・文字列の連続削除、配列指定
'----------------------------------------
Public Function DeleteArrayValue(ByVal Value As String, _
DeleteTableArray() As String) As String
    Dim Result As String
    Result = Value
    Dim I As Long
    For I = 0 To ArrayCount(DeleteTableArray) - 1
        Result = Replace(Result, DeleteTableArray(I), "")
    Next
    DeleteArrayValue = Result
End Function


'----------------------------------------
'◇配列指定(正規表現)処理
'----------------------------------------
'----------------------------------------
'・文字列の連続置き換え、正規表現、配列指定
'----------------------------------------
Public Function ReplaceArrayRegExp(ByVal Value As String, _
ByRef OldTableArray() As String, NewTableArray() As String, _
Optional ByVal CaseCompare As CaseCompare = CaseSensitive) As String
    Call Assert(ArrayCount(OldTableArray) = ArrayCount(NewTableArray), _
        "Error:ReplaceArrayValue:OldTableArray's Count is not same NewTableArray's Count")

    Dim Result As String
    Result = Value

    Dim RegExp As Object
    Set RegExp = CreateObject("VBScript.RegExp")

    Dim I As Long
    For I = 0 To ArrayCount(OldTableArray) - 1
        Result = ReplaceRegExp(Result, OldTableArray(I), NewTableArray(I), CaseCompare, RegExp)
    Next
    Set RegExp = Nothing

    ReplaceArrayRegExp = Result
End Function

'----------------------------------------
'・文字列の連続削除、正規表現、配列指定
'----------------------------------------
Public Function DeleteArrayRegExp(ByVal Value As String, _
DeleteTableArray() As String, _
Optional ByVal CaseCompare As CaseCompare = CaseSensitive) As String
    Dim Result As String
    Result = Value
    Dim I As Long
    For I = 0 To ArrayCount(DeleteTableArray) - 1
        Result = ReplaceRegExp(Result, DeleteTableArray(I), "", CaseCompare)
    Next
    DeleteArrayRegExp = Result
End Function


'----------------------------------------
'◇ 文字コード処理
'----------------------------------------

'----------------------------------------
'・ 文字列のShiftJIS以外の文字抽出
'----------------------------------------
'   ・  全てShiftJISに変換可能なら空文字が返る
'----------------------------------------
Public Function String_GetOutShiftJIS(ByVal Value As String) As String
    Dim Result As String: Result = ""
    Dim Character As String
    Dim I As Integer
    For I = 1 To Len(Value)
        Character = Mid(Value, I, 1)
        If StrConv(StrConv(Character, vbFromUnicode), vbUnicode) <> Character Then
            Result = Result + Character
        End If
    Next I
    String_GetOutShiftJIS = Result
End Function

Public Sub test_String_GetOutShiftJIS()
    Call Check("", String_GetOutShiftJIS("テスト"))
    Call Check(ChrW(&H33A5), String_GetOutShiftJIS("あいうえお" + ChrW(&H33A5)))
End Sub


'----------------------------------------
'・ 文字列のShiftJISの外字抽出
'----------------------------------------
'   ・  SJISには外字として
'       ＮＥＣ選定特文字
'       ＩＢＭ拡張文字
'       というものがあり、それを抽出する
'   ・  １．ＮＥＣ選定特文字
'           開始：まる数字の１「いち」（34624）
'　　　　　 終了：直角三角形「でるた」(34713)
'
'       ※  実際には、合併集合「しゅうごう」(34716)までだが
'           別の文字コードが優先される。
'　　　　　 下記のNo.３がそれ、
'
'　　   ２．ＩＢＭ拡張文字
'　　　　　 開始：ローマ数字の小文字の１「いち」（64064）
'　　　　　 終了：黒に似た漢字「ＩＭＥのコード一覧で表示して下さい。」（64587）
'
'　　   ３．ＮＥＣ選定特文字で１の範囲外
'　　　　　 (1)Ｕに似た記号「しゅうごう」（33214）
'　　　　　 (2)Ｕを逆さにした記号「しゅうごう」（33215）
'　　　　　 (3)点３つ「なぜならば」(33254)
'----------------------------------------
Public Function String_GetMachineDependentCharacter(ByVal Value As String) As String
    Dim Result As String
    Result = ""
    Dim CharIntValue As Integer
    Dim I As Integer
    For I = 1 To Len(Value)
        CharIntValue = Asc(Mid(Value, I, 1))
        If ((Asc(Chr(34624)) <= CharIntValue) And (CharIntValue <= Asc(Chr(34713)))) _
        Or ((Asc(Chr(64064)) <= CharIntValue) And (CharIntValue <= Asc(Chr(64587)))) _
        Or (CharIntValue = Asc(Chr(33214))) _
        Or (CharIntValue = Asc(Chr(33215))) _
        Or (CharIntValue = Asc(Chr(33254))) Then
            Result = Result & Chr(CharIntValue)
        End If
    Next I
    String_GetMachineDependentCharacter = Result
End Function

Public Sub test_String_GetMachineDependentCharacter()
    Call Check("", String_GetMachineDependentCharacter("高橋"))
    Call Check("髙", String_GetMachineDependentCharacter("髙橋"))
    Call Check("", String_GetMachineDependentCharacter("山崎"))
    Call Check("﨑", String_GetMachineDependentCharacter("山﨑"))
    Call Check("", String_GetMachineDependentCharacter(""))
    Call Check("髙", String_GetMachineDependentCharacter("髙橋"))
End Sub


'----------------------------------------
'◆日付時刻処理
'----------------------------------------

'----------------------------------------
'・月の最終日を取得
'----------------------------------------
Public Function MonthLastDay(ByVal DateValue As Date) As Date
    MonthLastDay = DateSerial(Year(DateValue), Month(DateValue) + 1, 0)
End Function

Private Sub testMonthLastDay()
    Call Check( _
        DateValue("2014/11/30"), _
        MonthLastDay(DateValue("2014/11/3")) _
        )
    Call Check( _
        DateSerial(2014, 11, 30), _
        MonthLastDay(DateSerial(2014, 11, 3)) _
        )
End Sub

'----------------------------------------
'・月の日数取得
'----------------------------------------
Public Function MonthDayCount(ByVal DateValue As Date) As Long
    MonthDayCount = _
        Day(MonthLastDay(DateValue))
End Function

Private Sub testMonthMonthDayCount()
    Call Check( _
        30, _
        MonthDayCount(DateValue("2014/11/3")) _
        )
    Call Check( _
        28, _
        MonthDayCount(DateValue("2014/2/3")) _
        )
End Sub


'----------------------------------------
'◇今週/先週/来週の曜日指定の日付取得
'----------------------------------------

Public Function ThisWeekDay(ByVal WeekDayValue As Long, ByVal DateValue As Date) As Date
    ThisWeekDay = _
        DateAdd("d", (WeekDayValue - Weekday(DateValue)), DateValue)
End Function

Public Sub testThisWeekDay()
    Call Check(CDate("2016/02/21"), ThisWeekDay(vbSunday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/22"), ThisWeekDay(vbMonday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/23"), ThisWeekDay(vbTuesday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/24"), ThisWeekDay(vbWednesday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/25"), ThisWeekDay(vbThursday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/26"), ThisWeekDay(vbFriday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/27"), ThisWeekDay(vbSaturday, CDate("2016/02/23")))
End Sub

Public Function LastWeekDay(ByVal WeekDayValue As Long, ByVal DateValue As Date) As Date
    LastWeekDay = _
        DateAdd("d", -7, ThisWeekDay(WeekDayValue, DateValue))
End Function

Public Sub testLastWeekDay()
    Call Check(CDate("2016/02/14"), LastWeekDay(vbSunday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/15"), LastWeekDay(vbMonday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/16"), LastWeekDay(vbTuesday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/17"), LastWeekDay(vbWednesday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/18"), LastWeekDay(vbThursday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/19"), LastWeekDay(vbFriday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/20"), LastWeekDay(vbSaturday, CDate("2016/02/23")))
End Sub

Public Function NextWeekDay(ByVal WeekDayValue As Long, ByVal DateValue As Date) As Date
    NextWeekDay = _
        DateAdd("d", 7, ThisWeekDay(WeekDayValue, DateValue))
End Function

Public Sub testNextWeekDay()
    Call Check(CDate("2016/02/28"), NextWeekDay(vbSunday, CDate("2016/02/23")))
    Call Check(CDate("2016/02/29"), NextWeekDay(vbMonday, CDate("2016/02/23")))
    Call Check(CDate("2016/03/01"), NextWeekDay(vbTuesday, CDate("2016/02/23")))
    Call Check(CDate("2016/03/02"), NextWeekDay(vbWednesday, CDate("2016/02/23")))
    Call Check(CDate("2016/03/03"), NextWeekDay(vbThursday, CDate("2016/02/23")))
    Call Check(CDate("2016/03/04"), NextWeekDay(vbFriday, CDate("2016/02/23")))
    Call Check(CDate("2016/03/05"), NextWeekDay(vbSaturday, CDate("2016/02/23")))
End Sub

'----------------------------------------
'◇日付時刻書式指定
'----------------------------------------

'----------------------------------------
'・日付書式
'----------------------------------------
Public Function FormatYYYYMMDD(ByVal DateValue As Date) As String
    FormatYYYYMMDD = FormatYYYY_MM_DD(DateValue, "")
End Function

Sub testFormatYYYYMMDD()
    Dim Value As Date: Value = CDate("2015/02/03")
    Call Check("20150203", FormatYYYYMMDD(Value))
End Sub

Public Function FormatYYYY_MM_DD( _
ByVal DateValue As Date, ByVal Delimiter As String) As String
    FormatYYYY_MM_DD = Format(DateValue, _
        "YYYY" + Delimiter + "MM" + Delimiter + "DD")
End Function

Public Function FormatYYYY_MM( _
ByVal DateValue As Date, ByVal Delimiter As String) As String
    FormatYYYY_MM = Format(DateValue, _
        "YYYY" + Delimiter + "MM")
End Function


'----------------------------------------
'・時刻書式
'----------------------------------------
Public Function FormatHHMMSS(ByVal TimeValue As Date) As String
    FormatHHMMSS = FormatHH_MM_SS(TimeValue, "")
End Function

Sub testFormatHHMMSS()
    Dim Value As Date: Value = CDate("2015/02/03 05:05")
    Call Check("05:05:00", FormatHH_MM_SS(Value, ":"))
End Sub

Public Function FormatHH_MM_SS( _
ByVal TimeValue As Date, ByVal Delimiter As String) As String
    FormatHH_MM_SS = Format(TimeValue, _
        "HH" + Delimiter + "NN" + Delimiter + "SS")
End Function

Public Function FormatHH_MM( _
ByVal TimeValue As Date, ByVal Delimiter As String)
    FormatHH_MM = Format(TimeValue, _
        "HH" + Delimiter + "NN")
End Function

'----------------------------------------
'・標準的な日付時刻書式文字列の取得
'----------------------------------------
Public Function FormatDateTimeNormal(DateValue As Date) As String
    FormatDateTimeNormal = _
        FormatYYYY_MM_DD(DateValue, "/") + _
        " " + _
        FormatHH_MM_SS(DateValue, ":")
End Function

'----------------------------------------
'・日付時刻書式
'----------------------------------------
Public Function FormatYYYYMMDDHHMMSS(ByVal DateTimeValue As Date) As String
    FormatYYYYMMDDHHMMSS = _
        FormatYYYYMMDD(DateTimeValue) + _
        FormatHHMMSS(DateTimeValue)
End Function


Public Function FormatYYYYMMDDHHMMSS_Hyphen(ByVal DateTimeValue)
    FormatYYYYMMDDHHMMSS_Hyphen = _
        FormatYYYY_MM_DD(DateTimeValue, "-") + "-" + _
        FormatHH_MM_SS(DateTimeValue, "-")
End Function

'----------------------------------------
'・Format文を変換してYYYYMMDDHHNNSS以外の変換を行わせない関数
'----------------------------------------
'   ・  日付時刻とは関係のないFomat関数の書式文字列を使えなくして
'       日付時刻書式だけを指定できるようにした
'----------------------------------------
Public Function Format_Date_UseOnlyYMDHNS( _
ByVal DateValue As Date, _
ByVal FormatStr As String) As String
    Dim Result As String: Result = ""
    Result = Format(DateValue, _
        ReplaceArrayValue(FormatStr, _
            ArrayStr("\", "0", "#", "%", "@", "&", "!", "<", ">", "."), _
            ArrayStr("\\", "\0", "\#", "\%", "\@", "\&", "\!", "\<", "\>", "\.")))
    Format_Date_UseOnlyYMDHNS = Result
End Function

Public Sub testFormat_Date_UseOnlyYMDHNS()
    Dim DateValue As Date
    DateValue = CDate("2017/03/21")
    
    Call Check("2017/03/21", Format(DateValue, "@"))
    Call Check("A-B-C", Format("ABC", "@-@-@"))
    Call Check("ABC", Format("ABC", "@@"))
    Call Check("ABC", Format("ABC", "@@@"))
    Call Check(" ABC", Format("ABC", "@@@@"))
    Call Check("ABC ", Format("ABC", "!@@@@"))
    Call Check(" abc", Format("ABC", "<@@@@"))
    Call Check(" ABC", Format("abc", ">@@@@"))
    Call Check("A-B-C", Format("ABC", "&-&-&"))
    Call Check("ABC", Format("ABC", "&&"))
    Call Check("ABC", Format("ABC", "&&&"))
    Call Check("ABC", Format("ABC", "&&&&"))
    Call Check("123.4.0.0.", Format("123.4", ".0.0.0."))
    Call Check(",123", Format("123456", ",0,0,0,"))
    
    Call Check("@", Format_Date_UseOnlyYMDHNS(DateValue, "@"))
    Call Check("@-@-@", Format_Date_UseOnlyYMDHNS(DateValue, "@-@-@"))
    Call Check("@@", Format_Date_UseOnlyYMDHNS(DateValue, "@@"))
    Call Check("@@@", Format_Date_UseOnlyYMDHNS(DateValue, "@@@"))
    Call Check("@@@@", Format_Date_UseOnlyYMDHNS(DateValue, "@@@@"))
    Call Check("!@@@@", Format_Date_UseOnlyYMDHNS(DateValue, "!@@@@"))
    Call Check("<@@@@", Format_Date_UseOnlyYMDHNS(DateValue, "<@@@@"))
    Call Check(">@@@@", Format_Date_UseOnlyYMDHNS(DateValue, ">@@@@"))
    Call Check("&-&-&", Format_Date_UseOnlyYMDHNS(DateValue, "&-&-&"))
    Call Check("&&", Format_Date_UseOnlyYMDHNS(DateValue, "&&"))
    Call Check("&&&", Format_Date_UseOnlyYMDHNS(DateValue, "&&&"))
    Call Check("&&&&", Format_Date_UseOnlyYMDHNS(DateValue, "&&&&"))
    Call Check(".0.0.0.", Format_Date_UseOnlyYMDHNS(DateValue, ".0.0.0."))
    Call Check(",0,0,0,", Format_Date_UseOnlyYMDHNS(DateValue, ",0,0,0,"))

End Sub

'----------------------------------------
'◆配列処理
'----------------------------------------

'----------------------------------------
'◇配列基本操作
'----------------------------------------

'----------------------------------------
'・要素無し配列に対してもエラーの起きないUBound/LBound
'----------------------------------------
'   ・  UBoundはArray()で返される要素無しの配列には-1を返すが
'       宣言しただけの動的配列ではエラーになるのでそれを防止する。
'   ・  Dimension:次元数は、多次元配列の場合その次元での結果を返す
'----------------------------------------
Public Function UBoundNoError(ByRef Value As Variant, _
Optional Dimension = 1) As Long
On Error Resume Next
    Call Assert(IsArray(Value), "Error:UBoundNoError:Value is not Array.")
    UBoundNoError = -1
    UBoundNoError = UBound(Value, Dimension)
End Function

Public Function LBoundNoError(ByRef Value As Variant, _
Optional Dimension = 1) As Long
On Error Resume Next
    Call Assert(IsArray(Value), "Error:LBoundNoError:Value is not Array.")
    LBoundNoError = 0
    LBoundNoError = LBound(Value, Dimension)
End Function

'----------------------------------------
'・配列の要素数を求める関数
'----------------------------------------
'   ・  LBound=0 でも 1 でも対応する。
'   ・  Dimension:次元数は、多次元配列の場合その次元での結果を返す
'----------------------------------------
Public Function ArrayCount(ByRef ArrayValue As Variant, _
Optional Dimension = 1) As Long
    Call Assert(IsArray(ArrayValue), "Error:ArrayCount:ArrayValue is not Array.")

    ArrayCount = _
        UBoundNoError(ArrayValue, Dimension) - _
        LBoundNoError(ArrayValue, Dimension) + 1
    '配列要素がない場合はUBound=-1/LBound=0になるので
    '配列要素数計算は正しく行われる。
End Function

Private Sub testArrayCount()
    Dim A() As String
    Call Check(0, ArrayCount(A))
    Call Check(0, ArrayCount(Array()))
    Call Check(1, ArrayCount(Split("123", ",")))
    Call Check(2, ArrayCount(Split("1,3", ",")))

    '二次元配列
    Dim B(3, 4) As String
    Call Check(4, ArrayCount(B, 1))
    Call Check(5, ArrayCount(B, 2))

    '三次元配列
    Dim C(5, 6, 7) As String
    Call Check(6, ArrayCount(C, 1))
    Call Check(7, ArrayCount(C, 2))
    Call Check(8, ArrayCount(C, 3))

End Sub

'----------------------------------------
'・配列を拡張する
'----------------------------------------
'   ・  固定配列は対応しない
'----------------------------------------
Public Sub SetArrayCount(ByRef ArrayValue As Variant, ByVal Count As Long)
    Call Assert(IsArray(ArrayValue), "Error:ArrayAdd:ArrayValue is not Array.")

    If ArrayCount(ArrayValue) < Count Then
        ReDim Preserve ArrayValue(Count - 1)
    ElseIf Count < ArrayCount(ArrayValue) Then
        ReDim Preserve ArrayValue(Count - 1)
    Else
    End If
End Sub

Public Sub testSetArrayCount()
    Dim A()
    A = Array("A", "B", "C")
    Call Check(3, ArrayCount(A))

    Call SetArrayCount(A, 5)
    Call Check(5, ArrayCount(A))
    A(3) = "D"
    A(4) = "E"
    
    Call Check("A,B,C,D,E", ArrayToString(A, ","))

    Call SetArrayCount(A, 2)
    Call Check("A,B", ArrayToString(A, ","))

End Sub

'----------------------------------------
'・配列の要素を追加する
'----------------------------------------
'   ・  オブジェクト値にも対応
'   ・  固定配列は対応しない
'----------------------------------------
Public Sub ArrayAdd(ByRef ArrayValue As Variant, ByVal Value As Variant)
    Call Assert(IsArray(ArrayValue), "Error:ArrayAdd:ArrayValue is not Array.")

    ReDim Preserve ArrayValue(ArrayCount(ArrayValue))
    Call SetValue(ArrayValue(UBound(ArrayValue)), Value)
End Sub

Private Sub testArrayAdd()
    Dim A()
    A = Array("A", "B", "C")

    Call ArrayAdd(A, "D")
    Call Check(4, ArrayCount(A))
    Call Check("D", A(3))

    'オブジェクト値にも対応
    Dim B()
    ReDim B(2)
    Set B(0) = CreateObject("VBScript.RegExp")
    Set B(1) = Shell
    Set B(2) = CreateObject("ADODB.Stream")
    Call ArrayAdd(B, fso)
    Call Check("test.txt", B(3).GetFileName("C:\temp\test.txt"))

    '二次元配列
    Dim C() As String
    ReDim Preserve C(3, 4)
    Call Check(4, ArrayCount(C, 1))
    Call Check(5, ArrayCount(C, 2))

    ReDim Preserve C(3, 5)
    Call Check(4, ArrayCount(C, 1))
    Call Check(6, ArrayCount(C, 2))

'    Call SetValue(C(UBound(C)), "abc")
End Sub

'----------------------------------------
'・配列の要素を重複チェックしてから追加する
'----------------------------------------
Public Sub ArrayAddNotDuplicate(ByRef ArrayValue As Variant, ByVal Value As Variant)
    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    If ArrayExists(ArrayValue, Value) = False Then
        Call ArrayAdd(ArrayValue, Value)
    End If
End Sub

'----------------------------------------
'・配列に配列を追加する
'----------------------------------------
Public Sub ArrayAddArray(ByRef ArrayValue As Variant, ByVal AddArrayValue As Variant)
    Call Assert(IsArray(ArrayValue), "Error:ArrayAdd:ArrayValue is not Array.")
    Call Assert(IsArray(AddArrayValue), "Error:ArrayAddArray:AddArrayValue is not Array.")

    Dim ArrayValue_Count As Long
    ArrayValue_Count = ArrayCount(ArrayValue)
    ReDim Preserve ArrayValue(ArrayValue_Count + ArrayCount(AddArrayValue) - 1)
    Dim I As Long
    For I = 0 To ArrayCount(AddArrayValue) - 1
        Call SetValue(ArrayValue(ArrayValue_Count + I), AddArrayValue(I))
    Next
End Sub

Private Sub testArrayAddArray()
    Dim A1()
    A1 = Array("A", "B", "C")
    Dim A2()
    A2 = Array("D", "E")

    Call ArrayAddArray(A1, A2)
    Call Check(5, ArrayCount(A1))
    Call Check("D", A1(3))
    Call Check("E", A1(4))

    '空配列に対しても使える
    Dim B1()
    Dim B2()
    B2 = Array("1", "2")
    Call ArrayAddArray(B1, B2)
    Call Check(2, ArrayCount(B1))
    Call Check("1", B1(0))
    Call Check("2", B1(1))

    
End Sub


'----------------------------------------
'・配列の要素を挿入する
'----------------------------------------
'   ・  オブジェクト値にも対応
'   ・  LBound(Array)=0でなくても対応。
'----------------------------------------
Sub ArrayInsert(ByRef ArrayValue As Variant, _
ByVal Index As Long, ByVal Value As Variant)
    Call Assert(IsArray(ArrayValue), "Error:ArrayInsert:ArrayValue is not Array.")
    Call Assert(InRange(LBound(ArrayValue), Index, UBound(ArrayValue)), _
        "Error:ArrayInsert:Index Range Over.")

    ReDim Preserve ArrayValue(LBound(ArrayValue) To UBound(ArrayValue) + 1)
    Dim I As Long
    For I = UBound(ArrayValue) To Index + 1 Step -1
        Call SetValue(ArrayValue(I), ArrayValue(I - 1))
    Next
    Call SetValue(ArrayValue(Index), Value)
End Sub

Private Sub testArrayInsert()
    Dim A
    A = Array("A", "B", "C")

    Call Check("B", A(1))
    Call Check(3, ArrayCount(A))
    Call ArrayInsert(A, 1, "1")
    Call Check(4, ArrayCount(A))
    Call Check("1", A(1))

    Dim B()
    ReDim B(2)
    Set B(0) = CreateObject("VBScript.RegExp")
    Set B(1) = Shell
    Set B(2) = CreateObject("ADODB.Stream")
    Call Check(Shell.CurrentDirectory, B(1).CurrentDirectory)
    Call ArrayInsert(B, 1, fso)
    Call Check("test.txt", B(1).GetFileName("C:\temp\test.txt"))
End Sub

'----------------------------------------
'・配列の要素を削除する
'----------------------------------------
'   ・  LBound(Array)=0でなくても対応。
'   ・  オブジェクト値にも対応
'----------------------------------------
Sub ArrayDelete(ArrayValue As Variant, Index As Long)
    Call Assert(IsArray(ArrayValue), "Error:ArrayDelete:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 1, "Error:ArrayDelete:Array Dimension != 1 .")
    Call Assert(ArrayCount(ArrayValue) <> 0, "Error:ArrayDelete:ArrayCount <> 0 .")
    Call Assert(InRange(LBound(ArrayValue), Index, UBound(ArrayValue)), _
        "Error:ArrayDelete:Index Range Over.")

    Dim I As Long
    For I = Index + 1 To UBound(ArrayValue)
        Call SetValue(ArrayValue(I - 1), ArrayValue(I))
    Next

    If LBound(ArrayValue) = UBound(ArrayValue) Then
        Erase ArrayValue
        '配列の初期化はEraseを使う
    Else
        ReDim Preserve ArrayValue(LBound(ArrayValue) To UBound(ArrayValue) - 1)
    End If
End Sub

Private Sub testArrayDelete()
    Dim A
    A = Array("A", "B", "C")

    Call Check("B", A(1))
    Call ArrayDelete(A, 1)
    Call Check(2, ArrayCount(A))
    Call Check("C", A(1))

    Dim B()
    ReDim B(2)
    Set B(0) = CreateObject("VBScript.RegExp")
    Set B(1) = Shell
    Set B(2) = fso
    Call Check(Shell.CurrentDirectory, B(1).CurrentDirectory)
    Call ArrayDelete(B, 1)
    Call Check("test.txt", B(1).GetFileName("C:\temp\test.txt"))
End Sub

'----------------------------------------
'・配列関数のテスト
'----------------------------------------
Private Sub testArrayFunctions()
  Dim A As Variant
  A = Array("A", "B", "C")
  Call Check(3, ArrayCount(A))
  Call Check("A B C", ArrayToString(A, " "))
  ArrayAdd A, "D"
  Call Check(4, ArrayCount(A))
  Call Check("A B C D", ArrayToString(A, " "))

  ArrayDelete A, 0
  Call Check("B C D", ArrayToString(A, " "))
  ArrayDelete A, 2
  Call Check("B C", ArrayToString(A, " "))

  A = Array("A", "B", "C")
  ArrayDelete A, 1
  Call Check("A C", ArrayToString(A, " "))

  ArrayInsert A, 1, "B"
  Call Check("A B C", ArrayToString(A, " "))
  ArrayInsert A, 0, "1"
  Call Check("1 A B C", ArrayToString(A, " "))
  ArrayInsert A, 3, "2"
  Call Check("1 A B 2 C", ArrayToString(A, " "))

  '要素なし配列
  Dim B() As String
  Call Check(0, ArrayCount(B))
  ArrayAdd B, "123"
  Call Check(1, ArrayCount(B))
  Call Check("123", B(0))
  ArrayAdd B, "789"
  ArrayInsert B, 1, "456"
  Call Check("123 456 789", ArrayToString(B, " "))
  ArrayDelete B, 0
  Call Check("456 789", ArrayToString(B, " "))

  'LBound(Array)=0ではない配列
  Dim C() As String
  ReDim C(4 To 6)
  C(4) = "A"
  C(5) = "B"
  C(6) = "C"
  Call Check(3, ArrayCount(C))
  Call Check("A B C", ArrayToString(C, " "))
  Call Check(3, ArrayCount(C))
  ArrayInsert C, 4, "Z"
  ArrayInsert C, 7, "B2"
  Call Check("Z A B B2 C", ArrayToString(C, " "))
'  ArrayAdd C, "D"
  ArrayDelete C, 4
  ArrayDelete C, 6
  Call Check("A B C", ArrayToString(C, " "))

  ArrayDelete C, 5
  Call Check("A C", ArrayToString(C, " "))
  ArrayDelete C, 4
  Call Check("C", ArrayToString(C, " "))

  ArrayDelete C, 4
  Call Check("", ArrayToString(C, " "))
  Call Check(0, ArrayCount(C))

End Sub


'----------------------------------------
'・配列内の値を検索してIndexを返す
'----------------------------------------
'   ・  LBound(Array)=0でなくても対応。
'   ・  大小文字比較対応
'   ・  完全一致/部分一致/ワイルドカード/正規表現対応
'----------------------------------------
Public Function ArrayIndexOf(ByRef ArrayValue As Variant, ByVal Value As Variant, _
Optional StartIndex As Long = -1, _
Optional CaseCompare As CaseCompare = CaseSensitive, _
Optional MatchType As MatchType = FullMatch) As Long

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Dim Result As Long: Result = -1

    Do
        If ArrayDimension(ArrayValue) <> 1 Then Exit Do
        If ArrayCount(ArrayValue) = 0 Then Exit Do
        If (StartIndex <> -1) _
        And ((StartIndex < LBound(ArrayValue)) _
                And (UBound(ArrayValue) < StartIndex)) Then Exit Do
        '↑範囲エラーの場合でもResult=-1を返すだけでエラーにはしない

        If StartIndex = -1 Then
            StartIndex = LBound(ArrayValue)
        End If

        Dim I As Long
        Select Case CaseCompare
        Case CaseSensitive
            Select Case MatchType
            Case FullMatch
                For I = StartIndex To UBound(ArrayValue)
                    If ArrayValue(I) = Value Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case PartMatch
                For I = StartIndex To UBound(ArrayValue)
                    If IsIncludeStr(ArrayValue(I), Value) Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case WildCardValue
                For I = StartIndex To UBound(ArrayValue)
                    If ArrayValue(I) Like Value Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case WildCardArray
                For I = StartIndex To UBound(ArrayValue)
                    If Value Like ArrayValue(I) Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case RegExpValue
                For I = StartIndex To UBound(ArrayValue)
                    If MatchRegExp(ArrayValue(I), Value, CaseSensitive) Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case RegExpArray
                For I = StartIndex To UBound(ArrayValue)
                    If MatchRegExp(Value, ArrayValue(I), CaseSensitive) Then
                        Result = I
                        Exit Do
                    End If
                Next
            End Select
        Case IgnoreCase
            Select Case MatchType
            Case FullMatch
                For I = StartIndex To UBound(ArrayValue)
                    If UCase(ArrayValue(I)) = UCase(Value) Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case PartMatch
                For I = StartIndex To UBound(ArrayValue)
                    If IsIncludeStr(UCase(ArrayValue(I)), UCase(Value)) Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case WildCardValue
                For I = StartIndex To UBound(ArrayValue)
                    If UCase(ArrayValue(I)) Like UCase(Value) Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case WildCardArray
                For I = StartIndex To UBound(ArrayValue)
                    If UCase(Value) Like UCase(ArrayValue(I)) Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case RegExpValue
                For I = StartIndex To UBound(ArrayValue)
                    If MatchRegExp(ArrayValue(I), Value, IgnoreCase) Then
                        Result = I
                        Exit Do
                    End If
                Next
            Case RegExpArray
                For I = StartIndex To UBound(ArrayValue)
                    If MatchRegExp(Value, ArrayValue(I), IgnoreCase) Then
                        Result = I
                        Exit Do
                    End If
                Next
            End Select
        End Select

    Loop While False
    ArrayIndexOf = Result
End Function

Sub testArrayIndexOf()
    'FullMatch
    Dim A As Variant
    A = Array("B", "C", "D")
    Call Check(0, ArrayIndexOf(A, "B"))
    Call Check(1, ArrayIndexOf(A, "C"))
    Call Check(2, ArrayIndexOf(A, "D"))
    Call Check(-1, ArrayIndexOf(A, "E"))

    Call Check(0, ArrayIndexOf(A, "B", 0))
    Call Check(1, ArrayIndexOf(A, "C", 1))
    Call Check(2, ArrayIndexOf(A, "D", 2))
    Call Check(-1, ArrayIndexOf(A, "B", 1))
    Call Check(-1, ArrayIndexOf(A, "C", 2))
    Call Check(2, ArrayIndexOf(A, "D", 2))
    Call Check(-1, ArrayIndexOf(A, "D", 3))

    'PartMatch IgnoreCase
    A = Array("ABC", "DEF", "123")
    Call Check(1, ArrayIndexOf(A, "DE", , CaseSensitive, PartMatch))
    Call Check(-1, ArrayIndexOf(A, "de", , CaseSensitive, PartMatch))
    Call Check(1, ArrayIndexOf(A, "de", , IgnoreCase, PartMatch))

    'Like WildCard Value
    A = Array("B", "C", "D")
    Call Check(0, ArrayIndexOf(A, "B", , , WildCardValue))
    Call Check(1, ArrayIndexOf(A, "C", , , WildCardValue))
    Call Check(2, ArrayIndexOf(A, "D", , , WildCardValue))
    Call Check(-1, ArrayIndexOf(A, "E", , , WildCardValue))

    Call Check(0, ArrayIndexOf(A, "B", 0, , WildCardValue))
    Call Check(1, ArrayIndexOf(A, "C", 1, , WildCardValue))
    Call Check(2, ArrayIndexOf(A, "D", 2, , WildCardValue))
    Call Check(-1, ArrayIndexOf(A, "B", 1, , WildCardValue))
    Call Check(-1, ArrayIndexOf(A, "C", 2, , WildCardValue))
    Call Check(-1, ArrayIndexOf(A, "D", 3, , WildCardValue))

    A = Array("ABC", "DEF", "123")
    Call Check(0, ArrayIndexOf(A, "A*", , , WildCardValue))
    Call Check(1, ArrayIndexOf(A, "D*", , , WildCardValue))
    Call Check(2, ArrayIndexOf(A, "1?3", , , WildCardValue))
    Call Check(-1, ArrayIndexOf(A, "A?B", , , WildCardValue))

    Call Check(0, ArrayIndexOf(A, "*C", 0, , WildCardValue))
    Call Check(1, ArrayIndexOf(A, "?E?", 1, , WildCardValue))
    Call Check(2, ArrayIndexOf(A, "?23", 2, , WildCardValue))
    Call Check(-1, ArrayIndexOf(A, "*C", 1, , WildCardValue))
    Call Check(-1, ArrayIndexOf(A, "?E?", 2, , WildCardValue))
    Call Check(-1, ArrayIndexOf(A, "?23", 3, , WildCardValue))

    'Like WildCard Value IgnoreCase
    Call Check(-1, ArrayIndexOf(A, "?e?", 1, , WildCardValue))
    Call Check(1, ArrayIndexOf(A, "?e?", 1, IgnoreCase, WildCardValue))

    'Like WildCard Value IgnoreCase 全角
    A = Array("ＡＢＣ", "ＤＥＦ", "１２３")
    Call Check(-1, ArrayIndexOf(A, "?ｅ?", 1, , WildCardValue))
    Call Check(1, ArrayIndexOf(A, "?ｅ?", 1, IgnoreCase, WildCardValue))

    'RegExp Value
    A = Array("ABC", "DEF", "123")
    Call Check(0, ArrayIndexOf(A, ".*C", 0, , RegExpValue))

    'RegExp Value IgnoreCase
    Call Check(-1, ArrayIndexOf(A, ".*c", 0, , RegExpValue))
    Call Check(0, ArrayIndexOf(A, ".*C", 0, , RegExpValue))

End Sub

'----------------------------------------
'・配列内の値存在チェック
'----------------------------------------
Public Function ArrayExists(ByRef ArrayValue As Variant, _
ByVal Value As Variant, _
Optional CaseCompare As CaseCompare = CaseSensitive, _
Optional MatchType As MatchType = FullMatch) As Boolean

    ArrayExists = Not (ArrayIndexOf(ArrayValue, Value, , CaseCompare, MatchType) = -1)
End Function

'----------------------------------------
'・配列内の値を検索してユニーク(同一値がない)かどうかを判断する
'----------------------------------------
Public Function ArrayIsUnique(ByRef ArrayValue As Variant) As Boolean
    Call Assert(IsArray(ArrayValue), "Error:ArrayIsUnique:ArrayValue is not array")
    Call Assert(ArrayDimension(ArrayValue) = 1, _
        "Error:ArrayIsUnique:ArrayValue Dimension is miss")

    Dim Result As Boolean: Result = True
    Do
        If OrValue(ArrayCount(ArrayValue), 0, 1) Then Exit Do

        Dim I As Long
        Dim J As Long
        For I = LBound(ArrayValue) To UBound(ArrayValue) - 1
            For J = I + 1 To UBound(ArrayValue)
                If ArrayValue(I) = ArrayValue(J) Then
                    Result = False
                    Exit Do
                End If
            Next
        Next
        Loop While False
    ArrayIsUnique = Result
End Function

Sub testArrayIsUnique()
    Dim A As Variant
    A = Array("B", "C", "D", "A", "B", "C")
    Call Check(False, ArrayIsUnique(A))

    A = Array("1", "2", "3", "A", "B", "C")
    Call Check(True, ArrayIsUnique(A))
End Sub


'----------------------------------------
'◇配列応用操作
'----------------------------------------

'----------------------------------------
'・配列内の値を検索して同一値を削除
'----------------------------------------
'   ・  LBound(Array)=0でなくても対応。
'       重複があればTrue/なければFalse
Public Function ArrayDeleteSameItem(ByRef ArrayValue As Variant, _
Optional StartIndex As Long = -1) As Boolean
    Dim Result As Boolean: Result = False
    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    If StartIndex <> -1 Then
        Call Assert(((StartIndex < LBound(ArrayValue)) _
                And (UBound(ArrayValue) < StartIndex) = False), "Error:ArrayDeleteSameItem:Range Over")
        '↑範囲エラーの場合もある。
    End If

    Do
        If ArrayDimension(ArrayValue) <> 1 Then Exit Do
        If ArrayCount(ArrayValue) = 0 Then Exit Do

        If StartIndex = -1 Then
            StartIndex = LBound(ArrayValue)
        End If

        Dim I As Long
        For I = UBound(ArrayValue) To StartIndex Step -1
            If ArrayIndexOf(ArrayValue, ArrayValue(I)) <> I Then
                Call ArrayDelete(ArrayValue, I)
                Result = True
            End If
        Next

    Loop While False
    ArrayDeleteSameItem = Result
End Function

Sub testArrayDeleteSameItem()
  Dim A As Variant
  A = Array("B", "C", "D", "A", "B", "C")

  Call Check("B C D A B C", ArrayToString(A, " "))

  Call ArrayDeleteSameItem(A)

  Call Check("B C D A", ArrayToString(A, " "))
End Sub

'----------------------------------------
'・配列の要素タイプを求める
'----------------------------------------
'   ・  LBound=0 でも 1 でも対応する。
'----------------------------------------
Public Function CheckArrayVarType(ByVal ArrayValue As Variant, TypeValue As VbVarType) As Boolean
    Dim Result As Boolean: Result = True
    Call Assert(IsArray(ArrayValue), "Error:IsArray")
    Dim I As Long
    For I = LBound(ArrayValue) To UBound(ArrayValue)
        If VarType(ArrayValue(I)) <> TypeValue Then
            Result = False
            Exit For
        End If
    Next
    CheckArrayVarType = Result
End Function

'----------------------------------------
'・文字列配列かどうか
'----------------------------------------
Public Function IsStrArray(ByVal ArrayValue As Variant) As Boolean
    IsStrArray = CheckArrayVarType(ArrayValue, vbString)
End Function

'----------------------------------------
'・配列を文字列にして出力する関数
'----------------------------------------
'   ・  要素がなくても対応。
'   ・  LBound(Array)=0でなくても対応。
'----------------------------------------
Public Function ArrayToString(ArrayValue As Variant, Delimiter As String) As String
    Call Assert(IsArray(ArrayValue), "配列ではありません")

    Dim Result As String
    Result = ""
    Do
        If ArrayCount(ArrayValue) = 0 Then Exit Do

        Call Assert(ArrayDimension(ArrayValue) = 1, "1次元配列ではありません")
        Dim I As Long
        For I = LBound(ArrayValue) To UBound(ArrayValue)
          Result = Result + ArrayValue(I) + Delimiter
        Next
    Loop While False
    Result = ExcludeLastStr(Result, Delimiter)
    ArrayToString = Result
End Function

'----------------------------------------
'・パラメータ配列を文字列配列にして返す関数
'----------------------------------------
Public Function ArrayStr(ParamArray Values()) As String()
    'パラメータ配列をString配列に代入している
    Dim Result() As String
    If 0 <= UBound(Values) Then
        ReDim Result(UBound(Values))
        Dim I As Long
        For I = 0 To UBound(Values) - LBound(Values)
            Result(I) = CStr(Values(I))
        Next
    End If
    ArrayStr = Result
End Function

'----------------------------------------
'・配列を文字列配列にして返す関数
'----------------------------------------
Public Function ArrayToStrArray(Values()) As String()
    Dim Result() As String
    If 0 <= UBound(Values) Then
        ReDim Result(UBound(Values))
        Dim I As Long
        For I = 0 To UBound(Values) - LBound(Values)
            Result(I) = CStr(Values(I))
        Next
    End If
    ArrayToStrArray = Result
End Function

'----------------------------------------
'・パラメータ配列をLong配列にして返す関数
'----------------------------------------
Public Function ArrayLong(ParamArray Values()) As Long()
    'パラメータ配列をLong配列に代入している
    Dim Result() As Long
    If 0 <= UBound(Values) Then
        ReDim Result(UBound(Values))
        Dim I As Long
        For I = 0 To UBound(Values) - LBound(Values)
            Result(I) = CLng(Values(I))
        Next
    End If
    ArrayLong = Result
End Function


'----------------------------------------
'◇配列ソート
'----------------------------------------

'----------------------------------------
'・クイックソート
'----------------------------------------
'   ・  IndexMin/IndexMaxを指定すると
'       指定範囲内の値をソートする
'----------------------------------------
Public Sub ArraySortQuick(ByRef ArrayValue As Variant, _
Optional ByVal SortOrder As SortOrder = SortOrder.Ascending, _
Optional ByVal IndexMin As Long = -1, Optional ByVal IndexMax As Long = -1)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 1, "Error:ArrayValue Dimension is miss")

    Call Assert(IndexMin <= IndexMax, "Error:IndexMin < IndexMax")
    Call Assert(InRange(-1, IndexMin, ArrayCount(ArrayValue) - 1), "Error:IndexMin Range is miss")
    Call Assert(InRange(-1, IndexMax, ArrayCount(ArrayValue) - 1), "Error:IndexMax Range is miss")

    '1以下ならソート不可能なのでExitする
    If ArrayCount(ArrayValue) <= 1 Then Exit Sub

    IndexMin = IIf(IndexMin = -1, 0, IndexMin)
    IndexMax = IIf(IndexMax = -1, ArrayCount(ArrayValue) - 1, IndexMax)

    'IndexMin=IndexMaxならソート不可能なのでExit
    If IndexMin = IndexMax Then Exit Sub

    Call ArraySortQuickBase(ArrayValue, SortOrder, IndexMin, IndexMax)
End Sub

'クイックソートのベース関数、再起呼び出しされる
Sub ArraySortQuickBase(ByRef ArrayValue As Variant, _
ByVal SortOrder As SortOrder, _
ByVal IndexMin As Long, ByVal IndexMax As Long)

    Dim IndexCenter As Long
    Dim Index1 As Long
    Dim Index2 As Long
    Dim Value1 As String
    Dim Value2 As String

    If IndexMax <= IndexMin Then Exit Sub

    IndexCenter = (IndexMin + IndexMax) \ 2

    '中央値をバッファ
    Value1 = ArrayValue(IndexCenter)
    '中央値に開始位置要素を代入
    ArrayValue(IndexCenter) = ArrayValue(IndexMin)

    Index2 = IndexMin

    Index1 = IndexMin + 1

    Select Case SortOrder
    Case Ascending
        Do While Index1 <= IndexMax
            If ArrayValue(Index1) < Value1 Then
                Index2 = Index2 + 1

                Value2 = ArrayValue(Index2)
                ArrayValue(Index2) = ArrayValue(Index1)
                ArrayValue(Index1) = Value2

            End If
            Index1 = Index1 + 1
        Loop
    Case Descending
        Do While Index1 <= IndexMax
            If Value1 < ArrayValue(Index1) Then
                Index2 = Index2 + 1

                Value2 = ArrayValue(Index2)
                ArrayValue(Index2) = ArrayValue(Index1)
                ArrayValue(Index1) = Value2

            End If
            Index1 = Index1 + 1
        Loop
    Case Else
        Call Assert(False, "Error:ArraySortQuickBase:SortOrder is miss.")
    End Select

    ArrayValue(IndexMin) = ArrayValue(Index2)
    ArrayValue(Index2) = Value1

    ' 分割前半を再帰呼び出しでSORT
    Call ArraySortQuickBase(ArrayValue, SortOrder, IndexMin, Index2 - 1)

    ' 分割後半を再帰呼び出しでSORT
    Call ArraySortQuickBase(ArrayValue, SortOrder, Index2 + 1, IndexMax)
End Sub

Sub testArrayQuickSort()
    Dim Array1(5) As Variant
    Array1(0) = "105"
    Array1(1) = "101"
    Array1(2) = "103"
    Array1(3) = "102"
    Array1(4) = "104"
    Array1(5) = "100"

    Call Check(Array1(0), "105")
    Call Check(Array1(1), "101")
    Call Check(Array1(2), "103")
    Call Check(Array1(3), "102")
    Call Check(Array1(4), "104")
    Call Check(Array1(5), "100")

    'Ascending
    Call ArraySortQuick(Array1, SortOrder.Ascending, 2, 4)
    Call Check(Array1(0), "105")
    Call Check(Array1(1), "101")
    Call Check(Array1(2), "102")
    Call Check(Array1(3), "103")
    Call Check(Array1(4), "104")
    Call Check(Array1(5), "100")

    Call ArraySortQuick(Array1, SortOrder.Ascending, 0, 2)
    Call Check(Array1(0), "101")
    Call Check(Array1(1), "102")
    Call Check(Array1(2), "105")
    Call Check(Array1(3), "103")
    Call Check(Array1(4), "104")
    Call Check(Array1(5), "100")

    Call ArraySortQuick(Array1)
    Call Check(Array1(0), "100")
    Call Check(Array1(1), "101")
    Call Check(Array1(2), "102")
    Call Check(Array1(3), "103")
    Call Check(Array1(4), "104")
    Call Check(Array1(5), "105")

    'Descending
    Array1(0) = "105"
    Array1(1) = "101"
    Array1(2) = "103"
    Array1(3) = "102"
    Array1(4) = "104"
    Array1(5) = "100"

    Call ArraySortQuick(Array1, SortOrder.Descending, 2, 4)
    Call Check(Array1(0), "105")
    Call Check(Array1(1), "101")
    Call Check(Array1(2), "104")
    Call Check(Array1(3), "103")
    Call Check(Array1(4), "102")
    Call Check(Array1(5), "100")

    Call ArraySortQuick(Array1, SortOrder.Descending, 0, 2)
    Call Check(Array1(0), "105")
    Call Check(Array1(1), "104")
    Call Check(Array1(2), "101")
    Call Check(Array1(3), "103")
    Call Check(Array1(4), "102")
    Call Check(Array1(5), "100")

    Call ArraySortQuick(Array1, SortOrder.Descending)
    Call Check(Array1(0), "105")
    Call Check(Array1(1), "104")
    Call Check(Array1(2), "103")
    Call Check(Array1(3), "102")
    Call Check(Array1(4), "101")
    Call Check(Array1(5), "100")
End Sub


'----------------------------------------
'・文字列長ソート
'----------------------------------------
Public Sub ArraySortStrLength(ByRef ArrayValue As Variant, _
Optional ByVal SortOrder As SortOrder = SortOrder.Ascending)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 1, "Error:ArrayValue Dimension is miss")
    Call Assert(OrValue(SortOrder, Ascending, Descending), "Error:SortOrder is miss.")

    Dim DigitArrayValue As Long
    DigitArrayValue = Len(CStr(ArrayCount(ArrayValue) - 1))
    Dim DigitStrLength As Long
    DigitStrLength = 0
    Dim MaxLength As Long
    MaxLength = 0

    Dim I As Long
    For I = 0 To ArrayCount(ArrayValue) - 1
        MaxLength = MaxValue(MaxLength, Len(ArrayValue(I)))
    Next
    DigitStrLength = Len(CStr(MaxLength))

    Select Case SortOrder
    Case Ascending
        For I = 0 To ArrayCount(ArrayValue) - 1
            ArrayValue(I) = _
                LongToStrDigitZero(Len(ArrayValue(I)), DigitStrLength) + _
                LongToStrDigitZero(I, DigitArrayValue) + _
                ArrayValue(I)
        Next
    Case Descending
        For I = 0 To ArrayCount(ArrayValue) - 1
            ArrayValue(I) = _
                LongToStrDigitZero(MaxLength - Len(ArrayValue(I)), DigitStrLength) + _
                LongToStrDigitZero(I, DigitArrayValue) + _
                ArrayValue(I)
        Next
    End Select
    Call ArraySortQuick(ArrayValue)

    For I = 0 To ArrayCount(ArrayValue) - 1
        ArrayValue(I) = _
            Mid$(ArrayValue(I), _
                DigitStrLength + DigitArrayValue + 1)
    Next
End Sub

Sub testArraySortStrLength()

    Dim Array1(5) As Variant
    Array1(0) = "1"
    Array1(1) = "12"
    Array1(2) = "1234"
    Array1(3) = "123"
    Array1(4) = "abc"
    Array1(5) = "a"

    Call ArraySortStrLength(Array1, Ascending)

    Call Check(Array1(0), "1")
    Call Check(Array1(1), "a")
    Call Check(Array1(2), "12")
    Call Check(Array1(3), "123")
    Call Check(Array1(4), "abc")
    Call Check(Array1(5), "1234")

    Call ArraySortStrLength(Array1, SortOrder.Descending)

    Call Check(Array1(0), "1234")
    Call Check(Array1(1), "123")
    Call Check(Array1(2), "abc")
    Call Check(Array1(3), "12")
    Call Check(Array1(4), "1")
    Call Check(Array1(5), "a")
End Sub


'----------------------------------------
'・独自並び順ソート
'----------------------------------------
'   ・  ソート指定配列の文字列に一致する順番に
'       並び替えをするソート
'   ・  s/m/l/xl/xxlとかそういう並び指定を行う
'----------------------------------------
Public Sub ArraySortCustomOrder(ByRef ArrayValue As Variant, _
ByRef OrderArrayWildCard() As String, _
Optional CaseCompare As CaseCompare = CaseCompare.IgnoreCase, _
Optional NoOrderValuePriority As Boolean = False)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 1)

    Dim DigitArrayValue As Long
    DigitArrayValue = Len(CStr(ArrayCount(ArrayValue) - 1))
    Dim DigitOrderArray As Long
    DigitOrderArray = Len(CStr(ArrayCount(OrderArrayWildCard) + 1))

    Dim I As Long

    For I = 0 To ArrayCount(ArrayValue) - 1
        Dim OrderArrayIndex As Long
        OrderArrayIndex = _
            ArrayIndexOf(OrderArrayWildCard, ArrayValue(I), , CaseCompare, WildCardArray)
        If OrderArrayIndex = -1 Then
            If NoOrderValuePriority = False Then
                ArrayValue(I) = _
                    LongToStrDigitZero(ArrayCount(OrderArrayWildCard) + 1, DigitOrderArray) + _
                    LongToStrDigitZero(I, DigitArrayValue) + _
                    ArrayValue(I)
            Else
                ArrayValue(I) = _
                    LongToStrDigitZero(0, DigitOrderArray) + _
                    LongToStrDigitZero(I, DigitArrayValue) + _
                    ArrayValue(I)
            End If
        Else
            ArrayValue(I) = _
                LongToStrDigitZero(OrderArrayIndex + 1, DigitOrderArray) + _
                LongToStrDigitZero(I, DigitArrayValue) + _
                ArrayValue(I)
        End If
    Next
    Call ArraySortQuick(ArrayValue)

    For I = 0 To ArrayCount(ArrayValue) - 1
        ArrayValue(I) = _
            Mid$(ArrayValue(I), _
                DigitOrderArray + DigitArrayValue + 1)
    Next
End Sub

Public Sub testArraySortCustomOrder()
    Dim Array1() As String

    Array1 = ArrayStr("b", "a", "s", "ss", "xl", "ll", "m")

    Call ArraySortCustomOrder(Array1, ArrayStr("ss*", "s*", "m*", "l*", "ll*", "xl*"))

    Call Check(Array1(0), "ss")
    Call Check(Array1(1), "s")
    Call Check(Array1(2), "m")
    Call Check(Array1(3), "ll")
    Call Check(Array1(4), "xl")
    Call Check(Array1(5), "b")
    Call Check(Array1(6), "a")

    Array1 = ArrayStr("Bサイズ", "Aサイズ", _
        "Sサイズ", "SSサイズ", "XLサイズ", "LLサイズ", "Mサイズ")

    Call ArraySortCustomOrder(Array1, ArrayStr("ss*", "s*", "m*", "l*", "ll*", "xl*"))

    Call Check(Array1(0), "SSサイズ")
    Call Check(Array1(1), "Sサイズ")
    Call Check(Array1(2), "Mサイズ")
    Call Check(Array1(3), "LLサイズ")
    Call Check(Array1(4), "XLサイズ")
    Call Check(Array1(5), "Bサイズ")
    Call Check(Array1(6), "Aサイズ")

    Array1 = ArrayStr("Bサイズ", "Aサイズ", _
        "Sサイズ", "SSサイズ", "XLサイズ", "LLサイズ", "Mサイズ")

    Call ArraySortCustomOrder(Array1, ArrayStr("ss*", "s*", "m*", "l*", "ll*", "xl*"), , True)

    Call Check(Array1(0), "Bサイズ")
    Call Check(Array1(1), "Aサイズ")
    Call Check(Array1(2), "SSサイズ")
    Call Check(Array1(3), "Sサイズ")
    Call Check(Array1(4), "Mサイズ")
    Call Check(Array1(5), "LLサイズ")
    Call Check(Array1(6), "XLサイズ")

End Sub


'----------------------------------------
'・配列を逆順にする
'----------------------------------------
'   ・  IndexMin/IndexMaxを指定すると
'       指定範囲内の値を逆順にする
'----------------------------------------
Public Sub ArrayReverse(ByRef ArrayValue As Variant, _
Optional ByVal IndexMin As Long = -1, Optional ByVal IndexMax As Long = -1)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 1)

    'IndexMin/Maxの指定が変ならエラーにする
    Call Assert(IndexMin <= IndexMax, "Error:IndexMin < IndexMax")
    Call Assert(InRange(-1, IndexMin, ArrayCount(ArrayValue) - 1), _
        "Error:ArrayReverse:IndexMin Range is miss.")
    Call Assert(InRange(-1, IndexMax, ArrayCount(ArrayValue) - 1), _
        "Error:ArrayReverse:IndexMax Range is miss.")

    '1以下ならソート不可能なのでExitする
    If ArrayCount(ArrayValue) <= 1 Then Exit Sub

    IndexMin = IIf(IndexMin = -1, 0, IndexMin)
    IndexMax = IIf(IndexMax = -1, ArrayCount(ArrayValue) - 1, IndexMax)

    'IndexMin=IndexMaxならソート不可能なのでExit
    If IndexMin = IndexMax Then Exit Sub

    Dim SortDataCount As Long
    SortDataCount = IndexMax - IndexMin + 1
    Dim DigitSortDataCount As Long
    DigitSortDataCount = Len(SortDataCount)

    Dim I As Long
    For I = IndexMin To IndexMax
        ArrayValue(I) = LongToStrDigitZero(I, DigitSortDataCount) + ArrayValue(I)
    Next
    Call ArraySortQuick(ArrayValue, Descending, IndexMin, IndexMax)
    For I = IndexMin To IndexMax
        ArrayValue(I) = Mid$(ArrayValue(I), _
            DigitSortDataCount + 1)
    Next

End Sub

Public Sub testArrayReverse()
    Dim Array1(5) As Variant
    Array1(0) = "105"
    Array1(1) = "101"
    Array1(2) = "103"
    Array1(3) = "102"
    Array1(4) = "104"
    Array1(5) = "100"

    Call Check(Array1(0), "105")
    Call Check(Array1(1), "101")
    Call Check(Array1(2), "103")
    Call Check(Array1(3), "102")
    Call Check(Array1(4), "104")
    Call Check(Array1(5), "100")

    Call ArrayReverse(Array1, 2, 4)
    Call Check(Array1(0), "105")
    Call Check(Array1(1), "101")
    Call Check(Array1(2), "104")
    Call Check(Array1(3), "102")
    Call Check(Array1(4), "103")
    Call Check(Array1(5), "100")

    Call ArrayReverse(Array1, 0, 2)
    Call Check(Array1(0), "104")
    Call Check(Array1(1), "101")
    Call Check(Array1(2), "105")
    Call Check(Array1(3), "102")
    Call Check(Array1(4), "103")
    Call Check(Array1(5), "100")

    Call ArrayReverse(Array1)
    Call Check(Array1(0), "100")
    Call Check(Array1(1), "103")
    Call Check(Array1(2), "102")
    Call Check(Array1(3), "105")
    Call Check(Array1(4), "101")
    Call Check(Array1(5), "104")

End Sub

'----------------------------------------
'◆2次元配列
'----------------------------------------
'   ・  2次元配列のDimension
'       Dimension=1:列
'       Dimension=2:行
'   ・  動的配列は列固定、行可変になる
'----------------------------------------

'----------------------------------------
'・次元数を取得する
'----------------------------------------
'   ・  要素がない配列の場合は次元数は0として返される
'----------------------------------------

Public Function ArrayDimension(ByRef ArrayValue As Variant) As Long
    Dim Result As Long
    Result = 0

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")

    Dim TempData As Variant
    Dim I As Long
    I = 0
    On Error Resume Next
    Do While Err.Number = 0
        I = I + 1
        TempData = UBound(ArrayValue, I)
    Loop
    On Error GoTo 0
    Result = I - 1

    ArrayDimension = Result
End Function

Public Sub testArrayDimension()

    Dim A() As String
    Call Check(0, ArrayDimension(A))

    Dim B()
    B = Array("A", "B", "C")
    Call Check(1, ArrayDimension(B))

    Dim C() As String
    C = ArrayStr("A", "B", "C")
    Call Check(1, ArrayDimension(C))

    Dim D() As String
    ReDim Preserve D(3, 4)

    Call Check(2, ArrayDimension(D))
End Sub

'----------------------------------------
'・2次元配列の列数を取得する
'----------------------------------------
'   ・  0オリジン/1オリジン、両方対応
'----------------------------------------
Public Function Array2dColumnsStartIndex(ByRef ArrayValue As Variant) As Long
    Call Assert(IsArray(ArrayValue), "Error:Array2dColumnsStartIndex:ArrayValue is not Array")
    Dim Dimension As Long: Dimension = ArrayDimension(ArrayValue)
    Call Assert(OrValue(Dimension, 0, 2), "Error:Array2dColumnsStartIndex:ArrayValue Dimension is miss")

    Dim Result As Long

    Select Case Dimension
    Case 2
        Result = LBoundNoError(ArrayValue, 1)
    Case 0
        '未定義配列
        Result = 0
    End Select
    Array2dColumnsStartIndex = Result
End Function

Public Function Array2dColumnsEndIndex(ByRef ArrayValue As Variant) As Long
    Call Assert(IsArray(ArrayValue), "Error:Array2dColumnsEndIndex:ArrayValue is not Array")
    Dim Dimension As Long: Dimension = ArrayDimension(ArrayValue)
    Call Assert(OrValue(Dimension, 0, 2), "Error:Array2dColumnsEndIndex:ArrayValue Dimension is miss")

    Dim Result As Long

    Select Case Dimension
    Case 2
        Result = UBoundNoError(ArrayValue, 1)
    Case 0
        '未定義配列
        Result = 0
    End Select
    Array2dColumnsEndIndex = Result
End Function

Public Function Array2dColumnsCount(ByRef ArrayValue As Variant) As Long
    Call Assert(IsArray(ArrayValue), "Error:Array2dColumnsCount:ArrayValue is not Array")
    Dim Dimension As Long: Dimension = ArrayDimension(ArrayValue)
    Call Assert(OrValue(Dimension, 0, 2), "Error:Array2dColumnsCount:ArrayValue Dimension is miss")

    Dim Result As Long

    Select Case Dimension
    Case 2
        Result = Array2dColumnsEndIndex(ArrayValue) - Array2dColumnsStartIndex(ArrayValue) + 1
    Case 0
        '未定義配列
        Result = 0
    End Select
    Array2dColumnsCount = Result
End Function

'----------------------------------------
'・2次元配列の行数を取得する
'----------------------------------------
'   ・  0オリジン/1オリジン、両方対応
'----------------------------------------
Public Function Array2dRowsStartIndex(ByRef ArrayValue As Variant) As Long
    Call Assert(IsArray(ArrayValue), "Error:Array2dRowsStartIndex:ArrayValue is not Array")
    Dim Dimension As Long: Dimension = ArrayDimension(ArrayValue)
    Call Assert(OrValue(Dimension, 0, 2), "Error:Array2dRowsStartIndex:ArrayValue Dimension is miss")

    Dim Result As Long

    Select Case Dimension
    Case 2
        Result = LBoundNoError(ArrayValue, 2)
    Case 0
        '未定義配列
        Result = 0
    End Select
    Array2dRowsStartIndex = Result
End Function

Public Function Array2dRowsEndIndex(ByRef ArrayValue As Variant) As Long
    Call Assert(IsArray(ArrayValue), "Error:Array2dRowsEndIndex:ArrayValue is not Array")
    Dim Dimension As Long: Dimension = ArrayDimension(ArrayValue)
    Call Assert(OrValue(Dimension, 0, 2), "Error:Array2dRowsEndIndex:ArrayValue Dimension is miss")

    Dim Result As Long

    Select Case Dimension
    Case 2
        Result = UBoundNoError(ArrayValue, 2)
    Case 0
        '未定義配列
        Result = 0
    End Select
    Array2dRowsEndIndex = Result
End Function

Public Function Array2dRowsCount(ByRef ArrayValue As Variant) As Long
    Call Assert(IsArray(ArrayValue), "Error:Array2dRowsCount:ArrayValue is not Array.")
    Dim Dimension As Long: Dimension = ArrayDimension(ArrayValue)
    Call Assert(OrValue(Dimension, 0, 2), "Error:Array2dRowsCount:ArrayValue Dimension is miss")

    Dim Result As Long

    Select Case ArrayDimension(ArrayValue)
    Case 2
        Result = Array2dRowsEndIndex(ArrayValue) - Array2dRowsStartIndex(ArrayValue) + 1
    Case 0
        '未定義配列
        Result = 0
    End Select
    Array2dRowsCount = Result
End Function

'----------------------------------------
'・2次元配列の列数(変更できない)をセットする
'----------------------------------------
'   ・  初期状態からのセットになるので
'       すでにセットされた配列に対して実行するとエラーになる
'   ・  行要素は最低1つは必要になる
'----------------------------------------
Public Sub Array2dSetColumn(ByRef ArrayValue As Variant, _
ByVal ColumnCount As Long)
    ReDim Preserve ArrayValue(ColumnCount - 1, 0)
End Sub

Public Sub testArray2dSetColumn()
    Dim A() As String
    Call Array2dSetColumn(A, 5)

    Call Check(5, ArrayCount(A, 1))
    Call Check(1, ArrayCount(A, 2))

'    Call Array2DSetColumn(A, 4)
    '2回実行するとエラーになる

End Sub


'----------------------------------------
'・2次元配列の行を設定する
'----------------------------------------
'   ・  列数が一致した配列を設定して行の値をセットする
'   ・  オブジェクト値にも対応
'   ・  0オリジン/1オリジン、両方対応
'----------------------------------------
Public Sub Array2dSetRowValues(ByRef ArrayValue As Variant, _
ByVal RowIndex As Long, _
ByRef Values As Variant)
    Call Assert(IsArray(ArrayValue), "Error:Array2dSetRowValues:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dSetRowValues:ArrayValue is not Array2D.")
    Call Assert(UBound(Values) - LBound(Values) + 1 = Array2dColumnsCount(ArrayValue), _
        "Error:Array2dSetRowValues:Values Count is miss.")
    Call Assert(InRange(LBound(ArrayValue, 2), RowIndex, UBound(ArrayValue, 2)), _
        "Error:Array2dSetRowValues:RowIndex range over.")

    Dim I As Long
    Dim StartIndex As Long: StartIndex = Array2dColumnsStartIndex(ArrayValue)
    For I = StartIndex To Array2dColumnsEndIndex(ArrayValue)
        Call SetValue(ArrayValue(I, RowIndex), Values(I - StartIndex))
    Next
End Sub

'----------------------------------------
'・2次元配列の行を取得する
'----------------------------------------
'   ・  オブジェクト値にも対応
'   ・  0オリジン/1オリジン、両方対応
'----------------------------------------
Public Function Array2dGetRowValues(ByRef ArrayValue As Variant, _
ByVal RowIndex As Long) As String()
    Call Assert(IsArray(ArrayValue), "Error:Array2dSetRowValues:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dSetRowValues:ArrayValue is not Array2D.")
    Call Assert(InRange(LBound(ArrayValue, 2), RowIndex, UBound(ArrayValue, 2)), _
        "Error:Array2dSetRowValues:RowIndex range over.")

    Dim Result() As String
    Result = ArrayStr()
    ReDim Preserve Result(Array2dColumnsCount(ArrayValue) - 1)

    Dim I As Long
    Dim StartIndex As Long: StartIndex = Array2dColumnsStartIndex(ArrayValue)
    For I = StartIndex To Array2dColumnsEndIndex(ArrayValue)
        Result(I - StartIndex) = ArrayValue(I, RowIndex)
    Next
    Array2dGetRowValues = Result
End Function


'----------------------------------------
'・2次元配列の列を設定する
'----------------------------------------
'   ・  列数が一致した配列を設定して行の値をセットする
'   ・  オブジェクト値にも対応
'   ・  0オリジン/1オリジン、両方対応
'----------------------------------------
Public Sub Array2dSetColumnValues(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long, _
ByRef Values As Variant)
    Call Assert(IsArray(ArrayValue), "Error:Array2dSetColumnValues:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dSetColumnValues:ArrayValue is not Array2D.")
    Call Assert(UBound(Values) - LBound(Values) + 1 = Array2dRowsCount(ArrayValue), _
        "Error:Array2dSetColumnValues:Values Count is miss.")
    Call Assert(InRange(LBound(ArrayValue, 1), ColumnIndex, UBound(ArrayValue, 1)), _
        "Error:Array2dSetColumnValues:ColumnIndex range over.")

    Dim I As Long
    Dim StartIndex As Long: StartIndex = Array2dRowsStartIndex(ArrayValue)
    For I = StartIndex To Array2dRowsEndIndex(ArrayValue)
        Call SetValue(ArrayValue(ColumnIndex, I), Values(I - StartIndex))
    Next
End Sub

'----------------------------------------
'・2次元配列の列を取得する
'----------------------------------------
'   ・  オブジェクト値にも対応
'   ・  0オリジン/1オリジン、両方対応
'----------------------------------------
Public Function Array2dGetColumnValues(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long) As String()
    Call Assert(IsArray(ArrayValue), "Error:Array2dGetColumnValues:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dGetColumnValues:ArrayValue is not Array2D.")
    Call Assert(InRange(LBound(ArrayValue, 1), ColumnIndex, UBound(ArrayValue, 1)), _
        "Error:Array2dGetColumnValues:ColumnIndex range over.")

    Dim Result() As String
    Result = ArrayStr()
    ReDim Preserve Result(Array2dRowsCount(ArrayValue) - 1)
    Dim I As Long
    Dim StartIndex As Long: StartIndex = Array2dRowsStartIndex(ArrayValue)
    For I = StartIndex To Array2dRowsEndIndex(ArrayValue)
        Result(I - StartIndex) = ArrayValue(ColumnIndex, I)
    Next
    Array2dGetColumnValues = Result
End Function

'----------------------------------------
'・配列の要素を追加する
'----------------------------------------
'   ・  列数が一致した配列を設定して行の値を追加する
'   ・  オブジェクト値にも対応
'----------------------------------------
Public Sub Array2dAdd(ByRef ArrayValue As Variant, _
ByRef Values As Variant)
    Call Assert(IsArray(ArrayValue), "Error:Array2dAdd:ArrayValue is not Array")
    Call Assert(IsArray(Values), "Error:Array2dAdd:Values is not Array")
    Call Assert(ArrayDimension(Values) = 1, "Error:Array2dAdd:Values Dimension is not 1")

    Select Case ArrayDimension(ArrayValue)
    Case 2
        Call Assert(UBound(Values) - LBound(Values) + 1 = Array2dColumnsCount(ArrayValue), _
            "Error:Array2dAdd:Values Count is miss.")
        ReDim Preserve ArrayValue(Array2dColumnsCount(ArrayValue) - 1, Array2dRowsCount(ArrayValue))
        Call Array2dSetRowValues(ArrayValue, UBound(ArrayValue, 2), Values)
    Case 0
        '未定義配列の場合
        '列数をセットして値を指定する
        Call Array2dSetColumn(ArrayValue, ArrayCount(Values))
        Call Array2dSetRowValues(ArrayValue, 0, Values)
    Case Else
        Call Assert(False, "Error:Array2dSetRowValues:ArrayValue Dimension is miss")
    End Select

End Sub

Public Sub testArray2dAdd()
    Dim A() As String

    Call Check(0, ArrayCount(A, 1))
    Call Check(0, ArrayCount(A, 2))

    Call Array2dSetColumn(A, 3)
    Call Check(3, ArrayCount(A, 1))
    Call Check(1, ArrayCount(A, 2))

    Call Array2dSetRowValues(A, 0, Array("A", "B", "C"))
    Call Array2dAdd(A, Array("D", "E", "F"))
    Call Array2dAdd(A, Array("G", "H", "I"))
    Call Array2dAdd(A, Array("1", "2", "3"))

    Dim B() As String
    Call Array2dAdd(B, Array("A", "B", "C"))
    Call Array2dAdd(B, Array("D", "E", "F"))
    Call Array2dAdd(B, Array("G", "H", "I"))
    Call Array2dAdd(B, Array("1", "2", "3"))

End Sub


'----------------------------------------
'・配列の要素を挿入する
'----------------------------------------
'   ・  オブジェクト値にも対応
'----------------------------------------
Public Sub Array2dInsert(ByRef ArrayValue As Variant, _
ByVal RowIndex As Long, ByVal Values As Variant)
    Call Assert(IsArray(ArrayValue), "Error:Array2dInsert:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dInsert:ArrayValue is not Array2D.")
    Call Assert(UBound(Values) - LBound(Values) + 1 = Array2dColumnsCount(ArrayValue), _
        "Error:Array2dInsert:Values Count is miss.")
    Call Assert(InRange(LBound(ArrayValue, 2), RowIndex, UBound(ArrayValue, 2)), _
        "Error:Array2dInsert:RowIndex range over.")

    ReDim Preserve ArrayValue(Array2dColumnsCount(ArrayValue) - 1, Array2dRowsCount(ArrayValue))
    Dim I As Long
    For I = UBound(ArrayValue, 2) To RowIndex + 1 Step -1
        Call Array2dSetRowValues(ArrayValue, I, _
            Array2dGetRowValues(ArrayValue, I - 1))
    Next
    Call Array2dSetRowValues(ArrayValue, RowIndex, Values)
End Sub

'----------------------------------------
'・配列の要素を削除する
'----------------------------------------
'   ・  オブジェクト値にも対応
'----------------------------------------
Public Sub Array2dDelete(ByRef ArrayValue As Variant, _
ByVal RowIndex As Long)
    Call Assert(IsArray(ArrayValue), "Error:Array2dInsert:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dInsert:ArrayValue is not Array2D.")
    Call Assert(InRange(LBound(ArrayValue, 2), RowIndex, UBound(ArrayValue, 2)), _
        "Error:Array2dInsert:RowIndex range over.")

    Dim I As Long
    For I = RowIndex + 1 To UBound(ArrayValue, 2)
        Call Array2dSetRowValues(ArrayValue, I - 1, _
            Array2dGetRowValues(ArrayValue, I))
    Next

    If LBound(ArrayValue, 2) = UBound(ArrayValue, 2) Then
        Erase ArrayValue
        '配列の初期化はEraseを使う
    Else
        ReDim Preserve ArrayValue(Array2dColumnsCount(ArrayValue) - 1, _
            LBound(ArrayValue, 2) To UBound(ArrayValue, 2) - 1)
    End If
End Sub

Public Sub testArray2dBasicFunction()
    Dim A()
    Call Check(0, ArrayCount(A, 1))
    Call Check(0, ArrayCount(A, 2))

    Call Array2dSetColumn(A, 3)
    Call Check(3, ArrayCount(A, 1))
    Call Check(1, ArrayCount(A, 2))

    Call Array2dSetRowValues(A, 0, Array("A", "B", "C"))
    Call Array2dAdd(A, Array("D", "E", "F"))
    Call Array2dAdd(A, Array("G", "H", "I"))
    Call Array2dAdd(A, Array("1", "2", "3"))

    Dim B() As String
    B = Array2dGetRowValues(A, 0)
    Call Check("A,B,C", ArrayToString(B, ","))
    Call Check("D,E,F", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("G,H,I", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("1,2,3", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check(3, Array2dColumnsCount(A))
    Call Check(4, Array2dRowsCount(A))

    Call Array2dInsert(A, 3, B)
    Call Check("A,B,C", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check("D,E,F", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("G,H,I", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("A,B,C", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check("1,2,3", ArrayToString(Array2dGetRowValues(A, 4), ","))
    Call Check(3, Array2dColumnsCount(A))
    Call Check(5, Array2dRowsCount(A))

    Call Array2dDelete(A, 0)
    Call Check("D,E,F", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check("G,H,I", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("A,B,C", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("1,2,3", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check(3, Array2dColumnsCount(A))
    Call Check(4, Array2dRowsCount(A))

    Call Array2dDelete(A, 3)
    Call Array2dDelete(A, 1)
    Call Array2dDelete(A, 0)
    Call Check("A,B,C", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check(3, Array2dColumnsCount(A))
    Call Check(1, Array2dRowsCount(A))

    Call Array2dDelete(A, 0)
    Call Check(0, Array2dColumnsCount(A))
    Call Check(0, Array2dRowsCount(A))

End Sub

'----------------------------------------
'・配列内の値を検索してユニーク(同一値がない)かどうかを判断する
'----------------------------------------
Public Function Array2dIsUnique(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long) As Boolean
    Call Assert(IsArray(ArrayValue), "Error:ArrayIsUnique:ArrayValue is not array")
    Call Assert(ArrayDimension(ArrayValue) = 2, _
        "Error:Array2dIsUnique:ArrayValue Dimension is miss")

    Dim Result As Boolean: Result = True
    Do
        If OrValue(Array2dRowsCount(ArrayValue), 0, 1) Then Exit Do

        Dim I As Long
        Dim J As Long
        For I = LBound(ArrayValue, 2) To UBound(ArrayValue, 2) - 1
            For J = I + 1 To UBound(ArrayValue, 2)
                If ArrayValue(ColumnIndex, I) = ArrayValue(ColumnIndex, J) Then
                    Result = False
                    Exit Do
                End If
            Next
        Next
        Loop While False
    Array2dIsUnique = Result
End Function

Sub testArray2dIsUnique()
    Dim A()
    Call Array2dSetColumn(A, 3)

    Call Array2dSetRowValues(A, 0, Array("A", "B", "C"))
    Call Array2dAdd(A, Array("D", "E", "C"))
    Call Array2dAdd(A, Array("G", "H", "C"))
    Call Array2dAdd(A, Array("1", "2", "C"))

    Call Check(True, Array2dIsUnique(A, 0))
    Call Check(True, Array2dIsUnique(A, 1))
    Call Check(False, Array2dIsUnique(A, 2))
End Sub

'----------------------------------------
'・クイックソート
'----------------------------------------
'   ・  IndexMin/IndexMaxを指定すると
'       指定範囲内の値をソートする
'----------------------------------------
Public Sub Array2dSortQuick(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long, _
Optional ByVal SortOrder As SortOrder = SortOrder.Ascending, _
Optional ByVal RowIndexMin As Long = -1, Optional ByVal RowIndexMax As Long = -1)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 2, _
        "Error:Array2dSortQuick:ArrayValue Dimension is miss.")
    Call Assert(InRange(LBound(ArrayValue, 1), ColumnIndex, UBound(ArrayValue, 1)), _
        "Error:Array2dSortQuick:ColumnIndex is range over.")

    Call Assert(RowIndexMin <= RowIndexMax, "Error:IndexMin < IndexMax")
    Call Assert(InRange(-1, RowIndexMin, Array2dRowsCount(ArrayValue) - 1), _
        "Error:ArrayReverse:RowIndexMin Range is miss.")
    Call Assert(InRange(-1, RowIndexMax, Array2dRowsCount(ArrayValue) - 1), _
        "Error:ArrayReverse:RowIndexMax Range is miss.")

    '1以下ならソート不可能なのでExitする
    If Array2dRowsCount(ArrayValue) <= 1 Then Exit Sub

    RowIndexMin = IIf(RowIndexMin = -1, LBoundNoError(ArrayValue, 2), RowIndexMin)
    RowIndexMax = IIf(RowIndexMax = -1, UBoundNoError(ArrayValue, 2), RowIndexMax)

    'IndexMin=IndexMaxならソート不可能なのでExit
    If RowIndexMin = RowIndexMax Then Exit Sub

    Call Array2dSortQuickBase(ArrayValue, ColumnIndex, SortOrder, RowIndexMin, RowIndexMax)
End Sub

'クイックソートのベース関数、再起呼び出しされる
Sub Array2dSortQuickBase(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long, _
ByVal SortOrder As SortOrder, _
ByVal RowIndexMin As Long, ByVal RowIndexMax As Long)

    Dim RowIndexCenter As Long
    Dim RowIndex1 As Long
    Dim RowIndex2 As Long
    Dim RowValue1 As Variant
    Dim RowValue2 As Variant

    If RowIndexMax <= RowIndexMin Then Exit Sub

    Dim RowOrigin As Long
    Dim ColOrigin As Long
    RowOrigin = LBoundNoError(ArrayValue, 2)
    ColOrigin = LBoundNoError(ArrayValue, 1)
    
    RowIndexCenter = (RowIndexMin + RowIndexMax) \ 2

    '中央値をバッファ
    RowValue1 = Array2dGetRowValues(ArrayValue, RowIndexCenter)
    '中央値に開始位置要素を代入
    Call Array2dSetRowValues(ArrayValue, RowIndexCenter, _
        Array2dGetRowValues(ArrayValue, RowIndexMin))

    RowIndex2 = RowIndexMin

    RowIndex1 = RowIndexMin + 1

    Select Case SortOrder
    Case Ascending
        Do While RowIndex1 <= RowIndexMax
            If ArrayValue(ColumnIndex, RowIndex1) < RowValue1(ColumnIndex - ColOrigin) Then
                RowIndex2 = RowIndex2 + 1

                RowValue2 = Array2dGetRowValues(ArrayValue, RowIndex2)
                Call Array2dSetRowValues(ArrayValue, RowIndex2, _
                    Array2dGetRowValues(ArrayValue, RowIndex1))
                Call Array2dSetRowValues(ArrayValue, RowIndex1, RowValue2)
            End If
            RowIndex1 = RowIndex1 + 1
        Loop
    Case Descending
        Do While RowIndex1 <= RowIndexMax
            If RowValue1(ColumnIndex - ColOrigin) < ArrayValue(ColumnIndex, RowIndex1) Then
                RowIndex2 = RowIndex2 + 1

                RowValue2 = Array2dGetRowValues(ArrayValue, RowIndex2)
                Call Array2dSetRowValues(ArrayValue, RowIndex2, _
                    Array2dGetRowValues(ArrayValue, RowIndex1))
                Call Array2dSetRowValues(ArrayValue, RowIndex1, RowValue2)
            End If
            RowIndex1 = RowIndex1 + 1
        Loop
    Case Else
        Call Assert(False, "Error:Array2dSortQuickBase:SortOrder is miss.")
    End Select

    Call Array2dSetRowValues(ArrayValue, RowIndexMin, _
        Array2dGetRowValues(ArrayValue, RowIndex2))
    Call Array2dSetRowValues(ArrayValue, RowIndex2, RowValue1)

    ' 分割前半を再帰呼び出しでSORT
    Call Array2dSortQuickBase(ArrayValue, ColumnIndex, SortOrder, RowIndexMin, RowIndex2 - 1)

    ' 分割後半を再帰呼び出しでSORT
    Call Array2dSortQuickBase(ArrayValue, ColumnIndex, SortOrder, RowIndex2 + 1, RowIndexMax)
End Sub

Sub testArray2dSortQuick()
    Dim Array1(2, 5) As Variant
    Array1(0, 0) = "A1"
    Array1(0, 1) = "A2"
    Array1(0, 2) = "A3"
    Array1(0, 3) = "A1"
    Array1(0, 4) = "A2"
    Array1(0, 5) = "A3"
    Array1(1, 0) = "100"
    Array1(1, 1) = "101"
    Array1(1, 2) = "102"
    Array1(1, 3) = "103"
    Array1(1, 4) = "104"
    Array1(1, 5) = "105"

    'クイックソートのためのキー項目作成
    Dim I As Long
    For I = 0 To Array2dRowsCount(Array1) - 1
        Array1(2, I) = Array1(0, I) + CStr(Array1(1, I))
    Next

    Call Check(Array1(0, 0), "A1")
    Call Check(Array1(0, 1), "A2")
    Call Check(Array1(0, 2), "A3")
    Call Check(Array1(0, 3), "A1")
    Call Check(Array1(0, 4), "A2")
    Call Check(Array1(0, 5), "A3")
    Call Check(Array1(1, 0), "100")
    Call Check(Array1(1, 1), "101")
    Call Check(Array1(1, 2), "102")
    Call Check(Array1(1, 3), "103")
    Call Check(Array1(1, 4), "104")
    Call Check(Array1(1, 5), "105")

    Call Array2dSortQuick(Array1, 0)
    Call Check(Array1(0, 0), "A1")
    Call Check(Array1(0, 1), "A1")
    Call Check(Array1(0, 2), "A2")
    Call Check(Array1(0, 3), "A2")
    Call Check(Array1(0, 4), "A3")
    Call Check(Array1(0, 5), "A3")
'    Call Check(Array1(1, 0), "100")
'    Call Check(Array1(1, 1), "103")
'    Call Check(Array1(1, 2), "101")
'    Call Check(Array1(1, 3), "104")
'    Call Check(Array1(1, 4), "102")
'    Call Check(Array1(1, 5), "105")
'クイックソートではキー項目がないと
'ソートがきれいに行われない。

    Call Array2dSortQuick(Array1, 1)
    Call Check(Array1(0, 0), "A1")
    Call Check(Array1(0, 1), "A2")
    Call Check(Array1(0, 2), "A3")
    Call Check(Array1(0, 3), "A1")
    Call Check(Array1(0, 4), "A2")
    Call Check(Array1(0, 5), "A3")
    Call Check(Array1(1, 0), "100")
    Call Check(Array1(1, 1), "101")
    Call Check(Array1(1, 2), "102")
    Call Check(Array1(1, 3), "103")
    Call Check(Array1(1, 4), "104")
    Call Check(Array1(1, 5), "105")

    Call Array2dSortQuick(Array1, 2)
    Call Check(Array1(0, 0), "A1")
    Call Check(Array1(0, 1), "A1")
    Call Check(Array1(0, 2), "A2")
    Call Check(Array1(0, 3), "A2")
    Call Check(Array1(0, 4), "A3")
    Call Check(Array1(0, 5), "A3")
    Call Check(Array1(1, 0), "100")
    Call Check(Array1(1, 1), "103")
    Call Check(Array1(1, 2), "101")
    Call Check(Array1(1, 3), "104")
    Call Check(Array1(1, 4), "102")
    Call Check(Array1(1, 5), "105")
    'キー項目に対してソートするときれいな結果になる
    
    
    '1開始の動的配列での動作確認
    Dim Array2(1 To 3, 1 To 6) As String
    Array2(1, 1) = "A1"
    Array2(1, 2) = "A2"
    Array2(1, 3) = "A3"
    Array2(1, 4) = "A1"
    Array2(1, 5) = "A2"
    Array2(1, 6) = "A3"
    Array2(2, 1) = "100"
    Array2(2, 2) = "101"
    Array2(2, 3) = "102"
    Array2(2, 4) = "103"
    Array2(2, 5) = "104"
    Array2(2, 6) = "105"

    'クイックソートのためのキー項目作成
    For I = LBoundNoError(Array2, 2) To UBoundNoError(Array2, 2)
        Array2(3, I) = Array2(1, I) + CStr(Array2(2, I))
    Next
    
    Call Array2dSortQuick(Array2, 3)
    Call Check(Array2(1, 1), "A1")
    Call Check(Array2(1, 2), "A1")
    Call Check(Array2(1, 3), "A2")
    Call Check(Array2(1, 4), "A2")
    Call Check(Array2(1, 5), "A3")
    Call Check(Array2(1, 6), "A3")
    Call Check(Array2(2, 1), "100")
    Call Check(Array2(2, 2), "103")
    Call Check(Array2(2, 3), "101")
    Call Check(Array2(2, 4), "104")
    Call Check(Array2(2, 5), "102")
    Call Check(Array2(2, 6), "105")
    
    
End Sub

'----------------------------------------
'・文字列長ソート
'----------------------------------------
Public Sub Array2dSortStrLength(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long, _
Optional ByVal SortOrder As SortOrder = SortOrder.Ascending)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 2, _
        "Error:Array2dSortStrLength:ArrayValue Dimension is miss.")
    Call Assert(InRange(LBound(ArrayValue, 1), ColumnIndex, UBound(ArrayValue, 1)), _
        "Error:Array2dSortStrLength:ColumnIndex is range over.")

    Dim DigitArrayRowsCount As Long
    Dim DigitStrLength As Long

    Dim Delimiter As String
    Delimiter = ""

    'ソートキー文字列の追加
    Call Array2dSortStrLengthSetKeyValue(ArrayValue, ColumnIndex, _
        ColumnIndex, FirstAdd, Delimiter, True, DigitStrLength, DigitArrayRowsCount, SortOrder)

    Call Array2dSortQuick(ArrayValue, ColumnIndex, Ascending)

    'ソートキー文字列の削除
    Dim I As Long
    For I = 0 To Array2dRowsCount(ArrayValue) - 1
        ArrayValue(ColumnIndex, I) = _
            Mid$(ArrayValue(ColumnIndex, I), _
                DigitStrLength + DigitArrayRowsCount + Len(Delimiter) + 1)
    Next

End Sub

Public Sub Array2dSortStrLengthSetKeyValue(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long, _
ByVal KeyColumnIndex As Long, _
ByVal KeyAddType As StrAddType, _
ByVal KeyDelimiter As String, _
ByVal OutputArrayRows As Boolean, _
ByRef Out_DigitStrLength As Long, _
ByRef Out_DigitArrayRowsCount As Long, _
ByVal SortOrder As SortOrder)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:ArrayValue Dimension is miss")
    Call Assert(InRange(LBound(ArrayValue, 1), ColumnIndex, UBound(ArrayValue, 1)), "Error:ColumnIndex is range over.")
    Call Assert(OrValue(SortOrder, Ascending, Descending), "Error:SortOrder is miss.")
    Call Assert(OrValue(KeyAddType, FirstAdd, LastAdd), "Error:KeyAddType is miss.")

    Out_DigitArrayRowsCount = Len(CStr(Array2dRowsCount(ArrayValue) - 1))
    Out_DigitStrLength = 0
    Dim MaxLength As Long
    MaxLength = 0

    Dim I As Long
    For I = 0 To Array2dRowsCount(ArrayValue) - 1
        MaxLength = MaxValue(MaxLength, Len(ArrayValue(ColumnIndex, I)))
    Next
    Out_DigitStrLength = Len(CStr(MaxLength))

    Select Case SortOrder
    Case Ascending
        Select Case KeyAddType
        Case FirstAdd
            For I = 0 To Array2dRowsCount(ArrayValue) - 1
                ArrayValue(KeyColumnIndex, I) = _
                    LongToStrDigitZero(Len(ArrayValue(ColumnIndex, I)), Out_DigitStrLength) + _
                    IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "") + _
                    KeyDelimiter + _
                    ArrayValue(KeyColumnIndex, I)
            Next
        Case LastAdd
            For I = 0 To Array2dRowsCount(ArrayValue) - 1
                ArrayValue(KeyColumnIndex, I) = _
                    ArrayValue(KeyColumnIndex, I) + _
                    KeyDelimiter + _
                    LongToStrDigitZero(Len(ArrayValue(ColumnIndex, I)), Out_DigitStrLength) + _
                    IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "")
            Next
        End Select
    Case Descending
        Select Case KeyAddType
        Case FirstAdd
            For I = 0 To Array2dRowsCount(ArrayValue) - 1
                ArrayValue(KeyColumnIndex, I) = _
                    LongToStrDigitZero(MaxLength - Len(ArrayValue(ColumnIndex, I)), Out_DigitStrLength) + _
                    IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "") + _
                    KeyDelimiter + _
                    ArrayValue(KeyColumnIndex, I)
            Next
        Case LastAdd
            For I = 0 To Array2dRowsCount(ArrayValue) - 1
                ArrayValue(KeyColumnIndex, I) = _
                    ArrayValue(KeyColumnIndex, I) + _
                    KeyDelimiter + _
                    LongToStrDigitZero(MaxLength - Len(ArrayValue(ColumnIndex, I)), Out_DigitStrLength) + _
                    IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "")
            Next
        End Select
    End Select
End Sub

Public Sub testArray2dSortStrLength()
    Dim A()

    Call Array2dAdd(A, Array("A", "B", "C", "123"))
    Call Array2dAdd(A, Array("D", "E", "F", "12"))
    Call Array2dAdd(A, Array("G", "H", "I", "1"))
    Call Array2dAdd(A, Array("1", "2", "3", "1"))
    Call Array2dAdd(A, Array("4", "5", "6", "22"))
    Call Array2dAdd(A, Array("7", "8", "9", "333"))

    Call Array2dSortStrLength(A, 3, Ascending)

    Call Check("G,H,I,1", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check("1,2,3,1", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("D,E,F,12", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("4,5,6,22", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check("A,B,C,123", ArrayToString(Array2dGetRowValues(A, 4), ","))
    Call Check("7,8,9,333", ArrayToString(Array2dGetRowValues(A, 5), ","))

    Erase A
    Call Array2dAdd(A, Array("A", "B", "C", "123"))
    Call Array2dAdd(A, Array("D", "E", "F", "12"))
    Call Array2dAdd(A, Array("G", "H", "I", "1"))
    Call Array2dAdd(A, Array("1", "2", "3", "1"))
    Call Array2dAdd(A, Array("4", "5", "6", "22"))
    Call Array2dAdd(A, Array("7", "8", "9", "333"))

    Call Array2dSortStrLength(A, 3, Descending)
    Call Check("A,B,C,123", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check("7,8,9,333", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("D,E,F,12", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("4,5,6,22", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check("G,H,I,1", ArrayToString(Array2dGetRowValues(A, 4), ","))
    Call Check("1,2,3,1", ArrayToString(Array2dGetRowValues(A, 5), ","))

End Sub

'----------------------------------------
'・独自並び順ソート
'----------------------------------------
'   ・  ソート指定配列の文字列に一致する順番に
'       並び替えをするソート
'   ・  s/m/l/xl/xxlとかそういう並び指定を行う
'----------------------------------------
Public Sub Array2dSortCustomOrder(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long, _
ByRef OrderArrayWildCard() As String, _
Optional CaseCompare As CaseCompare = CaseCompare.IgnoreCase, _
Optional NoOrderValuePriority As Boolean = False)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 2, _
        "Error:Array2dSortStrLength:ArrayValue Dimension is miss.")
    Call Assert(InRange(LBound(ArrayValue, 1), ColumnIndex, UBound(ArrayValue, 1)), _
        "Error:Array2dSortStrLength:ColumnIndex is range over.")

    Dim DigitArrayRowsCount As Long
    Dim DigitOrderCount As Long

    Dim Delimiter As String
    Delimiter = ""

    'ソートキー文字列の追加
    Call Array2dSortCustomOrderSetKeyValue(ArrayValue, ColumnIndex, _
        ColumnIndex, FirstAdd, Delimiter, True, DigitOrderCount, DigitArrayRowsCount, _
        OrderArrayWildCard, CaseCompare, NoOrderValuePriority)

    Call Array2dSortQuick(ArrayValue, ColumnIndex, Ascending)

    'ソートキー文字列の削除
    Dim I As Long
    Dim StartIndex As Long: StartIndex = Array2dRowsStartIndex(ArrayValue)
    For I = StartIndex To Array2dRowsEndIndex(ArrayValue)
        ArrayValue(ColumnIndex, I) = _
            Mid$(ArrayValue(ColumnIndex, I), _
                DigitOrderCount + DigitArrayRowsCount + Len(Delimiter) + 1)
    Next

End Sub

Public Sub Array2dSortCustomOrderSetKeyValue(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long, _
ByVal KeyColumnIndex As Long, _
ByVal KeyAddType As StrAddType, _
ByVal KeyDelimiter As String, _
ByVal OutputArrayRows As Boolean, _
ByRef Out_DigitOrderCount As Long, _
ByRef Out_DigitArrayRowsCount As Long, _
ByRef OrderArrayWildCard() As String, _
Optional CaseCompare As CaseCompare = CaseCompare.IgnoreCase, _
Optional NoOrderValuePriority As Boolean = False)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 2, _
        "Error:Array2dSortCustomOrderSetKeyValue:ArrayValue Dimension is miss.")
    Call Assert(InRange(LBound(ArrayValue, 1), ColumnIndex, UBound(ArrayValue, 1)), _
        "Error:Array2dSortCustomOrderSetKeyValue:ColumnIndex is range over.")
    Call Assert(OrValue(KeyAddType, FirstAdd, LastAdd), _
        "Error:Array2dSortCustomOrderSetKeyValue:KeyAddType is miss.")

    Out_DigitArrayRowsCount = Len(CStr(Array2dRowsCount(ArrayValue) - 1))
    Out_DigitOrderCount = Len(CStr(ArrayCount(OrderArrayWildCard) + 1))

    Dim I As Long
    Dim StartIndex As Long: StartIndex = Array2dRowsStartIndex(ArrayValue)

    Dim OrderArrayIndex As Long

    Select Case KeyAddType
    Case FirstAdd
        For I = StartIndex To Array2dRowsEndIndex(ArrayValue)
            OrderArrayIndex = _
                ArrayIndexOf(OrderArrayWildCard, ArrayValue(ColumnIndex, I), , CaseCompare, WildCardArray)
            If OrderArrayIndex = -1 Then
                If NoOrderValuePriority = False Then
                    ArrayValue(KeyColumnIndex, I) = _
                        LongToStrDigitZero(ArrayCount(OrderArrayWildCard) + 1, Out_DigitOrderCount) + _
                        IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "") + _
                        KeyDelimiter + _
                        ArrayValue(KeyColumnIndex, I)
                Else
                    ArrayValue(KeyColumnIndex, I) = _
                        LongToStrDigitZero(0, Out_DigitOrderCount) + _
                        IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "") + _
                        KeyDelimiter + _
                        ArrayValue(KeyColumnIndex, I)
                End If
            Else
                ArrayValue(KeyColumnIndex, I) = _
                    LongToStrDigitZero(OrderArrayIndex + 1, Out_DigitOrderCount) + _
                    IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "") + _
                    KeyDelimiter + _
                    ArrayValue(KeyColumnIndex, I)
            End If
        Next
    Case LastAdd
        For I = 0 To Array2dRowsCount(ArrayValue) - 1
            OrderArrayIndex = _
                ArrayIndexOf(OrderArrayWildCard, ArrayValue(ColumnIndex, I), , CaseCompare, WildCardArray)
            If OrderArrayIndex = -1 Then
                If NoOrderValuePriority = False Then
                    ArrayValue(KeyColumnIndex, I) = _
                        ArrayValue(KeyColumnIndex, I) + _
                        KeyDelimiter + _
                        LongToStrDigitZero(ArrayCount(OrderArrayWildCard) + 1, Out_DigitOrderCount) + _
                        IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "")
                Else
                    ArrayValue(KeyColumnIndex, I) = _
                        ArrayValue(KeyColumnIndex, I) + _
                        KeyDelimiter + _
                        LongToStrDigitZero(0, Out_DigitOrderCount) + _
                        IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "")
                End If
            Else
                ArrayValue(KeyColumnIndex, I) = _
                    ArrayValue(KeyColumnIndex, I) + _
                    KeyDelimiter + _
                    LongToStrDigitZero(OrderArrayIndex + 1, Out_DigitOrderCount) + _
                    IIf(OutputArrayRows, LongToStrDigitZero(I, Out_DigitArrayRowsCount), "")
            End If
        Next
    End Select

End Sub


Public Sub testArray2dSortCustomOrder()
    Dim A()

    Call Array2dAdd(A, Array("01", "02", "03", "b"))
    Call Array2dAdd(A, Array("04", "05", "06", "a"))
    Call Array2dAdd(A, Array("07", "08", "09", "s"))
    Call Array2dAdd(A, Array("11", "12", "13", "ss"))
    Call Array2dAdd(A, Array("14", "15", "16", "l"))
    Call Array2dAdd(A, Array("17", "18", "19", "ll"))
    Call Array2dAdd(A, Array("21", "22", "23", "m"))

    Call Array2dSortCustomOrder(A, 3, _
        ArrayStr("ss*", "s*", "m*", "l*", "ll*"), CaseSensitive, False)

    Call Check("11,12,13,ss", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check("07,08,09,s", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("21,22,23,m", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("14,15,16,l", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check("17,18,19,ll", ArrayToString(Array2dGetRowValues(A, 4), ","))
    Call Check("01,02,03,b", ArrayToString(Array2dGetRowValues(A, 5), ","))
    Call Check("04,05,06,a", ArrayToString(Array2dGetRowValues(A, 6), ","))

    Erase A

    Call Array2dAdd(A, Array("01", "02", "03", "b"))
    Call Array2dAdd(A, Array("04", "05", "06", "a"))
    Call Array2dAdd(A, Array("07", "08", "09", "s"))
    Call Array2dAdd(A, Array("11", "12", "13", "ss"))
    Call Array2dAdd(A, Array("14", "15", "16", "l"))
    Call Array2dAdd(A, Array("17", "18", "19", "ll"))
    Call Array2dAdd(A, Array("21", "22", "23", "m"))

    Call Array2dSortCustomOrder(A, 3, _
        ArrayStr("ss*", "s*", "m*", "l*", "ll*"), CaseSensitive, True)

    Call Check("01,02,03,b", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check("04,05,06,a", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("11,12,13,ss", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("07,08,09,s", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check("21,22,23,m", ArrayToString(Array2dGetRowValues(A, 4), ","))
    Call Check("14,15,16,l", ArrayToString(Array2dGetRowValues(A, 5), ","))
    Call Check("17,18,19,ll", ArrayToString(Array2dGetRowValues(A, 6), ","))

End Sub

'----------------------------------------
'・シートRangeに対応した 独自並び順ソート
'----------------------------------------
Public Sub SheetRangeSortCustomOrder(ByRef Sheet As Worksheet, _
ByVal Range As Range, _
ByVal ColumnIndex As Long, _
ByRef OrderArrayWildCard() As String, _
Optional CaseCompare As CaseCompare = CaseCompare.IgnoreCase, _
Optional NoOrderValuePriority As Boolean = False)
    
    Dim Array2D As Variant
    Array2D = Range
    
    Array2D = Array2dTranspose(Array2D)
    
    Call Array2dSortCustomOrder(Array2D, ColumnIndex, OrderArrayWildCard, _
        CaseCompare, NoOrderValuePriority)

    Array2D = Array2dTranspose(Array2D)
    
    Range = Array2D
End Sub

Public Function Array2dTranspose(ByRef ArrayValue As Variant) As Variant

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 2, _
        "Error:Array2dSortCustomOrderSetKeyValue:ArrayValue Dimension is miss.")

    Dim LBound1 As Long: LBound1 = LBoundNoError(ArrayValue, 1)
    Dim UBound1 As Long: UBound1 = UBoundNoError(ArrayValue, 1)
    Dim LBound2 As Long: LBound2 = LBoundNoError(ArrayValue, 2)
    Dim UBound2 As Long: UBound2 = UBoundNoError(ArrayValue, 2)
 
    Dim Result As Variant
    ReDim Result(LBound2 To UBound2, LBound1 To UBound1)
    Dim I As Long
    For I = LBound1 To UBound1
        Dim J As Long
        For J = LBound2 To UBound2
            Result(J, I) = ArrayValue(I, J)
        Next J
    Next I
    
    Array2dTranspose = Result
End Function


'----------------------------------------
'◆ファイル名処理
'----------------------------------------

'----------------------------------------
'◇ 終端パス区切り
'----------------------------------------

'----------------------------------------
'・終端にパス区切りを追加する関数
'----------------------------------------
Public Function IncludeLastPathDelim(ByVal Path As String) As String
    IncludeLastPathDelim = IncludeLastStr(Path, Application.PathSeparator)
End Function

'----------------------------------------
'・終端からパス区切りを削除する関数
'----------------------------------------
Public Function ExcludeLastPathDelim(ByVal Path As String) As String
    ExcludeLastPathDelim = ExcludeLastStr(Path, Application.PathSeparator)
End Function

'----------------------------------------
'◇ ドライブパス、ネットワークパス
'----------------------------------------

'----------------------------------------
'・ドライブパス"C:"を取り出す関数
'----------------------------------------
'   ・  [C:]という文字列が取得できる
'----------------------------------------
Public Function GetDrivePath(ByVal Path As String) As String
    GetDrivePath = IncludeLastStr( _
        FirstStrFirstDelim(Path, ":"), ":")
End Function

'----------------------------------------
'・ドライブパスが含まれているかどうか確認する関数
'[:]が2文字目以降にあるかどうかで判定
'----------------------------------------
Public Function IsDrivePath(ByVal Path As String) As Boolean
    Dim Result As Boolean
    Result = (OrValue(InStr(Path, ":"), 2, 3))
    IsDrivePath = Result
End Function
'
'----------------------------------------
'・ネットワークドライブかどうか確認する関数
'----------------------------------------
Public Function IsNetworkPath(ByVal Path As String) As Boolean
    Dim Result As Boolean: Result = False
    If IsFirstStr(Path, "\\") Then
        If 3 <= Len(Path) Then
            Result = True
        End If
    End If
    IsNetworkPath = Result
End Function


'----------------------------------------
'◇ ファイル名に許容できない文字
'----------------------------------------

'----------------------------------------
'・ ファイル名に許容できない文字が含まれているかどうか調べる
'----------------------------------------
'   ・  ファイル名として許容できない文字は次の通り
'       \/:*?"<>|
'----------------------------------------

Public Function FilePath_IsIncludeFileNameOutString(ByVal FileName As String) As Boolean
    Dim Result As Boolean: Result = False
    Result = Result Or IsIncludeStr(FileName, "\")
    Result = Result Or IsIncludeStr(FileName, "/")
    Result = Result Or IsIncludeStr(FileName, ":")
    Result = Result Or IsIncludeStr(FileName, "*")
    Result = Result Or IsIncludeStr(FileName, "?")
    Result = Result Or IsIncludeStr(FileName, """")
    Result = Result Or IsIncludeStr(FileName, "<")
    Result = Result Or IsIncludeStr(FileName, ">")
    Result = Result Or IsIncludeStr(FileName, "|")
    
    FilePath_IsIncludeFileNameOutString = Result
End Function

Public Sub test_FilePath_IsIncludeFileNameOutString()
    Call Check(False, FilePath_IsIncludeFileNameOutString("testtest.txt"))
    Call Check(True, FilePath_IsIncludeFileNameOutString("test\test.txt"))
    Call Check(True, FilePath_IsIncludeFileNameOutString("test|test.txt"))
    Call Check(True, FilePath_IsIncludeFileNameOutString("t:esttest.txt"))
    Call Check(True, FilePath_IsIncludeFileNameOutString("test<test>.txt"))
    Call Check(True, FilePath_IsIncludeFileNameOutString("test""test.txt"))
End Sub

'----------------------------------------
'・ ファイル名に許容できない文字が含まれていたら文字列置き換え
'----------------------------------------
'   ・  ファイル名として許容できない文字は次の通り
'       \/:*?"<>|
'----------------------------------------

Public Function FilePath_ReplaceFileNameOutString( _
ByVal FileName As String, _
Optional ByVal ReplaceStr As String = "") As String

    Dim Result As String: Result = ""
    Result = ReplaceArrayValue(FileName, _
        ArrayStr("\", "/", ":", "*", "?", """", "<", ">", "|"), _
        ArrayStr(ReplaceStr, ReplaceStr, ReplaceStr, ReplaceStr, ReplaceStr, ReplaceStr, ReplaceStr, ReplaceStr, ReplaceStr))
        
    FilePath_ReplaceFileNameOutString = Result
End Function

Public Sub test_FilePath_ReplaceFileNameOutString()
    Call Check("testtest.txt", FilePath_ReplaceFileNameOutString("testtest.txt"))
    Call Check("testtest.txt", FilePath_ReplaceFileNameOutString("test\test.txt"))
    Call Check("testtest.txt", FilePath_ReplaceFileNameOutString("test|test.txt"))
    Call Check("testtest.txt", FilePath_ReplaceFileNameOutString("t:esttest.txt"))
    Call Check("testtest.txt", FilePath_ReplaceFileNameOutString("test<test>.txt"))
    Call Check("testtest.txt", FilePath_ReplaceFileNameOutString("test""test.txt"))
End Sub


'----------------------------------------
'◇ 空白を含むファイルパス
'----------------------------------------
'----------------------------------------
'・空白を含むファイルパスをダブルクウォートで囲む
'----------------------------------------
Public Function InSpacePlusDoubleQuote(ByVal Path As String) As String
    Dim Result As String
    If 1 <= InStr(Path, " ") Then
        Result = IncludeBothEndsStr(Path, """")
    Else
        Result = Path
    End If
    InSpacePlusDoubleQuote = Result
End Function

'----------------------------------------
'◇ 拡張子の変更
'----------------------------------------

'----------------------------------------
'・拡張子の取得
'----------------------------------------
'   ・  fso.GetExtensionNameでは取得できない
'       最後がピリオドで終わるファイルでも
'       値を取得することができる
'----------------------------------------
Public Function GetExtensionIncludePeriod(ByVal Path As String) As String
    Dim Result As String
    Result = _
        LastStrLastDelim(Path, ".")
    If Result = Path Then
        Result = ""
    Else
        Result = IncludeFirstStr(Result, ".")
    End If
    GetExtensionIncludePeriod = Result
End Function

Private Sub testGetExtensionIncludePeriod()
    Call Check("txt", fso.GetExtensionName("C:\Test\test.txt"))
    Call Check(".txt", GetExtensionIncludePeriod("C:\Test\test.txt"))
    Call Check("", fso.GetExtensionName("C:\Test\test"))
    Call Check("", GetExtensionIncludePeriod("C:\Test\test"))
    Call Check("", fso.GetExtensionName("C:\Test\test."))
    Call Check(".", GetExtensionIncludePeriod("C:\Test\test."))
End Sub

'----------------------------------------
'・拡張子の変更
'----------------------------------------
'   ・  NewExtに空文字を指定すると拡張子なしになる
'   ・  NewExtにはピリオドは必須
'----------------------------------------
Public Function ChangeFileExtension(ByVal Path As String, _
ByVal NewExt As String) As String
    Dim Result As String: Result = ""
    Result = _
        ExcludeLastStr( _
            Path, GetExtensionIncludePeriod(Path)) _
        + NewExt
    ChangeFileExtension = Result
End Function

Private Sub testChangeFileExtension()
    Call Check("C:\temp\text.csv", _
        ChangeFileExtension("C:\temp\text.txt", ".csv"))
        
    Call Check("C:\temp\textcsv", _
        ChangeFileExtension("C:\temp\text", "csv"))
        
    Call Check("C:\temp\text_csv", _
        ChangeFileExtension("C:\temp\text.", "_csv"))
        
    Call Check("C:\temp\text", _
        ChangeFileExtension("C:\temp\text.", ""))
        
End Sub

'----------------------------------------
'◇ パスの結合
'----------------------------------------

'----------------------------------------
'・ パスの結合
'----------------------------------------
Public Function PathCombine(ParamArray Values()) As String
    'パラメータ配列を他のパラメータ配列に渡す事はできないので
    'パラメータ配列をString配列に代入している
    Dim Parameter() As String
    ReDim Parameter(UBound(Values))
    Dim I As Long
    For I = 0 To UBound(Values)
        Parameter(I) = Values(I)
    Next

    PathCombine = StringCombineArray( _
        Application.PathSeparator, Parameter)
End Function

Private Sub testPathCombine()

    Call Check("C:\Temp\Temp\temp.txt", PathCombine("C:", "Temp", "Temp", "temp.txt"))
    Call Check("C:\Temp\Temp\temp.txt", PathCombine("C:\Temp", "Temp\temp.txt"))
    Call Check("C:\Temp\Temp\temp.txt", PathCombine("C:\Temp\Temp\", "\temp.txt"))
    Call Check("\Temp\Temp\", PathCombine("\Temp\", "\Temp\"))

    Call Check("C:\work\bbb\a.txt", PathCombine("C:\work", "bbb\a.txt"))
    Call Check("C:\work\bbb\a.txt", PathCombine("C:\work\", "bbb\a.txt"))
    Call Check("C:\work\bbb\a.txt", PathCombine("C:\work", "\bbb\a.txt"))
    Call Check("C:\work\bbb\a.txt", PathCombine("C:\work\", "\bbb\a.txt"))

    Call Check("C:\work\bbb\a.txt", PathCombine("C:\work", "bbb", "a.txt"))
    Call Check("C:\work\bbb\a.txt", PathCombine("C:\work\", "\bbb\", "\a.txt"))

    Call Check("\C:\work\bbb\a.txt\", PathCombine("\C:\work\", "\bbb\", "\a.txt\"))

End Sub


'----------------------------------------
'◆ファイルフォルダパス取得
'----------------------------------------

'----------------------------------------
'・特殊フォルダ名
'----------------------------------------
Public Function GetSpecialFolderPath( _
ByVal SpecialFolderType As SpecialFolderType) As String

    Dim Result As String
    Select Case SpecialFolderType
    Case Desktop
        Result = Shell.SpecialFolders("Desktop")
    Case MyDocument
        Result = Shell.SpecialFolders("MyDocuments")
    Case StartMenu
        Result = Shell.SpecialFolders("STARTMENU")
    Case StartMenuProgram
        Result = Shell.SpecialFolders("PROGRAMS")
    Case StartMenuStartup
        Result = Shell.SpecialFolders("STARTUP")
    Case SendTo
        Result = Shell.SpecialFolders("SENDTO")
    Case AppData
        Result = Shell.SpecialFolders("Appdata")

    Case AllUsersDesktop
        Result = Shell.SpecialFolders("AllUsersDesktop")
    Case AllUsersStartMenu
        Result = Shell.SpecialFolders("AllUsersStartMenu")
    Case AllUsersStartMenuProgram
        Result = Shell.SpecialFolders("AllUsersPrograms")
    Case AllUsersStartMenuStartup
        Result = Shell.SpecialFolders("AllUsersStartup")

    Case TaskbarPin
        Result = PathCombine(Shell.SpecialFolders("Appdata"), _
            "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
    Case Windows
        Result = fso.GetSpecialFolder(WindowsFolder)
        'C:\Windows
    Case System
        Result = fso.GetSpecialFolder(SystemFolder)
        'C:\Windows\System32
    Case Temporary
        Result = fso.GetSpecialFolder(TemporaryFolder)

    Case Else
        Call Assert(False, "Error:GetSpecialFolderPath")
    End Select

    GetSpecialFolderPath = Result
End Function

Sub testGetSpecialFolderPath()
    Call MsgBox(GetSpecialFolderPath(Desktop))
    Call MsgBox(GetSpecialFolderPath(MyDocument))
    Call MsgBox(GetSpecialFolderPath(StartMenu))
    Call MsgBox(GetSpecialFolderPath(StartMenuProgram))
    Call MsgBox(GetSpecialFolderPath(StartMenuStartup))
    Call MsgBox(GetSpecialFolderPath(SendTo))
    Call MsgBox(GetSpecialFolderPath(AppData))
    Call MsgBox(GetSpecialFolderPath(AllUsersDesktop))
    Call MsgBox(GetSpecialFolderPath(AllUsersStartMenu))
    Call MsgBox(GetSpecialFolderPath(AllUsersStartMenuProgram))
    Call MsgBox(GetSpecialFolderPath(AllUsersStartMenuStartup))
End Sub


'----------------------------------------
'◆ファイル処理
'----------------------------------------

'------------------------------
'・ファイル存在確認
'------------------------------
'   ・ Win/Mac両対応版
'------------------------------
Function FileExists(ByVal AFileName As String) As Boolean
    On Error GoTo Catch

    FileSystem.FileLen AFileName

    FileExists = True

    GoTo Finally

Catch:
        FileExists = False
Finally:
End Function

'----------------------------------------
'・相対パスから絶対パス取得
'----------------------------------------
'   ・  ドライブパスとネットワークパスの場合をわけて
'       ドライブパスの場合は ChDrive と ChDir を組み合わせて
'       実装していたが、そんな必要もなく
'       Shellオブジェクトを使う方が楽に設定できる。
'----------------------------------------
Public Function AbsolutePath(ByVal BasePath As String, _
ByVal RelativePath As String) As String
    Dim CurDirBuffer As String
    CurDirBuffer = Shell.CurrentDirectory

    Call Assert(fso.FolderExists(BasePath) Or fso.FileExists(BasePath), _
        "Error:AbsolutePath")
        
    Call Assert(IsDrivePath(BasePath) Or IsNetworkPath(BasePath), _
        "Error:AbsolutePath")

    Shell.CurrentDirectory = BasePath
    AbsolutePath = TrimLastSpace(fso.GetAbsolutePathName(RelativePath))
    '終端に改行コードが含まれる場合があるので削除する
    Shell.CurrentDirectory = CurDirBuffer
    
End Function

Private Sub testAbsolutePath()
    Shell.CurrentDirectory = "C:\Program Files"

    Call Check("C:\Program Files", AbsolutePath("C:\", ".\Program Files"))
    Call Check("C:\", AbsolutePath("C:\Program Files", "..\"))
    Call Check("C:\Windows", AbsolutePath("C:\", ".\Program Files\..\Windows"))

    'ネットワークパスの場合、大小文字が一致しない時がある
    Call Check(LCase("\\vmware-host\Shared Folders\この Mac 上の satoshi_yamamoto\MyFolder\MyData"), _
        LCase(AbsolutePath("\\vmware-host\Shared Folders\この Mac 上の satoshi_yamamoto\MyFolder", ".\MyData")))

    '"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    Call Check("C:\Program Files (x86)\Google\Chrome", AbsolutePath("C:\Program Files (x86)\Microsoft Office\root\Office16", "..\..\..\Google\Chrome"))
    
    '存在しないフォルダでも相対アドレスで指定できた
    Call Check("C:\Program Files (x86)\abc\def", AbsolutePath("C:\Program Files (x86)\Microsoft Office\root\Office16", "..\..\..\abc\def"))
    
    '先頭にピリオドがなくても指定できる
    Call Check("C:\Program Files (x86)\Microsoft Office\root\Office16\abc\def", _
        AbsolutePath("C:\Program Files (x86)\Microsoft Office\root\Office16", "abc\def"))


    
    Shell.CurrentDirectory = "\\vmware-host\Shared Folders"
    
    Call Check("C:\Program Files", AbsolutePath("C:\", ".\Program Files"))
    Call Check("C:\", AbsolutePath("C:\Program Files", "..\"))
    Call Check("C:\Windows", AbsolutePath("C:\", ".\Program Files\..\Windows"))
    Call Check(LCase("\\vmware-host\Shared Folders\この Mac 上の satoshi_yamamoto\MyFolder\MyData"), _
        LCase(AbsolutePath("\\vmware-host\Shared Folders\この Mac 上の satoshi_yamamoto\MyFolder", ".\MyData")))
    Call Check("C:\Program Files (x86)\Google\Chrome", AbsolutePath("C:\Program Files (x86)\Microsoft Office\root\Office16", "..\..\..\Google\Chrome"))
    Call Check("C:\Program Files (x86)\abc\def", AbsolutePath("C:\Program Files (x86)\Microsoft Office\root\Office16", "..\..\..\abc\def"))
    Call Check("C:\Program Files (x86)\Microsoft Office\root\Office16\abc\def", _
        AbsolutePath("C:\Program Files (x86)\Microsoft Office\root\Office16", "abc\def"))

End Sub

'----------------------------------------
'プログラムの設定などでパスを取得する関数
'----------------------------------------
'   ・  相対アドレスなどに対応
'----------------------------------------
Public Function SettingFullPath( _
ByVal SettingPath As String, _
Optional ByVal BasePath As String = "") As String
    Dim Result As String

    If SettingPath = "" Then
        Result = ThisWorkbook.Path
    Else
        If BasePath = "" Then BasePath = ThisWorkbook.Path

        If IsDrivePath(BasePath) Then
            'ファイルダイアログを開いた後
            'カレントディレクトリが変になる場合があるので
            'カレントディレクトリをリセットする
            Call ChDrive(ExcludeLastStr(BasePath, ":\"))
            Call ChDir(BasePath)

            Result = AbsolutePath(BasePath, SettingPath)
        Else
            Result = SettingPath
        End If
    End If
    SettingFullPath = Result
End Function

'----------------------------------------
'・ファイルが作成されるのをしばらく待つ関数
'----------------------------------------
'   ・  作成されたらTrueを返す
'----------------------------------------
Public Function FileExistsWait(ByVal FilePath As String, _
Optional ByVal ExistsFlag As Boolean = True) As Boolean
    FileExistsWait = False
    Dim I As Long: I = 0
    Do While (fso.FileExists(FilePath) = Not ExistsFlag)
        I = I + 1
        If I = 10 Then Exit Function
    Loop
    FileExistsWait = True
End Function


'----------------------------------------
'・ファイルコピー上書き失敗を検知するための関数
'----------------------------------------
'   ・  Success:=True / Fail:=False
'----------------------------------------
Public Function CopyFile( _
ByVal SourceFilePath, ByVal DestFilePath) As Boolean
On Error GoTo Err:
    Call fso.CopyFile(SourceFilePath, DestFilePath, True)
    CopyFile = True
    Exit Function
Err:
    CopyFile = False
End Function


'----------------------------------------
'◇空フォルダの削除
'----------------------------------------

'----------------------------------------
'・ 指定したフォルダにファイルや下位フォルダがない場合は削除する関数
'----------------------------------------
Public Sub Folder_DeleteIfNoFile( _
ByVal FolderPath As String)

    If Folder_HasSubItem(FolderPath) = False Then
        Call fso.DeleteFolder(FolderPath)
    End If

End Sub

'----------------------------------------
'・ 指定したフォルダにファイルや下位フォルダがない場合は削除し
'   さらに上位フォルダを見て下位が無い場合は削除していく関数
'----------------------------------------
Public Sub Folder_DeleteIfNoFileToUpFolder( _
ByVal FolderPath As String, _
Optional BaseFolderPath As String = "")

    Do
        If BaseFolderPath = FolderPath Then
            Exit Do
        End If
        Call Folder_DeleteIfNoFile(FolderPath)
        If fso.FolderExists(FolderPath) = False Then
            FolderPath = fso.GetParentFolderName(FolderPath)
        Else
            Exit Do
        End If
    Loop While True

End Sub


'----------------------------------------
'◇子項目関連処理
'----------------------------------------

'----------------------------------------
'・ 指定したフォルダのファイルや下位フォルダの有無を調べる関数
'----------------------------------------
Public Function Folder_HasSubItem( _
ByVal FolderPath As String) As Boolean

    Dim Result As Boolean: Result = True
    Call Assert(fso.FolderExists(FolderPath), "Error:Folder_HasSubItem")
    
    If FilePathListTopFolder(FolderPath) = "" Then
        If FolderPathListTopFolder(FolderPath) = "" Then
            Result = False
        End If
    End If
    
    Folder_HasSubItem = Result
End Function

'----------------------------------------
'・ 指定したフォルダにファイルや下位フォルダがある場合は削除する関数
'----------------------------------------
'   ・  戻り値、成功=True/失敗=False
'----------------------------------------
Public Function Folder_DeleteSubItem( _
ByVal FolderPath As String) As Boolean

On Error GoTo ErrorLabel
    Call Assert(fso.FolderExists(FolderPath), "Error:Folder_HasSubItem")
    
    Dim I As Long
    
    Dim FileList() As String
    FileList = Split(FilePathListTopFolder(FolderPath), vbCrLf)
    For I = 0 To ArrayCount(FileList) - 1
        Call fso.DeleteFile(FileList(I))
    Next
    
    Dim FolderList() As String
    FolderList = Split(FolderPathListTopFolder(FolderPath), vbCrLf)

    For I = 0 To ArrayCount(FolderList) - 1
        Call fso.DeleteFolder(FolderList(I), True)
    Next
    
    Folder_DeleteSubItem = True
    Exit Function
ErrorLabel:
    Folder_DeleteSubItem = False
End Function


'----------------------------------------
'◇Force/Recrate
'----------------------------------------

'----------------------------------------
'・深い階層のフォルダでも一気に作成する関数
'----------------------------------------
Public Sub ForceCreateFolder(ByVal FolderPath As String)
    Dim ParentFolderPath As String
    ParentFolderPath = fso.GetParentFolderName(FolderPath)
    If fso.FolderExists(ParentFolderPath) = False Then
        Call ForceCreateFolder(ParentFolderPath)
    End If
    If fso.FolderExists(FolderPath) = False Then
        Call fso.CreateFolder(FolderPath)
    End If
End Sub

'----------------------------------------
'・フォルダを再生成する関数
'----------------------------------------
Public Sub ReCreateFolder( _
ByVal FolderPath As String)

    If fso.FolderExists(FolderPath) Then
        Call fso.DeleteFolder(FolderPath)
    End If

    'フォルダが消えるまでループ
    Do: Loop While fso.FolderExists(FolderPath)

    On Error Resume Next
    Do
        Call ForceCreateFolder(FolderPath)
    Loop Until fso.FolderExists(FolderPath)
    'フォルダが作成できるまでループ
End Sub

'----------------------------------------
'◆ファイルフォルダ列挙
'----------------------------------------

Sub testGetFilePathListTopFolder()
    Call MsgBox(FilePathListTopFolder(AbsolutePath(ThisWorkbook.Path, "..\..\")))
    Call MsgBox(FolderPathListTopFolder(AbsolutePath(ThisWorkbook.Path, "..\..\")))
    Call MsgBox(FolderPathListSubFolder(AbsolutePath(ThisWorkbook.Path, "..\..\")))
    Call MsgBox(FilePathListSubFolder(AbsolutePath(ThisWorkbook.Path, "..\..\")))
End Sub

Sub testFileFolderPathList()
    Dim Path As String: Path = AbsolutePath( _
        ThisWorkbook.Path, ".\Test\TestFileFolderPathList")
    Dim PathList

    PathList = Replace(UCase(FolderPathListTopFolder(Path)), UCase(Path), "")
'    Call MsgBox(PathList)
    Call Check(PathList, _
        StringCombine(vbCrLf, _
            "\AAA", _
            "\BBB" _
        ))

    PathList = Replace(UCase(FolderPathListSubFolder(Path)), UCase(Path), "")
'    Call MsgBox(PathList)
    Call Check(PathList, _
        StringCombine(vbCrLf, _
            "\AAA", _
            "\AAA\AAA-1", _
            "\AAA\AAA-2", _
            "\AAA\AAA-2\AAA-2-1", _
            "\AAA\AAA-2\AAA-2-2", _
            "\AAA\AAA-2\AAA-2-2\AAA-2-2-1", _
            "\BBB", _
            "\BBB\BBB-1", _
            "\BBB\BBB-1\BBB-1-1", _
            "\BBB\BBB-1\BBB-1-1\BBB-1-1-1", _
            "\BBB\BBB-1\BBB-1-2", _
            "\BBB\BBB-2", "" _
        ))

    PathList = Replace(UCase(FilePathListTopFolder(Path)), UCase(Path), "")
'    Call MsgBox(PathList)
    Call Check(PathList, _
        StringCombine(vbCrLf, _
            "\AAA.TXT", _
            "\BBB.TXT" _
        ))

    PathList = Replace(UCase(FilePathListSubFolder(Path)), UCase(Path), "")
'    Call MsgBox(PathList)
    Call Check(PathList, _
        StringCombine(vbCrLf, _
            "\AAA\AAA-1.TXT", _
            "\AAA\AAA-2.TXT", _
            "\AAA\AAA-1\AAA-1-1.TXT", _
            "\AAA\AAA-2\AAA-2-1.TXT", _
            "\AAA\AAA-2\AAA-2-2.TXT", _
            "\AAA\AAA-2\AAA-2-1\AAA-2-1-1.TXT", _
            "\AAA\AAA-2\AAA-2-2\AAA-2-2-1.TXT", _
            "\AAA\AAA-2\AAA-2-2\AAA-2-2-1\AAA-2-2-1-1.TXT", _
            "\BBB\BBB-1.TXT", _
            "\BBB\BBB-2.TXT", _
            "\BBB\BBB-1\BBB-1-1.TXT", _
            "\BBB\BBB-1\BBB-1-2.TXT", _
            "\BBB\BBB-1\BBB-1-1\BBB-1-1-1.TXT", _
            "\BBB\BBB-1\BBB-1-1\BBB-1-1-1\BBB-1-1-1-1.TXT", _
            "\BBB\BBB-1\BBB-1-2\BBB-1-2-1.TXT", _
            "\BBB\BBB-2\BBB-2-1.TXT", _
            "\AAA.TXT", _
            "\BBB.TXT" _
        ))
End Sub


'----------------------------------------
'◇フォルダ
'----------------------------------------

'----------------------------------------
'◇トップレベルのフォルダリストを取得
'----------------------------------------
'・ 存在しなければ空文字を返す。
'・ パスは改行コードで区切られている
'----------------------------------------
Public Function FolderPathListTopFolder(ByVal FolderPath As String) As String
    Call Assert(fso.FolderExists(FolderPath), _
        "Error:FolderPathListTopFolder:Folder no Exists")
    Dim Result As String: Result = ""
    Dim SubFolder As Folder
    For Each SubFolder In fso.GetFolder(FolderPath).SubFolders
        Result = StringCombine(vbCrLf, Result, SubFolder.Path)
    Next
    FolderPathListTopFolder = Result
End Function

'----------------------------------------
'◇サブフォルダのフォルダリストを取得
'----------------------------------------
'・ 存在しなければ空文字を返す。
'・ パスは改行コードで区切られている
'----------------------------------------
Function FolderPathListSubFolder(FolderPath As String) As String
    Call Assert(fso.FolderExists(FolderPath), _
        "Error:FolderPathListSubFolder:Folder no Exists")
    Dim Result As String: Result = ""
    Dim SubFolder As Folder
    For Each SubFolder In fso.GetFolder(FolderPath).SubFolders
        Result = StringCombine(vbCrLf, _
            Result, SubFolder.Path, _
            FolderPathListSubFolder(SubFolder.Path))
    Next
    FolderPathListSubFolder = Result
End Function

'----------------------------------------
'◇ファイル
'----------------------------------------

'----------------------------------------
'◇トップレベルのファイルリストを取得
'----------------------------------------
'・ 存在しなければ空文字を返す。
'・ パスは改行コードで区切られている
'----------------------------------------
Function FilePathListTopFolder(FolderPath As String) As String
    Call Assert(fso.FolderExists(FolderPath), _
        "Error:FilePathListTopFolder:Folder no Exists")
    Dim Result As String: Result = ""
    Dim File As File
    For Each File In fso.GetFolder(FolderPath).Files
        Result = StringCombine(vbCrLf, Result, File.Path)
    Next
    FilePathListTopFolder = ExcludeLastStr(Result, vbCrLf)
End Function

'----------------------------------------
'◇サブフォルダのファイルリストを取得
'存在しなければ空文字を返す。
'パスの最後には必ず改行コードが付属
Function FilePathListSubFolder(FolderPath As String) As String
    Call Assert(fso.FolderExists(FolderPath), _
        "Error:FilePathListSubFolder:Folder no Exists")
    Dim Result As String: Result = ""
    Dim FolderPathList() As String
    FolderPathList = Split( _
        FolderPathListSubFolder(FolderPath) + vbCrLf + FolderPath, vbCrLf)
    Dim I As Long
    For I = 0 To ArrayCount(FolderPathList) - 1
        If fso.FolderExists(FolderPathList(I)) Then
            Result = StringCombine(vbCrLf, _
                Result, FilePathListTopFolder(FolderPathList(I)))
        End If
    Next
    FilePathListSubFolder = ExcludeLastStr(Result, vbCrLf)
End Function

'----------------------------------------
'◆ファイル日時
'----------------------------------------

'----------------------------------------
'・UTCファイルタイム変換関数
'----------------------------------------
Private Function DateToApiFILETIME(ByVal datPARAM As Date) As FILETIME
    Dim Result As FILETIME
    Dim SysTime As SYSTEMTIME
    Dim LocalTime As FILETIME

    SysTime.wYear = Year(datPARAM)
    SysTime.wMonth = Month(datPARAM)
    SysTime.wDayOfWeek = Weekday(datPARAM)
    SysTime.wDay = Day(datPARAM)
    SysTime.wHour = Hour(datPARAM)
    SysTime.wMinute = Minute(datPARAM)
    SysTime.wSecond = Second(datPARAM)
    SysTime.wMilliseconds = 0

    Call SystemTimeToFileTime(SysTime, LocalTime)
    Call LocalFileTimeToFileTime(LocalTime, Result)

    DateToApiFILETIME = Result
End Function

'----------------------------------------
'・ファイル/フォルダの作成日時/更新日時/最終アクセス日時の取得
'----------------------------------------
Public Function GetFileFolderTime( _
ByVal Path As String) As FileFolderTime

    Dim Result As FileFolderTime
    If fso.FileExists(Path) Then
        Dim File As File
        Set File = fso.GetFile(Path)
        Result.CreataionTime = File.DateCreated
        Result.LastWriteTime = File.DateLastModified
        Result.LastAccessTime = File.DateLastAccessed
    ElseIf fso.FolderExists(Path) Then
        Dim Folder As Folder
        Set Folder = fso.GetFolder(Path)
        Result.CreataionTime = Folder.DateCreated
        Result.LastWriteTime = Folder.DateLastModified
        Result.LastAccessTime = Folder.DateLastAccessed
    Else
        Call Assert(False, "Error:GetFileFolderTime")
    End If
    GetFileFolderTime = Result
End Function

'----------------------------------------
'・ファイル/フォルダの作成日時/更新日時/最終アクセス日時の設定
'----------------------------------------
Public Function SetFileFolderTime( _
ByVal Path As String, _
FileFolderTime As FileFolderTime) As Boolean

    Dim Result As Boolean: Result = False

    Dim FileHandle As Long
    Dim CreateFileFlag   As Long
    Dim ReturnSetFileTime    As Long
    Dim CreateFILETIME As FILETIME
    Dim AccessFILETIME As FILETIME
    Dim ModifyFILETIME As FILETIME
    Dim SecurityAttr As SECURITY_ATTRIBUTES

    Do
        '// 対象の存在チェックとdwFlagsAndAttributes の設定
        If fso.FileExists(Path) Then
            'ファイルの場合
            CreateFileFlag = FILE_ATTRIBUTE_NORMAL
        ElseIf fso.FolderExists(Path) Then
            'フォルダの場合(NT系のOSのみ可能)
            If InStr(Application.OperatingSystem, "NT") > 0 Then
                CreateFileFlag = FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_BACKUP_SEMANTICS
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If

        Dim SetTime As FileFolderTime
        SetTime = GetFileFolderTime(Path)

        '// オプション引数が省略された場合は現状のものを補完
        If FileFolderTime.CreataionTime <> 0 Then
            SetTime.CreataionTime = FileFolderTime.CreataionTime
        End If
        If FileFolderTime.LastWriteTime <> 0 Then
            SetTime.LastWriteTime = FileFolderTime.LastWriteTime
        End If
        If FileFolderTime.LastAccessTime <> 0 Then
            SetTime.LastAccessTime = FileFolderTime.LastAccessTime
        End If

        '// SECURITY_ATTRIBUTES構造体初期化

        SecurityAttr.nLength = LenB(SecurityAttr)
        SecurityAttr.lpSecurityDescriptor = 0&
        SecurityAttr.bInheritHandle = 0&


        '// ファイルまたはフォルダハンドルを取得
        FileHandle = CreateFile(Path, GENERIC_WRITE, _
            FILE_SHARE_READ, SecurityAttr, OPEN_EXISTING, CreateFileFlag, vbNull)
        If FileHandle = INVALID_HANDLE_VALUE Then Exit Do

        '// ファイルタイムに変換し、設定する
        CreateFILETIME = DateToApiFILETIME(SetTime.CreataionTime)
        AccessFILETIME = DateToApiFILETIME(SetTime.LastAccessTime)
        ModifyFILETIME = DateToApiFILETIME(SetTime.LastWriteTime)
        ReturnSetFileTime = SetFileTime(FileHandle, CreateFILETIME, AccessFILETIME, ModifyFILETIME)
        If ReturnSetFileTime <> 0 Then
            Result = True
        End If

        '// ファイルまたはフォルダハンドル開放
        Call CloseHandle(FileHandle)
    Loop While False

    SetFileFolderTime = Result
End Function


'----------------------------------------
'◆ショートカットファイル操作
'----------------------------------------

'----------------------------------------
'・ショートカットファイル判定(拡張子)
'----------------------------------------

Public Function IsShortcutLinkFile(ByVal FilePath As String)
    Dim Result As Boolean: Result = False
    If LCase(GetExtensionIncludePeriod(FilePath)) = ".lnk" Then
        Result = True
    End If
    IsShortcutLinkFile = Result
End Function

'----------------------------------------
'・ショートカットファイルの作成
'----------------------------------------
Public Sub CreateShortcutFile( _
ByVal ShortcutFilePath As String, _
ByVal TargetFilePath As String, _
ByVal IconFilePath As String, _
ByVal Description As String)

    Dim ShortcutFile As IWshRuntimeLibrary.WshShortcut
    Set ShortcutFile = Shell.CreateShortcut(ShortcutFilePath)
    ShortcutFile.TargetPath = TargetFilePath
    ShortcutFile.Description = Description
    ShortcutFile.IconLocation = IconFilePath
    ShortcutFile.RelativePath = ""
    ShortcutFile.WorkingDirectory = ""
    ShortcutFile.Hotkey = ""
    ShortcutFile.Save
End Sub

'----------------------------------------
'・ショートカットファイルの作成/削除
'----------------------------------------
Public Sub SetShortcutIcon(ByVal Value As Boolean, _
ByVal ShortcutFilePath As String, ByVal LinkTargetFilePath As String, _
ByVal IconFilePath As String, _
ByVal Description As String, _
ByVal FolderDeleteFlag As Boolean)

    Dim ShortcutFileParentFolderPath As String
    ShortcutFileParentFolderPath = fso.GetParentFolderName(ShortcutFilePath)

    Dim FileExistsFlag As Boolean
    FileExistsFlag = fso.FileExists(ShortcutFilePath)
    If (Value) And (FileExistsFlag = False) Then
        Call ForceCreateFolder(ShortcutFileParentFolderPath)
        Call CreateShortcutFile(ShortcutFilePath, LinkTargetFilePath, _
            IconFilePath, Description)
    ElseIf (Value = False) And (FileExistsFlag) Then
        Call fso.DeleteFile(ShortcutFilePath)
        '↓フラグONなら空フォルダになった場合はフォルダ削除する
        If FolderDeleteFlag _
        And fso.GetFolder(ShortcutFileParentFolderPath).SubFolders.Count = 0 Then
            Call fso.DeleteFolder(ShortcutFileParentFolderPath)
        End If
    End If
End Sub


'----------------------------------------
'◆Iniファイル処理
'----------------------------------------
Public Function IniFile_GetString(ByVal Path As String, _
ByVal Section As String, ByVal Name As String, _
Optional ByVal DefaultValue As String = "") As String
    Dim Result As String

    ' 値を取得するバッファを確保する
    Dim ReturnValue As String * 256

    If 0 < GetPrivateProfileString(Section, Name, DefaultValue, _
        ReturnValue, Len(ReturnValue), Path) Then
        If InStr(ReturnValue, Chr$(0)) > 0 Then
            Result = FirstStrFirstDelim(ReturnValue, Chr$(0))
        Else
            Result = ReturnValue
        End If
    Else
        Result = DefaultValue
    End If

    IniFile_GetString = Result
End Function

Private Sub testIniFile_GetString()
    Dim Value As String
    Value = _
        IniFile_GetString( _
            ThisWorkbook.Path + Application.PathSeparator + "test.ini", _
            "Option", "Name", "Defalut")
    MsgBox Value
End Sub

Public Sub IniFile_SetString(ByVal Path As String, _
ByVal Section As String, ByVal Name As String, _
ByVal Value As String)
    Call WritePrivateProfileString(Section, Name, Value, Path)
End Sub

Private Sub testIniFile_SetString()
    Call IniFile_SetString( _
        ThisWorkbook.Path + Application.PathSeparator + "test.ini", _
        "Option", "Name", "TestValue01")
End Sub

'----------------------------------------
'◆テキストファイル読み書き
'----------------------------------------

Public Function CheckEncodeName(EncodeName As String) As Boolean
    CheckEncodeName = OrValue(UCase$(EncodeName), _
        "SHIFT_JIS", _
        "UNICODE", "UNICODEFFFE", "UTF-16", _
        "UTF-16LE", _
        "UNICODEFEFF", _
        "UTF-16BE", _
        "UTF-8", _
        "ISO-2022-JP", _
        "EUC-JP", _
        "UTF-7")
End Function

'----------------------------------------
'・テキストファイル読込
'----------------------------------------
'   ・  エンコード指定は下記の通り
'           エンコード          指定文字
'           ShiftJIS            SHIFT_JIS
'           UTF-16LE BOM有/無   UNICODEFFFE/UNICODE/UTF-16/UTF-16LE
'                               BOMの有無に関わらず読込可能
'           UTF-16BE _BOM_ON    UNICODEFEFF
'           UTF-16BE _BOM_OFF   UTF-16BE
'           UTF-8 BOM有/無      UTF-8/UTF-8N
'                               BOMの有無に関わらず読込可能
'           JIS                 ISO-2022-JP
'           EUC-JP              EUC-JP
'           UTF-7               UTF-7
'   ・  UTF-16LEとUTF-8は、BOMの有無にかかわらず読み込める
'----------------------------------------
Public Function ADOStream_LoadTextFile( _
ByVal TextFilePath As String, ByVal EncodeName As String) As String
    If CheckEncodeName(EncodeName) = False Then
        Call Assert(False, "Error:ADOStream_LoadTextFile")
    End If

    Dim ADOStream As ADODB.Stream
    Set ADOStream = New ADODB.Stream
    ADOStream.Type = adTypeText
    ADOStream.Charset = EncodeName
    ADOStream.Open
    ADOStream.LoadFromFile (TextFilePath)
    ADOStream_LoadTextFile = ADOStream.ReadText
    ADOStream.Close
End Function

'----------------------------------------
'・テキストファイル保存
'----------------------------------------
'   ・  エンコード指定は下記の通り
'           エンコード          指定文字
'           ShiftJIS            SHIFT_JIS
'           UTF-16LE _BOM_ON    UNICODEFFFE/UNICODE/UTF-16
'           UTF-16LE _BOM_OFF    UTF-16LE
'           UTF-16BE _BOM_ON    UNICODEFEFF
'           UTF-16BE _BOM_OFF    UTF-16BE
'           UTF-8 _BOM_ON       UTF-8
'           UTF-8 _BOM_OFF       UTF-8N
'           JIS                 ISO-2022-JP
'           EUC-JP              EUC-JP
'           UTF-7               UTF-7
'   ・  UTF-16LEとUTF-8はそのままだと_BOM_ONになるので
'       BON無し指定の場合は特殊処理をしている
'----------------------------------------
Public Sub ADOStream_SaveTextFile(ByVal Text As String, _
ByVal TextFilePath As String, ByVal EncodeName As String, _
Optional ByVal BOM As Boolean = True)
    If CheckEncodeName(EncodeName) = False Then
        Call Assert(False, "Error:ADOStream_LoadTextFile")
    End If

    Dim ADOStream As New ADODB.Stream
    ADOStream.Type = adTypeText
    ADOStream.Charset = EncodeName
    ADOStream.Open
    Call ADOStream.WriteText(Text)

    Dim ByteData() As Byte
    Select Case UCase$(EncodeName)
    Case "UNICODE", "UNICODEFFFE", "UTF-16LE", "UTF-16"
        If BOM = False Then
            ADOStream.Position = 0
            ADOStream.Type = adTypeBinary
            ADOStream.Position = 2
            ByteData = ADOStream.Read
            ADOStream.Close
            ADOStream.Open
            Call ADOStream.Write(ByteData)
        End If
    Case "UTF-8"
        If BOM = False Then
            ADOStream.Position = 0
            ADOStream.Type = adTypeBinary
            ADOStream.Position = 3
            ByteData = ADOStream.Read
            ADOStream.Close
            ADOStream.Open
            Call ADOStream.Write(ByteData)
        End If
    End Select
    Call ADOStream.SaveToFile(TextFilePath, adSaveCreateOverWrite)
    ADOStream.Close
End Sub

'----------------------------------------
'◇テキストファイル読み書き Enum指定版
'----------------------------------------
'   ・  ADODB.Stream が許容する文字列は次の通り
'           使用可能文字列                  エンコード
'           SHIFT_JIS
'           UNICODEFFFE/UNICODE/UTF-16      UTF-16LE_BOM_ON
'           UTF-16LE                        UTF-16LE_BOM_OFF
'           UNICODEFEFF                     UTF-16BE_BOM_ON
'           UTF-16BE                        UTF-16BE_BOM_OFF
'           UTF-8
'           ISO-2022-JP
'           EUC-JP
'           UTF-7
'       このうち、変わった挙動をするのは UTF-16LE と UFT-8
'
'       UTF-16LEは、読み込み時に指定すると
'       テキストファイルが UTF-16LE のBOMありなし関わらず読み込み可能
'       これは特に問題にはならないのだが
'       書き込み時に UTF-16LE を指定しても
'       BOMありとして書き込まれてしまう。
'       つまり、UTF-16LEは UNICODEFFFE/UNICODE/UTF-16 と同じ機能になる
'       それでは機能不足なので、String_SaveToFile ではBOMを除外する処理をしている。
'
'       UFT-8は、常にBOMありとして書き込まれる。
'       BOM無しのUTF-8なんて世の中に存在しない方がいいのだが
'       そうはいっても、UTF-8BOM無しで出力したい場合もあるので
'       ADODB.Stream はBOM無しUTF-8を許容しないのだが
'       UFT-8N という文字列によって
'       UTF-8のBOMなしの文字として表現して、
'       String_SaveToFile ではBOMを除外する処理をしている。
'----------------------------------------
Public Function GetEncodingTypeJpCharCode( _
ByVal EncodingTypeName As String) As EncodingTypeJpCharCode
    
    Dim Result As EncodingTypeJpCharCode
    Result = EncodingTypeJpCharCode.NONE
    Select Case UCase(EncodingTypeName)
    Case "SHIFT_JIS"
        Result = EncodingTypeJpCharCode.Shift_JIS
    
    Case "UNICODE", "UNICODEFFFE", "UTF-16"
        Result = EncodingTypeJpCharCode.UTF16_LE_BOM
    Case "UTF-16LE"
        Result = EncodingTypeJpCharCode.UTF16_LE_BOM_NO
    
    Case "UNICODEFEFF"
        Result = EncodingTypeJpCharCode.UTF16_BE_BOM
    Case "UTF-16BE"
        Result = EncodingTypeJpCharCode.UTF16_BE_BOM_NO
    
    Case "UTF-8"
        Result = EncodingTypeJpCharCode.UTF8_BOM
    Case "UTF-8N"
        Result = EncodingTypeJpCharCode.UTF8_BOM_NO
    
    Case "ISO-2022-JP"
        Result = EncodingTypeJpCharCode.JIS
        
    Case "EUC-JP"
        Result = EncodingTypeJpCharCode.EUC_JP
    
    Case "UTF-7"
        Result = EncodingTypeJpCharCode.UTF_7
    
    End Select
    
End Function


Public Function GetEncodingTypeName( _
ByVal EncodingType As EncodingTypeJpCharCode) As String
    Dim Result As String: Result = ""
    
    Select Case EncodingType
    Case EncodingTypeJpCharCode.Shift_JIS
        Result = "SHIFT_JIS"
    
    Case EncodingTypeJpCharCode.UTF16_LE_BOM
        Result = "UNICODEFFFE"
    Case EncodingTypeJpCharCode.UTF16_LE_BOM_NO
        Result = "UTF-16LE"
    
    Case EncodingTypeJpCharCode.UTF16_BE_BOM
        Result = "UNICODEFEFF"
    Case EncodingTypeJpCharCode.UTF16_BE_BOM_NO
        Result = "UTF-16BE"
    
    Case EncodingTypeJpCharCode.UTF8_BOM
        Result = "UTF-8"
    Case EncodingTypeJpCharCode.UTF8_BOM_NO
        Result = "UTF-8N"
    
    Case EncodingTypeJpCharCode.JIS
        Result = "ISO-2022-JP"
        
    Case EncodingTypeJpCharCode.EUC_JP
        Result = "EUC-JP"
    
    Case EncodingTypeJpCharCode.UTF_7
        Result = "UTF-7"
    
    End Select
    GetEncodingTypeName = Result
End Function

Public Function String_LoadFromFile( _
ByVal FilePath As String, _
ByVal EncodingType As EncodingTypeJpCharCode) As String

    Dim EncordingName As String
    EncordingName = GetEncodingTypeName(EncodingType)
    Call Assert(EncordingName <> "", "Error:Encoding No Select")

    Dim Stream As New ADODB.Stream
    Stream.Type = adTypeText
    Select Case EncodingType
    Case EncodingTypeJpCharCode.UTF8_BOM_NO
        Stream.Charset = GetEncodingTypeName(EncodingTypeJpCharCode.UTF8_BOM)
    Case Else
        Stream.Charset = EncordingName
    End Select
    Stream.Open
    Stream.LoadFromFile (FilePath)
    String_LoadFromFile = Stream.ReadText
    Stream.Close
   
End Function

Private Sub testString_LoadFromFile()
    Dim FolderPath As String
    FolderPath = PathCombine( _
        ThisWorkbook.Path, "Test", "ADOStream")
    Call ForceCreateFolder(FolderPath)

    Call Assert("Shift-JIS ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_Shift-JIS.txt"), _
            EncodingTypeJpCharCode.Shift_JIS))

    Call Assert("UTF-16LE-BOM ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
            EncodingTypeJpCharCode.UTF16_LE_BOM))
    Call Assert("UTF-16LE-BOM-NO ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
            EncodingTypeJpCharCode.UTF16_LE_BOM_NO))
    
    Call Assert("UTF-16BE-BOM ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
            EncodingTypeJpCharCode.UTF16_BE_BOM))
    Call Assert("UTF-16BE-BOM-NO ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
            EncodingTypeJpCharCode.UTF16_BE_BOM_NO))
        
    Call Assert("UTF-8-BOM ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
            EncodingTypeJpCharCode.UTF8_BOM))
    Call Assert("UTF-8-BOM-NO ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
            EncodingTypeJpCharCode.UTF8_BOM_NO))
            
    Call Assert("JIS ISO-2022-JP ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_JIS.txt"), _
            EncodingTypeJpCharCode.JIS))
        
    Call Assert("EUC-JP ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_EUC-JP.txt"), _
            EncodingTypeJpCharCode.EUC_JP))
        
    Call Assert("UTF-7 ＡＢＣ１２３" = _
        String_LoadFromFile( _
            PathCombine(FolderPath, "test_UTF-7.txt"), _
            EncodingTypeJpCharCode.UTF_7))
End Sub

Public Sub String_SaveToFile( _
ByVal Text As String, _
ByVal FilePath As String, _
ByVal EncodingType As EncodingTypeJpCharCode)

    Dim EncordingName As String
    EncordingName = GetEncodingTypeName(EncodingType)
    Call Assert(EncordingName <> "", "Error:Encoding No Select")

    Dim Stream As New ADODB.Stream
    Stream.Type = adTypeText
    Select Case EncodingType
    Case EncodingTypeJpCharCode.UTF8_BOM_NO
        Stream.Charset = GetEncodingTypeName(EncodingTypeJpCharCode.UTF8_BOM)
    Case Else
        Stream.Charset = EncordingName
    End Select
    Stream.Open
    Call Stream.WriteText(Text)

    Dim ByteData() As Byte
    Select Case EncodingType
    Case EncodingTypeJpCharCode.UTF16_LE_BOM_NO
        Stream.Position = 0
        Stream.Type = adTypeBinary
        Stream.Position = 2
        ByteData = Stream.Read
        Stream.Close
        Stream.Open
        Call Stream.Write(ByteData)
    Case EncodingTypeJpCharCode.UTF8_BOM_NO
        Stream.Position = 0
        Stream.Type = adTypeBinary
        Stream.Position = 3
        ByteData = Stream.Read
        Stream.Close
        Stream.Open
        Call Stream.Write(ByteData)
    End Select
    Call Stream.SaveToFile(FilePath, adSaveCreateOverWrite)
    Stream.Close
End Sub

Private Sub testString_SaveToFile()
    Dim FolderPath As String
    FolderPath = PathCombine( _
        ThisWorkbook.Path, "Test", "ADOStream")
    Call ForceCreateFolder(FolderPath)

    Call String_SaveToFile( _
        "Shift-JIS ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_Shift-JIS.txt"), _
        EncodingTypeJpCharCode.Shift_JIS)

    Call String_SaveToFile( _
        "UTF-16LE-BOM ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_UTF-16LE-BOM.txt"), _
        EncodingTypeJpCharCode.UTF16_LE_BOM)
    Call String_SaveToFile( _
        "UTF-16LE-BOM-NO ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_UTF-16LE-BOM-NO.txt"), _
        EncodingTypeJpCharCode.UTF16_LE_BOM_NO)
    
    Call String_SaveToFile( _
        "UTF-16BE-BOM ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_UTF-16BE-BOM.txt"), _
        EncodingTypeJpCharCode.UTF16_BE_BOM)
    Call String_SaveToFile( _
        "UTF-16BE-BOM-NO ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_UTF-16BE-BOM-NO.txt"), _
        EncodingTypeJpCharCode.UTF16_BE_BOM_NO)
        
    Call String_SaveToFile( _
        "UTF-8-BOM ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_UTF-8-BOM.txt"), _
        EncodingTypeJpCharCode.UTF8_BOM)
    Call String_SaveToFile( _
        "UTF-8-BOM-NO ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_UTF-8-BOM-NO.txt"), _
        EncodingTypeJpCharCode.UTF8_BOM_NO)
        
    Call String_SaveToFile( _
        "JIS ISO-2022-JP ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_JIS.txt"), _
        EncodingTypeJpCharCode.JIS)
        
    Call String_SaveToFile( _
        "EUC-JP ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_EUC-JP.txt"), _
        EncodingTypeJpCharCode.EUC_JP)
        
    Call String_SaveToFile( _
        "UTF-7 ＡＢＣ１２３", _
        PathCombine(FolderPath, "test_UTF-7.txt"), _
        EncodingTypeJpCharCode.UTF_7)
        
End Sub

'----------------------------------------
'◆画像ファイル
'----------------------------------------


'----------------------------------------
'・Jpegファイル判定(拡張子)
'----------------------------------------
Public Function IsJpegImageFile(ByVal FilePath As String)
    Dim Result As Boolean: Result = False
    If OrValue(LCase(GetExtensionIncludePeriod(FilePath)), ".jpg", ".jpeg") Then
        Result = True
    End If
    IsJpegImageFile = Result
End Function

'----------------------------------------
'・JpegExif含むファイル判定
'----------------------------------------
'   ・  Exifの撮影日時取得可能かどうかを判定なので
'       ファイルが実際に存在することも確認される
'----------------------------------------
Public Function IsJpegExifFile(ByVal FilePath As String)
    Dim Result As Boolean: Result = False

    If IsJpegImageFile(FilePath) Then
        If GetJpegExifDateTime(FilePath) <> 0 Then
            Result = True
        End If
    End If

    IsJpegExifFile = Result
End Function

'----------------------------------------
'・JpegExif情報撮影日時取得
'----------------------------------------
'   ・  取得できない場合はCDate(0)を返す
'----------------------------------------
Public Function GetJpegExifDateTime(ByVal FilePath As String) As Date
On Error GoTo Err:
    Dim Result As Date: Result = 0
    If IsJpegImageFile(FilePath) Then

        Dim WIA_ImageFile As Object
        Set WIA_ImageFile = CreateObject("Wia.ImageFile")
        Call WIA_ImageFile.LoadFile(FilePath)


        '撮影日時
        Dim ExifDateTime As String
        ExifDateTime = WIA_ImageFile.Properties("36867")
        ExifDateTime = Replace(ExifDateTime, ":", "/", , 2)
        Result = CDate(ExifDateTime)
    End If
Err:
    GetJpegExifDateTime = Result
End Function


'----------------------------------------
'◆シェル起動
'----------------------------------------

'----------------------------------------
'・ コマンド実行
'----------------------------------------
'   ・  %ComSpec% /c は
'       いろんなところで推奨されているけど
'       Exeが長いネットワークパスにあるときは動かない場面があった
'       なので、DOSコマンドではなく実行ファイルを指定して動かすのなら
'       直接、Shell.Run(Command, のようにしたほうがいいかも
'----------------------------------------
Public Sub CommandExecute(Command As String)
    Dim Result As String: Result = ""

    Call Shell.Run( _
        "%ComSpec% /c " + Command, _
         VBA.VbAppWinStyle.vbHide, True)

End Sub

Private Sub testCommandExecute()
    Call CommandExecute("ping")
End Sub

'----------------------------------------
'・ コマンド実行後の結果取得
'----------------------------------------
Public Function CommandExecuteReturn(Command As String, _
Optional ByVal EncodeName As String = "Shift_JIS") As String
    Dim Result As String: Result = ""

    'テンポラリファイルパスを取得
    Const TemporaryFolder = 2
    Dim TempFilePath As String
    Do
        TempFilePath = fso.BuildPath( _
            fso.GetSpecialFolder(TemporaryFolder), fso.GetTempName)
    Loop While fso.FileExists(TempFilePath)

    Call Shell.Run( _
        "%ComSpec% /c " + Command + ">" + TempFilePath + " 2>&1", _
         VBA.VbAppWinStyle.vbHide, True)

    If fso.FileExists(TempFilePath) Then
        Result = ADOStream_LoadTextFile(TempFilePath, EncodeName)
        Kill TempFilePath
    End If

    CommandExecuteReturn = Result
End Function

Private Sub testCommandExecuteReturn()
    Call MsgBox(CommandExecuteReturn("ping"))
End Sub

'----------------------------------------
'◆クリップボード
'----------------------------------------
'   ・  参照設定[Microsoft Forms 2.0 Object Library]で
'       DataObjectが使用可能
'       Macでも可能
'----------------------------------------

'----------------------------------------
'・テキストデータ取得
'----------------------------------------
'   ・  Win/Mac両対応動作確認隋
'----------------------------------------
Public Function GetClipboardText()
    Dim DataObject1 As New MSForms.DataObject

    DataObject1.GetFromClipboard
    GetClipboardText = DataObject1.GetText
End Function

'----------------------------------------
'・テキストデータ設定
'----------------------------------------
'   ・  Win/Mac両対応動作確認隋
'----------------------------------------
Public Sub SetClipboardText(ByVal ClipboardToText)
    Dim DataObject1 As New MSForms.DataObject

    Call DataObject1.SetText(ClipboardToText)
    DataObject1.PutInClipboard
End Sub

Public Sub testGetSetClipboard()
    Call SetClipboardText("ABC")
    Call Check("ABC", GetClipboardText)
End Sub


'----------------------------------------
'◆Excel
'----------------------------------------

'----------------------------------------
'・進捗表示
'----------------------------------------
Public Sub Application_StatusBar_Progress(ByVal Message As String, _
ByVal Delimiter As String, _
ByVal StartValue As Long, ByVal Value As Long, ByVal EndValue As Long, _
Optional ReverseFlag As Boolean = False)

    Application.StatusBar = ProgressText( _
        Message, Delimiter, _
        StartValue, Value, EndValue, _
        ReverseFlag)

End Sub

Public Function ProgressText( _
ByVal Message As String, ByVal Delimiter As String, _
ByVal StartValue As Long, ByVal Value As Long, ByVal EndValue As Long, _
Optional ReverseFlag As Boolean = False, _
Optional PercentVisible As Boolean = True)
    Dim Result As String
    If ReverseFlag = False Then
        Result = _
            Message + Delimiter + _
            CStr(Value - StartValue + 1) + "/" + _
            CStr(EndValue - StartValue + 1)
        If PercentVisible Then
            Result = Result + Delimiter + _
            CStr(Format((Value - StartValue + 1) / (EndValue - StartValue + 1) * 100, "0.00")) + "%"
        End If
            
    Else
        Result = _
            Message + Delimiter + _
            CStr(Value - StartValue + 1) + "/" + _
            CStr(EndValue - StartValue + 1)
        If PercentVisible Then
            Result = Result + Delimiter + _
            CStr(Format(100 - ((Value - StartValue + 1) / (EndValue - StartValue + 1) * 100), "0.00")) + "%"
    End If
    End If
    ProgressText = Result
End Function


'----------------------------------------
'・列番号から列名を取得する
'----------------------------------------
Public Function ColumnText(ByVal ColumnNumber As Long) As String
    ColumnText = _
        FirstStrFirstDelim( _
            Application.Columns(ColumnNumber).Address(False, False, xlA1), _
            ":")
End Function

Private Sub testColumnText()
    Call Check("C", ColumnText(3))
    Call Check("AX", ColumnText(50))
End Sub

'----------------------------------------
'・列名(A,B,C,…)から列番号を取得する
'----------------------------------------
'   ・  A→1, B→2, …, Z→26, AA→27, AB→28
'----------------------------------------
'Function ColumnNumber(ColumnText As String) As Long
'    ColumnNumber = Columns(ColumnText).Column
'End Function

Public Function ColumnNumber(ColumnText As String) As Long
    Dim Result As Long: Result = 0
    Dim CharNumber As Long
    Dim I As Long
    For I = 0 To Len(ColumnText) - 1
        CharNumber = Asc(UCase(Mid(ColumnText, Len(ColumnText) - I, 1))) - 64
        If I = 0 Then
            Result = CharNumber
        Else
            Result = Result + (CharNumber * (I * 26))
        End If
    Next
    ColumnNumber = Result
End Function

Sub testColumnNumber()
    Call Check(ColumnNumber("A"), 1)
    Call Check(ColumnNumber("b"), 2)
    Call Check(ColumnNumber("Z"), 26)
    Call Check(ColumnNumber("AA"), 27)
    Call Check(ColumnNumber("AB"), 28)
End Sub


'----------------------------------------
'◇タイトル行/列指定処理
'----------------------------------------

'----------------------------------------
'・タイトル行の列名から列番号を返す関数
'----------------------------------------
'   ・  日本語タイトル行などに対してタイトル文字列で行番号を返す
'----------------------------------------
Public Function Sheet_ColumnNumberByTitle(ByRef Sheet As Worksheet, _
ByVal TitleRowIndex As Long, _
ByVal ColumnTitleWildCard As String, _
Optional TitleMatchCount As Long = 1)
    Dim Result As Long: Result = 0
    Dim Counter As Long: Counter = 0
    Dim I As Long
    For I = Col_A To Sheet_DataLastColumn(Sheet, TitleRowIndex)
        If Sheet.Cells(TitleRowIndex, I).Value Like ColumnTitleWildCard Then
            Counter = Counter + 1
            If Counter = TitleMatchCount Then
            Result = I
            Exit For
        End If
        End If
    Next
    Sheet_ColumnNumberByTitle = Result
End Function


'----------------------------------------
'・タイトル列の行名から行番号を返す関数
'----------------------------------------
'   ・  日本語タイトル行などに対してタイトル文字列で行番号を返す
'----------------------------------------
Public Function Sheet_RowNumberByTitle(ByRef Sheet As Worksheet, _
ByVal TitleColIndex As Long, _
ByVal RowTitleWildCard As String, _
Optional TitleMatchCount As Long = 1)
    Dim Result As Long: Result = 0
    Dim Counter As Long: Counter = 0
    Dim I As Long
    For I = 1 To Sheet_DataLastRow(Sheet, TitleColIndex)
        If Sheet.Cells(I, TitleColIndex).Value Like RowTitleWildCard Then
            Counter = Counter + 1
            If Counter = TitleMatchCount Then
            Result = I
            Exit For
        End If
        End If
    Next
    Sheet_RowNumberByTitle = Result
End Function

'----------------------------------------
'・タイトル列の行名から行番号を返す関数
'----------------------------------------
'   ・  Sheet_RowNumberByTitle に グループ名を追加した
'----------------------------------------
Public Function Sheet_RowNumberByGroupTitle(ByVal Sheet As Worksheet, _
ByVal GroupColIndex As Long, _
ByVal TitleColIndex As Long, _
ByVal GroupName As String, _
ByVal RowTitleWildCard As String)
    Dim Result As Long: Result = 0
    Dim I As Long
    For I = 1 To Sheet_DataLastRow(Sheet, TitleColIndex)
        If Sheet.Cells(I, GroupColIndex).Value = GroupName Then
            If Sheet.Cells(I, TitleColIndex).Value Like RowTitleWildCard Then
                Result = I
                Exit For
            End If
        End If
    Next
    Sheet_RowNumberByGroupTitle = Result
End Function


'----------------------------------------
'◇最終行/列
'----------------------------------------
'----------------------------------------
'・最終行/列
'----------------------------------------
'   ・  データがない場合は1を戻す
'----------------------------------------

'・データ最終行
Public Function Sheet_DataLastRow(ByVal Sheet As Worksheet, _
Optional ByVal ColumnNumber As Long = -1) As Long
On Error Resume Next
    Dim Result As Long: Result = 1
    Call Assert(-1 <= ColumnNumber, "Error:DataLastRow")
    If ColumnNumber = -1 Then
        Result = Sheet.UsedRange.Find("*", _
            , xlFormulas, , xlByRows, xlPrevious).Row
    Else
        Result = Sheet.Cells(Sheet.Rows.Count, ColumnNumber).End(xlUp).Row
    End If
    Sheet_DataLastRow = Result
End Function

'・データ最終列
Public Function Sheet_DataLastColumn(ByVal Sheet As Worksheet, _
Optional ByVal RowNumber As Long = -1) As Long
On Error Resume Next
    Dim Result As Long: Result = 1
    Call Assert(-1 <= RowNumber, "Error:DataLastCol")
    If RowNumber = -1 Then
        Result = Sheet.UsedRange.Find("*", _
            , xlFormulas, , xlByColumns, xlPrevious).Column
    Else
        Result = Sheet.Cells(RowNumber, Sheet.Columns.Count).End(xlToLeft).Column
    End If
    Sheet_DataLastColumn = Result
End Function

Public Function Sheet_DataLastCellRange(ByVal Sheet As Worksheet) As Range
    Set Sheet_DataLastCellRange = Sheet.Cells( _
        Sheet_DataLastRow(Sheet), Sheet_DataLastColumn(Sheet))
End Function

'----------------------------------------
'◇最終行/列削除
'----------------------------------------
'   ・  RangeClearTypeは
'       Clear/ClearContents/ClearFormats
'----------------------------------------
Public Sub Range_Clear(ByRef Range As Range, _
ByVal RangeClearType As RangeClearType, _
Optional ByVal MergeCellOption As Boolean = False)
    Call Assert(OrValue(RangeClearType, _
        rcClear, rcClearContents, rcClearFormats), _
        "Error:RangeClear:Args RangeClear")

    If MergeCellOption Then
        Dim Cell As Range
        Select Case RangeClearType
        Case rcClear
            For Each Cell In Range
                If Cell.MergeCells Then
                    Cell.MergeArea.Clear
                Else
                    Cell.Clear
                End If
            Next
        Case rcClearContents
            For Each Cell In Range
                If Cell.MergeCells Then
                    Cell.MergeArea.ClearContents
                Else
                    Cell.ClearContents
                End If
            Next
        Case rcClearFormats
            For Each Cell In Range
                If Cell.MergeCells Then
                    Cell.MergeArea.ClearFormats
                Else
                    Cell.ClearFormats
                End If
            Next
        End Select
    Else
        Select Case RangeClearType
        Case rcClear
            Range.Clear
        Case rcClearContents
            Range.ClearContents
        Case rcClearFormats
            Range.ClearFormats
        End Select
    End If
End Sub

Public Sub Sheet_ClearRangeLastData( _
ByVal Sheet As Worksheet, _
ByVal RowIndex As Long, ByVal ColumnIndex As Long, _
Optional ByVal RangeClearType As RangeClearType = rcClear, _
Optional ByVal MergeCellOption As Boolean = False)
    If (RowIndex <= Sheet_DataLastRow(Sheet)) _
    And (ColumnIndex <= Sheet_DataLastColumn(Sheet)) Then
        Call Range_Clear( _
            Sheet.Range( _
                Sheet.Cells(RowIndex, ColumnIndex), _
                Sheet.Cells(Sheet_DataLastRow(Sheet), Sheet_DataLastColumn(Sheet))), _
            RangeClearType, MergeCellOption)
    End If
End Sub

'・列のクリア、最終行まで
Public Sub Sheet_ClearColumnLastRow( _
ByVal Sheet As Worksheet, _
ByVal RowIndex As Long, ByVal ColumnIndex As Long, _
Optional ByVal RangeClearType As RangeClearType = rcClear, _
Optional ByVal MergeCellOption As Boolean = False)
    Dim LastRow As Long: LastRow = Sheet_DataLastRow(Sheet, ColumnIndex)
    If (RowIndex <= LastRow) Then
        Call Range_Clear( _
            Sheet.Range( _
                Sheet.Cells(RowIndex, ColumnIndex), _
                Sheet.Cells(Sheet_DataLastRow(Sheet, ColumnIndex), ColumnIndex)), _
            RangeClearType, MergeCellOption)
    End If
End Sub

'・行のクリア、最終列まで
Public Sub Sheet_ClearRowLastColumn( _
ByVal Sheet As Worksheet, _
ByVal RowIndex As Long, ByVal ColumnIndex As Long, _
Optional ByVal RangeClearType As RangeClearType = rcClear, _
Optional ByVal MergeCellOption As Boolean = False)
    Dim LastCol As Long: LastCol = Sheet_DataLastColumn(Sheet, RowIndex)
    If (ColumnIndex <= LastCol) Then
        Call Range_Clear( _
            Sheet.Range( _
                Sheet.Cells(RowIndex, ColumnIndex), _
                Sheet.Cells(RowIndex, Sheet_DataLastColumn(Sheet, RowIndex))), _
            RangeClearType, MergeCellOption)
    End If
End Sub

'----------------------------------------
'◇数式
'----------------------------------------

'----------------------------------------
'・数式を削除する関数
'----------------------------------------
Public Sub Sheet_RangeDeleteFormula( _
ByRef Sheet As Worksheet, ByRef Range As Range)

    '数式に影響が出ないように指定範囲の後方から値を指定している
    '=SUBTOTAL(9, …
    'とかの数式は、数式を無視して値に対して合算するというものなので
    '上部の数式が数値になった場合に値が変化してしまう
    Dim RowIndex As Long
    Dim ColIndex As Long
    For RowIndex = Range.Row + Range.Rows.Count To Range.Row Step -1
        For ColIndex = Range.Column + Range.Columns.Count To Range.Column Step -1
            If Sheet.Cells(RowIndex, ColIndex).HasFormula Then
                Sheet.Cells(RowIndex, ColIndex).Value = _
                    Sheet.Cells(RowIndex, ColIndex).Value
            End If
        Next
    Next
End Sub


'----------------------------------------
'◇Sheet.Rangeのコピー処理
'----------------------------------------

'----------------------------------------
'・数値書式のコピー
'----------------------------------------
'   ・  Excelの書式のコピーがバグっているので修正のために作成
'   ・  Excelのコピーでは
'       【#,##0_);[赤](#,##0)】が【#,##0_);[赤]-#,##0】に
'       なってしまう場合がある。
'       ファイルが破損しているのかもしれないが解消できなかったので
'       この関数を作成
'----------------------------------------
Public Sub Range_CopyNumberFormat( _
ByRef RangeSource As Range, _
ByRef RangeDest As Range)
    Dim FormatText As String
    Dim CellRangeSource As Range
    For Each CellRangeSource In RangeSource
        FormatText = CellRangeSource.NumberFormatLocal

        RangeDest.Parent.Cells( _
            RangeDest.Row + (CellRangeSource.Row - RangeSource.Row), _
            RangeDest.Column + (CellRangeSource.Column - RangeSource.Column) _
        ).NumberFormatLocal = FormatText

    Next
End Sub

'----------------------------------------
'・値など全てのコピー
'----------------------------------------
Public Sub Range_CopyAll( _
ByRef RangeSource As Range, _
ByRef RangeDest As Range)
    RangeSource.Copy
    Call RangeDest.PasteSpecial(Paste:=xlPasteAll)
    Call Range_CopyNumberFormat(RangeSource, RangeDest)
End Sub

'----------------------------------------
'・書式のコピー
'----------------------------------------
Public Sub Range_CopyFormat( _
ByRef RangeSource As Range, _
ByRef RangeDest As Range)
    RangeSource.Copy
    Call RangeDest.PasteSpecial(Paste:=xlPasteFormats)
    Call Range_CopyNumberFormat(RangeSource, RangeDest)
End Sub

'----------------------------------------
'・値のコピー
'----------------------------------------
Public Sub Range_CopyValue( _
ByRef RangeSource As Range, _
ByRef RangeDest As Range)
    RangeSource.Copy
    Call RangeDest.PasteSpecial(Paste:=xlPasteAllUsingSourceTheme)
    Call Range_CopyNumberFormat(RangeSource, RangeDest)
End Sub


'----------------------------------------
'◇セル
'----------------------------------------




'----------------------------------------
'◇範囲
'----------------------------------------

'----------------------------------------
'・範囲の上の1行
'----------------------------------------
Public Function Range_UpRow(ByRef SourceRange As Range) As Range
    Set Range_UpRow = _
        SourceRange.Resize(1, SourceRange.Columns.Count).Offset(-1, 0)
End Function

'----------------------------------------
'・範囲の下の1行
'----------------------------------------
Public Function Range_DownRow(ByRef SourceRange As Range) As Range
    Set Range_DownRow = _
        SourceRange.Resize(1, SourceRange.Columns.Count).Offset( _
            SourceRange.Rows.Count, 0)
End Function

'----------------------------------------
'◇範囲移動
'----------------------------------------

'----------------------------------------
'・範囲を上に1、移動する
'----------------------------------------
Public Sub Range_MoveUpRowOne(ByRef SourceRange As Range)
    '複数の選択範囲には非対応
    Call Assert(SourceRange.Areas.Count = 1, _
        "Error:RangeMoveUpRowOne:Areas.Count != 1")

    Dim EnableEventsBuffer As Boolean
    EnableEventsBuffer = _
        Application.EnableEvents
    Application.EnableEvents = False

    Dim SelectionFlag As Boolean
    If Selection.Address = SourceRange.Address Then
        SelectionFlag = True
    Else
        SelectionFlag = False
    End If

    '選択範囲の下1セルをあける
    Call Range_DownRow(SourceRange).Insert(xlDown)

    '上のセルを下のセルにコピーする
    Call Range_UpRow(SourceRange).Copy( _
        Destination:=Range_DownRow(SourceRange))

    '上のセルを1つ削除
    Call Range_UpRow(SourceRange).Delete(xlUp)

    '選択位置を1つ上にする
    If SelectionFlag Then
        SourceRange.Select
    End If

    Application.EnableEvents = EnableEventsBuffer

End Sub

Public Sub Range_MoveDownRowOne(ByRef SourceRange As Range)
    '複数の選択範囲には非対応
    Call Assert(SourceRange.Areas.Count = 1, _
        "Error:RangeMoveUpRowOne:Areas.Count != 1")

    Dim EnableEventsBuffer As Boolean
    EnableEventsBuffer = _
        Application.EnableEvents
    Application.EnableEvents = False

    Dim SelectionFlag As Boolean
    If Selection.Address = SourceRange.Address Then
        SelectionFlag = True
    Else
        SelectionFlag = False
    End If

    '選択範囲の上1セルをあける
    Call SourceRange.Resize(1, Selection.Columns.Count).Insert(xlDown)

    '下のセルを上のセルにコピーする
    Call Range_DownRow(SourceRange).Copy( _
        Destination:=Range_UpRow(SourceRange))

    '下のセルを1つ削除
    Call Range_DownRow(SourceRange).Delete(xlUp)

    '選択位置を1つ上にする
    If SelectionFlag Then
        SourceRange.Select
    End If

    Application.EnableEvents = EnableEventsBuffer

End Sub


'----------------------------------------
'◆Excel オブジェクト
'----------------------------------------


'----------------------------------------
'◇ Applicationによるワークブック操作
'----------------------------------------


'----------------------------------------
'・ワークブックの存在確認
'----------------------------------------
Public Function App_GetOpenedBook( _
ByVal App As Application, _
ByVal BookNameWildCard As String, _
Optional ByVal BookFolderPath As String = "") As Workbook

    Dim Result As Workbook: Set Result = Nothing
    Dim Book As Workbook
    If BookFolderPath = "" Then
        For Each Book In App.Workbooks
            If Book.Name Like BookNameWildCard Then
                Set Result = Book
                Exit For
            End If
        Next
    Else
        For Each Book In App.Workbooks
            If (Book.Name Like BookNameWildCard) _
            And (Book.Path = BookFolderPath) Then
                Set Result = Book
                Exit For
            End If
        Next
    End If
    Set App_GetOpenedBook = Result
End Function

Public Function App_OpenedBookExists( _
ByVal App As Application, _
ByVal BookNameWildCard As String, _
Optional ByVal BookFolderPath As String = "") As Boolean

    Dim Result As Boolean: Result = False
    If IsNotNothing( _
        App_GetOpenedBook(App, BookNameWildCard, BookFolderPath)) Then
        Result = True
    End If

    App_OpenedBookExists = Result
End Function

Public Sub testBook_Exists()
    Call Check(True, App_OpenedBookExists(Application, "st_vba.xlsm"))
    Call Check(True, App_OpenedBookExists(Application, "st_vba*"))
    Call Check(False, App_OpenedBookExists(Application, "st_vba.xls"))
End Sub

'----------------------------------------
'・ワークブックが開いていれば取得
'  開いていなければ開く
'----------------------------------------

Public Function App_GetOpenedBookOrOpenBook( _
ByVal App As Application, _
ByVal FilePath As String, _
Optional ByVal CheckFullPath As Boolean = False, _
Optional ByVal OpenReadOnlyFlag As Boolean = False, _
Optional ByRef ResultOpen As Boolean) As Workbook
    Dim Result As Workbook
    
    If CheckFullPath Then
        Set Result = App_GetOpenedBook(App, _
            fso.GetFileName(FilePath), _
            fso.GetParentFolderName(FilePath))
    Else
        Set Result = App_GetOpenedBook(App, _
            fso.GetFileName(FilePath))
    End If
    
    If (IsNothing(Result)) Then
        Set Result = App.Workbooks.Open(FilePath, , OpenReadOnlyFlag)
        ResultOpen = True
    Else
        ResultOpen = False
    End If
    
    Set App_GetOpenedBookOrOpenBook = Result
End Function


'----------------------------------------
'◇ ワークブック
'----------------------------------------

'----------------------------------------
'・ワークブックのフルパス取得
'----------------------------------------
Public Function Book_FullPath( _
ByVal Book As Workbook) As String
'    Book_FullPath = _
'        PathCombine( _
'            Book.Path, _
'            Book.Name)
    Book_FullPath = Book.FullName
End Function

Public Sub test_Book_FullPath()
    Call Check(ThisWorkbook.FullName, Book_FullPath(ThisWorkbook))
End Sub

'----------------------------------------
'・ワークブックを確認ダイアログなど無しで保存する
'----------------------------------------
'   ・  旧形式(XLS)の拡張子に自動対応
'----------------------------------------
Public Sub Book_SaveAs( _
ByVal Book As Workbook, _
ByVal FilePath As String)

    If Val(Application.Version) < 12 Then
        '以前のバージョンならそのまま保存
        Call Book.SaveAs(FilePath)
    Else
        '拡張子XLS なら古い形式で保存
        If LCase(GetExtensionIncludePeriod(FilePath)) = ".xls" Then
        
            '旧バージョンの確認ダイアログのようなものを出させない
            
            Dim Application_DisplayAlerts_Flag As Boolean
            Application_DisplayAlerts_Flag = Application.DisplayAlerts
            Application.DisplayAlerts = False
            Call Book.SaveAs(FilePath, XlFileFormat.xlExcel8)
            Application.DisplayAlerts = Application_DisplayAlerts_Flag
        Else
            Call Book.SaveAs(FilePath)
        End If
    End If

End Sub

'----------------------------------------
'・ワークブックを確認ダイアログなど無しで閉じる
'----------------------------------------
Public Sub Book_CloseSilence(ByVal Book As Workbook)
    Call Book.Close(SaveChanges:=False)
End Sub

'----------------------------------------
'・ワークブックBeforCloseイベント時に保存するかしないかを記述する
'----------------------------------------
Public Sub Book_BeforeClose_NoSave( _
ByVal Book As Workbook)
    '終了時に保存しないためにSavedフラグをTrueにする
    Book.Saved = True
End Sub

Public Sub Book_BeforeClose_Save( _
ByVal Book As Workbook)
    Call Book.Save
End Sub

'----------------------------------------
'・ワークシートの存在確認
'----------------------------------------

Public Function Book_GetSheet( _
ByVal Book As Workbook, _
ByVal SheetNameWildCard As String) As Worksheet

    Dim Result As Worksheet: Set Result = Nothing
    Dim I As Long
    For I = 1 To Book.Sheets.Count
        If Book.Sheets(I).Name Like SheetNameWildCard Then
            Set Result = Book.Sheets(I)
            Exit For
        End If
    Next

    Set Book_GetSheet = Result
End Function

Public Function Book_SheetExists( _
ByVal Book As Workbook, _
ByVal SheetNameWildCard As String) As Boolean

    Dim Result As Boolean: Result = False
    If (Book_GetSheet(Book, SheetNameWildCard) Is Nothing) = False Then
        Result = True
    End If

    Book_SheetExists = Result
End Function

Public Sub testBook_SheetExists()
    Call Check(True, Book_SheetExists(ThisWorkbook, "Sheet1"))
    Call Check(True, Book_SheetExists(ThisWorkbook, "Sheet*"))
    Call Check(False, Book_SheetExists(ThisWorkbook, "Sheet"))
End Sub

'----------------------------------------
'・ワークシートの削除
'----------------------------------------
Public Sub Book_DefaultSheetsDelete( _
ByVal Book As Workbook)
    Call Book_SheetsDelete(Book, "Sheet*", , False)
End Sub

'----------------------------------------
'・ワークシート(複数)の削除
'----------------------------------------
'   ・  次のようにして使うと便利
'           Call DeleteSheets("*(?)")
'       Sheet1(2) という名前のシートだけ削除される
'----------------------------------------
Public Sub Book_SheetsDelete( _
ByVal Book As Workbook, _
ByVal SheetNameWildCard As String, _
Optional MatchUnDelete As Boolean = False, _
Optional MsgBoxFlag As Boolean = True)

    Dim MessageText As String
    MessageText = ""

    Dim Sheet As Worksheet
    Dim I As Long
    For I = Book.Sheets.Count To 1 Step -1
        If MatchUnDelete = False Then
            If (Book.Sheets(I).Name Like SheetNameWildCard) Then
                MessageText = MessageText + _
                    Book.Sheets(I).Name + vbCrLf
            End If
        Else
            If Not (Book.Sheets(I).Name Like SheetNameWildCard) Then
                MessageText = MessageText + _
                    Book.Sheets(I).Name + vbCrLf
            End If
        End If
    Next
    
    If MessageText = "" Then
        If MsgBoxFlag Then
            Call MsgBox("削除対象シートはありません。")
        End If
        Exit Sub
    End If
    
    If MsgBoxFlag Then
        If MsgBox(StringCombine(vbCrLf, _
            "次のシートを削除しますか？", MessageText), _
            vbYesNo, "シート削除") _
            <> VbMsgBoxResult.vbYes Then
            Exit Sub
        End If
    End If
    
    Dim Application_DisplayAlerts_Flag As Boolean
    Application_DisplayAlerts_Flag = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    For I = Book.Sheets.Count To 1 Step -1
        If MatchUnDelete = False Then
            If (Book.Sheets(I).Name Like SheetNameWildCard) Then
                Book.Sheets(I).Delete
            End If
        Else
            If Not (Book.Sheets(I).Name Like SheetNameWildCard) Then
                Book.Sheets(I).Delete
            End If
        End If
    Next
        
    Application.DisplayAlerts = Application_DisplayAlerts_Flag
    
End Sub

'----------------------------------------
'◇ワークシート
'----------------------------------------

'----------------------------------------
'・ワークシートを確認ダイアログ無しで削除する
'----------------------------------------
Public Sub Sheet_DeleteSilence(ByVal Sheet As Worksheet)
    Dim Application_DisplayAlerts_Flag As Boolean
    Application_DisplayAlerts_Flag = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Call Sheet.Delete
    Application.DisplayAlerts = Application_DisplayAlerts_Flag
End Sub


'----------------------------------------
'・ワークシートへのテキスト配置
'----------------------------------------

Public Sub Sheet_SetText(ByVal Sheet As Worksheet, _
ByVal RowIndex As Long, ByVal ColumnIndex As Long, _
ByVal DocumentText As String)

    DocumentText = Replace(DocumentText, vbCrLf, vbCr)
    DocumentText = Replace(DocumentText, vbLf, vbCr)

    Dim Lines() As String
    Lines = Split(DocumentText, vbCr)
    Dim LineIndex As Long: LineIndex = RowIndex
    Dim I As Long
    For I = 0 To ArrayCount(Lines) - 1
        If IsIncludeStr(Lines(I), vbTab) Then
            Dim Columns() As String
            Columns = Split(Lines(I), vbTab)
            Dim J As Long
            For J = 0 To ArrayCount(Columns) - 1
                Sheet.Cells(LineIndex, ColumnIndex + J).Value = Columns(J)
            Next
        Else
        Sheet.Cells(LineIndex, ColumnIndex).Value = Lines(I)
        End If
        LineIndex = LineIndex + 1
    Next
End Sub

'----------------------------------------
'・値の増加
'----------------------------------------
Public Sub Sheet_CellValueIncrement( _
ByVal Sheet As Worksheet, _
ByVal Row As Long, ByVal Col As Long, _
ByVal Increment As Long)
    Sheet.Cells(Row, Col).Value = Sheet.Cells(Row, Col).Value + Increment
End Sub

'----------------------------------------
'・ シートの特定の列をチェックボックスのUIにする
'----------------------------------------
'   ・  ダブルクリックするとON/OFFが切り替えられる列になる
'   ・  タイトル列をダブルクリックすると全行がON/OFF切り替わる
'   ・  使い方は次の通り
'           Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'               Cancel = Sheet_CheckBoxColumn(Me, Target, Col_A, Col_D, 2)
'           End Sub
'       Col_Aは対象列
'       Col_Dはデータの有無をチェックする列
'       2は、タイトル行、2+1以上がデータ列になる
'----------------------------------------
Public Function Sheet_CheckBoxColumn(ByVal Sheet As Worksheet, _
ByVal Target As Range, _
ByVal Col_CheckBox As Long, _
ByVal Col_Data As Long, _
ByVal Row_Title As Long) As Boolean


    Dim Result As Boolean: Result = False
    
    If Target.Column <> Col_CheckBox Then Exit Function
    If Target.Columns.Count <> 1 Then Exit Function
    If Target.Rows.Count <> 1 Then Exit Function
    
    If Target.Row = Row_Title Then

        '全てONされているかどうかを調べて
        '全てONにしたり、OFFにしたりする制御
        Dim AllOnFlag As Boolean
        AllOnFlag = True
        Dim I As Long
        For I = Row_Title + 1 To Sheet_DataLastRow(Sheet)
            If Sheet.Cells(I, Col_Data).Value <> "" Then
                If Sheet.Cells(I, Col_CheckBox).Value <> "ON" Then
                    AllOnFlag = False
                End If
            End If
        Next
        For I = Row_Title + 1 To Sheet_DataLastRow(Sheet)
            If Sheet.Cells(I, Col_Data).Value = "" Then
                Sheet.Cells(I, Col_CheckBox).Value = ""
            Else
                Sheet.Cells(I, Col_CheckBox).Value = IIf(AllOnFlag, "OFF", "ON")
            End If
        Next
        
        Result = True
    ElseIf Row_Title + 1 <= Target.Row Then
    
        If Target.Value = "" Then
            Target.Value = "ON"
        ElseIf Target.Value = "ON" Then
            Target.Value = "OFF"
        ElseIf Target.Value = "OFF" Then
            Target.Value = ""
        End If
        Result = True
    End If
    
    Sheet_CheckBoxColumn = Result
End Function

'----------------------------------------
'◇チェックボックス
'----------------------------------------
'   ・  フォントWindingsのチェックボックス表示の文字列を返す
'   ・  ChrWの反対はAscW
'----------------------------------------
Function Wingdings_Checkbox_Checked() As String
    Wingdings_Checkbox_Checked = _
        ChrW(254)
End Function

Function Wingdings_Checkbox_UnChecked() As String
    Wingdings_Checkbox_UnChecked = _
        ChrW(168)
End Function

'----------------------------------------
'◇オブジェクト
'----------------------------------------

'----------------------------------------
'・ChartObjectの存在確認
'----------------------------------------
Public Function ChartObjectExists( _
ByVal ChartObjectName As String, _
Optional ByVal Sheet As Worksheet = Nothing) As Boolean

    If Sheet Is Nothing Then Set Sheet = ActiveSheet

    Dim Result As Boolean: Result = False
    Dim ChartObject As ChartObject
    For Each ChartObject In Sheet.ChartObjects
        If ChartObject.Name = ChartObjectName Then
            Result = True
        End If
    Next
    ChartObjectExists = Result
End Function

Private Sub testChartObjectExists()
    Call Check(True, ChartObjectExists("Graph01"))
    Call Check(False, ChartObjectExists("Graph02"))
End Sub

'----------------------------------------
'・OLEObjectの存在確認
'----------------------------------------
Public Function OLEObjectExists( _
ByVal OLEObjectName As String, _
Optional ByVal Sheet As Worksheet = Nothing) As Boolean

    If Sheet Is Nothing Then Set Sheet = ActiveSheet

    Dim Result As Boolean: Result = False
    Dim OLEObject As OLEObject
    For Each OLEObject In Sheet.OLEObjects
        If OLEObject.Name = OLEObjectName Then
            Result = True
        End If
    Next
    OLEObjectExists = Result
End Function

'----------------------------------------
'◇Shape
'----------------------------------------

'----------------------------------------
'・Shapesの存在確認
'----------------------------------------
Public Function ShapeExists( _
ByVal ShapeName As String, _
Optional ByVal Sheet As Worksheet = Nothing) As Boolean

    If Sheet Is Nothing Then Set Sheet = ActiveSheet

    Dim Result As Boolean: Result = False
    Dim Shape As Shape
    For Each Shape In Sheet.Shapes
        If Shape.Name = ShapeName Then
            Result = True
        End If
    Next
    ShapeExists = Result
End Function

Private Sub testShapeExists()
    Call Check(True, ShapeExists("Graph01"))
    Call Check(False, ShapeExists("Graph02"))
End Sub


'----------------------------------------
'・セル範囲に当てはまるように画像ファイルを貼り付ける処理
'----------------------------------------
Public Function GetShapeFromImageFile(ByVal Sheet As Worksheet, _
    ByVal ImageFilePath As String, _
    ByVal SheetRange As Range, _
    Optional ByVal Margin As Long = 1, _
    Optional HorizontalAlign As AlineHorizontal = AlineHorizontal.alCenter, _
    Optional VerticalAlign As AlineVertical = AlineVertical.alCenter) _
    As Shape

    If fso.FileExists(ImageFilePath) = False Then
        Set GetShapeFromImageFile = Nothing
        Exit Function
    End If

    'マージンをとるために値を設定
    Dim Rect As Rect
    Rect.Left = SheetRange.Left + Margin
    Rect.Top = SheetRange.Top + Margin
    Call SetRectWidth(Rect, SheetRange.Width - (Margin * 2))
    Call SetRectHeight(Rect, SheetRange.Height - (Margin * 2))

    Dim Shape As Shape
    Set Shape = Sheet.Shapes.AddPicture( _
        FileName:=ImageFilePath, LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=Rect.Left, _
        Top:=Rect.Top, _
        Width:=0, _
        Height:=0)

    '元画像サイズに戻す
    Call Shape.ScaleHeight(1#, True)
    Call Shape.ScaleWidth(1#, True)

    '縦横比を保持したまま、高さを調整する
    Shape.LockAspectRatio = True
    Shape.Height = GetRectHeight(Rect)

    '画像横サイズが範囲内に収まっているかどうか確認
    If Shape.Width > GetRectWidth(Rect) Then
        '横サイズがはみ出ているなら横を合わせる
        Shape.Width = GetRectWidth(Rect)

        '左右位置はぴったりなので上下位置調整をする
        Select Case VerticalAlign
        Case AlineVertical.alCenter
            Shape.Top = Shape.Top + (GetRectHeight(Rect) - Shape.Height) / 2
        Case AlineVertical.alBottom
            Shape.Top = Shape.Top + (GetRectHeight(Rect) - Shape.Height)
        End Select
    Else
        '上下位置はぴったりなので左右位置調整をする
        Select Case HorizontalAlign
        Case AlineHorizontal.alCenter
            Shape.Left = Shape.Left + (GetRectWidth(Rect) - Shape.Width) / 2
        Case AlineHorizontal.alRight
            Shape.Left = Shape.Left + (GetRectWidth(Rect) - Shape.Width)
        End Select
    End If

    Set GetShapeFromImageFile = Shape
End Function

'----------------------------------------
'・Shape画像を圧縮する
'----------------------------------------
'   ・  クリップボードを経由する方法しか無いらしい
'----------------------------------------
Public Sub ShapeCompressUseClipboard(ByVal Sheet As Worksheet, ByVal Shape As Shape)
On Error Resume Next
    Dim Point As Point
    Dim RectSize As RectSize
    Point.X = Shape.Left
    Point.Y = Shape.Top
    RectSize.Width = Shape.Width
    RectSize.Height = Shape.Height

    Shape.Cut
    If Err.Number <> 0 Then
        Err.Clear
        Shape.Cut
        If Err.Number <> 0 Then
            Err.Clear
            Shape.Cut
        End If
    End If

    Sheet.Select
    Sheet.Activate

'    Sheet.PasteSpecial Format:="図 (拡張メタファイル)", Link:=False, DisplayAsIcon:=False

    Sheet.PasteSpecial Format:="図 (JPEG)", Link:=False, DisplayAsIcon:=False
    If Err.Number <> 0 Then
        Err.Clear
        Sheet.PasteSpecial Format:="図 (JPEG)", Link:=False, DisplayAsIcon:=False
        If Err.Number <> 0 Then
            Err.Clear
            Sheet.PasteSpecial Format:="図 (JPEG)", Link:=False, DisplayAsIcon:=False
        End If
    End If

    Selection.ShapeRange.Width = RectSize.Width
    Selection.ShapeRange.Height = RectSize.Height
    Selection.Left = Point.X
    Selection.Top = Point.Y
End Sub

'----------------------------------------
'・座標位置に対するセル位置を返す関数
'----------------------------------------
'   ・  Shape.TopLeftCell/.BottomRightCell はあるが
'       中心位置のセルを求める方法はなかったので作成。
'       速度は速くない。
'----------------------------------------
Public Function TopLeftCell(ByRef Sheet As Worksheet, _
ByVal Top As Long, ByVal Left As Long) As Range
    Call Assert(0 <= Top, "Error:TopLeftCell:Top < 0")
    Call Assert(0 <= Left, "Error:TopLeftCell:Left < 0")

    Dim Row As Long
    Row = 0
    Do
        If Top < Sheet.Rows(Row + 1).Top Then Exit Do
        Row = Row + 1
    Loop While True

    Dim Col As Long
    Col = 0
    Do
        If Left < Sheet.Columns(Col + 1).Left Then Exit Do
        Col = Col + 1
    Loop While True

    Set TopLeftCell = Sheet.Cells(Row, Col)
End Function

Public Sub testTopLeftCell()
    Call Check(Sheets(1).Cells(1, 1), TopLeftCell(Sheets(1), 0, 0))
End Sub

'----------------------------------------
'・ 指定範囲(Range)にあるShapeを削除する
'----------------------------------------
'   ・  Shapeは中心がRangeに含まれている事で判断する
'----------------------------------------
Public Sub Range_DeleteShape( _
ByVal Sheet As Worksheet, _
ByVal Range As Range)

    Dim ShapeRange As Range
    Dim Shape As Shape
    For Each Shape In Sheet.Shapes
    Do
        Shape.Placement = xlMove
        Dim Point As Point
        Point.X = Shape.Left + (Shape.Width / 2)
        Point.Y = Shape.Top + (Shape.Height / 2)
        Set ShapeRange = TopLeftCell(Sheet, Point.Y, Point.X)
        'Shapeの中心が位置するセルを求める

        If InRange(Range.Top, ShapeRange.Top, Range.Top + Range.Rows.Count - 1) Then
            If InRange(Range.Left, ShapeRange.Left, Range.Left + Range.Columns.Count - 1) Then
                Shape.Delete
            End If
        End If
        
    Loop While False
    Next
    
End Sub

'----------------------------------------
'◇ ファイル選択ダイアログ
'----------------------------------------

'----------------------------------------
'・ ファイル選択ダイアログ
'----------------------------------------
'   ・  ダイアログのタイトルには「参照」と表示される
'   ・  Application.FileDialog(msoFileDialogFilePicker)
'       をラッピング
'   ・  Filtersは
'           "Excelブック|*.xls; *.xlsx; *.xlsm"
'           "Textファイル|*.txt"
'       という形式にする。
'       [*.txt;]という形式にはしないこと。エラーになる。
'   ・  戻り値は改行区切りのフルパス
'       キャンセルが押された場合には空文字が返る
'----------------------------------------
Public Function FileDialog_FilePicker(ByVal FilePath As String, _
ByVal OptionInitialView As MsoFileDialogView, _
ByVal OptionAllowMultiSelect As Boolean, _
ParamArray Filters())

    Dim Result As String: Result = ""

    Dim Dialog As FileDialog
    Set Dialog = Application.FileDialog(msoFileDialogFilePicker)
    
    Call Dialog.Filters.Clear
    Dim I As Long
    For I = LBound(Filters) To UBound(Filters)
        Call Dialog.Filters.Add( _
            FirstStrFirstDelim(Filters(I), "|"), _
            LastStrFirstDelim(Filters(I), "|"), I + 1)
            
'        Call Dialog.Filters.Add("Excelブック", "*.xls; *.xlsx; *.xlsm", 1)
'        Call Dialog.Filters.Add("Textファイル", "*.txt", 1)
    Next
        
    Dialog.InitialFileName = FilePath
    Dialog.InitialView = OptionInitialView
    Dialog.AllowMultiSelect = OptionAllowMultiSelect
    If Dialog.Show = True Then

        For I = 1 To Dialog.SelectedItems.Count
            Result = StringCombine(vbCrLf, Result, Dialog.SelectedItems(I))
        Next

    End If
    
    FileDialog_FilePicker = Result
End Function

Public Sub testFileDialog_FilePicker()
    Dim Result As String
    Result = FileDialog_FilePicker(Book_FullPath(ThisWorkbook), _
        msoFileDialogViewDetails, True)
    Call MsgBox(Result)
    
    Result = FileDialog_FilePicker(Book_FullPath(ThisWorkbook), _
        msoFileDialogViewDetails, False, _
        "Excelブック|*.xls; *.xlsx; *.xlsm", _
        "Textファイル|*.txt")
    Call MsgBox(Result)
    
    Result = FileDialog_FilePicker("", _
        msoFileDialogViewDetails, True)
    Call MsgBox(Result)
End Sub


'----------------------------------------
'・ ファイルオープンダイアログ
'----------------------------------------
'   ・  ダイアログのタイトルには「ファイルを開く」と表示される
'   ・  Application.FileDialog(msoFileDialogOpen)
'       をラッピング
'   ・  Filtersは
'           "Excelブック|*.xls; *.xlsx; *.xlsm"
'           "Textファイル|*.txt"
'       という形式にする。
'       [*.txt;]という形式にはしないこと。エラーになる。
'   ・  戻り値は改行区切りのフルパス
'       キャンセルが押された場合には空文字が返る
'----------------------------------------
Public Function FileDialog_Open(ByVal FilePath As String, _
ByVal OptionInitialView As MsoFileDialogView, _
ByVal OptionAllowMultiSelect As Boolean, _
ParamArray Filters())

    Dim Result As String: Result = ""

    Dim Dialog As FileDialog
    Set Dialog = Application.FileDialog(msoFileDialogOpen)
    
    Call Dialog.Filters.Clear
    Dim I As Long
    For I = LBound(Filters) To UBound(Filters)
        Call Dialog.Filters.Add( _
            FirstStrFirstDelim(Filters(I), "|"), _
            LastStrFirstDelim(Filters(I), "|"), I + 1)
            
'        Call Dialog.Filters.Add("Excelブック", "*.xls; *.xlsx; *.xlsm", 1)
'        Call Dialog.Filters.Add("Textファイル", "*.txt", 1)
    Next
        
    Dialog.InitialFileName = FilePath
    Dialog.InitialView = OptionInitialView
    Dialog.AllowMultiSelect = OptionAllowMultiSelect
    If Dialog.Show = True Then

        For I = 1 To Dialog.SelectedItems.Count
            Result = StringCombine(vbCrLf, Result, Dialog.SelectedItems(I))
        Next

    End If
    
    FileDialog_Open = Result
End Function

Public Sub testFileDialog_Open()
    Dim Result As String
    Result = FileDialog_Open(Book_FullPath(ThisWorkbook), _
        msoFileDialogViewDetails, True)
    Call MsgBox(Result)
    
    Result = FileDialog_Open(Book_FullPath(ThisWorkbook), _
        msoFileDialogViewDetails, False, _
        "Excelブック|*.xls; *.xlsx; *.xlsm", _
        "Textファイル|*.txt")
    Call MsgBox(Result)
    
    Result = FileDialog_Open("", _
        msoFileDialogViewDetails, True)
    Call MsgBox(Result)
End Sub

'----------------------------------------
'・ 名前を付けて保存ダイアログ
'----------------------------------------
'   ・  ダイアログのタイトルには「名前を付けて保存」と表示される
'   ・  Application.FileDialog(msoFileDialogSaveAs)
'       をラッピング
'   ・  Filterは指定できない。Excel標準の保存形式で固定。
'   ・  複数選択もできない。(しても単独選択のみ)
'   ・  戻り値はフルパス
'       キャンセルが押された場合には空文字が返る
'----------------------------------------
Public Function FileDialog_SaveAs(ByVal FilePath As String, _
ByVal OptionInitialView As MsoFileDialogView)

    Dim Result As String: Result = ""

    Dim Dialog As FileDialog
    Set Dialog = Application.FileDialog(msoFileDialogSaveAs)
    
    Dialog.InitialFileName = FilePath
    Dialog.InitialView = OptionInitialView
    If Dialog.Show = True Then
        Dim I As Long
        For I = 1 To Dialog.SelectedItems.Count
            Result = StringCombine(vbCrLf, Result, Dialog.SelectedItems(I))
        Next

    End If
    
    FileDialog_SaveAs = Result
End Function

Public Sub testFileDialog_SaveAs()
    Dim Result As String
    Result = FileDialog_SaveAs(Book_FullPath(ThisWorkbook), _
        msoFileDialogViewDetails)
    Call MsgBox(Result)
    
    Result = FileDialog_SaveAs(Book_FullPath(ThisWorkbook), _
        msoFileDialogViewDetails)
    Call MsgBox(Result)
    
    Result = FileDialog_SaveAs("", _
        msoFileDialogViewDetails)
    Call MsgBox(Result)
End Sub


'----------------------------------------
'・ フォルダ選択ダイアログ
'----------------------------------------
'   ・  ダイアログのタイトルには「参照」と表示される
'   ・  Application.FileDialog(msoFileDialogFolderPicker)
'       をラッピング
'   ・  Filterは指定できない。
'   ・  複数選択もできない。(しても単独選択のみ)
'   ・  戻り値はフルパス
'       キャンセルが押された場合には空文字が返る
'   ・  InitialFileName はフォルダ選択としてはいまいちな動作。
'       指定したフォルダが開くが
'       そのままOKを押しても、
'       そのフォルダ内に同名フォルダが無いとなる
'       回避不可能な不具合と思える
'----------------------------------------
Public Function FileDialog_FolderPicker(ByVal FolderPath As String, _
ByVal OptionInitialView As MsoFileDialogView)

    Dim Result As String: Result = ""

    Dim Dialog As FileDialog
    Set Dialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    Dim I As Long
        
    Dialog.InitialFileName = FolderPath
    Dialog.InitialView = OptionInitialView
    If Dialog.Show = True Then

        For I = 1 To Dialog.SelectedItems.Count
            Result = StringCombine(vbCrLf, Result, Dialog.SelectedItems(I))
        Next

    End If
    
    FileDialog_FolderPicker = Result
End Function

Public Sub testFileDialog_FolderPicker()
    Dim Result As String
    Result = FileDialog_FolderPicker(ThisWorkbook.Path, _
        msoFileDialogViewDetails)
    Call MsgBox(Result)
    
    Result = FileDialog_FolderPicker(ThisWorkbook.Path, _
        msoFileDialogViewDetails)
    Call MsgBox(Result)
    
    Result = FileDialog_FolderPicker("", _
        msoFileDialogViewDetails)
    Call MsgBox(Result)
End Sub



'----------------------------------------
'◆Excel アプリケーション
'----------------------------------------

'----------------------------------------
'・Excel ウィンドウタイトルバー表示
'----------------------------------------
Public Sub SetExcelWindowTitle( _
ByVal AppTitle As String, _
Optional ByVal ActTitle As String = "")

    Application.Caption = AppTitle
    ActiveWindow.Caption = ActTitle
    'Application.Caption = "" の場合、Excel という文字が自動で入る
    'ActionWindow.Caption <> "" の場合
    '  ウィンドウタイトル - アプリケーションタイトル
    'というようにハイフンで接続される

    'なので単に文字列を入れたい場合は
    'Application.Captionに文字設定して
    'ActiveWindow.Caption = "" にするとよい
End Sub

Public Sub ApplicationModeOn()
    Call ApplicationMode(ThisWorkbook.ActiveSheet, True)
End Sub

Public Sub ApplicationModeOff()
    Call ApplicationMode(ThisWorkbook.ActiveSheet, False)
End Sub

Public Sub ApplicationMode(ByVal Sheet As Worksheet, ByVal Switch As Boolean)
    Dim ScreenUpdatingBuffer As Boolean
    ScreenUpdatingBuffer = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Application.DisplayStatusBar = Not Switch
    Application.DisplayFormulaBar = Not Switch

    ActiveWindow.DisplayGridlines = Not Switch
    ActiveWindow.DisplayHeadings = Not Switch
    ActiveWindow.DisplayHeadings = Not Switch
    ActiveWindow.DisplayWorkbookTabs = Not Switch
    ActiveWindow.DisplayHorizontalScrollBar = Not Switch
    ActiveWindow.DisplayVerticalScrollBar = Not Switch
'
    If Switch Then
        Call Application.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",False)")

        Sheet.Unprotect
        Call Sheet.Protect(userinterfaceonly:=True)
        Sheet.EnableSelection = xlUnlockedCells
        ActiveSheet.ScrollArea = "$A$1"
    Else
        Call Application.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",True)")

        Sheet.Unprotect
        Sheet.EnableSelection = xlNoRestrictions
        Sheet.ScrollArea = ""
    End If

    Application.ScreenUpdating = ScreenUpdatingBuffer
End Sub


'----------------------------------------
'◆メニュー処理
'----------------------------------------
Public Function GetCheckFaceId(ByVal Value As Boolean) As Long
    Dim Result As Long: Result = 0
    If Value Then
        Result = 990
    End If
    GetCheckFaceId = Result
End Function

Public Function PopupMenu_ActionText(ByVal ReturnValue As String) As String
    PopupMenu_ActionText = _
        "'PopupMenu_ActionReturn """ + ReturnValue + """'"
End Function

Public Function PopupMenu_PopupReturn( _
ByRef PopupMenu As CommandBar, _
ByVal X As Long, ByVal Y As Long) As String
    PopupMenu_Return = ""
    Call PopupMenu.ShowPopup(X, Y)
    PopupMenu_PopupReturn = PopupMenu_Return
End Function

Public Function PopupMenu_PopupReturn_NoPosition( _
ByRef PopupMenu As CommandBar) As String
    PopupMenu_Return = ""
    Call PopupMenu.ShowPopup
    PopupMenu_PopupReturn_NoPosition = PopupMenu_Return
End Function

Public Sub PopupMenu_ActionReturn(ByVal ReturnValue As String)
    PopupMenu_Return = ReturnValue
End Sub


'----------------------------------------
'◆グラフ処理
'----------------------------------------

'----------------------------------------
'・GraphFormulaDataを取得と設定
'----------------------------------------

'Chart.SeriesCollection.Item(I).Formulaメソッドで得られる文字列の例
'   =SERIES(,Sheet1!$A$2:$A$32,Sheet1!$B$2:$B$32,1)
'   =SERIES(系列名,X軸項目軸,データ,系列番号)

Public Function GetGraphFormulaData(Chart As Object, SeriesNumber As Long) As GraphFormulaData
    Dim Result As GraphFormulaData
    Dim FormulaStr As String
    Dim FormulaSeriesArgsStr() As String

    FormulaStr = Chart.SeriesCollection.Item(SeriesNumber).Formula
    FormulaStr = ExcludeLastStr(ExcludeFirstStr(FormulaStr, "=SERIES("), ")")
    FormulaSeriesArgsStr = Split(FormulaStr, ",")

    Result.SeriesName = FormulaSeriesArgsStr(0)
    Result.ItemXAxisRangeStr = FormulaSeriesArgsStr(1)
    Result.DataRangeStr = FormulaSeriesArgsStr(2)
    Result.SeriesNumber = CLng(FormulaSeriesArgsStr(3))

    GetGraphFormulaData = Result
End Function

Public Sub SetGraphFormulaData(Chart As Object, SeriesNumber As Long, Data As GraphFormulaData)
    Dim FormulaStr As String
    FormulaStr = _
        Data.SeriesName + "," + _
        Data.ItemXAxisRangeStr + "," + _
        Data.DataRangeStr + "," + _
        CStr(Data.SeriesNumber)

        Chart.SeriesCollection.Item(SeriesNumber).Formula = _
            "=SERIES(" + FormulaStr + ")"
End Sub

Public Function GetGraphDataRange(Chart As Object, SeriesNumber As Long) As Range
    Set GetGraphDataRange = _
        Application.Range(GetGraphFormulaData(Chart, SeriesNumber).DataRangeStr)
End Function

Public Function GetGraphXAxisRange(Chart As Object, SeriesNumber As Long) As Range
    Set GetGraphXAxisRange = _
        Application.Range(GetGraphFormulaData(Chart, SeriesNumber).ItemXAxisRangeStr)
End Function

'----------------------------------------
'◇グラフ単独データの範囲変更
'----------------------------------------

'----------------------------------------
'・GraphFormulaDataの終端操作
'   Value＝正、終端広がる
'   Value＝負、終端狭まる
'----------------------------------------
Public Sub GraphSeriesLastRangeUp(ByRef Data As GraphFormulaData, Value As Long)
    Dim R1 As Range
    Dim SheetName As String

    If Data.ItemXAxisRangeStr <> "" Then
        Set R1 = Application.Range(Data.ItemXAxisRangeStr)
'        SheetName = FirstStrFirstDelim(Data.ItemXAxisRangeStr, "!")
        SheetName = FirstStrFirstDelim(Data.ItemXAxisRangeStr, "!")
        Data.ItemXAxisRangeStr = _
            IncludeFirstStr( _
                R1.Resize(R1.Rows.Count + Value, R1.Columns.Count).Address, _
                SheetName + "!")
    End If

    Set R1 = Application.Range(Data.DataRangeStr)
    SheetName = FirstStrFirstDelim(Data.DataRangeStr, "!")
    Data.DataRangeStr = _
        IncludeFirstStr( _
            R1.Resize(R1.Rows.Count + Value, R1.Columns.Count).Address, _
            SheetName + "!")
End Sub

'----------------------------------------
'・GraphFormulaDataの先頭を操作する
'   Value＝正、先頭広がる
'   Value＝負、先頭狭まる
'----------------------------------------
Public Sub GraphSeriesFirstRangeUp(ByRef Data As GraphFormulaData, Value As Long)
    Call GraphSeriesMove(Data, -Value)
    Call GraphSeriesLastRangeUp(Data, Value)
End Sub

'----------------------------------------
'・GraphFormulaDataの範囲を移動する
'   Value＝正、後方移動
'   Value＝負、前方移動
'----------------------------------------
Public Sub GraphSeriesMove(ByRef Data As GraphFormulaData, Value As Long)
    Dim R1 As Range
    Dim SheetName As String

    If Data.ItemXAxisRangeStr <> "" Then
        Set R1 = Application.Range(Data.ItemXAxisRangeStr)
        SheetName = FirstStrFirstDelim(Data.ItemXAxisRangeStr, "!")
        Data.ItemXAxisRangeStr = _
            IncludeFirstStr( _
                R1.Offset(Value, 0).Address, _
                SheetName + "!")
    End If

    Set R1 = Application.Range(Data.DataRangeStr)
    SheetName = FirstStrFirstDelim(Data.DataRangeStr, "!")
    Data.DataRangeStr = _
        IncludeFirstStr( _
            R1.Offset(Value, 0).Address, _
            SheetName + "!")
End Sub

'----------------------------------------
'・GraphFormulaDataのサイズ操作
'   Value＝行数
'----------------------------------------
Public Sub GraphSeriesResize(ByRef Data As GraphFormulaData, Value As Long)
    Dim R1 As Range
    Dim SheetName As String

    If Data.ItemXAxisRangeStr <> "" Then
        Set R1 = Application.Range(Data.ItemXAxisRangeStr)
        SheetName = FirstStrFirstDelim(Data.ItemXAxisRangeStr, "!")
        Data.ItemXAxisRangeStr = _
            IncludeFirstStr( _
                R1.Resize(Value, R1.Columns.Count).Address, _
                SheetName + "!")
    End If

    Set R1 = Application.Range(Data.DataRangeStr)
    SheetName = FirstStrFirstDelim(Data.DataRangeStr, "!")
    Data.DataRangeStr = _
        IncludeFirstStr( _
            R1.Resize(Value, R1.Columns.Count).Address, _
            SheetName + "!")
End Sub

'----------------------------------------
'・グラフの範囲を取得する
'----------------------------------------
Public Function GetGraphRowCount(Chart As Chart) As Long
    Dim Result As Long
    Result = 0
    If 1 <= Chart.SeriesCollection.Count Then
        Dim Data As GraphFormulaData
        Data = GetGraphFormulaData(Chart, 1)

        Result = Application.Range(Data.ItemXAxisRangeStr).Rows.Count
    End If
    GetGraphRowCount = Result
End Function

'----------------------------------------
'・GraphFormulaDataの指定列の変更
'----------------------------------------
Public Sub SetGraphFormulaDataColumn(ByRef Data As GraphFormulaData, ColumnIndex As Long)
    Dim R1 As Range
    Dim SheetName As String
    Set R1 = Application.Range(Data.DataRangeStr)
    SheetName = FirstStrFirstDelim(Data.DataRangeStr, "!")

    Data.DataRangeStr = _
        IncludeFirstStr( _
            R1.Offset(0, ColumnIndex - R1.Column).Address, _
            SheetName + "!")
End Sub

'----------------------------------------
'・グラフのシリーズデータの指定列変更
'----------------------------------------
Public Sub SetChartSeriesColumn(Chart As Chart, ChartSeriesNumber As Long, ColumnIndex As Long)
    If ChartSeriesNumber <= Chart.SeriesCollection.Count Then
        Dim Data As GraphFormulaData
        Data = GetGraphFormulaData(Chart, ChartSeriesNumber)
        Call SetGraphFormulaDataColumn(Data, ColumnIndex)
        Call SetGraphFormulaData(Chart, ChartSeriesNumber, Data)
    End If
End Sub

'----------------------------------------
'◇グラフ全系列データの範囲変更
'----------------------------------------

'----------------------------------------
'・終端操作
'----------------------------------------
'   ・  Value＝正、終端広がる
'       Value＝負、終端狭まる
'----------------------------------------
Public Sub GraphAllSeriesLastRangeUp(Chart As Object, Value As Long)
On Error GoTo Err:
    Dim I As Long
    For I = 1 To Chart.SeriesCollection.Count
        Dim Data As GraphFormulaData
        Data = GetGraphFormulaData(Chart, I)
        Call GraphSeriesLastRangeUp(Data, Value)
        Call SetGraphFormulaData(Chart, I, Data)
    Next I
    Exit Sub
Err:
    Call MsgBox("範囲指定が正しくありません")
End Sub

'----------------------------------------
'・先頭操作
'----------------------------------------
'   ・  Value＝正、先頭広がる
'       Value＝負、先頭狭まる
'----------------------------------------
Public Sub GraphAllSeriesFirstRangeUp(Chart As Object, Value As Long)
On Error GoTo Err:
    Dim I As Long
    For I = 1 To Chart.SeriesCollection.Count
        Dim Data As GraphFormulaData
        Data = GetGraphFormulaData(Chart, I)
        Call GraphSeriesFirstRangeUp(Data, Value)
        Call SetGraphFormulaData(Chart, I, Data)
    Next I
    Exit Sub
Err:
    Call MsgBox("範囲指定が正しくありません")
End Sub

'----------------------------------------
'・範囲移動
'----------------------------------------
'   ・  Value＝正、後方移動
'       Value＝負、前方移動
'----------------------------------------
Public Sub GraphAllSeriesMove(Chart As Object, Value As Long)
On Error GoTo Err:
    Dim I As Long
    For I = 1 To Chart.SeriesCollection.Count
        Dim Data As GraphFormulaData
        Data = GetGraphFormulaData(Chart, I)
        Call GraphSeriesMove(Data, Value)
        Call SetGraphFormulaData(Chart, I, Data)
    Next I
    Exit Sub
Err:
    Call MsgBox("範囲指定が正しくありません")
End Sub

'----------------------------------------
'・サイズ操作
'----------------------------------------
'   ・  Value＝行数
'----------------------------------------
Public Sub GraphAllSeriesResize(Chart As Object, Value As Long)
On Error GoTo Err:
    Dim I As Long
    For I = 1 To Chart.SeriesCollection.Count
        Dim Data As GraphFormulaData
        Data = GetGraphFormulaData(Chart, I)
        Call GraphSeriesResize(Data, Value)
        Call SetGraphFormulaData(Chart, I, Data)
    Next I
    Exit Sub
Err:
    Call MsgBox("範囲指定が正しくありません")
End Sub

'----------------------------------------
'◆UserForm処理
'----------------------------------------

'----------------------------------------
'・WindowStyle
'----------------------------------------
Public Sub SetWindowStyle(hWnd As Long, _
ByVal TitleBar As Boolean, _
ByVal SystemMenu As Boolean, _
ByVal ResizeFrame As Boolean, _
ByVal MinimizeButton As Boolean, _
ByVal MaximizeButton As Boolean)
    Dim Style As Long
    Style = GetWindowLong(hWnd, GWL_STYLE)
    If TitleBar Then
        Style = Style Or WS_CAPTION
    Else
        Style = Style And (Not WS_CAPTION)
    End If
    If SystemMenu Then
        Style = Style Or WS_SYSMENU
    Else
        Style = Style And (Not WS_SYSMENU)
    End If
    If ResizeFrame Then
        Style = Style Or WS_THICKFRAME
    Else
        Style = Style And (Not WS_THICKFRAME)
    End If
    If MinimizeButton Then
        Style = Style Or WS_MINIMIZEBOX
    Else
        Style = Style And (Not WS_MINIMIZEBOX)
    End If
    If MaximizeButton Then
        Style = Style Or WS_MAXIMIZEBOX
    Else
        Style = Style And (Not WS_MAXIMIZEBOX)
    End If
    Call SetWindowLong(hWnd, GWL_STYLE, Style)
End Sub

Public Sub GetWindowStyle(hWnd As Long, _
ByRef TitleBar As Boolean, _
ByRef SystemMneu As Boolean, _
ByRef ResizeFrame As Boolean, _
ByRef MinimizeButton As Boolean, _
ByRef MaximizeButton As Boolean)
    Dim Style As Long
    Style = GetWindowLong(hWnd, GWL_STYLE)
    TitleBar = _
        Style = (Style Or WS_CAPTION)
    SystemMneu = _
        Style = (Style Or WS_SYSMENU)
    ResizeFrame = _
        Style = (Style Or WS_THICKFRAME)
    MinimizeButton = _
        Style = (Style Or WS_MINIMIZEBOX)
    MaximizeButton = _
        Style = (Style Or WS_MAXIMIZEBOX)
End Sub

'----------------------------------------
'・WindowExStyle
'----------------------------------------
Public Sub SetWindowExStyle(hWnd As Long, _
ByVal TaskBarButton As Boolean)
    Dim ExStyle As Long
    ExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    If TaskBarButton Then
        ExStyle = ExStyle Or WS_EX_APPWINDOW
    Else
        ExStyle = ExStyle And (Not WS_EX_APPWINDOW)
    End If
    Call SetWindowLong(hWnd, GWL_EXSTYLE, ExStyle)
End Sub

Public Sub GetWindowExStyle(hWnd As Long, _
ByRef TaskBarButton As Boolean)

    Dim ExStyle As Long
    ExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)

    TaskBarButton = _
        ExStyle = (ExStyle Or WS_EX_APPWINDOW)
End Sub

'----------------------------------------
'・CloseButton
'----------------------------------------
Public Sub SetWindowCloseButton(hWnd As Long, _
ByVal Enabled As Boolean)

    Dim hMenu As Long
    Dim rc As Long

    If Enabled Then
        'メニューをリセット
        hMenu = GetSystemMenu(hWnd, True)
    Else
        hMenu = GetSystemMenu(hWnd, False)
        rc = DeleteMenu(hMenu, 5, MF_BYPOSITION)
        rc = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    End If
    rc = DrawMenuBar(hWnd)

    'EnableMenuItemAPIを使って制御しようとしても
    'システムメニュー表示時に
    'メニューが勝手に有効化してしまう不具合があるようなので
    'DeleteMenu不採用とする
End Sub

Public Function GetWindowCloseButton(ByVal hWnd As Long) As Boolean
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, False)
    GetWindowCloseButton = (GetMenuItemID(hMenu, 6) <> -1)
End Function

'----------------------------------------
'・TopMost
'----------------------------------------
Public Sub SetWindowTopMost(hWnd As Long, _
ByVal TopMost As Boolean)

    If TopMost Then
        Call SetWindowPos(hWnd, _
            HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE Or SWP_NOSIZE)
    Else
        Call SetWindowPos(hWnd, _
            HWND_NOTOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
    End If
End Sub

Public Function GetWindowTopMost(hWnd As Long) As Boolean
    GetWindowTopMost = _
        WS_EX_TOPMOST = (GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_TOPMOST)
End Function

'----------------------------------------
'・WindowState
'----------------------------------------
Public Function GetWindowState(ByVal hWnd As Long) As Excel.XlWindowState
    Dim Result As Excel.XlWindowState
    Dim wp As WINDOWPLACEMENT
    Call GetWindowPlacement(hWnd, wp)
    Select Case wp.showCmd
    Case SW_SHOWNORMAL
        Result = xlNormal
    Case SW_SHOWMINIMIZED
        Result = xlMinimized
    Case SW_SHOWMAXIMIZED
        Result = xlMaximized
    Case Else
        Call Assert(False, "Error:GetWindowState")
    End Select
    GetWindowState = Result
End Function

'----------------------------------------
'・PixelRect
'----------------------------------------

Public Function Form_GetRectPixel(Form As Object) As Rect
    Dim Result As Rect: Result = NewRect(0, 0, 0, 0)
    Result.Left = PointToPixel(Form.Left)
    Result.Top = PointToPixel(Form.Top)
    Result.Right = PointToPixel(Form.Left + Form.Width)
    Result.Bottom = PointToPixel(Form.Top + Form.Height)
    Form_GetRectPixel = Result
End Function

Public Sub Form_SetRectPixel(ByVal Form As Object, _
ByRef RectValue As Rect)
    Form.Left = PixelToPoint(RectValue.Left)
    Form.Top = PixelToPoint(RectValue.Top)
    Form.Width = PixelToPoint(RectValue.Right - RectValue.Left)
    Form.Height = PixelToPoint(RectValue.Bottom - RectValue.Top)
End Sub

'----------------------------------------
'・Iniファイル位置保存復帰
'----------------------------------------
Public Sub Form_IniWritePosition(Form As Object, _
ByVal IniFilePath As String, _
ByVal Section As String, ByVal Name As String)
    Call IniFile_SetString( _
        IniFilePath, Section, Name, _
        RectToStr(Form_GetRectPixel(Form)))
End Sub

Public Sub Form_IniReadPosition(Form As Object, _
ByVal IniFilePath As String, _
ByVal Section As String, ByVal Name As String, _
ByVal PositionOnly As Boolean)

    Dim RectStr As String
    RectStr = IniFile_GetString( _
        IniFilePath, Section, Name, "")
    Dim r As Rect
    If CanStrToRect(RectStr) Then
        r = StrToRect(RectStr)
        If PositionOnly Then
            r = NewRect_PositionSize( _
                    NewPoint(r.Left, r.Top), _
                    GetRectSize(Form_GetRectPixel(Form)))
        End If
    Else
        r = Form_GetRectPixel(Form)
        r = GetRectMoveCenter(r, GetPointRectCenter(GetRectWorkArea))
    End If
    Call Form_SetRectPixel(Form, GetRectInsideDesktopRect(r, GetRectWorkArea))
End Sub


'----------------------------------------
'◆ComboBox
'----------------------------------------

'----------------------------------------
'・Combobox.Textをクリアせずに項目だけクリアする
'----------------------------------------
Public Sub Combobox_ClearList(ComboBox As ComboBox)
    Dim I As Long
    For I = ComboBox.ListCount - 1 To 0 Step -1
        ComboBox.RemoveItem (I)
    Next
End Sub

'----------------------------------------
'◇ComboBoxと文字列配列との変換
'----------------------------------------
'   ・  項目をタブで区切る
'   ・  タブ区切り文字列とColumnCountは一致させておくこと
'----------------------------------------
Public Function ComboBox_GetStrings(ByVal ComboBox As ComboBox) As String()
    Dim Result() As String
    Dim Item As String
    Do
        If ComboBox.ListCount = 0 Then Exit Do
        ReDim Result(ComboBox.ListCount - 1)
        Dim I As Long
        Dim J As Long
        For I = 0 To ComboBox.ListCount - 1
            Item = ""
            For J = 0 To ComboBox.ColumnCount - 1
                Item = Item + ComboBox.List(I, J) + vbTab
            Next
            Result(I) = ExcludeLastStr(Item, vbTab)
        Next
    Loop While False
    ComboBox_GetStrings = Result
End Function

Sub ComboBox_SetStrings(ComboBox As ComboBox, Strings() As String)
    Dim I As Long
    Dim J As Long
    Dim Data() As String
    Dim Line() As String

    'エラーチェック
    'タブ区切り文字がComboBoxの列数とあっているかどうか
    For I = 0 To ArrayCount(Strings) - 1
        Call Assert(ArrayCount(Split(Strings(I), vbTab)) = ComboBox.ColumnCount, "Error:ComboBox_SetStrings")
    Next

    'ComboBox.Clear
    Call Combobox_ClearList(ComboBox)
    If ArrayCount(Strings) = 0 Then Exit Sub
    ReDim Data(ComboBox.ColumnCount - 1, ArrayCount(Strings) - 1)

    For I = 0 To ArrayCount(Strings) - 1
        Line = Split(Strings(I), vbTab)
        For J = 0 To ComboBox.ColumnCount - 1
            Data(J, I) = Line(J)
        Next
    Next
    ComboBox.Column() = Data
End Sub

Sub testComboBox_GetSetStrings(ComboBox1 As ComboBox)
    Dim myData(2, 2) As String
        myData(0, 0) = "A"
        myData(0, 1) = "B"
        myData(0, 2) = "C"
        myData(1, 0) = "あああ"
        myData(1, 1) = "いいい"
        myData(1, 2) = "ううう"
        myData(2, 0) = "1"
        myData(2, 1) = "2"
        myData(2, 2) = "3"

        With ComboBox1
            .ColumnCount = 3
            .ColumnWidths = "50;50;50"
            .Column() = myData
        End With

    Dim Data() As String
    Data = ComboBox_GetStrings(ComboBox1)
    Call Check( _
        "A" + vbTab + "あああ" + vbTab + "1" + vbCrLf + _
        "B" + vbTab + "いいい" + vbTab + "2" + vbCrLf + _
        "C" + vbTab + "ううう" + vbTab + "3", _
        ArrayToString(Data, vbCrLf))
    Call ArrayAdd(Data, "D" + vbTab + "えええ" + vbTab + "4")
    Call ComboBox_SetStrings(ComboBox1, Data)
    Data = ComboBox_GetStrings(ComboBox1)
    Call Check( _
        "A" + vbTab + "あああ" + vbTab + "1" + vbCrLf + _
        "B" + vbTab + "いいい" + vbTab + "2" + vbCrLf + _
        "C" + vbTab + "ううう" + vbTab + "3" + vbCrLf + _
        "D" + vbTab + "えええ" + vbTab + "4", _
        ArrayToString(Data, vbCrLf))
End Sub



'----------------------------------------
'◆アイコン用API操作
'----------------------------------------

Public Sub SetWindowIcon(ByVal hWnd As Long, _
ByVal IconPath As String, ByVal IconIndex As Long)
    Dim hIcon As Long
    hIcon = ExtractIcon(0, IconPath, IconIndex)
    If hIcon <> 0 Then
        Call SendMessage(hWnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
        Call SendMessage(hWnd, WM_SETICON, ICON_BIG, ByVal hIcon)
        Call DrawMenuBar(hWnd)
        Call DestroyIcon(hIcon)
    End If
End Sub

Public Function GetWindowIcon(ByVal hWnd As Long) As Boolean
    Dim hIcon As Long
    hIcon = SendMessage(hWnd, WM_GETICON, ICON_SMALL, ByVal 0&)
    GetWindowIcon = (hIcon <> 0)
End Function

Public Sub ResetWindowIcon(ByVal hWnd As Long)
    Call SendMessage(hWnd, WM_SETICON, ICON_SMALL, ByVal 0&)
    Call SendMessage(hWnd, WM_SETICON, ICON_BIG, ByVal 0&)
    Call DrawMenuBar(hWnd)
End Sub

Public Function SystemIconFilePath() As String
    SystemIconFilePath = _
        PathCombine( _
            GetSpecialFolderPath(System), _
            "imageres.dll")
End Function

Public Function NewIconFilePathIndex( _
ByVal Path As String, ByVal Index As Long) As IconFilePathIndex
    NewIconFilePathIndex.Path = Path
    NewIconFilePathIndex.Index = Index
End Function

Public Function GetBitmapDrawIcon( _
ByRef IconInfo As IconFilePathIndex, _
ByRef IconSize As RectSize) As Long

    Dim hIcon As Long
    hIcon = ExtractIcon(0, IconInfo.Path, IconInfo.Index)

    Dim hDC As Long
    Dim hBitmap As Long
    Dim hBitmapOld As Long
    hDC = CreateCompatibleDC(GetDC(0))
    hBitmap = CreateCompatibleBitmap(GetDC(0&), IconSize.Width, IconSize.Height)
    hBitmapOld = SelectObject(hDC, hBitmap)

    Dim r As Rect
    r = NewRect(0, 0, IconSize.Width, IconSize.Height)
    Call FillRect(hDC, r, GetStockObject(0))
    Call DrawIcon(hDC, 0, 0, hIcon)

    Call SelectObject(hDC, hBitmapOld)
    Call DeleteObject(hDC)
    Call DestroyIcon(hIcon)

    GetBitmapDrawIcon = hBitmap
End Function

Public Sub Image_Picture_SetBitmap( _
ByRef Image As Image, _
ByVal hBitmap As Long)
    Dim Pic As StdPicture
    Dim hPalette As Long
    Dim pd As PictDesc
    Dim g As guid
    pd.cbSizeOfStruct = Len(pd)
    pd.picType = 1
    pd.hImage = hBitmap
    pd.Option1 = hPalette
    g.Data1 = &H20400
    g.Data4(0) = &HC0
    g.Data4(7) = &H46
    Call OleCreatePictureIndirect(pd, g, 0&, Pic)
    Image.Picture = Pic
End Sub

'----------------------------------------
'◆Windows情報
'----------------------------------------

'----------------------------------------
'・デスクトップ/WorkAreaサイズ
'----------------------------------------
Public Function GetRectDesktop() As Rect
    GetRectDesktop = GetRectDesktop1
End Function

Public Function GetRectDesktop1() As Rect
    Dim Result As Rect
    Call GetWindowRect(GetDesktopWindow, Result)
    GetRectDesktop1 = Result
End Function

Public Function GetRectDesktop2() As Rect
    Dim Result As Rect
    Result = NewRect(0, 0, _
        GetSystemMetrics(SM_CXSCREEN), _
        GetSystemMetrics(SM_CYSCREEN))
    GetRectDesktop2 = Result
End Function

Sub testGetRectDesktop()
    Call Check(True, RectEqual( _
        GetRectDesktop1, GetRectDesktop2))
End Sub


Public Function GetRectWorkArea() As Rect
    Dim Result As Rect
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, _
        Result, 0)
    GetRectWorkArea = Result
End Function

Sub testGetRectWorkArea()
    Call MsgBox(RectToStr(GetRectWorkArea))
End Sub

'----------------------------------------
'・Windowsバージョン
'----------------------------------------

'Windows 8.1 →     Windows (32-bit) NT 6.03(たぶん)
'Windows 8 →       Windows (32-bit) NT 6.02
'Windows 7 →       Windows (32-bit) NT 6.01
'Windows Vista →   Windows (32-bit) NT 6.00
'Windows XP →      Windows (32-bit) NT 5.01
'Windows 2000 →    Windows (32-bit) NT 5.00
'Windows Me →      Windows (32-bit) 4.90
'Windows 98 →      Windows (32-bit) 4.10
'Windows 95 →      Windows (32-bit) 4.00
Public Function IsWindowsOffice64bit() As Boolean
    IsWindowsOffice64bit = _
        (1 <= InStr(Application.OperatingSystem, _
            "Windows (64-bit)"))
End Function

Public Function IsWindowsOffice32bit() As Boolean
    IsWindowsOffice32bit = _
        (1 <= InStr(Application.OperatingSystem, _
            "Windows (32-bit)"))
End Function

Public Function WindowsMajorVersion() As Long
    WindowsMajorVersion = _
        CLng(FirstStrFirstDelim( _
            LastStrLastDelim( _
                Application.OperatingSystem, " "), "."))
End Function

Public Function WindowsMinorVersion() As Long
    WindowsMinorVersion = _
        CLng(LastStrLastDelim( _
            LastStrLastDelim( _
                Application.OperatingSystem, " "), "."))
End Function


Sub testWindowsOfficeVersion()
    MsgBox BoolToStr(IsWindowsOffice32bit)
    MsgBox WindowsMajorVersion
    MsgBox WindowsMinorVersion
End Sub

'----------------------------------------
'◆タスクバーピンアイコン登録用
'----------------------------------------

Public Function IsTaskbarPinWindows() As Boolean
    If (6 <= WindowsMajorVersion) _
    And (1 <= WindowsMinorVersion) Then
        IsTaskbarPinWindows = True
    Else
        IsTaskbarPinWindows = False
    End If
End Function

'----------------------------------------
'・タスクバーボタン用のAppIDの登録
'----------------------------------------
Public Sub SetTaskbarButtonAppID(ByVal AppID As String)
    If IsTaskbarPinWindows Then
        Call SetCurrentProcessExplicitAppUserModelID( _
            StrPtr(AppID))
    End If
End Sub

'----------------------------------------
'・タスクバーピン止め用コマンド
'----------------------------------------
Public Sub SetTaskbarPin(ByVal FilePath As String, ByVal Value As Boolean)
    Dim CommandVerb As String
    If Value Then
        Call CreateObject("Shell.Application"). _
            Namespace(fso.GetParentFolderName(FilePath)). _
            ParseName(fso.GetFileName(FilePath)).InvokeVerb("taskbarpin")
    Else
        Call CreateObject("Shell.Application"). _
            Namespace(fso.GetParentFolderName(FilePath)). _
            ParseName(fso.GetFileName(FilePath)).InvokeVerb("taskbarunpin")
    End If
    'InvokeVerbの後の文字列は変数ではダメ
    'なぜか定数でないといけない。
End Sub

'----------------------------------------
'・タスクバーピン用ショートカットファイルの作成/削除
'----------------------------------------
Public Sub SetTaskbarPinShortcutIcon(ByVal Value As Boolean, _
ByVal ShortcutFilePath As String, ByVal LinkTargetFilePath As String, _
ByVal IconFilePath As String, _
ByVal Description As String, _
ByVal DummyLinkTargetFilePath As String, _
ByVal DummyTaskbarPinFileName As String, _
ByVal TaskbarPinCommandProgramPath As String, _
ByVal AppID As String)

    Dim FileExistsFlag As Boolean
    FileExistsFlag = fso.FileExists(ShortcutFilePath)
    If (Value) And (FileExistsFlag = False) Then

        'タスクバーにピン止め
        Call SetTaskbarPin(DummyLinkTargetFilePath, True)
        Call fso.MoveFile( _
            PathCombine(GetSpecialFolderPath(TaskbarPin), DummyTaskbarPinFileName), _
            ShortcutFilePath)
        Call CommandExecuteReturn( _
            InSpacePlusDoubleQuote(TaskbarPinCommandProgramPath) + _
            " " + _
            InSpacePlusDoubleQuote(ShortcutFilePath) + _
            " " + _
            AppID)

        'ショートカットファイルのリンク先変更
        If FileExistsWait(ShortcutFilePath) Then
            Call CreateShortcutFile(ShortcutFilePath, _
                LinkTargetFilePath, _
                IconFilePath, Description)
        End If
        'スクリプトファイルを直接はタスクバーピン止めできないので
        '一度ダミーのプログラムを登録して
        'その後でショートカットファイルのリンク先を書き換えている。

    ElseIf (Value = False) And (FileExistsFlag) Then

        'タスクバーピン解除
        Call SetTaskbarPin(ShortcutFilePath, False)
    End If
End Sub

'----------------------------------------
'◆キーボード操作
'----------------------------------------

Sub NumLockOn()
  Dim NumLockState As Boolean
  Dim keys(0 To 255) As Byte

  Call GetKeyboardState(keys(0))
  NumLockState = keys(VK_NUMLOCK)
       
  '「NumLock」キーがオフの場合はオンにする。
  If NumLockState <> True Then
    'キーを押す
    Call keybd_event(VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0)
    'キーを放す
    Call keybd_event(VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
  End If
End Sub

'----------------------------------------
'◆マウス操作
'----------------------------------------

Public Sub MouseMove(ByRef Position As Point)
    Dim DesktopRect As Rect
    DesktopRect = GetRectDesktop

    Call mouse_event(MOUSE_MOVED Or MOUSEEVENTF_ABSOLUTE, _
        Position.X * (65535 / GetRectWidth(DesktopRect)), _
        Position.Y * (65535 / GetRectHeight(DesktopRect)), 0, 0)
    '↑クリック位置を画面解像度から補正する
End Sub

Public Sub MouseClick()
   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
End Sub

'----------------------------------------
'◆Internet系関数
'----------------------------------------
'----------------------------------------
'・URL指定のファイルダウンロード
'----------------------------------------
'   ・  APIのURLDownloadToFileを使いやすくした
'----------------------------------------
Public Function URLDownloadFile(ByVal URL As String, ByVal FilePath As String) As Boolean
    Dim Result As Long
    Result = URLDownloadToFile(0, URL, FilePath, 0, 0)
    URLDownloadFile = (Result = 0)
End Function

'----------------------------------------
'・日本語文字列のURLエンコード
'----------------------------------------
Public Function UrlEncode(ByVal Word As String) As String
    Dim HtmlFile As Object
    Dim Element As Object
    Word = Replace(Word, "\", "\\")
    Word = Replace(Word, "'", "\'")
    Set HtmlFile = CreateObject("htmlfile")
    Set Element = HtmlFile.createElement("span")
    Call Element.setAttribute("id", "result")
    Call HtmlFile.appendChild(Element)
    Call HtmlFile.parentWindow.execScript("document.getElementById('result').innerText = encodeURIComponent('" & Word & "');", "JScript")
    UrlEncode = Element.InnerText
End Function


'----------------------------------------
'◆IPアドレス
'----------------------------------------

'----------------------------------------
'・文字列がIPアドレスなのかどうかを判定する関数
'----------------------------------------
Public Function IsIPAddress( _
ByVal Value As String) As Boolean
    Dim Result As Boolean
    Result = False
    If StrCount(Value, ".") = 3 Then
        Dim S() As String
        S = Split(Value, ".")
        
        If IsLong(S(0)) _
        And IsLong(S(1)) _
        And IsLong(S(2)) _
        And IsLong(S(3)) Then
            
            Result = InRange(0, CLng(S(0)), 255) _
                And InRange(0, CLng(S(1)), 255) _
                And InRange(0, CLng(S(2)), 255) _
                And InRange(0, CLng(S(3)), 255) _
        
        End If
    End If
    
    IsIPAddress = Result
End Function

'----------------------------------------
'・IPアドレス文字列をCurrency型にする
'----------------------------------------
Public Function IPAddressToCurrency( _
ByVal IPAddressText As String) As Currency
    Dim Result As Currency
    Call Assert(IsIPAddress(IPAddressText), "Error:IPAddressToCurrency")
    
    Dim S() As String
    S = Split(IPAddressText, ".")
    
    Result = _
        CCur(LongToStrDigitZero(S(0), 3)) * 1000000000 + _
        CCur(LongToStrDigitZero(S(1), 3)) * 1000000 + _
        CCur(LongToStrDigitZero(S(2), 3)) * 1000 + _
        CCur(LongToStrDigitZero(S(3), 3)) * 1

    IPAddressToCurrency = Result
End Function

'----------------------------------------
'・IPアドレスが指定範囲内にあるかどうかを確認する
'----------------------------------------
Public Function InRangeIPAddress( _
ByVal MinValue As String, _
ByVal Value As String, _
ByVal MaxValue As String) As Boolean

    InRangeIPAddress = _
    ( _
        (IPAddressToCurrency(MinValue) <= IPAddressToCurrency(Value)) _
        And _
        (IPAddressToCurrency(Value) <= IPAddressToCurrency(MaxValue)) _
    )
    
End Function


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
'・ Book_FullPath の修正
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
'--------------------------------------------------
