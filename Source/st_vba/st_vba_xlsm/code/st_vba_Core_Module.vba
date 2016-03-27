'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    Base Module
'ObjectName:    st_vba_Core
'--------------------------------------------------
'Discription:   Standard Software Library For Windows Excel VBA
'--------------------------------------------------
'OpenSource:    https://github.com/standard-software/st_vba/
'License:       MIT License
'   URL:        https://github.com/standard-software/st_vba/blob/master/Document/Readme_jp.txt
'All Right Reserved:
'   Name:       Standard Software
'   URL:        http://standard-software.net/
'--------------------------------------------------
'Version:       2016/03/27
'--------------------------------------------------

'--------------------------------------------------
'���}�[�N
'--------------------------------------------------

    '--------------------------------------------------
    '��
    '--------------------------------------------------

    '----------------------------------------
    '��
    '----------------------------------------
    '��
    '----------------------------------------
    '�E
    '----------------------------------------

'--------------------------------------------------
'���Q�Ɛݒ�
'�E Microsoft Scripting Runtime
'       FileSystemObject
'�E Windows Script Host Object Model
'       WshShell
'�E Microsoft AxtiveX Data Objects 6.1 Library
'       ADODB.Stream
'�E Microsoft Forms 2.0 Object Library
'       Image
'       ComboBox
'�E Microsoft Internet Controls
'       InternetExplorer
'�E Microsoft Windows Common Controls 6.0 (SP6)
'       ListView
'       32bit Windows / 32bit Excel
'           C:\Windows\system32\MSCOMCTL.OCX
'       64bit Windows / 32bit Excel
'           C:\Windows\SysWOW64\mscomctl.ocx
'       64bit Windows / 64bit Excel
'           �d�l�s��
'--------------------------------------------------
'�E Microsoft Windows Common Controls 6.0 (SP6)
'       64bit Windows / 32bit Excel
'           C:\Windows\SysWOW64\mscomctl.ocx
'   �E  http://www.microsoft.com/ja-jp/download/details.aspx?id=10019
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
'���萔/�^�錾
'--------------------------------------------------

'----------------------------------------
'���ʒu�E�T�C�Y
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
'��FileSystemObject
'----------------------------------------
Public fso As New FileSystemObject

'----------------------------------------
'��Shell
'----------------------------------------
Public Shell As New WshShell

'----------------------------------------
'���������r
'----------------------------------------
Public Enum MatchType
    FullMatch = 0   '���S��v
    PartMatch = 1   '������v
    WildCardValue = 2
    WildCardArray = 3
    RegExpValue = 4
    RegExpArray = 5
End Enum

Public Enum CaseCompare
    CaseSensitive
    IgnoreCase
End Enum

'----------------------------------------
'���z��
'----------------------------------------
Public Enum SortOrder
    Ascending
    Descending
End Enum

'----------------------------------------
'��Excel
'----------------------------------------

'----------------------------------------
'����w��
'----------------------------------------

Public Const Col__A = 1, Col__B = 2, Col__C = 3, Col__D = 4, Col__E = 5, Col__F = 6
Public Const Col__G = 7, Col__H = 8, Col__I = 9, Col__J = 10, Col__K = 11, Col__L = 12
Public Const Col__M = 13, Col__N = 14, Col__O = 15, Col__P = 16, Col__Q = 17, Col__R = 18
Public Const Col__S = 19, Col__T = 20, Col__U = 21, Col__V = 22, Col__W = 23, Col__X = 24
Public Const Col__Y = 25, Col__Z = 26
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
'��Cell�폜����
'----------------------------------------
'   �E  ClearComments/ClearOutline��
'       ���ɗp�r���Ȃ������Ȃ̂Ŏ������Ȃ�����
'   �E  rcClear:            �S�ăN���A
'       rcClearContents:    �����E������̃N���A
'       rcClearFormats:     �����̃N���A
'----------------------------------------
Enum RangeClearType
    rcClear
    rcClearContents
    rcClearFormats
    
End Enum

'----------------------------------------
'���O���t����
'----------------------------------------
Public Type GraphFormulaData
    SeriesName As String
    ItemXAxisRangeStr As String
    DataRangeStr As String
    SeriesNumber As Long
End Type

'----------------------------------------
'�����j���[����
'----------------------------------------
Private PopupMenu_Return As String

'----------------------------------------
'���t�@�C���t�H���_�p�X�擾
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
'���V�X�e��
'----------------------------------------
#If VBA7 And Win64 Then
    Const Excel64bit As Boolean = True
#Else
    Const Excel64bit As Boolean = False
#End If

'--------------------------------------------------
'��API
'--------------------------------------------------

'----------------------------------------
'���t�@�C������
'----------------------------------------

'�t�@�C�����쐬�܂��̓I�[�v��
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

'�V�X�e���^�C�����t�@�C���^�C���ɕϊ�����
Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32.dll" ( _
    ByRef lpSystemTime As SYSTEMTIME, _
    ByRef lpFileTime As FILETIME) As Long

'���[�J���t�@�C���^�C����UTC�t�@�C���^�C���`���Ŏ擾����
Private Declare PtrSafe Function LocalFileTimeToFileTime Lib "kernel32.dll" ( _
    ByRef lpLocalFileTime As FILETIME, _
    ByRef lpFileTime As FILETIME) As Long

'�t�@�C���̃t�@�C���^�C����ݒ肷��
Private Declare PtrSafe Function SetFileTime Lib "kernel32.dll" ( _
    ByVal hFile As Long, _
    ByRef lpCreationTime As FILETIME, _
    ByRef lpLastAccessTime As FILETIME, _
    ByRef lpLastWriteTime As FILETIME) As Long

'FILETIME �\����
Private Type FILETIME
    dwLowDateTime As Long    '����32�r�b�g�l
    dwHighDateTime As Long   '���32�r�b�g�l
End Type

'SECURITY_ATTRIBUTES �\����
Private Type SECURITY_ATTRIBUTES
    nLength              As LongPtr '�\���̂̃o�C�g��
    lpSecurityDescriptor As LongPtr '�Z�L�����e�B�f�X�N���v�^(Win95,98�ł͖���)
    bInheritHandle       As LongPtr '1�̂Ƃ��������p������
End Type

'CreateFile�Ŏg�p����萔
Private Const FILE_FLAG_BACKUP_SEMANTICS As Long = &H2000000  'NT�nOS�̂�
Private Const GENERIC_READ               As Long = &H80000000
Private Const GENERIC_WRITE              As Long = &H40000000
Private Const FILE_SHARE_READ            As Long = &H1
Private Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
Private Const OPEN_EXISTING              As Long = 3
Private Const OPEN_ALWAYS                As Long = 4
Private Const INVALID_HANDLE_VALUE       As Long = &HFFFFFFFF

'SYSTEMTIME �\����
Private Type SYSTEMTIME
    wYear         As Integer '�N
    wMonth        As Integer '��
    wDayOfWeek    As Integer '�j��(��=0, ��=1 ...)
    wDay          As Integer '��
    wHour         As Integer '��
    wMinute       As Integer '��
    wSecond       As Integer '�b
    wMilliseconds As Integer '�~���b
End Type

Public Type FileFolderTime
    CreataionTime As Date
    LastWriteTime As Date
    LastAccessTime As Date
End Type

'----------------------------------------
'��Ini�t�@�C��
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
'���L�[�{�[�h����
'----------------------------------------
Public Declare PtrSafe Function GetAsyncKeyState _
    Lib "User32.dll" (ByVal vKey As Long) As Long

'----------------------------------------
'���}�E�X
'----------------------------------------
Public Declare PtrSafe Function GetCursorPos _
    Lib "user32" (lpPoint As Point) As Long

Public Declare PtrSafe Sub mouse_event _
    Lib "user32" ( _
    ByVal dwFlags As Long, _
    ByVal dx As Long, ByVal dy As Long, _
    ByVal cButtons As Long, _
    ByVal dwExtraInfo As Long)

Public Const MOUSE_MOVED = &H1              '�}�E�X���ړ�����(���΍��W)
Public Const MOUSEEVENTF_ABSOLUTE = &H8000& 'MOUSE_MOVED or �Ő�΍��W���w��
Public Const MOUSEEVENTF_LEFTUP = &H4       '���{�^��UP
Public Const MOUSEEVENTF_LEFTDOWN = &H2     '���{�^��Down
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20  '�����{�^��Down
Public Const MOUSEEVENTF_MIDDLEUP = &H40    '�����{�^��UP
Public Const MOUSEEVENTF_RIGHTDOWN = &H8    '�E�{�^��Down
Public Const MOUSEEVENTF_RIGHTUP = &H10     '�E�{�^��UP

'----------------------------------------
'���}�E�X�{�^���C�x���g
'----------------------------------------
Enum MouseButton
    fmButtonLeft = 1       '���t�g�{�^���N���b�N
    fmButtonRight = 2      '���C�g�{�^���N���b�N
    fmButtonLeftRight = 3  '���t�g+���C�g�{�^���𓯎��N���b�N
    fmButtonMiddle = 4     '���{�^���N���b�N
End Enum


'----------------------------------------
'��Form
'----------------------------------------

'----------------------------------------
'��Window�n���h��
'----------------------------------------
Public Declare PtrSafe Function WindowFromAccessibleObject _
    Lib "oleacc.dll" ( _
    ByVal IAcessible As Object, _
    ByRef hWnd As Long) As Long

'----------------------------------------
'��Window�X�^�C��
'----------------------------------------
Public Const GWL_HINSTANCE = (-6) '�C���X�^���X�n���h�����擾
Public Const GWL_HWNDPARENT = (-8) '�e�E�C���h�E�̃n���h�����擾
Public Const GWL_ID = (-12) '�E�C���h�E��ID���擾
Public Const GWL_EXSTYLE = (-20) '�g���E�C���h�E�X�^�C�����擾
Public Const GWL_STYLE = (-16) '�E�C���h�E�X�^�C�����擾
Public Const GWL_WNDPROC = (-16) '�E�C���h�E�֐��̃A�h���X���擾
Public Const GWL_USERDATA = (-21) '���[�U�[��`��32�r�b�g�l���擾

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
'��SystemMenu/Close�{�^��
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
'��FormIcon
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
'��Window�ʒu/TopMost
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
'��Window�ʒu/TopMost
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
'���V�X�e��
'----------------------------------------

'----------------------------------------
'�EOS�o�[�W����
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
'�ESleep
'----------------------------------------
#If VBA7 And Win64 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

'----------------------------------------
'���^�X�N�o�[�{�^���o�^
'----------------------------------------
Public Declare PtrSafe Function SetCurrentProcessExplicitAppUserModelID _
    Lib "shell32.dll" ( _
    ByVal lpString As LongPtr) As Long

'----------------------------------------
'��Window
'----------------------------------------
Public Declare PtrSafe Function GetDesktopWindow _
    Lib "user32" () As Long

'----------------------------------------
'���`��
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
'���A�C�R��
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
'��Rect
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
'���C���^�[�l�b�g
'----------------------------------------
Public Declare PtrSafe Function URLDownloadToFile _
    Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long


'--------------------------------------------------
'������
'--------------------------------------------------

'----------------------------------------
'���������f
'----------------------------------------
Public Sub Assert(ByVal Value As Boolean, Optional ByVal Message As String)
    If Value = False Then
        Call Err.Raise(9999, , Message)
    End If
End Sub

Private Sub testAssert()
    Call Assert(False, "�e�X�g")
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
'�EOrValue
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
'�EIsNothing/IsNotNothing
'----------------------------------------
Public Function IsNothing(ByRef Value As Object) As Boolean
    IsNothing = (Value Is Nothing)
End Function

Public Function IsNotNothing(ByRef Value As Object) As Boolean
    IsNotNothing = Not (Value Is Nothing)
End Function

'----------------------------------------
'�E�l���󕶎��̏ꍇ�����ʂ̒l��Ԃ��֐�
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
'���^�A�^�ϊ�
'----------------------------------------

'----------------------------------------
'�E�ϐ��ɒl��I�u�W�F�N�g���Z�b�g����
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
'��Long
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
'���J���}�t�������̕ϊ�
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
'��Boolean
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
'��Point
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
'��Rect
'----------------------------------------

'----------------------------------------
'�ERect������ϊ�
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
'�ERect����
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
'�ERect��r
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
'�ERect Width/Height�l�擾
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
'��Rect Get�n
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

'�͂ݏo���Ă����璆�ɂ����
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
'��RectSize
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
'��Pixel Point ���ݕϊ�
'----------------------------------------

Public Function GetDPI() As Long
    Dim Result As Long: Result = 0

    Dim hWnd As Long
    Dim hDC As Long
    hWnd = Excel.Application.hWnd
    hDC = GetDC(hWnd)
    '��������DPI
    Result = GetDeviceCaps(hDC, LOGPIXELSX)
    '��������DPI
    Result = GetDeviceCaps(hDC, LOGPIXELSY)
    Call ReleaseDC(hWnd, hDC)

    GetDPI = Result
 End Function

'96�����擾�ł��Ȃ��I
Public Function GetDPI1() As Long
    GetDPI1 = ActiveWorkbook.WebOptions.PixelsPerInch
End Function

'120�����擾�ł��Ȃ��I
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
'�����l����
'----------------------------------------

'----------------------------------------
'�E�ő�l�ŏ��l
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
'�E�l�͈�
'----------------------------------------
Public Function InRange(ByVal MinValue As Long, _
ByVal Value As Long, ByVal MaxValue As Long) As Boolean
    InRange = ((MinValue <= Value) And (Value <= MaxValue))
End Function

'----------------------------------------
'�������񏈗�
'----------------------------------------

'----------------------------------------
'�EStrCount
'----------------------------------------
'   �E  ������̐��𐔂���֐�
'       AAA����AA�𐔂����2��Ԃ�
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
'�E�A���X�y�[�X��P�ƃX�y�[�X�ɕϊ�
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
'�EIsInclude
'----------------------------------------
Public Function IsIncludeStr(ByVal str As String, ByVal SubStr As String)
    IsIncludeStr = _
        (1 <= InStr(str, SubStr))
End Function

'----------------------------------------
'��First / Last
'----------------------------------------

'----------------------------------------
'�EFirst
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
'�ELast
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
'�EBoth
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
'��First / Last Delim
'----------------------------------------

'----------------------------------------
'�EFirstStrFirstDelim
'----------------------------------------
'   �E  �擪�Ō�����΋󕶎���Ԃ�
'   �E  ������Ȃ���Ε��������̂܂ܕԂ�
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
'�EFirstStrLastDelim
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
'�ELastStrFirstDelim
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
'�ELastStrLastDelim
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
'��Tag����
'----------------------------------------

'----------------------------------------
'�E�^�O�̓���������
'----------------------------------------
Public Function TagInnerText(ByVal Text As String, _
    ByVal StartTag As String, ByVal EndTag As String) As String
    
    Dim Result As String
    Result = IfEmptyStr(LastStrFirstDelim(Text, StartTag), Text)
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
End Sub

'----------------------------------------
'�E�^�O���܂񂾓���������
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
'�E�w��̃^�O�ł͂��܂ꂽ������̃��X�g���o�͂���
'----------------------------------------
'   �E ���ʂ͉��s�R�[�h�ŋ�؂��ďo�͂����
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
'�EHTML�^�O���폜����֐�
'----------------------------------------
 Function ReplaceHTMLTag(ByVal txt As String) As String
 With CreateObject("vbscript.regexp")
     .Pattern = "<[^>]*>"
     .Global = True
     ReplaceHTMLTag = .Replace(txt, "")
 End With
 End Function

'----------------------------------------
'��Trim
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
    TrimFirstSpace = TrimFirstStrs(str, ArrayStr(" ", vbCr, vbLf, vbTab))
End Function

Public Function TrimLastSpace(ByVal str As String) As String
    TrimLastSpace = TrimLastStrs(str, ArrayStr(" ", vbCr, vbLf, vbTab))
End Function

Public Function TrimBothEndsSpace(ByVal str As String) As String
    TrimBothEndsSpace = _
        TrimFirstSpace(TrimLastSpace(str))
End Function


'----------------------------------------
'�������񌋍�
'----------------------------------------

'----------------------------------------
'�E�����񌋍�
'----------------------------------------
'   �E  ���Ȃ��Ƃ�1��Delimiter���Ԃɓ����Đڑ������B
'   �E  Delimiter�������̗��[�ɕt������ꍇ��1�ɂȂ�B
'   �E  2�A���Ō����̗��[�ɂ���ꍇ��1���폜�����
'       (�e�X�g�ł̓���Q��)
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
'��Byte�w�蕶���񏈗�
'----------------------------------------
'   �E  SJIS�ɕϊ����Ă��邽��Unicode�ŗL�����͔�Ή�
'       �W����LenB�֐���Unicode���������Ȃ̂�SJIS�Ή��ł͂Ȃ�
'----------------------------------------

'----------------------------------------
'�EByte���擾
'----------------------------------------
Public Function LengthSjisByte(ByVal S As String) As Long
    LengthSjisByte = LenB(StrConv(S, vbFromUnicode))
End Function

'----------------------------------------
'�EByte���Ő؂�o��Left�֐�
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
'�EByte���Ő؂�o��Right�֐�
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
'�EByte���Ő؂�o��Mid�֐�
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
'�������񐳋K�\��
'----------------------------------------

'----------------------------------------
'�E���K�\���̈�v���m�F����
'----------------------------------------
'   �E  ����Ώۂ�1�s�e�L�X�g
'   �E  RegExp�I�u�W�F�N�g�͊O������w��\
'   �E  Match�I�u�W�F�N�g�͖߂�l�Ƃ��ė��p�\
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
        
        '���K�\���p�I�u�W�F�N�g�p��
        Dim RegCreateFlag As Boolean
        RegCreateFlag = False
        If RegExp Is Nothing Then
            RegCreateFlag = True
            Set RegExp = CreateObject("VBScript.RegExp")
        End If
        
        '���K�\���}�b�`����
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
'�E���K�\���ł̒u������
'----------------------------------------
'   �E  ����Ώۂ�1�s�e�L�X�g
'   �E  RegExp�I�u�W�F�N�g�͊O������w��\
'----------------------------------------
Public Function ReplaceRegExp(ByVal Value As String, ByVal Pattern As String, _
ByVal NewValue As String, _
Optional ByVal CaseCompare As CaseCompare = CaseSensitive, _
Optional RegExp As Object = Nothing) As String

On Error GoTo Err:
    Dim Result As String: Result = Value
    Do
        If (Pattern = "") Or (Value = "") Then Exit Do
        
        '���K�\���p�I�u�W�F�N�g�p��
        Dim RegCreateFlag As Boolean
        RegCreateFlag = False
        If RegExp Is Nothing Then
            RegCreateFlag = True
            Set RegExp = CreateObject("VBScript.RegExp")
        End If
        
        '���K�\���}�b�`����
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
'���z��w�菈��
'----------------------------------------

'----------------------------------------
'�E������̘A���u�������A�z��w��
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
'�E������̘A���폜�A�z��w��
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
'���z��w��(���K�\��)����
'----------------------------------------
'----------------------------------------
'�E������̘A���u�������A���K�\���A�z��w��
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
'�E������̘A���폜�A���K�\���A�z��w��
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
'�����t��������
'----------------------------------------

'----------------------------------------
'�E���̍ŏI�����擾
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
'�E���̓����擾
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
'�����T/��T/���T�̗j���w��̓��t�擾
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
'�����t���������w��
'----------------------------------------

'----------------------------------------
'�E���t����
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
'�E��������
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
'�E�W���I�ȓ��t��������������̎擾
'----------------------------------------
Public Function FormatDateTimeNormal(DateValue As Date) As String
    FormatDateTimeNormal = _
        FormatYYYY_MM_DD(DateValue, "/") + _
        " " + _
        FormatHH_MM_SS(DateValue, ":")
End Function

'----------------------------------------
'�E���t��������
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
'���z�񏈗�
'----------------------------------------

'----------------------------------------
'���z���{����
'----------------------------------------

'----------------------------------------
'�E�v�f�����z��ɑ΂��Ă��G���[�̋N���Ȃ�UBound/LBound
'----------------------------------------
'   �E  UBound��Array()�ŕԂ����v�f�����̔z��ɂ�-1��Ԃ���
'       �錾���������̓��I�z��ł̓G���[�ɂȂ�̂ł����h�~����B
'   �E  Dimension:�������́A�������z��̏ꍇ���̎����ł̌��ʂ�Ԃ�
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
'�E�z��̗v�f�������߂�֐�
'----------------------------------------
'   �E  LBound=0 �ł� 1 �ł��Ή�����B
'   �E  Dimension:�������́A�������z��̏ꍇ���̎����ł̌��ʂ�Ԃ�
'----------------------------------------
Public Function ArrayCount(ByRef ArrayValue As Variant, _
Optional Dimension = 1) As Long
    Call Assert(IsArray(ArrayValue), "Error:ArrayCount:ArrayValue is not Array.")

    ArrayCount = _
        UBoundNoError(ArrayValue, Dimension) - _
        LBoundNoError(ArrayValue, Dimension) + 1
    '�z��v�f���Ȃ��ꍇ��UBound=-1/LBound=0�ɂȂ�̂�
    '�z��v�f���v�Z�͐������s����B
End Function

Private Sub testArrayCount()
    Dim A() As String
    Call Check(0, ArrayCount(A))
    Call Check(0, ArrayCount(Array()))
    Call Check(1, ArrayCount(Split("123", ",")))
    Call Check(2, ArrayCount(Split("1,3", ",")))
    
    '�񎟌��z��
    Dim B(3, 4) As String
    Call Check(4, ArrayCount(B, 1))
    Call Check(5, ArrayCount(B, 2))
    
    '�O�����z��
    Dim C(5, 6, 7) As String
    Call Check(6, ArrayCount(C, 1))
    Call Check(7, ArrayCount(C, 2))
    Call Check(8, ArrayCount(C, 3))
    
End Sub


'----------------------------------------
'�E�z��̗v�f��ǉ�����
'----------------------------------------
'   �E  �I�u�W�F�N�g�l�ɂ��Ή�
'   �E  ReDim Preserve�ɂ����
'       LBound(Array)=0�ɂȂ��Ă��܂�
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

    Dim B()
    ReDim B(2)
    Set B(0) = CreateObject("VBScript.RegExp")
    Set B(1) = Shell
    Set B(2) = CreateObject("ADODB.Stream")
    Call ArrayAdd(B, fso)
    Call Check("test.txt", B(3).GetFileName("C:\temp\test.txt"))
    
    '�񎟌��z��
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
'�E�z��̗v�f���d���`�F�b�N���Ă���ǉ�����
'----------------------------------------
Public Sub ArrayAddNotDuplicate(ByRef ArrayValue As Variant, ByVal Value As Variant)
    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    If ArrayExists(ArrayValue, Value) = False Then
        Call ArrayAdd(ArrayValue, Value)
    End If
End Sub


'----------------------------------------
'�E�z��̗v�f��}������
'----------------------------------------
'   �E  �I�u�W�F�N�g�l�ɂ��Ή�
'   �E  LBound(Array)=0�łȂ��Ă��Ή��B
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
'�E�z��̗v�f���폜����
'----------------------------------------
'   �E  LBound(Array)=0�łȂ��Ă��Ή��B
'   �E  �I�u�W�F�N�g�l�ɂ��Ή�
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
        '�z��̏�������Erase���g��
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
'�E�z��֐��̃e�X�g
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

  '�v�f�Ȃ��z��
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

  'LBound(Array)=0�ł͂Ȃ��z��
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
'�E�z����̒l����������Index��Ԃ�
'----------------------------------------
'   �E  LBound(Array)=0�łȂ��Ă��Ή��B
'   �E  �召������r�Ή�
'   �E  ���S��v/������v/���C���h�J�[�h/���K�\���Ή�
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
        '���͈̓G���[�̏ꍇ�ł�Result=-1��Ԃ������ŃG���[�ɂ͂��Ȃ�

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

    'Like WildCard Value IgnoreCase �S�p
    A = Array("�`�a�b", "�c�d�e", "�P�Q�R")
    Call Check(-1, ArrayIndexOf(A, "?��?", 1, , WildCardValue))
    Call Check(1, ArrayIndexOf(A, "?��?", 1, IgnoreCase, WildCardValue))

    'RegExp Value
    A = Array("ABC", "DEF", "123")
    Call Check(0, ArrayIndexOf(A, ".*C", 0, , RegExpValue))
    
    'RegExp Value IgnoreCase
    Call Check(-1, ArrayIndexOf(A, ".*c", 0, , RegExpValue))
    Call Check(0, ArrayIndexOf(A, ".*C", 0, , RegExpValue))

End Sub

'----------------------------------------
'�E�z����̒l���݃`�F�b�N
'----------------------------------------
Public Function ArrayExists(ByRef ArrayValue As Variant, _
ByVal Value As Variant, _
Optional CaseCompare As CaseCompare = CaseSensitive, _
Optional MatchType As MatchType = FullMatch) As Boolean

    ArrayExists = Not (ArrayIndexOf(ArrayValue, Value, , CaseCompare, MatchType) = -1)
End Function

'----------------------------------------
'�E�z����̒l���������ă��j�[�N(����l���Ȃ�)���ǂ����𔻒f����
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
'���z�񉞗p����
'----------------------------------------

'----------------------------------------
'�E�z����̒l���������ē���l���폜
'----------------------------------------
'   �E  LBound(Array)=0�łȂ��Ă��Ή��B
'       �d���������True/�Ȃ����False
Public Function ArrayDeleteSameItem(ByRef ArrayValue As Variant, _
Optional StartIndex As Long = -1) As Boolean
    Dim Result As Boolean: Result = False
    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    If StartIndex <> -1 Then
        Call Assert(((StartIndex < LBound(ArrayValue)) _
                And (UBound(ArrayValue) < StartIndex) = False), "Error:ArrayDeleteSameItem:Range Over")
        '���͈̓G���[�̏ꍇ������B
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
'�E�z��̗v�f�^�C�v�����߂�
'----------------------------------------
'   �E  LBound=0 �ł� 1 �ł��Ή�����B
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
'�E������z�񂩂ǂ���
'----------------------------------------
Public Function IsStrArray(ByVal ArrayValue As Variant) As Boolean
    IsStrArray = CheckArrayVarType(ArrayValue, vbString)
End Function

'----------------------------------------
'�E�z��𕶎���ɂ��ďo�͂���֐�
'----------------------------------------
'   �E  �v�f���Ȃ��Ă��Ή��B
'   �E  LBound(Array)=0�łȂ��Ă��Ή��B
'----------------------------------------
Public Function ArrayToString(ArrayValue As Variant, Delimiter As String) As String
    Call Assert(IsArray(ArrayValue), "�z��ł͂���܂���")

    Dim Result As String
    Result = ""
    Do
        If ArrayCount(ArrayValue) = 0 Then Exit Do

        Call Assert(ArrayDimension(ArrayValue) = 1, "1�����z��ł͂���܂���")
        Dim I As Long
        For I = LBound(ArrayValue) To UBound(ArrayValue)
          Result = Result + ArrayValue(I) + Delimiter
        Next
    Loop While False
    Result = ExcludeLastStr(Result, Delimiter)
    ArrayToString = Result
End Function

'----------------------------------------
'�E�p�����[�^�z��𕶎���z��ɂ��ĕԂ��֐�
'----------------------------------------
Public Function ArrayStr(ParamArray Values()) As String()
    '�p�����[�^�z���String�z��ɑ�����Ă���
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
'�E�z��𕶎���z��ɂ��ĕԂ��֐�
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
'�E�p�����[�^�z���Long�z��ɂ��ĕԂ��֐�
'----------------------------------------
Public Function ArrayLong(ParamArray Values()) As Long()
    '�p�����[�^�z���Long�z��ɑ�����Ă���
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
'���z��\�[�g
'----------------------------------------

'----------------------------------------
'�E�N�C�b�N�\�[�g
'----------------------------------------
'   �E  IndexMin/IndexMax���w�肷���
'       �w��͈͓��̒l���\�[�g����
'----------------------------------------
Public Sub ArraySortQuick(ByRef ArrayValue As Variant, _
Optional ByVal SortOrder As SortOrder = SortOrder.Ascending, _
Optional ByVal IndexMin As Long = -1, Optional ByVal IndexMax As Long = -1)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 1)
    
    Call Assert(IndexMin <= IndexMax, "Error:IndexMin < IndexMax")
    Call Assert(InRange(-1, IndexMin, ArrayCount(ArrayValue) - 1), _
        "Error:ArrayReverse:IndexMin Range is miss.")
    Call Assert(InRange(-1, IndexMax, ArrayCount(ArrayValue) - 1), _
        "Error:ArrayReverse:IndexMax Range is miss.")
    
    '1�ȉ��Ȃ�\�[�g�s�\�Ȃ̂�Exit����
    If ArrayCount(ArrayValue) <= 1 Then Exit Sub
    
    IndexMin = IIf(IndexMin = -1, 0, IndexMin)
    IndexMax = IIf(IndexMax = -1, ArrayCount(ArrayValue) - 1, IndexMax)
    
    'IndexMin=IndexMax�Ȃ�\�[�g�s�\�Ȃ̂�Exit
    If IndexMin = IndexMax Then Exit Sub
    
    Call ArraySortQuickBase(ArrayValue, SortOrder, IndexMin, IndexMax)
End Sub

'�N�C�b�N�\�[�g�̃x�[�X�֐��A�ċN�Ăяo�������
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
    
    '�����l���o�b�t�@
    Value1 = ArrayValue(IndexCenter)
    '�����l�ɊJ�n�ʒu�v�f����
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
    
    ' �����O�����ċA�Ăяo����SORT
    Call ArraySortQuickBase(ArrayValue, SortOrder, IndexMin, Index2 - 1)

    ' �����㔼���ċA�Ăяo����SORT
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
'�E�����񒷃\�[�g
'----------------------------------------
Public Sub ArraySortStrLength(ByRef ArrayValue As Variant, _
Optional ByVal SortOrder As SortOrder = SortOrder.Ascending)

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
'�E�Ǝ����я��\�[�g
'----------------------------------------
'   �E  �\�[�g�w��z��̕�����Ɉ�v���鏇�Ԃ�
'       ���ёւ�������\�[�g
'   �E  s/m/l/xl/xxl�Ƃ������������юw����s��
'----------------------------------------
Public Sub ArraySortCustomOrder(ByRef ArrayValue As Variant, _
ByRef OrderArrayWildCard() As String, _
Optional CaseCompare As CaseCompare = CaseCompare.IgnoreCase, _
Optional NoOrderValuePriority As Boolean = False)

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
    
    Array1 = ArrayStr("B�T�C�Y", "A�T�C�Y", _
        "S�T�C�Y", "SS�T�C�Y", "XL�T�C�Y", "LL�T�C�Y", "M�T�C�Y")
    
    Call ArraySortCustomOrder(Array1, ArrayStr("ss*", "s*", "m*", "l*", "ll*", "xl*"))
    
    Call Check(Array1(0), "SS�T�C�Y")
    Call Check(Array1(1), "S�T�C�Y")
    Call Check(Array1(2), "M�T�C�Y")
    Call Check(Array1(3), "LL�T�C�Y")
    Call Check(Array1(4), "XL�T�C�Y")
    Call Check(Array1(5), "B�T�C�Y")
    Call Check(Array1(6), "A�T�C�Y")
    
    Array1 = ArrayStr("B�T�C�Y", "A�T�C�Y", _
        "S�T�C�Y", "SS�T�C�Y", "XL�T�C�Y", "LL�T�C�Y", "M�T�C�Y")
    
    Call ArraySortCustomOrder(Array1, ArrayStr("ss*", "s*", "m*", "l*", "ll*", "xl*"), , True)
    
    Call Check(Array1(0), "B�T�C�Y")
    Call Check(Array1(1), "A�T�C�Y")
    Call Check(Array1(2), "SS�T�C�Y")
    Call Check(Array1(3), "S�T�C�Y")
    Call Check(Array1(4), "M�T�C�Y")
    Call Check(Array1(5), "LL�T�C�Y")
    Call Check(Array1(6), "XL�T�C�Y")
    
End Sub


'----------------------------------------
'�E�z����t���ɂ���
'----------------------------------------
'   �E  IndexMin/IndexMax���w�肷���
'       �w��͈͓��̒l���t���ɂ���
'----------------------------------------
Public Sub ArrayReverse(ByRef ArrayValue As Variant, _
Optional ByVal IndexMin As Long = -1, Optional ByVal IndexMax As Long = -1)

    Call Assert(IsArray(ArrayValue), "Error:ArrayValue is not Array")
    Call Assert(ArrayDimension(ArrayValue) = 1)
    
    'IndexMin/Max�̎w�肪�ςȂ�G���[�ɂ���
    Call Assert(IndexMin <= IndexMax, "Error:IndexMin < IndexMax")
    Call Assert(InRange(-1, IndexMin, ArrayCount(ArrayValue) - 1), _
        "Error:ArrayReverse:IndexMin Range is miss.")
    Call Assert(InRange(-1, IndexMax, ArrayCount(ArrayValue) - 1), _
        "Error:ArrayReverse:IndexMax Range is miss.")
    
    '1�ȉ��Ȃ�\�[�g�s�\�Ȃ̂�Exit����
    If ArrayCount(ArrayValue) <= 1 Then Exit Sub
    
    IndexMin = IIf(IndexMin = -1, 0, IndexMin)
    IndexMax = IIf(IndexMax = -1, ArrayCount(ArrayValue) - 1, IndexMax)
    
    'IndexMin=IndexMax�Ȃ�\�[�g�s�\�Ȃ̂�Exit
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
'��2�����z��
'----------------------------------------

'----------------------------------------
'�E���������擾����
'----------------------------------------
'   �E  �v�f���Ȃ��z��̏ꍇ�͎�������0�Ƃ��ĕԂ����
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
'�E2�����z��̗�(�ύX�ł��Ȃ�)���Z�b�g����
'----------------------------------------
'   �E  ������Ԃ���̃Z�b�g�ɂȂ�̂�
'       ���łɃZ�b�g���ꂽ�z��ɑ΂��Ď��s����ƃG���[�ɂȂ�
'   �E  �s�v�f�͍Œ�1�͕K�v�ɂȂ�
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
    '2����s����ƃG���[�ɂȂ�
    
End Sub


'----------------------------------------
'�E2�����z��̍s��ݒ肷��
'----------------------------------------
'   �E  �񐔂���v�����z���ݒ肵�čs�̒l���Z�b�g����
'   �E  �I�u�W�F�N�g�l�ɂ��Ή�
'----------------------------------------
Public Sub Array2dSetRowValues(ByRef ArrayValue As Variant, _
ByVal RowIndex As Long, _
ByRef Values As Variant)
    Call Assert(IsArray(ArrayValue), "Error:Array2dSetRowValues:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dSetRowValues:ArrayValue is not Array2D.")
    Call Assert(UBound(Values) - LBound(Values) + 1 = ArrayCount(ArrayValue, 1), _
        "Error:Array2dSetRowValues:Values Count is miss.")
    Call Assert(InRange(LBound(ArrayValue, 2), RowIndex, UBound(ArrayValue, 2)), _
        "Error:Array2dSetRowValues:RowIndex range over.")
        
    Dim I As Long
    For I = 0 To ArrayCount(ArrayValue, 1) - 1
        Call SetValue(ArrayValue(I, RowIndex), Values(I))
    Next
End Sub

'----------------------------------------
'�E2�����z��̍s���擾����
'----------------------------------------
'   �E  �I�u�W�F�N�g�l�ɂ��Ή�
'----------------------------------------
Public Function Array2dGetRowValues(ByRef ArrayValue As Variant, _
ByVal RowIndex As Long) As Variant
    Call Assert(IsArray(ArrayValue), "Error:Array2dSetRowValues:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dSetRowValues:ArrayValue is not Array2D.")
    Call Assert(InRange(LBound(ArrayValue, 2), RowIndex, UBound(ArrayValue, 2)), _
        "Error:Array2dSetRowValues:RowIndex range over.")
        
    Dim Result As Variant
    Result = Array()
    ReDim Preserve Result(ArrayCount(ArrayValue, 1) - 1)
    Dim I As Long
    For I = 0 To ArrayCount(ArrayValue, 1) - 1
        Result(I) = ArrayValue(I, RowIndex)
    Next
    Array2dGetRowValues = Result
End Function

'----------------------------------------
'�E�z��̗v�f��ǉ�����
'----------------------------------------
'   �E  �񐔂���v�����z���ݒ肵�čs�̒l��ǉ�����
'   �E  �I�u�W�F�N�g�l�ɂ��Ή�
'----------------------------------------
Public Sub Array2dAdd(ByRef ArrayValue As Variant, _
ByRef Values As Variant)
    Call Assert(IsArray(ArrayValue), "Error:Array2dAdd:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dAdd:ArrayValue is not Array2D.")
    Call Assert(UBound(Values) - LBound(Values) + 1 = ArrayCount(ArrayValue, 1), _
        "Error:Array2dAdd:Values Count is miss.")

    ReDim Preserve ArrayValue(ArrayCount(ArrayValue, 1) - 1, ArrayCount(ArrayValue, 2))
    Call Array2dSetRowValues(ArrayValue, UBound(ArrayValue, 2), Values)
End Sub


'----------------------------------------
'�E�z��̗v�f��}������
'----------------------------------------
'   �E  �I�u�W�F�N�g�l�ɂ��Ή�
'----------------------------------------
Public Sub Array2dInsert(ByRef ArrayValue As Variant, _
ByVal RowIndex As Long, ByVal Values As Variant)
    Call Assert(IsArray(ArrayValue), "Error:Array2dInsert:ArrayValue is not Array.")
    Call Assert(ArrayDimension(ArrayValue) = 2, "Error:Array2dInsert:ArrayValue is not Array2D.")
    Call Assert(UBound(Values) - LBound(Values) + 1 = ArrayCount(ArrayValue, 1), _
        "Error:Array2dInsert:Values Count is miss.")
    Call Assert(InRange(LBound(ArrayValue, 2), RowIndex, UBound(ArrayValue, 2)), _
        "Error:Array2dInsert:RowIndex range over.")

    ReDim Preserve ArrayValue(ArrayCount(ArrayValue, 1) - 1, ArrayCount(ArrayValue, 2))
    Dim I As Long
    For I = UBound(ArrayValue, 2) To RowIndex + 1 Step -1
        Call Array2dSetRowValues(ArrayValue, I, _
            Array2dGetRowValues(ArrayValue, I - 1))
    Next
    Call Array2dSetRowValues(ArrayValue, RowIndex, Values)
End Sub

'----------------------------------------
'�E�z��̗v�f���폜����
'----------------------------------------
'   �E  �I�u�W�F�N�g�l�ɂ��Ή�
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
        '�z��̏�������Erase���g��
    Else
        ReDim Preserve ArrayValue(ArrayCount(ArrayValue, 1) - 1, _
            LBound(ArrayValue, 2) To UBound(ArrayValue, 2) - 1)
    End If
End Sub

Public Sub testArray2dBasicFunction()
    Dim A()
    Call Check(0, ArrayCount(A, 1))
    Call Check(0, ArrayCount(A, 2))
    
    Call Array2dSetColumn(A, 3)
    
    Call Array2dSetRowValues(A, 0, Array("A", "B", "C"))
    Call Array2dAdd(A, Array("D", "E", "F"))
    Call Array2dAdd(A, Array("G", "H", "I"))
    Call Array2dAdd(A, Array("1", "2", "3"))

    Dim B()
    B = Array2dGetRowValues(A, 0)
    Call Check("A,B,C", ArrayToString(B, ","))
    Call Check("D,E,F", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("G,H,I", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("1,2,3", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check(3, ArrayCount(A, 1))
    Call Check(4, ArrayCount(A, 2))
    
    Call Array2dInsert(A, 3, B)
    Call Check("A,B,C", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check("D,E,F", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("G,H,I", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("A,B,C", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check("1,2,3", ArrayToString(Array2dGetRowValues(A, 4), ","))
    Call Check(3, ArrayCount(A, 1))
    Call Check(5, ArrayCount(A, 2))
    
    Call Array2dDelete(A, 0)
    Call Check("D,E,F", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check("G,H,I", ArrayToString(Array2dGetRowValues(A, 1), ","))
    Call Check("A,B,C", ArrayToString(Array2dGetRowValues(A, 2), ","))
    Call Check("1,2,3", ArrayToString(Array2dGetRowValues(A, 3), ","))
    Call Check(3, ArrayCount(A, 1))
    Call Check(4, ArrayCount(A, 2))

    Call Array2dDelete(A, 3)
    Call Array2dDelete(A, 1)
    Call Array2dDelete(A, 0)
    Call Check("A,B,C", ArrayToString(Array2dGetRowValues(A, 0), ","))
    Call Check(3, ArrayCount(A, 1))
    Call Check(1, ArrayCount(A, 2))

    Call Array2dDelete(A, 0)
    Call Check(0, ArrayCount(A, 1))
    Call Check(0, ArrayCount(A, 2))

End Sub



'----------------------------------------
'�E�z����̒l���������ă��j�[�N(����l���Ȃ�)���ǂ����𔻒f����
'----------------------------------------
Public Function Array2dIsUnique(ByRef ArrayValue As Variant, _
ByVal ColumnIndex As Long) As Boolean
    Call Assert(IsArray(ArrayValue), "Error:ArrayIsUnique:ArrayValue is not array")
    Call Assert(ArrayDimension(ArrayValue) = 2, _
        "Error:Array2dIsUnique:ArrayValue Dimension is miss")
    
    Dim Result As Boolean: Result = True
    Do
        If OrValue(ArrayCount(ArrayValue, 2), 0, 1) Then Exit Do

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
'�E�N�C�b�N�\�[�g
'----------------------------------------
'   �E  IndexMin/IndexMax���w�肷���
'       �w��͈͓��̒l���\�[�g����
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
    Call Assert(InRange(-1, RowIndexMin, ArrayCount(ArrayValue, 2) - 1), _
        "Error:ArrayReverse:RowIndexMin Range is miss.")
    Call Assert(InRange(-1, RowIndexMax, ArrayCount(ArrayValue, 2) - 1), _
        "Error:ArrayReverse:RowIndexMax Range is miss.")
    
    '1�ȉ��Ȃ�\�[�g�s�\�Ȃ̂�Exit����
    If ArrayCount(ArrayValue, 2) <= 1 Then Exit Sub
    
    RowIndexMin = IIf(RowIndexMin = -1, 0, RowIndexMin)
    RowIndexMax = IIf(RowIndexMax = -1, ArrayCount(ArrayValue, 2) - 1, RowIndexMax)
    
    'IndexMin=IndexMax�Ȃ�\�[�g�s�\�Ȃ̂�Exit
    If RowIndexMin = RowIndexMax Then Exit Sub
    
    Call Array2dSortQuickBase(ArrayValue, ColumnIndex, SortOrder, RowIndexMin, RowIndexMax)
End Sub

'�N�C�b�N�\�[�g�̃x�[�X�֐��A�ċN�Ăяo�������
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
    
    RowIndexCenter = (RowIndexMin + RowIndexMax) \ 2
    
    '�����l���o�b�t�@
    RowValue1 = Array2dGetRowValues(ArrayValue, RowIndexCenter)
    '�����l�ɊJ�n�ʒu�v�f����
    Call Array2dSetRowValues(ArrayValue, RowIndexCenter, _
        Array2dGetRowValues(ArrayValue, RowIndexMin))
    
    RowIndex2 = RowIndexMin
    
    RowIndex1 = RowIndexMin + 1
    
    Select Case SortOrder
    Case Ascending
        Do While RowIndex1 <= RowIndexMax
            If ArrayValue(ColumnIndex, RowIndex1) < RowValue1(ColumnIndex) Then
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
            If RowValue1(ColumnIndex) < ArrayValue(ColumnIndex, RowIndex1) Then
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
    
    ' �����O�����ċA�Ăяo����SORT
    Call Array2dSortQuickBase(ArrayValue, ColumnIndex, SortOrder, RowIndexMin, RowIndex2 - 1)

    ' �����㔼���ċA�Ăяo����SORT
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

    '�N�C�b�N�\�[�g�̂��߂̃L�[���ڍ쐬
    Dim I As Long
    For I = 0 To ArrayCount(Array1, 2) - 1
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
'�N�C�b�N�\�[�g�ł̓L�[���ڂ��Ȃ���
'�\�[�g�����ꂢ�ɍs���Ȃ��B

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
    '�L�[���ڂɑ΂��ă\�[�g����Ƃ��ꂢ�Ȍ��ʂɂȂ�
End Sub

'----------------------------------------
'���t�@�C��������
'----------------------------------------

'----------------------------------------
'�E�I�[�Ƀp�X��؂��ǉ�����֐�
'----------------------------------------
Public Function IncludeLastPathDelim(ByVal Path As String) As String
    IncludeLastPathDelim = IncludeLastStr(Path, Application.PathSeparator)
End Function

'----------------------------------------
'�E�I�[����p�X��؂���폜����֐�
'----------------------------------------
Public Function ExcludeLastPathDelim(ByVal Path As String) As String
    ExcludeLastPathDelim = ExcludeLastStr(Path, Application.PathSeparator)
End Function

'----------------------------------------
'�E�h���C�u�p�X"C:"�����o���֐�
'----------------------------------------
Public Function GetDrivePath(ByVal Path As String) As String
    GetDrivePath = IncludeLastStr( _
        FirstStrFirstDelim(Path, ":"), ":")
End Function

'----------------------------------------
'�E�h���C�u�p�X���܂܂�Ă��邩�ǂ����m�F����֐�
'[:]��2�����ڈȍ~�ɂ��邩�ǂ����Ŕ���
'----------------------------------------
Public Function IsDrivePath(ByVal Path As String) As String
    Dim Result As String
    Result = (OrValue(InStr(Path, ":"), 2, 3))
    IsDrivePath = Result
End Function
'
'----------------------------------------
'�E�l�b�g���[�N�h���C�u���ǂ����m�F����֐�
'----------------------------------------
Public Function IsNetworkPath(ByVal Path As String) As String
    Dim Result As String: Result = False
    If IsFirstStr(Path, "\\") Then
        If 3 <= Len(Path) Then
            Result = True
        End If
    End If
    IsNetworkPath = Result
End Function


'----------------------------------------
'�E�󔒂��܂ރt�@�C���p�X���_�u���N�E�H�[�g�ň͂�
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
'�E�g���q�̎擾
'----------------------------------------
'   �E  fso.GetExtensionName�ł͎擾�ł��Ȃ�
'       �Ōオ�s���I�h�ŏI���t�@�C���ł�
'       �l���擾���邱�Ƃ��ł���
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
'�E�g���q�̕ύX
'----------------------------------------
'   �E  NewExt�ɂ͐擪�s���I�h�������Ă��Ȃ��Ă��悢
'----------------------------------------
Public Function ChangeFileExtension(ByVal Path As String, _
ByVal NewExt As String) As String
    Dim Result As String: Result = ""
    Result = _
        IncludeLastStr( _
            ExcludeLastStr( _
                Path, GetExtensionIncludePeriod(Path)), _
            IncludeFirstStr(NewExt, "."))
    ChangeFileExtension = Result
End Function

Private Sub testChangeFileExtension()
    Call Check("C:\temp\text.csv", _
        ChangeFileExtension("C:\temp\text.txt", ".csv"))
    Call Check("C:\temp\text.csv", _
        ChangeFileExtension("C:\temp\text", "csv"))
    Call Check("C:\temp\text.csv", _
        ChangeFileExtension("C:\temp\text.", ".csv"))
End Sub

'----------------------------------------
'�E�p�X�̌���
'----------------------------------------
Public Function PathCombine(ParamArray Values()) As String
    '�p�����[�^�z��𑼂̃p�����[�^�z��ɓn�����͂ł��Ȃ��̂�
    '�p�����[�^�z���String�z��ɑ�����Ă���
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
'���t�@�C���t�H���_�p�X�擾
'----------------------------------------

'----------------------------------------
'�E����t�H���_��
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
'���t�@�C������
'----------------------------------------

'------------------------------
'�E�t�@�C�����݊m�F
'------------------------------
'   �E Win/Mac���Ή���
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
'�E���΃p�X�����΃p�X�擾
'----------------------------------------
Public Function AbsolutePath(ByVal BasePath As String, _
ByVal RelativePath As String) As String
    Dim CurDirBuffer As String
    CurDirBuffer = CurDir

    Call Assert(fso.FolderExists(BasePath) Or fso.FileExists(BasePath), _
        "Error:AbsolutePath")
    Call Assert(IsDrivePath(BasePath), "Error:AbsolutePath")

    '�J�����g�h���C�u/�f�B���N�g����BasePath�ɍ��킹��
    Call ChDrive(ExcludeLastStr(BasePath, ":\"))
    Call ChDir(BasePath)

    '���΃p�XRelativePath�ŃJ�����g�f�B���N�g����ݒ肷��

    AbsolutePath = fso.GetAbsolutePathName(RelativePath)

    '�o�b�t�@���Ă����l�ŃJ�����g�h���C�u/�f�B���N�g����ݒ肷��
    Call ChDrive(ExcludeLastStr(CurDirBuffer, ":\"))
    Call ChDir(CurDirBuffer)
End Function

Private Sub testAbsolutePath()
    Call Check("C:\Program Files", AbsolutePath("C:\", "..\Program Files"))
End Sub

'----------------------------------------
'�v���O�����̐ݒ�ȂǂŃp�X���擾����֐�
'----------------------------------------
'   �E  ���΃A�h���X�ȂǂɑΉ�
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
            '�t�@�C���_�C�A���O���J������
            '�J�����g�f�B���N�g�����ςɂȂ�ꍇ������̂�
            '�J�����g�f�B���N�g�������Z�b�g����
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
'�E�t�@�C�����쐬�����̂����΂炭�҂֐�
'----------------------------------------
'   �E  �쐬���ꂽ��True��Ԃ�
'----------------------------------------
Public Function FileCreateWait(ByVal FilePath As String) As Boolean
    FileCreateWait = False
    Dim I As Long: I = 0
    Do While (fso.FileExists(FilePath) = False)
        I = I + 1
        If I = 10 Then Exit Function
    Loop
    FileCreateWait = True
End Function


'----------------------------------------
'�E�t�@�C���R�s�[�㏑�����s�����m���邽�߂̊֐�
'----------------------------------------
'   �E  Success:=True / Fail:=False
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
'��Force/Recrate
'----------------------------------------

'----------------------------------------
'�E�[���K�w�̃t�H���_�ł���C�ɍ쐬����֐�
'----------------------------------------
Public Sub ForceCreateFolder(ByVal FolderPath As String)
    Dim ParentFolderPath As String
    ParentFolderPath = fso.GetParentFolderName(FolderPath)
    If fso.FolderExists(ParentFolderPath) = False Then
        Call ForceCreateFolder(ParentFolderPath)
    Else
        If fso.FolderExists(FolderPath) = False Then
            Call fso.CreateFolder(FolderPath)
        End If
    End If
End Sub

'----------------------------------------
'�E�t�H���_���Đ�������֐�
'----------------------------------------
Public Sub ReCreateFolder( _
ByVal FolderPath As String)

    If fso.FolderExists(FolderPath) Then
        Call fso.DeleteFolder(FolderPath)
    End If

    '�t�H���_��������܂Ń��[�v
    Do: Loop While fso.FolderExists(FolderPath)

    On Error Resume Next
    Do
        Call ForceCreateFolder(FolderPath)
    Loop Until fso.FolderExists(FolderPath)
    '�t�H���_���쐬�ł���܂Ń��[�v
End Sub



'----------------------------------------
'���t�@�C���t�H���_��
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
'���t�H���_
'----------------------------------------

'----------------------------------------
'���g�b�v���x���̃t�H���_���X�g���擾
'----------------------------------------
'�E ���݂��Ȃ���΋󕶎���Ԃ��B
'�E �p�X�͉��s�R�[�h�ŋ�؂��Ă���
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
'���T�u�t�H���_�̃t�H���_���X�g���擾
'----------------------------------------
'�E ���݂��Ȃ���΋󕶎���Ԃ��B
'�E �p�X�͉��s�R�[�h�ŋ�؂��Ă���
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
'���t�@�C��
'----------------------------------------

'----------------------------------------
'���g�b�v���x���̃t�@�C�����X�g���擾
'----------------------------------------
'�E ���݂��Ȃ���΋󕶎���Ԃ��B
'�E �p�X�͉��s�R�[�h�ŋ�؂��Ă���
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
'���T�u�t�H���_�̃t�@�C�����X�g���擾
'���݂��Ȃ���΋󕶎���Ԃ��B
'�p�X�̍Ō�ɂ͕K�����s�R�[�h���t��
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
'���t�@�C������
'----------------------------------------

'----------------------------------------
'�EUTC�t�@�C���^�C���ϊ��֐�
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
'�E�t�@�C��/�t�H���_�̍쐬����/�X�V����/�ŏI�A�N�Z�X�����̎擾
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
'�E�t�@�C��/�t�H���_�̍쐬����/�X�V����/�ŏI�A�N�Z�X�����̐ݒ�
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
        '// �Ώۂ̑��݃`�F�b�N��dwFlagsAndAttributes �̐ݒ�
        If fso.FileExists(Path) Then
            '�t�@�C���̏ꍇ
            CreateFileFlag = FILE_ATTRIBUTE_NORMAL
        ElseIf fso.FolderExists(Path) Then
            '�t�H���_�̏ꍇ(NT�n��OS�̂݉\)
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

        '// �I�v�V�����������ȗ����ꂽ�ꍇ�͌���̂��̂�⊮
        If FileFolderTime.CreataionTime <> 0 Then
            SetTime.CreataionTime = FileFolderTime.CreataionTime
        End If
        If FileFolderTime.LastWriteTime <> 0 Then
            SetTime.LastWriteTime = FileFolderTime.LastWriteTime
        End If
        If FileFolderTime.LastAccessTime <> 0 Then
            SetTime.LastAccessTime = FileFolderTime.LastAccessTime
        End If

        '// SECURITY_ATTRIBUTES�\���̏�����

        SecurityAttr.nLength = LenB(SecurityAttr)
        SecurityAttr.lpSecurityDescriptor = 0&
        SecurityAttr.bInheritHandle = 0&


        '// �t�@�C���܂��̓t�H���_�n���h�����擾
        FileHandle = CreateFile(Path, GENERIC_WRITE, _
            FILE_SHARE_READ, SecurityAttr, OPEN_EXISTING, CreateFileFlag, vbNull)
        If FileHandle = INVALID_HANDLE_VALUE Then Exit Do

        '// �t�@�C���^�C���ɕϊ����A�ݒ肷��
        CreateFILETIME = DateToApiFILETIME(SetTime.CreataionTime)
        AccessFILETIME = DateToApiFILETIME(SetTime.LastAccessTime)
        ModifyFILETIME = DateToApiFILETIME(SetTime.LastWriteTime)
        ReturnSetFileTime = SetFileTime(FileHandle, CreateFILETIME, AccessFILETIME, ModifyFILETIME)
        If ReturnSetFileTime <> 0 Then
            Result = True
        End If

        '// �t�@�C���܂��̓t�H���_�n���h���J��
        Call CloseHandle(FileHandle)
    Loop While False

    SetFileFolderTime = Result
End Function


'----------------------------------------
'���V���[�g�J�b�g�t�@�C������
'----------------------------------------

'----------------------------------------
'�E�V���[�g�J�b�g�t�@�C������(�g���q)
'----------------------------------------

Public Function IsShortcutLinkFile(ByVal FilePath As String)
    Dim Result As Boolean: Result = False
    If LCase(GetExtensionIncludePeriod(FilePath)) = ".lnk" Then
        Result = True
    End If
    IsShortcutLinkFile = Result
End Function

'----------------------------------------
'�E�V���[�g�J�b�g�t�@�C���̍쐬
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
'�E�V���[�g�J�b�g�t�@�C���̍쐬/�폜
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
        '���t���OON�Ȃ��t�H���_�ɂȂ����ꍇ�̓t�H���_�폜����
        If FolderDeleteFlag _
        And fso.GetFolder(ShortcutFileParentFolderPath).SubFolders.Count = 0 Then
            Call fso.DeleteFolder(ShortcutFileParentFolderPath)
        End If
    End If
End Sub


'----------------------------------------
'��Ini�t�@�C������
'----------------------------------------
Public Function IniFile_GetString(ByVal Path As String, _
ByVal Section As String, ByVal Name As String, _
Optional ByVal DefaultValue As String = "") As String
    Dim Result As String

    ' �l���擾����o�b�t�@���m�ۂ���
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
'���e�L�X�g�t�@�C���ǂݏ���
'----------------------------------------

Public Function CheckEncodeName(EncodeName As String) As Boolean
    CheckEncodeName = OrValue(UCase$(EncodeName), _
        "SHIFT_JIS", _
        "UNICODE", "UNICODEFFFE", "UTF-16LE", "UTF-16", _
        "UNICODEFEFF", _
        "UTF-16BE", _
        "UTF-8", _
        "ISO-2022-JP", _
        "EUC-JP", _
        "UTF-7")
End Function

'----------------------------------------
'�E�e�L�X�g�t�@�C���Ǎ�
'----------------------------------------
'   �E  �G���R�[�h�w��͉��L�̒ʂ�
'           �G���R�[�h          �w�蕶��
'           ShiftJIS            SHIFT_JIS
'           UTF-16LE BOM�L/��   UNICODEFFFE/UNICODE/UTF-16/UTF-16LE
'                           BOM�̗L���Ɋւ�炸�Ǎ��\
'           UTF-16BE _BOM_ON    UNICODEFEFF
'           UTF-16BE _BOM_OFF   UTF-16BE
'           UTF-8 BOM�L/��      UTF-8/UTF-8N
'                           BOM�̗L���Ɋւ�炸�Ǎ��\
'           JIS                 ISO-2022-JP
'           EUC-JP              EUC-JP
'           UTF-7               UTF-7
'   �E  UTF-16LE��UTF-8�́ABOM�̗L���ɂ�����炸�ǂݍ��߂�
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

Private Sub testADOStream_LoadTextFile()
    MsgBox ADOStream_LoadTextFile( _
        ThisWorkbook.Path + "\test.ini", _
        "UTF-16LE")
End Sub

'----------------------------------------
'�E�e�L�X�g�t�@�C���ۑ�
'----------------------------------------
'   �E  �G���R�[�h�w��͉��L�̒ʂ�
'           �G���R�[�h          �w�蕶��
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
'   �E  UTF-16LE��UTF-8�͂��̂܂܂���_BOM_ON�ɂȂ�̂�
'       BON�����w��̏ꍇ�͓��ꏈ�������Ă���
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

Private Sub testADOStream_SaveTextFile()
    Call ADOStream_SaveTextFile( _
        "[Option]" + vbCrLf + "Name = TestValue02", _
        ThisWorkbook.Path + "\test.ini", _
        "UTF-16LE", False)
End Sub

'----------------------------------------
'���摜�t�@�C��
'----------------------------------------


'----------------------------------------
'�EJpeg�t�@�C������(�g���q)
'----------------------------------------
Public Function IsJpegImageFile(ByVal FilePath As String)
    Dim Result As Boolean: Result = False
    If OrValue(LCase(GetExtensionIncludePeriod(FilePath)), ".jpg", ".jpeg") Then
        Result = True
    End If
    IsJpegImageFile = Result
End Function

'----------------------------------------
'�EJpegExif�܂ރt�@�C������
'----------------------------------------
'   �E  Exif�̎B�e�����擾�\���ǂ����𔻒�Ȃ̂�
'       �t�@�C�������ۂɑ��݂��邱�Ƃ��m�F�����
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
'�EJpegExif���B�e�����擾
'----------------------------------------
'   �E  �擾�ł��Ȃ��ꍇ��CDate(0)��Ԃ�
'----------------------------------------
Public Function GetJpegExifDateTime(ByVal FilePath As String) As Date
On Error GoTo Err:
    Dim Result As Date: Result = 0
    If IsJpegImageFile(FilePath) Then

        Dim WIA_ImageFile As Object
        Set WIA_ImageFile = CreateObject("Wia.ImageFile")
        Call WIA_ImageFile.LoadFile(FilePath)

        
        '�B�e����
        Dim ExifDateTime As String
        ExifDateTime = WIA_ImageFile.Properties("36867")
        ExifDateTime = Replace(ExifDateTime, ":", "/", , 2)
        Result = CDate(ExifDateTime)
    End If
Err:
    GetJpegExifDateTime = Result
End Function


'----------------------------------------
'���V�F���N��
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

Public Function CommandExecuteReturn(Command As String, _
Optional ByVal EncodeName As String = "Shift_JIS") As String
    Dim Result As String: Result = ""

    '�e���|�����t�@�C���p�X���擾
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
'���N���b�v�{�[�h
'----------------------------------------
'   �E  �Q�Ɛݒ�[Microsoft Forms 2.0 Object Library]��
'       DataObject���g�p�\
'       Mac�ł��\
'----------------------------------------

'----------------------------------------
'�E�e�L�X�g�f�[�^�擾
'----------------------------------------
'   �E  Win/Mac���Ή�����m�F�@
'----------------------------------------
Public Function GetClipboardText()
    Dim DataObject1 As New MSForms.DataObject

    DataObject1.GetFromClipboard
    GetClipboardText = DataObject1.GetText
End Function

'----------------------------------------
'�E�e�L�X�g�f�[�^�ݒ�
'----------------------------------------
'   �E  Win/Mac���Ή�����m�F�@
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
'��Excel
'----------------------------------------

'----------------------------------------
'�E�i���\��
'----------------------------------------
Public Sub Application_StatusBar_Progress(ByVal Message As String, _
ByVal StartValue As Long, ByVal Value As Long, ByVal EndValue As Long, _
Optional ReverseFlag As Boolean = False)

    If ReverseFlag = False Then
        Application.StatusBar = _
            Message + ":" + _
            CStr(Value - StartValue + 1) + "/" + _
            CStr(EndValue - StartValue + 1) + ":" + _
            CStr(Format((Value - StartValue + 1) / (EndValue - StartValue + 1) * 100, "0.00")) + "%"
    Else
        Application.StatusBar = _
            Message + ":" + _
            CStr(Value - StartValue + 1) + "/" + _
            CStr(EndValue - StartValue + 1) + ":" + _
            CStr(Format(100 - ((Value - StartValue + 1) / (EndValue - StartValue + 1) * 100), "0.00")) + "%"
    End If

End Sub

'----------------------------------------
'�E��ԍ�����񖼂��擾����
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
'�E��(A,B,C,�c)�����ԍ����擾����
'----------------------------------------
'   �E  A��1, B��2, �c, Z��26, AA��27, AB��28
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
'���^�C�g���s/��w�菈��
'----------------------------------------

'----------------------------------------
'�E�^�C�g���s�̗񖼂����ԍ���Ԃ��֐�
'----------------------------------------
'   �E  ���{��^�C�g���s�Ȃǂɑ΂��ă^�C�g��������ōs�ԍ���Ԃ�
'----------------------------------------
Public Function ColumnNumberByTitle(ByRef Sheet As Worksheet, _
ByVal TitleRowIndex As Long, _
ByVal ColumnTitle As String, _
Optional TitleMatchCount As Long = 1)
    Dim Result As Long: Result = 0
    Dim Counter As Long: Counter = 0
    Dim I As Long
    For I = Col__A To DataLastCol(Sheet, TitleRowIndex)
        If Sheet.Cells(TitleRowIndex, I).Value = ColumnTitle Then
            Counter = Counter + 1
            If Counter = TitleMatchCount Then
            Result = I
            Exit For
        End If
        End If
    Next
    ColumnNumberByTitle = Result
End Function


'----------------------------------------
'�E�^�C�g����̍s������s�ԍ���Ԃ��֐�
'----------------------------------------
'   �E  ���{��^�C�g���s�Ȃǂɑ΂��ă^�C�g��������ōs�ԍ���Ԃ�
'----------------------------------------
Public Function RowNumberByTitle(ByRef Sheet As Worksheet, _
ByVal TitleColIndex As Long, _
ByVal RowTitle As String, _
Optional TitleMatchCount As Long = 1)
    Dim Result As Long: Result = 0
    Dim Counter As Long: Counter = 0
    Dim I As Long
    For I = 1 To DataLastRow(Sheet, TitleColIndex)
        If Sheet.Cells(I, TitleColIndex).Value = RowTitle Then
            Counter = Counter + 1
            If Counter = TitleMatchCount Then
            Result = I
            Exit For
        End If
        End If
    Next
    RowNumberByTitle = Result
End Function

'----------------------------------------
'���ŏI�s/��
'----------------------------------------
'----------------------------------------
'�E�ŏI�s/��
'----------------------------------------
'   �E  �f�[�^���Ȃ��ꍇ��1��߂�
'----------------------------------------

'�E�f�[�^�ŏI�s
Public Function DataLastRow(ByVal Sheet As Worksheet, _
Optional ByVal ColumnNumber As Long = -1) As Long
On Error Resume Next
    DataLastRow = 1
    Call Assert(-1 <= ColumnNumber, "Error:DataLastRow")
    If ColumnNumber = -1 Then
        DataLastRow = Sheet.UsedRange.Find("*", _
            , xlFormulas, , xlByRows, xlPrevious).Row
    Else
        DataLastRow = Sheet.Cells(Sheet.Rows.Count, ColumnNumber).End(xlUp).Row
    End If
End Function

'�E�f�[�^�ŏI��
Public Function DataLastCol(ByVal Sheet As Worksheet, _
Optional ByVal RowNumber As Long = -1) As Long
On Error Resume Next
    DataLastCol = 1
    Call Assert(-1 <= RowNumber, "Error:DataLastCol")
    If RowNumber = -1 Then
        DataLastCol = Sheet.UsedRange.Find("*", _
            , xlFormulas, , xlByColumns, xlPrevious).Column
    Else
        DataLastCol = Sheet.Cells(RowNumber, Sheet.Columns.Count).End(xlToLeft).Column
    End If
End Function

Public Function DataLastCell(ByVal Sheet As Worksheet) As Range
    Set DataLastCell = Sheet.Cells( _
        DataLastRow(Sheet), DataLastCol(Sheet))
End Function

'----------------------------------------
'���ŏI�s/��폜
'----------------------------------------
'   �E  RangeClearType��
'       Clear/ClearContents/ClearFormats
'----------------------------------------
Public Sub RangeClear(ByRef Range As Range, _
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

Public Sub ClearRangeLastData(ByVal Sheet As Worksheet, _
ByVal RowIndex As Long, ByVal ColumnIndex As Long, _
Optional ByVal RangeClearType As RangeClearType = rcClear, _
Optional ByVal MergeCellOption As Boolean = False)
    If (RowIndex <= DataLastRow(Sheet)) _
    And (ColumnIndex <= DataLastCol(Sheet)) Then
        Call RangeClear( _
            Sheet.Range( _
                Sheet.Cells(RowIndex, ColumnIndex), _
                Sheet.Cells(DataLastRow(Sheet), DataLastCol(Sheet))), _
            RangeClearType, MergeCellOption)
    End If
End Sub

'�E��̃N���A�A�ŏI�s�܂�
Public Sub ClearColumnLastRow(ByVal Sheet As Worksheet, _
ByVal RowIndex As Long, ByVal ColumnIndex As Long, _
Optional ByVal RangeClearType As RangeClearType = rcClear, _
Optional ByVal MergeCellOption As Boolean = False)
    Dim LastRow As Long: LastRow = DataLastRow(Sheet, ColumnIndex)
    If (RowIndex <= LastRow) Then
        Call RangeClear( _
            Sheet.Range( _
                Sheet.Cells(RowIndex, ColumnIndex), _
                Sheet.Cells(DataLastRow(Sheet, ColumnIndex), ColumnIndex)), _
            RangeClearType, MergeCellOption)
    End If
End Sub

'�E�s�̃N���A�A�ŏI��܂�
Public Sub ClearRowLastColumn(ByVal Sheet As Worksheet, _
ByVal RowIndex As Long, ByVal ColumnIndex As Long, _
Optional ByVal RangeClearType As RangeClearType = rcClear, _
Optional ByVal MergeCellOption As Boolean = False)
    Dim LastCol As Long: LastCol = DataLastCol(Sheet, RowIndex)
    If (ColumnIndex <= LastCol) Then
        Call RangeClear( _
            Sheet.Range( _
                Sheet.Cells(RowIndex, ColumnIndex), _
                Sheet.Cells(RowIndex, DataLastCol(Sheet, RowIndex))), _
            RangeClearType, MergeCellOption)
    End If
End Sub

'----------------------------------------
'������
'----------------------------------------

'----------------------------------------
'�E�������폜����֐�
'----------------------------------------
Public Sub RangeDeleteFormula(ByRef Sheet As Worksheet, ByRef Range As Range)

    '�����ɉe�����o�Ȃ��悤�Ɏw��͈͂̌������l���w�肵�Ă���
    '=SUBTOTAL(9, �c
    '�Ƃ��̐����́A�����𖳎����Ēl�ɑ΂��č��Z����Ƃ������̂Ȃ̂�
    '�㕔�̐��������l�ɂȂ����ꍇ�ɒl���ω����Ă��܂�
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
'��Sheet.Range�̃R�s�[����
'----------------------------------------

'----------------------------------------
'�E���l�����̃R�s�[
'----------------------------------------
'   �E  Excel�̏����̃R�s�[���o�O���Ă���̂ŏC���̂��߂ɍ쐬
'   �E  Excel�̃R�s�[�ł�
'       �y#,##0_);[��](#,##0)�z���y#,##0_);[��]-#,##0�z��
'       �Ȃ��Ă��܂��ꍇ������B
'       �t�@�C�����j�����Ă���̂�������Ȃ��������ł��Ȃ������̂�
'       ���̊֐����쐬
'----------------------------------------
Public Sub RangeCopyNumberFormat( _
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
'�E�l�ȂǑS�ẴR�s�[
'----------------------------------------
Public Sub RangeCopyAll( _
ByRef RangeSource As Range, _
ByRef RangeDest As Range)
    RangeSource.Copy
    Call RangeDest.PasteSpecial(Paste:=xlPasteAll)
    Call RangeCopyNumberFormat(RangeSource, RangeDest)
End Sub

'----------------------------------------
'�E�����̃R�s�[
'----------------------------------------
Public Sub RangeCopyFormat( _
ByRef RangeSource As Range, _
ByRef RangeDest As Range)
    RangeSource.Copy
    Call RangeDest.PasteSpecial(Paste:=xlPasteFormats)
    Call RangeCopyNumberFormat(RangeSource, RangeDest)
End Sub

'----------------------------------------
'�E�l�̃R�s�[
'----------------------------------------
Public Sub RangeCopyValue( _
ByRef RangeSource As Range, _
ByRef RangeDest As Range)
    RangeSource.Copy
    Call RangeDest.PasteSpecial(Paste:=xlPasteAllUsingSourceTheme)
    Call RangeCopyNumberFormat(RangeSource, RangeDest)
End Sub



'----------------------------------------
'���͈�
'----------------------------------------

'----------------------------------------
'�E�͈͂̏��1�s
'----------------------------------------
Public Function RangeUpRow(ByRef SourceRange As Range) As Range
    Set RangeUpRow = _
        SourceRange.Resize(1, SourceRange.Columns.Count).Offset(-1, 0)
End Function

'----------------------------------------
'�E�͈͂̉���1�s
'----------------------------------------
Public Function RangeDownRow(ByRef SourceRange As Range) As Range
    Set RangeDownRow = _
        SourceRange.Resize(1, SourceRange.Columns.Count).Offset( _
            SourceRange.Rows.Count, 0)
End Function

'----------------------------------------
'���͈͈ړ�
'----------------------------------------

'----------------------------------------
'�E�͈͂����1�A�ړ�����
'----------------------------------------
Public Sub RangeMoveUpRowOne(ByRef SourceRange As Range)
    '�����̑I��͈͂ɂ͔�Ή�
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
    
    '�I��͈͂̉�1�Z����������
    Call RangeDownRow(SourceRange).Insert(xlDown)
    
    '��̃Z�������̃Z���ɃR�s�[����
    Call RangeUpRow(SourceRange).Copy( _
        Destination:=RangeDownRow(SourceRange))
    
    '��̃Z����1�폜
    Call RangeUpRow(SourceRange).Delete(xlUp)
    
    '�I���ʒu��1��ɂ���
    If SelectionFlag Then
        SourceRange.Select
    End If
    
    Application.EnableEvents = EnableEventsBuffer
    
End Sub

Public Sub RangeMoveDownRowOne(ByRef SourceRange As Range)
    '�����̑I��͈͂ɂ͔�Ή�
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
        
    '�I��͈͂̏�1�Z����������
    Call SourceRange.Resize(1, Selection.Columns.Count).Insert(xlDown)
    
    '���̃Z������̃Z���ɃR�s�[����
    Call RangeDownRow(SourceRange).Copy( _
        Destination:=RangeUpRow(SourceRange))
    
    '���̃Z����1�폜
    Call RangeDownRow(SourceRange).Delete(xlUp)
    
    '�I���ʒu��1��ɂ���
    If SelectionFlag Then
        SourceRange.Select
    End If

    Application.EnableEvents = EnableEventsBuffer

End Sub


'----------------------------------------
'��Excel �I�u�W�F�N�g
'----------------------------------------

'----------------------------------------
'�����[�N�u�b�N
'----------------------------------------
'----------------------------------------
'�E���[�N�u�b�N�̑��݊m�F
'----------------------------------------
Public Function GetWorkbook( _
ByVal WorkbookNameWildCard As String, _
Optional ByVal WorkbookFolderPath As String = "", _
Optional ByVal App As Application = Nothing) As Workbook

    If App Is Nothing Then Set App = Application

    Dim Result As Workbook: Set Result = Nothing
    Dim Book As Workbook
    If WorkbookFolderPath = "" Then
        For Each Book In App.Workbooks
            If Book.Name Like WorkbookNameWildCard Then
                Set Result = Book
                Exit For
            End If
        Next
    Else
        For Each Book In App.Workbooks
            If (Book.Name Like WorkbookNameWildCard) _
            And (Book.Path = WorkbookFolderPath) Then
                Set Result = Book
                Exit For
            End If
        Next
    End If
    Set GetWorkbook = Result
End Function

Public Function WorkbookExists( _
ByVal WorkbookNameWildCard As String, _
Optional ByVal WorkbookFolderPath As String = "", _
Optional ByVal App As Application = Nothing) As Boolean

    Dim Result As Boolean: Result = False
    If (GetWorkbook(WorkbookNameWildCard) Is Nothing) = False Then
        Result = True
    End If

    WorkbookExists = Result
End Function

Public Sub testWorkbookExists()
    Call Check(True, WorkbookExists("st_vba.xlsm"))
    Call Check(True, WorkbookExists("st_vba*"))
    Call Check(False, WorkbookExists("st_vba.xls"))
End Sub

'----------------------------------------
'�����[�N�V�[�g
'----------------------------------------

'----------------------------------------
'�E���[�N�V�[�g�̑��݊m�F
'----------------------------------------

Public Function GetWorksheet(ByVal SheetNameWildCard As String, _
Optional ByVal Book As Workbook = Nothing) As Worksheet

    If Book Is Nothing Then
        Set Book = ThisWorkbook
    End If

    Dim Result As Worksheet: Set Result = Nothing
    Dim I As Long
    For I = 1 To Book.Sheets.Count
        If Book.Sheets(I).Name Like SheetNameWildCard Then
            Set Result = Book.Sheets(I)
        End If
    Next
    
    Set GetWorksheet = Result
End Function

Public Function WorksheetExists(ByVal SheetNameWildCard As String, _
Optional ByVal Book As Workbook = Nothing) As Boolean

    Dim Result As Boolean: Result = False
    If (GetWorksheet(SheetNameWildCard) Is Nothing) = False Then
        Result = True
    End If
        
    WorksheetExists = Result
End Function

Public Sub testWorksheetExists()
    Call Check(True, WorksheetExists("Sheet1"))
    Call Check(True, WorksheetExists("Sheet*"))
    Call Check(False, WorksheetExists("Sheet"))
End Sub

'----------------------------------------
'�E���[�N�V�[�g�̍폜
'----------------------------------------

Public Sub DeleteSheet(ByVal SheetNameWildCard As String, _
Optional MatchUnDelete As Boolean = False, _
Optional ByVal Book As Workbook = Nothing)

    If Book Is Nothing Then
        Set Book = ThisWorkbook
    End If

    Dim Sheet As Worksheet
    Dim I As Long
    For I = Book.Sheets.Count To 1 Step -1
        If MatchUnDelete Then
            If Not (Book.Sheets(I).Name Like SheetNameWildCard) Then
                Book.Sheets(I).Delete
            End If
        Else
            If (Book.Sheets(I).Name Like SheetNameWildCard) Then
            Book.Sheets(I).Delete
        End If
        End If
    Next
    
End Sub

Public Sub DeleteDefaultSheet()
    Call DeleteSheet("Sheet*")
End Sub

'----------------------------------------
'�E���[�N�V�[�g�ւ̃e�L�X�g�z�u
'----------------------------------------

Public Sub SetTextSheet(ByVal Sheet As Worksheet, _
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
'���`�F�b�N�{�b�N�X
'----------------------------------------
'   �E  �t�H���gWindings�̃`�F�b�N�{�b�N�X�\���̕������Ԃ�
'   �E  ChrW�̔��΂�AscW
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
'���I�u�W�F�N�g
'----------------------------------------

'----------------------------------------
'�EChartObject�̑��݊m�F
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
'�EOLEObject�̑��݊m�F
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
'��Shape
'----------------------------------------

'----------------------------------------
'�EShapes�̑��݊m�F
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
'�E�Z���͈͂ɓ��Ă͂܂�悤�ɉ摜�t�@�C����\��t���鏈��
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
    
    '�}�[�W�����Ƃ邽�߂ɒl��ݒ�
    Dim Rect As Rect
    Rect.Left = SheetRange.Left + Margin
    Rect.Top = SheetRange.Top + Margin
    Call SetRectWidth(Rect, SheetRange.Width - (Margin * 2))
    Call SetRectHeight(Rect, SheetRange.Height - (Margin * 2))
    
    Dim Shape As Shape
    Set Shape = Sheet.Shapes.AddPicture( _
        Filename:=ImageFilePath, LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=Rect.Left, _
        Top:=Rect.Top, _
        Width:=0, _
        Height:=0)
    
    '���摜�T�C�Y�ɖ߂�
    Call Shape.ScaleHeight(1#, True)
    Call Shape.ScaleWidth(1#, True)
    
    '�c�����ێ������܂܁A�����𒲐�����
    Shape.LockAspectRatio = True
    Shape.Height = GetRectHeight(Rect)
    
    '�摜���T�C�Y���͈͓��Ɏ��܂��Ă��邩�ǂ����m�F
    If Shape.Width > GetRectWidth(Rect) Then
        '���T�C�Y���͂ݏo�Ă���Ȃ牡�����킹��
        Shape.Width = GetRectWidth(Rect)
        
        '���E�ʒu�͂҂�����Ȃ̂ŏ㉺�ʒu����������
        Select Case VerticalAlign
        Case AlineVertical.alCenter
            Shape.Top = Shape.Top + (GetRectHeight(Rect) - Shape.Height) / 2
        Case AlineVertical.alBottom
            Shape.Top = Shape.Top + (GetRectHeight(Rect) - Shape.Height)
        End Select
    Else
        '�㉺�ʒu�͂҂�����Ȃ̂ō��E�ʒu����������
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
'�EShape�摜�����k����
'----------------------------------------
'   �E  �N���b�v�{�[�h���o�R������@���������炵��
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
    
'    Sheet.PasteSpecial Format:="�} (�g�����^�t�@�C��)", Link:=False, DisplayAsIcon:=False

    Sheet.PasteSpecial Format:="�} (JPEG)", Link:=False, DisplayAsIcon:=False
    If Err.Number <> 0 Then
        Err.Clear
        Sheet.PasteSpecial Format:="�} (JPEG)", Link:=False, DisplayAsIcon:=False
        If Err.Number <> 0 Then
            Err.Clear
            Sheet.PasteSpecial Format:="�} (JPEG)", Link:=False, DisplayAsIcon:=False
        End If
    End If
    
    Selection.ShapeRange.Width = RectSize.Width
    Selection.ShapeRange.Height = RectSize.Height
    Selection.Left = Point.X
    Selection.Top = Point.Y
End Sub

'----------------------------------------
'�E���W�ʒu�ɑ΂���Z���ʒu��Ԃ��֐�
'----------------------------------------
'   �E  Shape.TopLeftCell/.BottomRightCell �͂��邪
'       ���S�ʒu�̃Z�������߂���@�͂Ȃ������̂ō쐬�B
'       ���x�͑����Ȃ��B
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
'��Excel �A�v���P�[�V����
'----------------------------------------

'----------------------------------------
'�EExcel �E�B���h�E�^�C�g���o�[�\��
'----------------------------------------
Public Sub SetExcelWindowTitle( _
ByVal AppTitle As String, _
Optional ByVal ActTitle As String = "")

    Application.Caption = AppTitle
    ActiveWindow.Caption = ActTitle
    'Application.Caption = "" �̏ꍇ�AExcel �Ƃ��������������œ���
    'ActionWindow.Caption <> "" �̏ꍇ
    '  �E�B���h�E�^�C�g�� - �A�v���P�[�V�����^�C�g��
    '�Ƃ����悤�Ƀn�C�t���Őڑ������

    '�Ȃ̂ŒP�ɕ��������ꂽ���ꍇ��
    'Application.Caption�ɕ����ݒ肵��
    'ActiveWindow.Caption = "" �ɂ���Ƃ悢
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
'�����j���[����
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
'���O���t����
'----------------------------------------

'----------------------------------------
'�EGraphFormulaData���擾�Ɛݒ�
'----------------------------------------

'Chart.SeriesCollection.Item(I).Formula���\�b�h�œ����镶����̗�
'   =SERIES(,Sheet1!$A$2:$A$32,Sheet1!$B$2:$B$32,1)
'   =SERIES(�n��,X�����ڎ�,�f�[�^,�n��ԍ�)

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
'���O���t�P�ƃf�[�^�͈͕̔ύX
'----------------------------------------

'----------------------------------------
'�EGraphFormulaData�̏I�[����
'   Value�����A�I�[�L����
'   Value�����A�I�[���܂�
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
'�EGraphFormulaData�̐擪�𑀍삷��
'   Value�����A�擪�L����
'   Value�����A�擪���܂�
'----------------------------------------
Public Sub GraphSeriesFirstRangeUp(ByRef Data As GraphFormulaData, Value As Long)
    Call GraphSeriesMove(Data, -Value)
    Call GraphSeriesLastRangeUp(Data, Value)
End Sub

'----------------------------------------
'�EGraphFormulaData�͈̔͂��ړ�����
'   Value�����A����ړ�
'   Value�����A�O���ړ�
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
'�EGraphFormulaData�̃T�C�Y����
'   Value���s��
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
'�E�O���t�͈̔͂��擾����
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
'�EGraphFormulaData�̎w���̕ύX
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
'�E�O���t�̃V���[�Y�f�[�^�̎w���ύX
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
'���O���t�S�n��f�[�^�͈͕̔ύX
'----------------------------------------

'----------------------------------------
'�E�I�[����
'----------------------------------------
'   �E  Value�����A�I�[�L����
'       Value�����A�I�[���܂�
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
    Call MsgBox("�͈͎w�肪����������܂���")
End Sub

'----------------------------------------
'�E�擪����
'----------------------------------------
'   �E  Value�����A�擪�L����
'       Value�����A�擪���܂�
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
    Call MsgBox("�͈͎w�肪����������܂���")
End Sub

'----------------------------------------
'�E�͈͈ړ�
'----------------------------------------
'   �E  Value�����A����ړ�
'       Value�����A�O���ړ�
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
    Call MsgBox("�͈͎w�肪����������܂���")
End Sub

'----------------------------------------
'�E�T�C�Y����
'----------------------------------------
'   �E  Value���s��
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
    Call MsgBox("�͈͎w�肪����������܂���")
End Sub

'----------------------------------------
'��UserForm����
'----------------------------------------

'----------------------------------------
'�EWindowStyle
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
'�EWindowExStyle
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
'�ECloseButton
'----------------------------------------
Public Sub SetWindowCloseButton(hWnd As Long, _
ByVal Enabled As Boolean)

    Dim hMenu As Long
    Dim rc As Long

    If Enabled Then
        '���j���[�����Z�b�g
        hMenu = GetSystemMenu(hWnd, True)
    Else
        hMenu = GetSystemMenu(hWnd, False)
        rc = DeleteMenu(hMenu, 5, MF_BYPOSITION)
        rc = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    End If
    rc = DrawMenuBar(hWnd)

    'EnableMenuItemAPI���g���Đ��䂵�悤�Ƃ��Ă�
    '�V�X�e�����j���[�\������
    '���j���[������ɗL�������Ă��܂��s�������悤�Ȃ̂�
    'DeleteMenu�s�̗p�Ƃ���
End Sub

Public Function GetWindowCloseButton(ByVal hWnd As Long) As Boolean
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, False)
    GetWindowCloseButton = (GetMenuItemID(hMenu, 6) <> -1)
End Function

'----------------------------------------
'�ETopMost
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
'�EWindowState
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
'�EPixelRect
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
'�EIni�t�@�C���ʒu�ۑ����A
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
'��ComboBox
'----------------------------------------

'----------------------------------------
'�ECombobox.Text���N���A�����ɍ��ڂ����N���A����
'----------------------------------------
Public Sub Combobox_ClearList(ComboBox As ComboBox)
    Dim I As Long
    For I = ComboBox.ListCount - 1 To 0 Step -1
        ComboBox.RemoveItem (I)
    Next
End Sub

'----------------------------------------
'��ComboBox�ƕ�����z��Ƃ̕ϊ�
'----------------------------------------
'   �E  ���ڂ��^�u�ŋ�؂�
'   �E  �^�u��؂蕶�����ColumnCount�͈�v�����Ă�������
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

    '�G���[�`�F�b�N
    '�^�u��؂蕶����ComboBox�̗񐔂Ƃ����Ă��邩�ǂ���
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
        myData(1, 0) = "������"
        myData(1, 1) = "������"
        myData(1, 2) = "������"
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
        "A" + vbTab + "������" + vbTab + "1" + vbCrLf + _
        "B" + vbTab + "������" + vbTab + "2" + vbCrLf + _
        "C" + vbTab + "������" + vbTab + "3", _
        ArrayToString(Data, vbCrLf))
    Call ArrayAdd(Data, "D" + vbTab + "������" + vbTab + "4")
    Call ComboBox_SetStrings(ComboBox1, Data)
    Data = ComboBox_GetStrings(ComboBox1)
    Call Check( _
        "A" + vbTab + "������" + vbTab + "1" + vbCrLf + _
        "B" + vbTab + "������" + vbTab + "2" + vbCrLf + _
        "C" + vbTab + "������" + vbTab + "3" + vbCrLf + _
        "D" + vbTab + "������" + vbTab + "4", _
        ArrayToString(Data, vbCrLf))
End Sub



'----------------------------------------
'���A�C�R���pAPI����
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
'��Windows���
'----------------------------------------

'----------------------------------------
'�E�f�X�N�g�b�v/WorkArea�T�C�Y
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
'�EWindows�o�[�W����
'----------------------------------------

'Windows 8.1 ��     Windows (32-bit) NT 6.03(���Ԃ�)
'Windows 8 ��       Windows (32-bit) NT 6.02
'Windows 7 ��       Windows (32-bit) NT 6.01
'Windows Vista ��   Windows (32-bit) NT 6.00
'Windows XP ��      Windows (32-bit) NT 5.01
'Windows 2000 ��    Windows (32-bit) NT 5.00
'Windows Me ��      Windows (32-bit) 4.90
'Windows 98 ��      Windows (32-bit) 4.10
'Windows 95 ��      Windows (32-bit) 4.00
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
'���^�X�N�o�[�s���A�C�R���o�^�p
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
'�E�^�X�N�o�[�{�^���p��AppID�̓o�^
'----------------------------------------
Public Sub SetTaskbarButtonAppID(ByVal AppID As String)
    If IsTaskbarPinWindows Then
        Call SetCurrentProcessExplicitAppUserModelID( _
            StrPtr(AppID))
    End If
End Sub

'----------------------------------------
'�E�^�X�N�o�[�s���~�ߗp�R�}���h
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
    'InvokeVerb�̌�̕�����͕ϐ��ł̓_��
    '�Ȃ����萔�łȂ��Ƃ����Ȃ��B
End Sub

'----------------------------------------
'�E�^�X�N�o�[�s���p�V���[�g�J�b�g�t�@�C���̍쐬/�폜
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

        '�^�X�N�o�[�Ƀs���~��
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

        '�V���[�g�J�b�g�t�@�C���̃����N��ύX
        If FileCreateWait(ShortcutFilePath) Then
            Call CreateShortcutFile(ShortcutFilePath, _
                LinkTargetFilePath, _
                IconFilePath, Description)
        End If
        '�X�N���v�g�t�@�C���𒼐ڂ̓^�X�N�o�[�s���~�߂ł��Ȃ��̂�
        '��x�_�~�[�̃v���O������o�^����
        '���̌�ŃV���[�g�J�b�g�t�@�C���̃����N������������Ă���B

    ElseIf (Value = False) And (FileExistsFlag) Then

        '�^�X�N�o�[�s������
        Call SetTaskbarPin(ShortcutFilePath, False)
    End If
End Sub

'----------------------------------------
'���}�E�X����
'----------------------------------------

Public Sub MouseMove(ByRef Position As Point)
    Dim DesktopRect As Rect
    DesktopRect = GetRectDesktop

    Call mouse_event(MOUSE_MOVED Or MOUSEEVENTF_ABSOLUTE, _
        Position.X * (65535 / GetRectWidth(DesktopRect)), _
        Position.Y * (65535 / GetRectHeight(DesktopRect)), 0, 0)
    '���N���b�N�ʒu����ʉ𑜓x����␳����
End Sub

Public Sub MouseClick()
   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
End Sub

'----------------------------------------
'��Internet�n�֐�
'----------------------------------------
'----------------------------------------
'�EURL�w��̃t�@�C���_�E�����[�h
'----------------------------------------
'   �E  API��URLDownloadToFile���g���₷������
'----------------------------------------
Public Function URLDownloadFile(ByVal URL As String, ByVal FilePath As String) As Boolean
    Dim Result As Long
    Result = URLDownloadToFile(0, URL, FilePath, 0, 0)
    URLDownloadFile = (Result = 0)
End Function

'----------------------------------------
'�E���{�ꕶ�����URL�G���R�[�h
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
'��VBE����
'----------------------------------------


'----------------------------------------
'���Q�Ɛݒ�ǉ�
'----------------------------------------

'----------------------------------------
'�EMicrosoft Scripting Runtime
'----------------------------------------
'   �E  FSO:FileSystemObject���g�p����̂ɕK�v
'----------------------------------------
Sub ReferenceAdd_ScriptingRuntime(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\system32\scrrun.dll")
End Sub

Sub Run_ReferenceAdd_ScriptingRuntime()
    Call ReferenceAdd_ScriptingRuntime(ThisWorkbook)
End Sub

'----------------------------------------
'�EWindows Script Host Object Model
'----------------------------------------
'   �E  WshShell���g�p����̂ɕK�v
'----------------------------------------
Sub ReferenceAdd_WshObjectModel(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\system32\wshom.ocx")
End Sub

Sub Run_ReferenceAdd_WshObjectModel()
    Call ReferenceAdd_WshObjectModel(ThisWorkbook)
End Sub

'----------------------------------------
'�EMicrosoft Windows Common Controls 6.0 (SP6)
'----------------------------------------
Sub ReferenceAdd_CommonControls(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\System32\MSCOMCTL.OCX")
End Sub

Sub Run_ReferenceAdd_CommonControls()
    Call ReferenceAdd_CommonControls(ThisWorkbook)
End Sub

'----------------------------------------
'�EMicrosoft Visual Basic for Applications Extensibility 5.3
'----------------------------------------
Sub ReferenceAdd_VBAExtensibility(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB")
End Sub

Sub Run_ReferenceAdd_VBAExtensibility()
    Call ReferenceAdd_VBAExtensibility(ThisWorkbook)
End Sub

'----------------------------------------
'�EMicrosoft AxtiveX Data Objects 2.8 Library
'----------------------------------------
'   �E  ADODB.Stream���g�p����̂ɕK�v
'----------------------------------------
Sub ReferenceAdd_ADO_2_8(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files\Common Files\System\ado\msado28.tlb")
End Sub

Sub Run_ReferenceAdd_ADO_2_8()
    Call ReferenceAdd_ADO_2_8(ThisWorkbook)
End Sub

'----------------------------------------
'�EMicrosoft AxtiveX Data Objects 6.1 Library
'----------------------------------------
'   �E  ADODB.Stream���g�p����̂ɕK�v
'----------------------------------------
Sub ReferenceAdd_ADO_6_1(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files\Common Files\System\ado\msado15.dll")
End Sub

Sub Run_ReferenceAdd_ADO_6_1()
    Call ReferenceAdd_ADO_6_1(ThisWorkbook)
End Sub



'--------------------------------------------------
'������
'�� ver 2014/11/03
'�E �쐬
'�E �����񏈗�First/Last/Delimiter
'�E �O���t����
'�E DataLastRow/Col
'�E ArrayCount
'�E Assert/Check/OrValue
'�E IncludeLastPathDelim
'�E IniFile_GetString/SetString
'�E GetAbsolutePath
'�E MaxValue/MinValue
'�E LongToStrDigitZero
'�E PixelToPoint/PointToPixel
'�E ADOStream
'�� ver 2014/11/06
'�E CommandExecuteReturn
'�E IncludeBothEndsStr/ExcludeBothEndsStr
'�E GetFirstStr---/GetLastStr---
'�E TrimLast/TrimFirst
'�E IsLong
'�� ver 2014/11/07
'�E ClearLast
'�E CommandExecuteReturn
'�� ver 2014/11/08
'�E ChartObjectExists/ShapeExists
'�� ver 2014/11/19
'�E ExcludeLastPathDelim�ǉ�
'�E UBound/LBound
'�E ArrayStr/StringArrayCombine/StringCombine/PathCombine
'�E GetExtensionIncludePeriod/ChangeFileExtension
'�E Get/SetWindowLong
'�E SetWindowStyle/SetWindowExStyle/SetWindowTopMost
'�� ver 2014/11/20
'�E GetAsyncKeyState
'�E BooleanToString
'�E FormatYYYY_MM_DD/FormatHH_MM_SS
'�E GetFolderPathListTopFolder
'�E ClearLineColumn
'�E SetTaskbarButtonAppID
'�� ver 2014/11/21
'�E SetIcon/ResetIcon
'�� ver 2014/11/24
'�E BooleanToString>>BoolToStr
'�E RectToStr/StrToRect
'�E NewRect/NewRectSize/NewPoint/NewRect_PositionSize
'   /GetRectSize/RectEqual
'   /GetRectInsideDesktopRect
'�E PopupMenu
'�E Form_GetRectPixel/Form_SetRectPixel
'�E GetDesktopWindow/GetWindowRect/SystemParametersInfo
'   GetRectDesktop/GetRectWorkArea
'�E GetSpecialFolderPath
'�E Form_IniWritePosition/Form_IniReadPosition
'�E TaskDialog
'�� ver 2014/11/26
'�E IsWindowsOffice64/32bit
'   WindowsMajor/MinorVersion
'   IsTaskbarPinWindows
'�E ForceCreateFolder
'�E CreateShortcutFile
'�E GetWindowState
'�E GetRectInsideDesktopRect�C��
'�� ver 2014/12/01
'�E TaskDialog�n�̏C��
'�E SetWindowIcon/ResetWindowIcon
'�E GetBitmapDrawIcon/Image_Picture_SetBitmap
'�E GetDC/FillRect/DrawIcon
'   /CreateCompatibleDC/CreateCompatibleBitmap
'   /SelectObject/DeleteObject/GetStockObject
'�E GetWindowCloseButton/GetWindowStyle/GetWindowExStyle
'   /GetWindowIcon
'�� ver 2014/12/02
'�E MouseMove/MouseClick
'�� ver 2014/12/04
'�E SetShortcutIcon/SetTaskbarPinShortcutIcon/SetTaskbarPin
'�� ver 2014/12/06
'�E StrToLongDefault
'�E ArrayAdd
'�E ApplicationMode/SetExcelWindowTitle
'�� ver 2015/02/02
'�E Microsoft Forms 2.0 Object Library�̎Q�Ɛݒ�ǉ�
'�E FirstStrFirstDelim/FirstStrLastDelim
'   /LastStrFirstDelim/LastStrLastDelim
'�� ver 2015/02/06
'�E ReCreateFolder�쐬
'�� ver 2015/02/13
'�E DataLastCol�C��
'   DataLastCell�쐬
'�� ver 2015/03/05
'�E �Q�Ɛݒ�ReferenceAdd�n�����ǉ�
'�E �z��֘A�����ǉ�
'   ArrayInsert/ArrayDelete
'   /ArrayIndexOf/ArrayDeleteSameItem
'   /ArrayDimension/ArrayToString
'�E ListView�֘A�����ǉ�
'   ListView_SelectedItemCount/ListView_CheckedItemCount
'   /ListView_SelectAll/ListView_CheckSelectedItem
'   /ListView_IsCheckSelectedItem/ListView_MultiSelectChecked
'   /ListView_IndexOfKey
'�E �t�@�C�������֘A�����ǉ�
'   DateToApiFILETIME/GetFileFolderTime/SetFileFolderTime
'�E FormatDateTimeNormal�ǉ�
'�E �t�@�C���t�H���_�ꗗ�����ǉ�
'   FolderPathListTopFolder/FolderPathListSubFolder
'   /FilePathListTopFolder/FilePathListSubFolder
'�E ComboBox�֘A�����ǉ�
'   ComboBox_GetStrings/ComboBox_SetStrings
'   /Combobox_ClearList
'�E ���O�ύX GetAbsolutePath>>AbsolutePath
'�E StringCombine/StringCombineArray
'   /PathCombine�C��
'�� ver 2015/03/11
'�E ArraySetValueObject��ǉ�
'   ArrayAdd/ArrayInsert���C��
'�� ver 2015/03/19
'�E ArrayAdd/ArrayInsert/ArrayDelete���C��
'�E �R�����g�̏C��
'�� ver 2015/07/23
'�E StarndardSoftwareLibrary����st_vba�ɖ��O�ύX
'�� ver 2015/07/29
'�E 64bit��Excel�ւ̎b��Ή�(������32bit��Excel�݂̂̑Ή�)
'   TaskDialogAPI���폜
'�E GetDPI�̐������������s�����B
'�� ver 2015/08/07
'�E FileExists(Win/Mac���Ή���)��ǉ�
'�E GetClipboardText/SetClipboardText(Win/Mac���Ή���)��ǉ�
'�� ver 2015/08/23
'�E CommandExecute��ǉ�
'�E PopupMenu_PopupReturn_NoPosition��ǉ�
'�E IsShortcutLinkFile�ǉ�
'�E IsJpegImageFile/IsJpegExifFile�ǉ�
'�E GetJpegExifDateTime�ǉ�
'�� ver 2015/12/12
'�E Excel64bit�萔�ǉ�
'�E SleepAPI�ǉ�
'�E IE_NewObject/IE_GetObject/IE_Navigate
'   /IE_NavigateWait/IE_RunJavaScript�ǉ�
'�E IsIncludeStr�ǉ�
'�� ver 2015/12/16
'�E ClearLastRange/ClearLastColumn/ClearLastRow
'   /ClearLastRangeContents
'   /ClearLastColumnContents/ClearLastRowContents���C���ǉ�
'�� ver 2016/01/08
'�E ClearLastRange/ClearLastColumn/ClearLastRow���C��
'   ClearContents�@�\��ǉ�
'�E TrimFirstChar/TrimLastChar/TrimBothEndsChar��p�~
'   TrimFirstStrs/TrimLastStrs/TrimBothEndsStrs
'   /TrimFirstSpace/TrimLastSpace/TrimBothEndsSpace��ǉ�
'�E DataLastRow/DataLastCol���f�[�^���Ȃ��Ƃ��ɃG���[��������̂�
'   OnErrorResume����悤�ɏC��
'�� ver 2016/02/06
'�E Enum AlineHorizontal/AlineVertical �̒�`
'�E URLDownloadToFile API��URLDownloadFile�̒ǉ�
'�E ���t���������w��֐��̒ǉ�
'   FormatYYYYMMDD/FormatYYYY_MM
'   /FormatHHMMSS/FormatHH_MM
'   /FormatYYYYMMDDHHMMSS/FormatYYYYMMDDHHMMSS_Hyphen
'�E �N���A�`�����̖��O�ύX
'   ClearLastRange��ClearRangeLast
'   ClearLastColumn��ClearColumnLast
'   ClearLastRow��ClearRowLast
'�E Shape�����̒ǉ�
'   GetShapeFromImageFile/ShapeCompressUseClipboard
'�E IE�����̏C�� IE_NewObject/IE_Refresh
'   /IE_Navigate/IE_NavigateWait
'�� ver 2016/02/20
'�E GetWorkbook�ǉ�
'�E GetWorksheet/WorksheetExists�ǉ�
'�E DeleteSheet/DeleteDefaultSheet�ǉ�
'�E SetTextSheet�ǉ�
'�E TagInnerText/TagOuterText�ǉ�
'�E IfEmptyStr�ǉ�
'�E �Z���N���A�n�����̖��O�ύX
'   ClearRangeLast��ClearRangeLastData
'   ClearColumnLast��ClearColumnLastRow
'   ClearRowLast��ClearRowLastColumn
'�E URLDownloadFile�̖߂�l��Boolean�ɕύX
'�� ver 2016/02/21
'�E IsNothing/IsNotNothing�ǉ�
'�E CastExcludeComma�ǉ�
'�E IE_GetElementByTagNameClassName/IE_GetElementByTagNameInnerHTML�̒ǉ�
'�E FormulaDeleteRange�ǉ�
'�E ColumnNumberByTitle�ǉ�
'�E ColumnNumber�ǉ�
'�E CopyFile�ǉ�
'�� ver 2016/02/23
'�E ThisWeekDay/LastWeekDay/NextWeekDay�ǉ�
'�� ver 2016/02/24
'�E IsDrivePath/IsNetworkPath�ǉ�
'�E SettingFullPath�ǉ�
'   AbsolutePath�C��
'�� ver 2016/02/28
'�E ThisWeekDay/LastWeekDay/NextWeekDay�C��
'�E ColumnNumberByTitle�C��
'�E RangeClear�@�\�ǉ�MergeCellOption�Ή�
'�E RangeCopyNumberFormat/RangeCopyFormat/RangeCopyAll�ǉ�
'�E FormulaDeleteRange��RangeDeleteFormula���O�ύX
'�E FirstStrFirstDelim/FirstStrLastDelim
'   /LastStrFirstDelim/LastStrLastDelim �̏C��
'�E DeleteSheet�̏C��
'�E SetTextSheet�̏C��
'�E IE_GetElementByTagNameId�ǉ�
'�� ver 2016/02/29
'�E ClearRangeLastData/ClearColumnLastRow/ClearRowLastColumn�C��
'�� ver 2016/03/04
'�E TagOuterText�̏C��
'�E TagOuterTextList�ǉ�
'�E ReplaceHTMLTag�ǉ�
'�� ver 2016/03/10
'�E Wingdings_Checkbox_Checked/UnChecked�ǉ�
'�E urlEncode�ǉ�
'�E ArrayAddNotDuplicate/ArrayExists�ǉ�
'�E ArraySortQuick�ǉ�
'�E RangeUpRow/RangeDownRow�ǉ�
'�E RangeMoveUpRowOne/RangeMoveDownRowOne�ǉ�
'�E LengthSjisByte
'   /LeftSjisByte/RightSjisByte
'   /MidSjisByte�ǉ�
'�� ver 2016/03/13
'�E urlEncode�C��
'�E TopLeftCell�ǉ�
'�E StrCount�ǉ�
'�E StrToBool�ǉ�
'�E st_vba_Base����Ast_vba_Core�ɖ��̕ύX
'�E ListView�������Ast_vba_ListView�Ɉڍs
'�E InternetExplorer�������Ast_vba_IE�Ɉڍs
'�� ver 2016/03/20
'�E IE_GetElementByTagName��ǉ�
'�E ReplaceContinuousSpace�ǉ�
'�E RangeCopyValue�ǉ�
'�E MatchRegExp�ǉ�
'�E ArrayIndexOf�Ɋ��S��v/������v/���C���h�J�[�h/���K�\��
'   �̋@�\��ǉ��BArrayExists���ǉ��B
'�� ver 2016/03/23
'�E ArrayIndexOf�����ǂ���
'   ���C���h�J�[�h�z��/���K�\���z��̋@�\��ǉ�
'�E ReplaceArrayValue/DeleteArrayValue��ǉ�
'�E ArraySortOrder��ǉ�
'�� ver 2016/03/26
'�E ArraySortOrder���C��
'   ArraySortCustomOrder�ɖ��̕ύX
'�E ReplaceRegExp��ǉ�
'�E ReplaceArrayRegExp��ǉ�
'�E DeleteArrayRegExp��ǉ�
'�E ArraySortQuick��SortOrder�@�\�ǉ�
'�E ArraySortStrLength�ǉ�
'�E ArrayReverse�ǉ�
'�E ShapeCompressUseClipboard�C��
'�E RowNumberByTitle�ǉ�
'�� ver 2016/03/27
'�E ArrayIsUnique�ǉ�
'�E 2�����z��n�̏�����ǉ�
'   Array2dSetColumn
'   /Array2dSetRowValues/Array2dGetRowValues
'   /Array2dAdd/Array2dInsert/Array2dDelete
'   /Array2dSortQuick/Array2dIsUnique
'--------------------------------------------------
