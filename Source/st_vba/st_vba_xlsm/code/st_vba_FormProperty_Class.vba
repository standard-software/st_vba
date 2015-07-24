'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    FormProperty Class
'ObjectName:    st_vba_FormProperty
'--------------------------------------------------
'Version:       2015/07/24
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'■変数宣言等
'--------------------------------------------------
Private FTaskbarButton As Boolean
Private FTitleBar As Boolean
Private FSystemMenu As Boolean
Private FFormIcon As Boolean
Private FMinimizeButton As Boolean
Private FMaximizeButton As Boolean
Private FCloseButton As Boolean
Private FResizeFrame As Boolean
Private FTopMost As Boolean
   
Private FForm As Object
Private FHandle As Long
Private FIconPath  As String
Private FIconIndex As Long

Public Initializing As Boolean

'--------------------------------------------------
'■初期化
'--------------------------------------------------
Private Sub Class_Initialize()
    Initializing = True
End Sub

Public Sub InitializeForm(ByVal Form As Object)
    Set FForm = Form
    Call WindowFromAccessibleObject(Form, FHandle)
End Sub

Public Sub InitializeProperty( _
ByVal TaskBarButton As Boolean, _
ByVal TitleBar As Boolean, _
ByVal SystemMenu As Boolean, _
ByVal FormIcon As Boolean, _
ByVal MinimizeButton As Boolean, _
ByVal MaximizeButton As Boolean, _
ByVal CloseButton As Boolean, _
ByVal ResizeFrame As Boolean, _
ByVal TopMost As Boolean)
    
    FTaskbarButton = TaskBarButton
    FTitleBar = TitleBar
    FSystemMenu = SystemMenu
    FFormIcon = FormIcon
    FMinimizeButton = MinimizeButton
    FMaximizeButton = MaximizeButton
    FCloseButton = CloseButton
    FResizeFrame = ResizeFrame
    FTopMost = TopMost
End Sub

'--------------------------------------------------
'■プロパティ
'--------------------------------------------------
Property Get Handle() As Long
    Handle = FHandle
End Property

'値セットは特殊なためにSetHandleメソッドとする
Public Sub SetHandle(ByVal Value As Long)
    FHandle = Value
End Sub

Property Get TaskBarButton() As Boolean
    TaskBarButton = FTaskbarButton
End Property

Property Let TaskBarButton(ByVal Value As Boolean)
    FTaskbarButton = Value
    Call SetWindowExStyle(FHandle, _
        FTaskbarButton)
End Property

Property Get TitleBar() As Boolean
    TitleBar = FTitleBar
End Property

Property Let TitleBar(ByVal Value As Boolean)
    FTitleBar = Value
    Call SetWindowStyle(FHandle, _
        FTitleBar, FSystemMenu, FResizeFrame, _
        FMinimizeButton, FMaximizeButton)
End Property

Property Get SystemMenu() As Boolean
    SystemMenu = FSystemMenu
End Property

Property Let SystemMenu(ByVal Value As Boolean)
    FSystemMenu = Value
    Call SetWindowStyle(FHandle, _
        FTitleBar, FSystemMenu, FResizeFrame, _
        FMinimizeButton, FMaximizeButton)
End Property

Property Get FormIcon() As Boolean
    FormIcon = FFormIcon
End Property

Property Let FormIcon(ByVal Value As Boolean)
    FFormIcon = Value
    Call SetFormIcon(Value)
End Property

Public Sub SetFormIcon(ByVal Value As Boolean)
    If Value Then
        Call SetWindowIcon(FHandle, FIconPath, FIconIndex)
    Else
        Call ResetWindowIcon(FHandle)
    End If
End Sub

Public Sub GetFormIcon(ByRef Value As Boolean)
    Value = GetWindowIcon(FHandle)
End Sub

Property Get MinimizeButton() As Boolean
    MinimizeButton = FMinimizeButton
End Property

Property Let MinimizeButton(ByVal Value As Boolean)
    FMinimizeButton = Value
    Call SetWindowStyle(FHandle, _
        FTitleBar, FSystemMenu, FResizeFrame, _
        FMinimizeButton, FMaximizeButton)
End Property

Property Get MaximizeButton() As Boolean
    MaximizeButton = FMaximizeButton
End Property

Property Let MaximizeButton(ByVal Value As Boolean)
    FMaximizeButton = Value
    Call SetWindowStyle(FHandle, _
        FTitleBar, FSystemMenu, FResizeFrame, _
        FMinimizeButton, FMaximizeButton)
End Property

Property Get CloseButton() As Boolean
    CloseButton = FCloseButton
End Property

Property Let CloseButton(ByVal Value As Boolean)
    FCloseButton = Value
    Call SetFormCloseButton(Value)
End Property

Public Sub SetFormCloseButton(ByVal Value As Boolean)
    Call SetWindowCloseButton(FHandle, _
        Value)
End Sub

Public Sub GetFormCloseButton(ByRef Value As Boolean)
    Value = GetWindowCloseButton(FHandle)
End Sub

Property Get ResizeFrame() As Boolean
    ResizeFrame = FResizeFrame
End Property

Property Let ResizeFrame(ByVal Value As Boolean)
    FResizeFrame = Value
    Call SetWindowStyle(FHandle, _
        FTitleBar, FSystemMenu, FResizeFrame, _
        FMinimizeButton, FMaximizeButton)
End Property

Property Get TopMost() As Boolean
    TopMost = FTopMost
End Property

Property Let TopMost(ByVal Value As Boolean)
    FTopMost = Value
    Call SetFormTopMost(Value)
End Property

Public Sub SetFormTopMost(ByVal Value As Boolean)
    Call SetWindowTopMost(FHandle, Value)
End Sub

Public Sub GetFormTopMost(ByRef Value As Boolean)
    Value = GetWindowTopMost(FHandle)
End Sub

Property Get IconPath() As String
    IconPath = FIconPath
End Property

Property Let IconPath(ByVal Value As String)
    FIconPath = Value
End Property

Property Get IconIndex() As Long
    IconIndex = FIconIndex
End Property

Property Let IconIndex(ByVal Value As Long)
    FIconIndex = Value
End Property

Property Get WindowState() As Excel.XlWindowState
    WindowState = _
        GetWindowState(Handle)
End Property

'--------------------------------------------------
'■Windowへの値設定/取得
'--------------------------------------------------
Public Sub SetWindowsProperty()
    Call SetWindowExStyle(FHandle, FTaskbarButton)
    Call SetWindowStyle(FHandle, _
        FTitleBar, FSystemMenu, FResizeFrame, _
        FMinimizeButton, FMaximizeButton)
    Call SetFormCloseButton(FCloseButton)
    Call SetFormTopMost(FTopMost)
    Call SetFormIcon(FFormIcon)
End Sub

Public Sub GetWindowsProperty()
    Call GetWindowExStyle(FHandle, FTaskbarButton)
    Call GetWindowStyle(FHandle, _
        FTitleBar, FSystemMenu, FResizeFrame, _
        FMinimizeButton, FMaximizeButton)
    Call GetFormCloseButton(FCloseButton)
    Call GetFormTopMost(FTopMost)
    Call GetFormIcon(FFormIcon)
End Sub

'--------------------------------------------------
'■マウスによる強制アクティブ化
'--------------------------------------------------
Public Sub ForceActiveMouseClick()
    Dim CursorPosBuffer As Point
    Call GetCursorPos(CursorPosBuffer)
    '↑カーソル位置を保持
    
    If FTopMost = False Then
        Call SetFormTopMost(True)
    End If
    
    '↓マウスをForm位置でクリック
    Call MouseMove(NewPoint( _
        PointToPixel(FForm.Left) + 1, _
        PointToPixel(FForm.Top) + 1))
    Call MouseClick
    
    If FTopMost = False Then
        Call SetFormTopMost(False)
    End If
    
    '↓カーソル位置を元に戻す
    Call MouseMove(CursorPosBuffer)
End Sub



