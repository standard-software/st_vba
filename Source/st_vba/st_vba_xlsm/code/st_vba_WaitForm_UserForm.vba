'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    WaitForm Module
'ObjectName:    st_vba_WaitForm_UserForm.vba
'--------------------------------------------------
'Version:       2017/04/16
'--------------------------------------------------
'   ・  処理を待たせるときにこのFormを使って
'       キャンセルボタンを表示して進捗表示する
'--------------------------------------------------
Option Explicit

Private FormProperty As New st_vba_FormProperty

Public RunCancelFlag As Boolean

Private Sub ButtonCancel_Click()
    RunCancelFlag = True
    Me.Hide
End Sub

'------------------------------
'◇初期化
'------------------------------
Public Sub Initialize( _
ByVal TaskBarButton As Boolean, _
ByVal TitleBar As Boolean, _
ByVal SystemMenu As Boolean, _
ByVal FormIcon As Boolean, _
ByVal MinimizeButton As Boolean, _
ByVal MaximizeButton As Boolean, _
ByVal CloseButton As Boolean, _
ByVal ResizeFrame As Boolean, _
ByVal TopMost As Boolean)

    With Nothing
        Call FormProperty.InitializeForm(Me)

        Call FormProperty.InitializeProperty( _
            TaskBarButton:=TaskBarButton, _
            TitleBar:=TitleBar, _
            SystemMenu:=SystemMenu, _
            FormIcon:=FormIcon, _
            MinimizeButton:=MinimizeButton, _
            MaximizeButton:=MaximizeButton, _
            CloseButton:=CloseButton, _
            ResizeFrame:=ResizeFrame, _
            TopMost:=TopMost)

    End With


End Sub


Private Sub UserForm_Activate()
    If FormProperty.Initializing Then
    
        RunCancelFlag = False

        If FormProperty.Handle = 0 Then
            Call FormProperty.InitializeForm(Me)
            FormProperty.GetWindowsProperty
        Else
            FormProperty.SetWindowsProperty
        End If

        FormProperty.Initializing = False
        
    End If
End Sub


'------------------------------
'◇ループ表示処理
'------------------------------

'------------------------------
'・ ループ表示の制御
'------------------------------
'   ・  PassLoopCount=10 とすることで
'       ループ中に10回に1回だけ
'       戻り値 = Result = True になり、
'       呼び出し元で戻り値をチェックして
'       DoEventsを行うようにするとループ中も
'       とまったように見えない
'   ・  PassLoopCount=1にすると毎回DoEventsされる
'       PassLoopCount=5にすると5回に1回DoEventsされる
'   ・  StartLimitを設定すると
'       処理開始時の初期のループ回数だけ毎回DoEventされる
'   ・  LabelText/StatusTextを指定することもできる
'       指定なければ 5/100 　5% という表示になる
'------------------------------

Public Function Update( _
ByVal LabelText As String, _
ByVal StatusText As String, _
ByVal PassLoopCount As Long, _
ByVal Value As Long, _
ByVal StartValue As Long, ByVal EndValue As Long, _
Optional ByVal StartLimit As Long = 1) As Boolean

    Dim Result As Boolean: Result = False
    
    If ((Value - StartValue + 1) = 1) Then Result = True
    If ((Value - StartValue + 1) <= StartLimit) Then Result = True
    If (((Value - StartValue + 1) Mod PassLoopCount) = 0) Then Result = True
    
    If Result Then
        Me.Label1.Caption = LabelText
        If Me.Visible = False Then Me.Show
        Application.StatusBar = StatusText
    End If
    
    Update = Result
End Function

Public Function Update_ProgressInfo( _
ByVal Message As String, _
ByVal PassLoopCount As Long, _
ByVal Value As Long, _
ByVal StartValue As Long, ByVal EndValue As Long, _
Optional ByVal StartLimit As Long = 1) As Boolean

    Update_ProgressInfo = _
        Update( _
            ProgressText(Message, vbCrLf, StartValue, Value, EndValue, , False), _
            ProgressText(Message, "|", StartValue, Value, EndValue), _
            PassLoopCount, _
            Value, StartValue, EndValue, StartLimit)
End Function


'------------------------------
'・ 使い方
'------------------------------
'   ・  下記のようにして使う
'------------------------------

'    Dim WaitForm As New st_vba_WaitForm
'
'    'キャンセル画面の初期化
'    WaitForm.Label1.Caption = ""
'    Call WaitForm.Initialize( _
'        TaskBarButton:=False, _
'        TitleBar:=True, _
'        SystemMenu:=False, _
'        FormIcon:=False, _
'        MinimizeButton:=False, _
'        MaximizeButton:=False, _
'        CloseButton:=False, _
'        ResizeFrame:=False, _
'        TopMost:=True)
'    WaitForm.RunCancelFlag = False
'
'    Application.ScreenUpdating = False
'
'    Dim I As Long
'    Dim StartIndex As Long
'    Dim EndIndex As Long
'
'    If WaitForm.RunCancelFlag = False Then
'
'        StartIndex = 1
'        EndIndex = 1000
'
'        For I = StartIndex To EndIndex
'        Do
'            If WaitForm.Update_ProgressInfo( _
'                "処理A:", _
'                10,
'                Row_WriteIndex, _
'                StartIndex, _
'                EndIndex) Then
'                DoEvents
'                If WaitForm.RunCancelFlag = True Then Exit For
'            End If
'
'        :
'
'        Loop While False
'        Next
'        Application.StatusBar = False
'    End If
'
'    If WaitForm.RunCancelFlag = False Then
'
'        StartIndex = 1
'        EndIndex = 1000
'
'        For I = StartIndex To EndIndex
'        Do
'            If WaitForm.Update_ProgressInfo(10, _
'                "処理B:", _
'                Row_WriteIndex, _
'                StartIndex, _
'                EndIndex) Then
'                DoEvents
'                If WaitForm.RunCancelFlag = True Then Exit For
'            End If
'
'        :
'
'        Loop While False
'        Next
'        Application.StatusBar = False
'    End If
'
'    Application.ScreenUpdating = True
'    Application.StatusBar = False
'    WaitForm.Hide
'
'    If WaitForm.RunCancelFlag = False Then
'        If (Check_IE_Visible = False) Then
'            ie.Quit
'        End If
'    End If
'    Set WaitForm = Nothing


