Option Explicit

Private FormProperty As New st_vba_FormProperty

Public RunCancelFlag As Boolean

Private Sub ButtonCancel_Click()
    RunCancelFlag = True
    Me.Hide
End Sub

'------------------------------
'��������
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
'�����[�v�\������
'------------------------------

Public Function Update_ProgressInfo( _
ByVal PassLoopCount As Long, _
ByVal Message As String, _
ByVal Value As Long, _
ByVal StartValue As Long, ByVal EndValue As Long) As Boolean

    Dim Result As Boolean: Result = False
    If ((Value - StartValue + 1) = 1) Or (((Value - StartValue + 1) Mod PassLoopCount) = 0) Then
        Me.Label1.Caption = _
            ProgressText(Message + vbCrLf, vbCrLf, StartValue, Value, EndValue)
        If Me.Visible = False Then Me.Show
        Call Application_StatusBar_Progress(Message, StartValue, Value, EndValue)
        Result = True
    End If
    Update_ProgressInfo = Result
End Function

'------------------------------
'�E �g����
'------------------------------
'   �E  ���L�̂悤�ɂ��Ďg��
'------------------------------

'    Dim WaitForm As New st_vba_WaitForm
'
'    '�L�����Z����ʂ̏�����
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
'            If WaitForm.Update_ProgressInfo(10, _
'                "����A:", _
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
'                "����B:", _
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


