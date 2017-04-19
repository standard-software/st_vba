'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    WaitForm Module
'ObjectName:    st_vba_WaitForm_UserForm.vba
'--------------------------------------------------
'Version:       2017/04/16
'--------------------------------------------------
'   �E  ������҂�����Ƃ��ɂ���Form���g����
'       �L�����Z���{�^����\�����Đi���\������
'--------------------------------------------------
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

'------------------------------
'�E ���[�v�\���̐���
'------------------------------
'   �E  ���[�v���A1000��̒���10���1��Ƃ�����
'       Result = True �ɂ��āA�O����DoEvents���s��
'   �E  PassLoopCount=1�ɂ���Ɩ���DoEvents
'       PassLoopCount=5�ɂ����5���1��DoEvents
'   �E  StartLimit��ݒ肷��Ə����̃��[�v�񐔂�������DoEvent
'------------------------------

Public Function Update_ProgressInfo( _
ByVal Message As String, _
ByVal PassLoopCount As Long, _
ByVal Value As Long, _
ByVal StartValue As Long, ByVal EndValue As Long, _
Optional ByVal StartLimit As Long = 1) As Boolean

    Dim Result As Boolean: Result = False
    
    If ((Value - StartValue + 1) = 1) Then Result = True
    If ((Value - StartValue + 1) <= StartLimit) Then Result = True
    If (((Value - StartValue + 1) Mod PassLoopCount) = 0) Then Result = True
    
    If Result Then
        Me.Label1.Caption = _
            ProgressText(Message, vbCrLf, StartValue, Value, EndValue, , False)
        If Me.Visible = False Then Me.Show
        Call Application_StatusBar_Progress(Message, "|", StartValue, Value, EndValue)
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
'            If WaitForm.Update_ProgressInfo( _
'                "����A:", _
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


