'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    StringBuilder Class
'ObjectName:    st_vba_StringBuilder
'--------------------------------------------------
'Version:       2017/03/19
'--------------------------------------------------
'   �E  ����������������Ĉ������߂̃��W���[��
'   �E  .NET �� StringBuilder �ƃ��\�b�h�̌݊���������킯�ł͂Ȃ���
'       ���ʓ������Ƃ��ł�����̂Ȃ̂ŁA���O���؂�Ă���
'--------------------------------------------------
Option Explicit

Const BufferBlockSize As Long = 1000000

Private Buffer As String
Public LengthIndex As Long

Sub Initialize()
    LengthIndex = 0
    Buffer = VBA.String(BufferBlockSize, " ")
End Sub

Sub Clear()
    Call Initialize
End Sub

Sub Add(Value As String)
    Dim NewLengthIndex As Long
    NewLengthIndex = LengthIndex + Len(Value)
    Do While (Len(Buffer) < NewLengthIndex)
        Buffer = Buffer + VBA.String(BufferBlockSize, " ")
    Loop

    Mid(Buffer, LengthIndex + 1) = Value
    LengthIndex = LengthIndex + Len(Value)
End Sub

Public Property Get Text() As String
    Text = Left(Buffer, LengthIndex)
End Property

'----------
'�g�����Ƥ����m�F�R�[�h
'Option Explicit
'
'    Sub ������A����Mid�X�e�[�g�����g�Ƃ̔�r()
'
'        Dim Str1 As String
'        Dim Str2 As String
'        Dim Str3 As String
'        Dim lng As Long
'        Dim strTime As String
'        Dim dblTimer As Double
'
'        Dim LoopCount As Long
'        LoopCount = 200000
'
'        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'        '�ʏ�̘A��
'        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'    ' dblTimer = Time
'    '
'    ' Str1 = ""
'    ' For lng = 1 To LoopCount
'    ' Str1 = Str1 & "��������������������"
'    ' Next lng
'    '
'    ' strTime = "�ʏ�F" & Time - dblTimer
'
'        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'        'Mid�X�e�[�g�����g�g�p
'        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'
'        dblTimer = Time
'
'        '���炩���ߕK�v�ȕ����������m��
'        Str2 = VBA.String(LoopCount * 10, "a")
'        Dim lngPos As Long
'        lngPos = 1 '���[�v���ŊJ�n���镶���ʒu
'        For lng = 1 To LoopCount
'            Mid(Str2, lngPos, 10) = "��������������������"
'            lngPos = lngPos + 10
'        Next lng
'        Str2 = Left(Str2, lngPos - 1)
'
'        strTime = strTime & vbCrLf & "Mid�F" & Time - dblTimer
'
'        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'        'StringBuilder
'        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'
'        dblTimer = Time
'
'        Dim StrBuilder As StringBuilder
'        Set StrBuilder = New StringBuilder
'
'        For lng = 1 To LoopCount
'            Call StrBuilder.Add("��������������������")
'        Next
'        Str3 = StrBuilder.Text
'
'        strTime = strTime & vbCrLf & "Mid�F" & Time - dblTimer
'
'        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'        '��r
'        '�\�\�\�\�\�\�\�\�\�\�\�\�\�\�\
'
'        If (Str1 <> "") And (Str1 <> Str2) Then
'            MsgBox "Str1<>Str2"
'        End If
'
'        If Str2 <> Str3 Then
'            MsgBox "Str1<>Str2"
'        End If
'
'
'        '���Ԕ��\
'        MsgBox strTime
'
'    End Sub

