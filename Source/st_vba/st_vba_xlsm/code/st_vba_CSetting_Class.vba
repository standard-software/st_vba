'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    Setting Class
'ObjectName:    st_vba_CSetting
'--------------------------------------------------
'Version:       2017/09/19
'--------------------------------------------------
Option Explicit

Private m_Sheet As Worksheet

Public Sub Initialize(ByVal Sheet As Worksheet)

    Set m_Sheet = Sheet

End Sub

'�V�[�g�����̂悤�ɂ��Ă���
'
'   |A          |B      |C
'1  |�ݒ薼     |�l     |����
'2  |Option1    |�lA    |
'3  |Option2    |�lB    |
'
'�������Ă����ƁASetting.Read("Option1")�ŒlA�̓��e���擾���邱�Ƃ��ł���

Public Function Read(ByVal Key As String) As String
    Read = m_Sheet.Cells( _
        Sheet_RowNumberByTitle(m_Sheet, Col_A, Key), _
        Col_B).Value
End Function
