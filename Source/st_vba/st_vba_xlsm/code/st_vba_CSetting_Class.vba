'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    Setting Class
'ObjectName:    st_vba_CSetting
'--------------------------------------------------
'Version:       2017/04/01
'--------------------------------------------------
Option Explicit

Public Option1 As String


Public Sub Initialize(ByVal Sheet As Worksheet)
    
    Option1 = Sheet.Cells( _
        Sheet_RowNumberByTitle(Sheet, Col_A, "Option1"), Col_B).Value

End Sub

'�V�[�g�����̂悤�ɂ��Ă���
'
'   |A          |B      |C
'1  |�ݒ薼     |�l     |����
'2  |Option1    |�lA    |
'3  |Option2    |�lB    |
'
'�������Ă����ƁAOption1�ŒlA�̓��e���擾���邱�Ƃ��ł���
