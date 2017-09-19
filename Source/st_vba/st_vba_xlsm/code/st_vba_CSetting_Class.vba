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

'シートを次のようにしておく
'
'   |A          |B      |C
'1  |設定名     |値     |説明
'2  |Option1    |値A    |
'3  |Option2    |値B    |
'
'こうしておくと、Setting.Read("Option1")で値Aの内容を取得することができる

Public Function Read(ByVal Key As String) As String
    Read = m_Sheet.Cells( _
        Sheet_RowNumberByTitle(m_Sheet, Col_A, Key), _
        Col_B).Value
End Function
