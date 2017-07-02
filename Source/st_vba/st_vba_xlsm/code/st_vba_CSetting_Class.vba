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

'シートを次のようにしておく
'
'   |A          |B      |C
'1  |設定名     |値     |説明
'2  |Option1    |値A    |
'3  |Option2    |値B    |
'
'こうしておくと、Option1で値Aの内容を取得することができる
