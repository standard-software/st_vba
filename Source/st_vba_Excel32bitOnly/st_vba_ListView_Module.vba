'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    ListView Module
'ObjectName:    st_vba_ListView
'--------------------------------------------------
'Version:       2017/03/19
'--------------------------------------------------
'   ・  ListViewをコントロールするためのモジュール
'   ・  ListViewは32bit版Excel(2013や2016時点で)でしか動かず
'       今後の追加も考えられないため、
'       64bit対応している st_vba_Core からは分離している
'--------------------------------------------------
Option Explicit


'----------------------------------------
'◆ListView
'----------------------------------------

'----------------------------------------
'・ListViewの選択項目個数
'----------------------------------------
Public Function ListView_SelectedItemCount(ByVal ListView As ListView) As Long
    Dim Result As Long: Result = 0
    Dim I As Long
    For I = 0 To ListView.ListItems.Count - 1
        If ListView.ListItems(I + 1).Selected Then
            Result = Result + 1
        End If
    Next
    ListView_SelectedItemCount = Result
End Function

'----------------------------------------
'・ListViewのチェック項目個数
'----------------------------------------
Public Function ListView_CheckedItemCount(ByVal ListView As ListView) As Long
    Dim Result As Long: Result = 0
    Dim I As Long
    For I = 0 To ListView.ListItems.Count - 1
        If ListView.ListItems(I + 1).Checked Then
            Result = Result + 1
        End If
    Next
    ListView_CheckedItemCount = Result
End Function

'----------------------------------------
'・ListView 全て選択
'----------------------------------------
Sub ListView_SelectAll(ListView As ListView, _
SelectValue As Boolean)
    Dim I As Long
    For I = 0 To ListView.ListItems.Count - 1
        ListView.ListItems(I + 1).Selected = SelectValue
    Next
End Sub

'----------------------------------------
'・ListView 全てチェック
'----------------------------------------
Sub ListView_CheckAll(ListView As ListView, _
CheckValue As Boolean)
    Dim I As Long
    For I = 0 To ListView.ListItems.Count - 1
        ListView.ListItems(I + 1).Checked = CheckValue
    Next
End Sub

'----------------------------------------
'・ListView 選択チェック
'----------------------------------------
Sub ListView_CheckSelectedItem(ListView As ListView, _
CheckValue As Boolean)
    Dim I As Long
    For I = 0 To ListView.ListItems.Count - 1
        If ListView.ListItems(I + 1).Selected Then
            ListView.ListItems(I + 1).Checked = CheckValue
        End If
    Next
End Sub

'----------------------------------------
'・ListView 選択項目が全てチェックされているかどうか確認
'----------------------------------------
Public Function ListView_IsCheckSelectedItem(ByVal ListView As ListView) As Boolean
    Dim Result As Boolean: Result = True
    Dim I As Long
    For I = 0 To ListView.ListItems.Count - 1
        If ListView.ListItems(I + 1).Selected Then
            If ListView.ListItems(I + 1).Checked = False Then
                Result = False
                Exit For
            End If
        End If
    Next
    ListView_IsCheckSelectedItem = Result
End Function


'----------------------------------------
'・複数選択時の同時チェック切り替え
'----------------------------------------
'   ・  例:
'       次のように使うとよい
'       Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'           Call ListView_MultiSelectChecked(ListView1, Item)
'       End Sub
'----------------------------------------
Sub ListView_MultiSelectChecked(ListView As ListView, _
CheckedItem As ListItem)
    Dim I As Long
    If CheckedItem.Selected Then
        'チェックした項目が選択されているなら
        '他のすべての選択項目もチェックをあわせる
        For I = 0 To ListView.ListItems.Count - 1
            If ListView.ListItems(I + 1).Selected Then
                ListView.ListItems(I + 1).Checked = CheckedItem.Checked
            End If
        Next
    Else
        'チェックした項目が選択されていないなら
        '選択を解除して、その項目を選択する
        For I = 0 To ListView.ListItems.Count - 1
            ListView.ListItems(I + 1).Selected = False
        Next
        CheckedItem.Selected = True
    End If
End Sub

'----------------------------------------
'・キーでサーチするIndexOf
'----------------------------------------
Public Function ListView_IndexOfKey(ByVal ListView As ListView, _
ByVal SearchKey As String, Optional StartIndex As Long = 0) As Long
    Dim Result As Long: Result = -1
    Dim I As Long
    For I = StartIndex To ListView.ListItems.Count - 1
        If ListView.ListItems(I + 1).Key = SearchKey Then
            Result = I
            Exit For
        End If
    Next
    ListView_IndexOfKey = Result
End Function


