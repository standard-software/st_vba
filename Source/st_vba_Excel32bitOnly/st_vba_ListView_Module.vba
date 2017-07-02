'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    ListView Module
'ObjectName:    st_vba_ListView
'--------------------------------------------------
'Version:       2017/03/19
'--------------------------------------------------
'   �E  ListView���R���g���[�����邽�߂̃��W���[��
'   �E  ListView��32bit��Excel(2013��2016���_��)�ł���������
'       ����̒ǉ����l�����Ȃ����߁A
'       64bit�Ή����Ă��� st_vba_Core ����͕������Ă���
'--------------------------------------------------
Option Explicit


'----------------------------------------
'��ListView
'----------------------------------------

'----------------------------------------
'�EListView�̑I�����ڌ�
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
'�EListView�̃`�F�b�N���ڌ�
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
'�EListView �S�đI��
'----------------------------------------
Sub ListView_SelectAll(ListView As ListView, _
SelectValue As Boolean)
    Dim I As Long
    For I = 0 To ListView.ListItems.Count - 1
        ListView.ListItems(I + 1).Selected = SelectValue
    Next
End Sub

'----------------------------------------
'�EListView �S�ă`�F�b�N
'----------------------------------------
Sub ListView_CheckAll(ListView As ListView, _
CheckValue As Boolean)
    Dim I As Long
    For I = 0 To ListView.ListItems.Count - 1
        ListView.ListItems(I + 1).Checked = CheckValue
    Next
End Sub

'----------------------------------------
'�EListView �I���`�F�b�N
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
'�EListView �I�����ڂ��S�ă`�F�b�N����Ă��邩�ǂ����m�F
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
'�E�����I�����̓����`�F�b�N�؂�ւ�
'----------------------------------------
'   �E  ��:
'       ���̂悤�Ɏg���Ƃ悢
'       Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'           Call ListView_MultiSelectChecked(ListView1, Item)
'       End Sub
'----------------------------------------
Sub ListView_MultiSelectChecked(ListView As ListView, _
CheckedItem As ListItem)
    Dim I As Long
    If CheckedItem.Selected Then
        '�`�F�b�N�������ڂ��I������Ă���Ȃ�
        '���̂��ׂĂ̑I�����ڂ��`�F�b�N�����킹��
        For I = 0 To ListView.ListItems.Count - 1
            If ListView.ListItems(I + 1).Selected Then
                ListView.ListItems(I + 1).Checked = CheckedItem.Checked
            End If
        Next
    Else
        '�`�F�b�N�������ڂ��I������Ă��Ȃ��Ȃ�
        '�I�����������āA���̍��ڂ�I������
        For I = 0 To ListView.ListItems.Count - 1
            ListView.ListItems(I + 1).Selected = False
        Next
        CheckedItem.Selected = True
    End If
End Sub

'----------------------------------------
'�E�L�[�ŃT�[�`����IndexOf
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


