'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    Setting KeyValue Class
'ObjectName:    st_vba_CSettingGroupKeyValue
'--------------------------------------------------
'Version:       2017/12/24
'--------------------------------------------------
Option Explicit

Private m_Sheet As Worksheet

Public Sub Initialize(ByVal Sheet As Worksheet)
    Set m_Sheet = Sheet
End Sub

'----------------------------------------
'列情報
'----------------------------------------
Public Property Get Col_Group() As Long
    Col_Group = Col_A
End Property

Public Property Get Col_Key() As Long
    Col_Key = Col_B
End Property

Public Property Get Col_Value() As Long
    Col_Value = Col_C
End Property

'グループ列/キー列/値列の場合
Public Function ReadValue( _
ByVal Group As String, ByVal Key As String) As String
    Dim Row As Long
    Row = RowByGroupTitle(m_Sheet, Col_Group, Col_Key, Group, Key)
    If Row <> 0 Then
        ReadValue = m_Sheet.Cells(Row, Col_Value).Value
    Else
        ReadValue = ""
    End If
End Function

Public Sub WriteValue(ByVal Group As String, ByVal Key As String, ByVal Value As String)
    Dim Row As Long
    Row = RowByGroupTitle(m_Sheet, Col_Group, Col_Key, Group, Key)
    If Row <> 0 Then
        Call CellText(m_Sheet.Cells(Row, Col_Value), Value)
    Else
        Row = DataLastRow(m_Sheet) + 1
        Call CellText(m_Sheet.Cells(Row, Col_Group), Group)
        Call CellText(m_Sheet.Cells(Row, Col_Key), Key)
        Call CellText(m_Sheet.Cells(Row, Col_Value), Value)
    End If
End Sub



