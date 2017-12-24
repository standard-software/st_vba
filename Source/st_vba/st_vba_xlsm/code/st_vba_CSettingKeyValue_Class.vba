'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    Setting KeyValue Class
'ObjectName:    st_vba_CSettingKeyValue
'--------------------------------------------------
'Version:       2017/12/24
'--------------------------------------------------
Option Explicit

Private m_Sheet As Worksheet

Public Sub Initialize(ByVal Sheet As Worksheet)
    Set m_Sheet = Sheet
End Sub

Public Property Get Col_Key() As Long
    Col_Key = Col_A
End Property

Public Property Get Col_Value() As Long
    Col_Value = Col_B
End Property

'ÉLÅ[óÒ/ílóÒÇÃèÍçá
Public Function ReadValue( _
ByVal Key As String) As String
    Dim Row As Long
    Row = RowByTitle(m_Sheet, Col_Key, Key)
    If Row <> 0 Then
        ReadValue = m_Sheet.Cells(Row, Col_Value).Value
    Else
        ReadValue = ""
    End If
End Function

Public Sub WriteValue( _
ByVal Key As String, ByVal Value As String)
    Dim Row As Long
    Row = RowByTitle(m_Sheet, Col_Key, Key)
    If Row <> 0 Then
        Call CellText(m_Sheet.Cells(Row, Col_Value), Value)
    Else
        Row = DataLastRow(m_Sheet) + 1
        Call CellText(m_Sheet.Cells(Row, Col_Key), Key)
        Call CellText(m_Sheet.Cells(Row, Col_Value), Value)
    End If
End Sub

