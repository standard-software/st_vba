'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    SheetData Class
'ObjectName:    st_vba_CSheetData_Sample
'--------------------------------------------------
'Version:       2017/03/19
'--------------------------------------------------
'   �E  �V�[�g�f�[�^�����Ă�����
'       �V�[�g�񖼒�`�Ȃǂ���������ăR�[�h�������₷���ł�
'   �E  �����V�[�g�̎�ނ��ƂɃN���X�𑝂₵�Ďg���Ă�������
'       �N���X���� st_vba_CSheetData_Sample ����
'       CSheetData_XXX �̂悤�ɕύX���Ďg���Ă�������
'   �E  Col_First / Col_Second �̃v���p�e�B����ύX������ǉ�����
'       Col_Name / Col_ID / Col_Mail �Ȃǂ����Ƃ悢�ł�
'       Col_XXX �� �Œ�l��Ԃ��Ă������ł����A
'       Initialize_Col ���\�b�h�Ƃ���
'       ��^�C�g�����̂��瓮�I�Ɏ擾����Ƃ�������������܂�
'--------------------------------------------------
Option Explicit

'----------------------------------------
'������
'----------------------------------------

Private m_Book As Workbook
Private m_Sheet As Worksheet

Private m_Col_Third As Long
Private m_Row_DataTitle As Long
Private m_Row_Start As Long
Private m_Row_End As Long

Private m_Row_Header_Start As Long
Private m_Row_Header_End As Long

Public Sub Initialize()
    Set m_Book = Nothing
    Set m_Sheet = Nothing
    
    m_Row_Header_Start = 2
    m_Row_Header_End = 8
    
    m_Row_DataTitle = 10
    m_Row_Start = 11
    m_Row_End = 0
End Sub

Public Sub Initialize_Col()
    m_Col_Third = ColumnNumberByTitle( _
        m_Sheet, m_Row_DataTitle, "3�Ԗڂ̗�")
End Sub

Public Sub Initialize_Row_EndIndex()
    Call Assert(IsNotNothing(m_Sheet), "Error:Sheet is Nothing")
    m_Row_End = DataLastRow(m_Sheet)
End Sub

'----------------------------------------
'Book/Sheet���
'----------------------------------------
Public Property Let Book(ByVal Value As Workbook)
    Set m_Book = Value
End Property

Public Property Get Book() As Workbook
    Set Book = m_Book
End Property

Public Property Let Sheet(ByVal Value As Worksheet)
    Set m_Sheet = Value
End Property

Public Property Get Sheet() As Worksheet
    Set Sheet = m_Sheet
End Property

'----------------------------------------
'����
'----------------------------------------
Public Property Get Col_First() As Long
    Col_First = Col_A
End Property
        
Public Property Get Col_Second() As Long
    Col_Second = Col_B
End Property

Public Property Get Col_Third() As Long
    Col_Third = m_Col_Third
End Property


'----------------------------------------
'�s���
'----------------------------------------
Public Property Get Row_Start() As Long
    Row_Start = m_Row_Start
End Property

Public Property Get Row_End() As Long
    Row_End = m_Row_End
End Property

'----------------------------------------
'�w�b�_�[����
'----------------------------------------
Public Property Get Col_Header_First() As Long
    Col_Header_First = Col_A
End Property

Public Property Get Col_Header_Second() As Long
    Col_Header_Second = Col_B
End Property


'----------------------------------------
'�w�b�_�[�s���
'----------------------------------------
Public Property Let Row_Header_Start(ByVal Value As Long)
    m_Row_Header_Start = Value
End Property

Public Property Get Row_Header_Start() As Long
    Row_Header_Start = m_Row_Header_Start
End Property

Public Property Let Row_Header_End(ByVal Value As Long)
    m_Row_Header_End = Value
End Property

Public Property Get Row_Header_End() As Long
    Row_Header_End = m_Row_Header_End
End Property

'----------------------------------------
'�w�b�_�[�Z�����
'----------------------------------------
Public Property Get SelectDateCellAddress() As String
    SelectDateCellAddress = "A1"
End Property

