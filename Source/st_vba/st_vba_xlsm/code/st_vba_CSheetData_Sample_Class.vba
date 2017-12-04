'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    SheetData Class
'ObjectName:    st_vba_CSheetData_Sample
'--------------------------------------------------
'Version:       2017/03/19
'--------------------------------------------------
'   ・  シートデータを入れておくと
'       シート列名定義などが分離されてコードが書きやすいです
'   ・  扱うシートの種類ごとにクラスを増やして使ってください
'       クラス名を st_vba_CSheetData_Sample から
'       CSheetData_XXX のように変更して使ってください
'   ・  Col_First / Col_Second のプロパティ名を変更したり追加して
'       Col_Name / Col_ID / Col_Mail などを作るとよいです
'       Col_XXX は 固定値を返してもいいですし、
'       Initialize_Col メソッドとして
'       列タイトル名称から動的に取得するというやり方もあります
'--------------------------------------------------
Option Explicit

'----------------------------------------
'初期化
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
        m_Sheet, m_Row_DataTitle, "3番目の列")
End Sub

Public Sub Initialize_Row_EndIndex()
    Call Assert(IsNotNothing(m_Sheet), "Error:Sheet is Nothing")
    m_Row_End = DataLastRow(m_Sheet)
End Sub

'----------------------------------------
'Book/Sheet情報
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
'列情報
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
'行情報
'----------------------------------------
Public Property Get Row_Start() As Long
    Row_Start = m_Row_Start
End Property

Public Property Get Row_End() As Long
    Row_End = m_Row_End
End Property

'----------------------------------------
'ヘッダー列情報
'----------------------------------------
Public Property Get Col_Header_First() As Long
    Col_Header_First = Col_A
End Property

Public Property Get Col_Header_Second() As Long
    Col_Header_Second = Col_B
End Property


'----------------------------------------
'ヘッダー行情報
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
'ヘッダーセル情報
'----------------------------------------
Public Property Get SelectDateCellAddress() As String
    SelectDateCellAddress = "A1"
End Property

