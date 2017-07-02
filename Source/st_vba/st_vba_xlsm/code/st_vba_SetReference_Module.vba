'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    SetReference Module
'ObjectName:    st_vba_SetReference
'--------------------------------------------------
'Version:       2017/03/19
'--------------------------------------------------
'   ・  参照設定を行うためのModule
'   ・  st_vba_Coreにこのコードを入れておくと
'       コンパイルエラーで動作させにくく
'       別Moduleにしておくと参照設定追加が楽になる
'--------------------------------------------------
Option Explicit


'----------------------------------------
'◆参照設定追加
'----------------------------------------

'----------------------------------------
'◇st_vba_Coreに必要
'----------------------------------------

'----------------------------------------
'・Microsoft Scripting Runtime
'----------------------------------------
'   ・  FSO:FileSystemObjectを使用するのに必要
'----------------------------------------
Sub ReferenceAdd_ScriptingRuntime(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\system32\scrrun.dll")
End Sub

Sub Run_ReferenceAdd_ScriptingRuntime()
    Call ReferenceAdd_ScriptingRuntime(ThisWorkbook)
End Sub

'----------------------------------------
'・Windows Script Host Object Model
'----------------------------------------
'   ・  WshShellを使用するのに必要
'----------------------------------------
Sub ReferenceAdd_WshObjectModel(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\system32\wshom.ocx")
End Sub

Sub Run_ReferenceAdd_WshObjectModel()
    Call ReferenceAdd_WshObjectModel(ThisWorkbook)
End Sub

'----------------------------------------
'・Microsoft Forms 2.0 Object Library
'----------------------------------------
'   ・  ComboBox に必要
'----------------------------------------
Sub ReferenceAdd_Form2_0(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\System32\FM20.DLL")
End Sub

Sub Run_ReferenceAdd_Form2_0()
    Call ReferenceAdd_Form2_0(ThisWorkbook)
End Sub

'----------------------------------------
'・Microsoft AxtiveX Data Objects 6.1 Library
'----------------------------------------
'   ・  ADODB.Streamを使用するのに必要
'----------------------------------------
Sub ReferenceAdd_ADO_6_1(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files\Common Files\System\ado\msado15.dll")
End Sub

Sub Run_ReferenceAdd_ADO_6_1()
    Call ReferenceAdd_ADO_6_1(ThisWorkbook)
End Sub

'----------------------------------------
'◇st_vba_IEに必要
'----------------------------------------

'----------------------------------------
'・ Microsoft Internet Controls
'----------------------------------------
'   ・  InternetExplorer に必要
'----------------------------------------
Sub ReferenceAdd_InternetControls(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\System32\ieframe.dll")
End Sub

Sub Run_ReferenceAdd_InternetControls()
    Call ReferenceAdd_InternetControls(ThisWorkbook)
End Sub


'----------------------------------------
'◇使用頻度低
'----------------------------------------

'----------------------------------------
'・Microsoft Windows Common Controls 6.0 (SP6)
'----------------------------------------
'   ・  ListViewのために必要
'   ・  ただし64bit版ExcelではListViewが使えないので
'       ほとんど必要ない
'----------------------------------------
Sub ReferenceAdd_CommonControls(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\System32\MSCOMCTL.OCX")
End Sub

Sub Run_ReferenceAdd_CommonControls()
    Call ReferenceAdd_CommonControls(ThisWorkbook)
End Sub

'----------------------------------------
'・Microsoft Visual Basic for Applications Extensibility 5.3
'----------------------------------------
'   ・  VBE(VBEditor)の操作に必要
'----------------------------------------
Sub ReferenceAdd_VBAExtensibility(Book As Workbook)

#If VBA7 And Win64 Then
    '64bit版Windows
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB")
#Else
    '32bit版Windows
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB")
#End If

End Sub

Sub Run_ReferenceAdd_VBAExtensibility()
    Call ReferenceAdd_VBAExtensibility(ThisWorkbook)
End Sub

'----------------------------------------
'・Microsoft AxtiveX Data Objects 2.8 Library
'----------------------------------------
'   ・  ADODB.Streamを使用するのに必要
'   ・  通常は 6.1 を使うのでこちらは使用しない
'----------------------------------------
Sub ReferenceAdd_ADO_2_8(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files\Common Files\System\ado\msado28.tlb")
End Sub

Sub Run_ReferenceAdd_ADO_2_8()
    Call ReferenceAdd_ADO_2_8(ThisWorkbook)
End Sub






