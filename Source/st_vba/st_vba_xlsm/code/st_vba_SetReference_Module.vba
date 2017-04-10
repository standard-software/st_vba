'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    SetReference Module
'ObjectName:    st_vba_SetReference
'--------------------------------------------------
'Version:       2017/03/19
'--------------------------------------------------
'   �E  �Q�Ɛݒ���s�����߂�Module
'   �E  st_vba_Core�ɂ��̃R�[�h�����Ă�����
'       �R���p�C���G���[�œ��삳���ɂ���
'       ��Module�ɂ��Ă����ƎQ�Ɛݒ�ǉ����y�ɂȂ�
'--------------------------------------------------
Option Explicit


'----------------------------------------
'���Q�Ɛݒ�ǉ�
'----------------------------------------

'----------------------------------------
'��st_vba_Core�ɕK�v
'----------------------------------------

'----------------------------------------
'�EMicrosoft Scripting Runtime
'----------------------------------------
'   �E  FSO:FileSystemObject���g�p����̂ɕK�v
'----------------------------------------
Sub ReferenceAdd_ScriptingRuntime(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\system32\scrrun.dll")
End Sub

Sub Run_ReferenceAdd_ScriptingRuntime()
    Call ReferenceAdd_ScriptingRuntime(ThisWorkbook)
End Sub

'----------------------------------------
'�EWindows Script Host Object Model
'----------------------------------------
'   �E  WshShell���g�p����̂ɕK�v
'----------------------------------------
Sub ReferenceAdd_WshObjectModel(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\system32\wshom.ocx")
End Sub

Sub Run_ReferenceAdd_WshObjectModel()
    Call ReferenceAdd_WshObjectModel(ThisWorkbook)
End Sub

'----------------------------------------
'�EMicrosoft Forms 2.0 Object Library
'----------------------------------------
'   �E  ComboBox �ɕK�v
'----------------------------------------
Sub ReferenceAdd_Form2_0(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\System32\FM20.DLL")
End Sub

Sub Run_ReferenceAdd_Form2_0()
    Call ReferenceAdd_Form2_0(ThisWorkbook)
End Sub

'----------------------------------------
'�EMicrosoft AxtiveX Data Objects 6.1 Library
'----------------------------------------
'   �E  ADODB.Stream���g�p����̂ɕK�v
'----------------------------------------
Sub ReferenceAdd_ADO_6_1(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files\Common Files\System\ado\msado15.dll")
End Sub

Sub Run_ReferenceAdd_ADO_6_1()
    Call ReferenceAdd_ADO_6_1(ThisWorkbook)
End Sub

'----------------------------------------
'��st_vba_IE�ɕK�v
'----------------------------------------

'----------------------------------------
'�E Microsoft Internet Controls
'----------------------------------------
'   �E  InternetExplorer �ɕK�v
'----------------------------------------
Sub ReferenceAdd_InternetControls(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\System32\ieframe.dll")
End Sub

Sub Run_ReferenceAdd_InternetControls()
    Call ReferenceAdd_InternetControls(ThisWorkbook)
End Sub


'----------------------------------------
'���g�p�p�x��
'----------------------------------------

'----------------------------------------
'�EMicrosoft Windows Common Controls 6.0 (SP6)
'----------------------------------------
'   �E  ListView�̂��߂ɕK�v
'   �E  ������64bit��Excel�ł�ListView���g���Ȃ��̂�
'       �قƂ�ǕK�v�Ȃ�
'----------------------------------------
Sub ReferenceAdd_CommonControls(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Windows\System32\MSCOMCTL.OCX")
End Sub

Sub Run_ReferenceAdd_CommonControls()
    Call ReferenceAdd_CommonControls(ThisWorkbook)
End Sub

'----------------------------------------
'�EMicrosoft Visual Basic for Applications Extensibility 5.3
'----------------------------------------
'   �E  VBE(VBEditor)�̑���ɕK�v
'----------------------------------------
Sub ReferenceAdd_VBAExtensibility(Book As Workbook)

#If VBA7 And Win64 Then
    '64bit��Windows
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB")
#Else
    '32bit��Windows
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB")
#End If

End Sub

Sub Run_ReferenceAdd_VBAExtensibility()
    Call ReferenceAdd_VBAExtensibility(ThisWorkbook)
End Sub

'----------------------------------------
'�EMicrosoft AxtiveX Data Objects 2.8 Library
'----------------------------------------
'   �E  ADODB.Stream���g�p����̂ɕK�v
'   �E  �ʏ�� 6.1 ���g���̂ł�����͎g�p���Ȃ�
'----------------------------------------
Sub ReferenceAdd_ADO_2_8(Book As Workbook)
    Call Book.VBProject.References.AddFromFile( _
        "C:\Program Files\Common Files\System\ado\msado28.tlb")
End Sub

Sub Run_ReferenceAdd_ADO_2_8()
    Call ReferenceAdd_ADO_2_8(ThisWorkbook)
End Sub






