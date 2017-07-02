'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    ControlAnchor Class
'ObjectName:    st_vba_ControlAnchor
'--------------------------------------------------
'Version:       2017/02/05
'--------------------------------------------------
Option Explicit

Enum HorizonAnchorType
    haLeft
    haRight
    haStretch
    haFloat
End Enum

Enum VerticalAnchorType
    vaTop
    vaBottom
    vaStretch
    vaFloat
End Enum

'--------------------------------------------------
'���ϐ��錾��
'--------------------------------------------------
Private TargetControl As Object
Private HorizonAnchor As HorizonAnchorType
Private VerticalAnchor As VerticalAnchorType

Private LeftMarginOrigin As Double
Private TopMarginOrigin As Double
Private RightMarginOrigin As Double
Private BottomMarginOrigin As Double
Private WidthOrigin As Double
Private HeightOrigin As Double
Private ParentWidthOrigin As Double
Private ParentHeightOrigin As Double

Sub Initialize(ByRef Control As Object, _
ByVal HorizonAnchorValue As HorizonAnchorType, _
ByVal RightOffset As Double, _
ByVal VerticalAnchorValue As VerticalAnchorType, _
ByVal BottomOffset As Double)

    Set TargetControl = Control
    HorizonAnchor = HorizonAnchorValue
    VerticalAnchor = VerticalAnchorValue

    LeftMarginOrigin = TargetControl.Left
    TopMarginOrigin = TargetControl.Top

    RightMarginOrigin = TargetControl.Parent.InsideWidth - _
        TargetControl.Left - TargetControl.Width + RightOffset
    '2016�̏ꍇ�AResize�\Form�ł����Ă�
    'RightOffset=0�ɂ���Ɨǂ��l�q
    '2013�̏ꍇ�AResize�\Form�̏ꍇ��
    'RightOffset=8�ɂ���Ɨǂ��l�q
    
    BottomMarginOrigin = TargetControl.Parent.InsideHeight - _
        TargetControl.Top - TargetControl.Height + BottomOffset
    
    WidthOrigin = TargetControl.Width
    HeightOrigin = TargetControl.Height
    
    ParentWidthOrigin = TargetControl.Parent.InsideWidth
    ParentHeightOrigin = TargetControl.Parent.InsideHeight

End Sub

Sub Layout()
    Select Case HorizonAnchor
    Case HorizonAnchorType.haRight
        TargetControl.Left = _
            TargetControl.Parent.InsideWidth - _
            RightMarginOrigin - TargetControl.Width
    Case HorizonAnchorType.haStretch
        TargetControl.Width = _
            MaxValue(1, _
                TargetControl.Parent.InsideWidth - _
                RightMarginOrigin - TargetControl.Left)
    Case HorizonAnchorType.haFloat
        TargetControl.Left = _
            (LeftMarginOrigin * TargetControl.Parent.InsideWidth) / ParentWidthOrigin
        TargetControl.Width = _
            (WidthOrigin * TargetControl.Parent.InsideWidth) / ParentWidthOrigin
    End Select
    
    Select Case VerticalAnchor
    Case VerticalAnchorType.vaBottom
        TargetControl.Top = _
            TargetControl.Parent.InsideHeight - _
            BottomMarginOrigin - TargetControl.Height
    Case VerticalAnchorType.vaStretch
        TargetControl.Height = _
            MaxValue(1, _
                TargetControl.Parent.InsideHeight - _
                BottomMarginOrigin - TargetControl.Top)
    Case VerticalAnchorType.vaFloat
        TargetControl.Top = _
            (TopMarginOrigin * TargetControl.Parent.InsideHeight) / ParentHeightOrigin
        TargetControl.Height = _
            (HeightOrigin * TargetControl.Parent.InsideHeight) / ParentHeightOrigin
    End Select
End Sub

'--------------------------------------------------
'������
'�� ver 2015/07/24
'�E �쐬
'�� ver 2017/02/05
'�E Form��̃R���g���[���ł� Parent.Width �����������삷�邪
'   MultiPage��̃R���g���[���ł� �G���[�ɂȂ�̂�
'   Parent.InsideWidth/InsideHeight �ɒu���������B
'--------------------------------------------------

