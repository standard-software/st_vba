'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    ControlAnchor Class
'ObjectName:    st_vba_ControlAnchor
'--------------------------------------------------
'Version:       2015/07/24
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
'Å°ïœêîêÈåæìô
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
VerticalAnchorValue As VerticalAnchorType, _
ByVal BottomOffset As Double)

    Set TargetControl = Control
    HorizonAnchor = HorizonAnchorValue
    VerticalAnchor = VerticalAnchorValue

    LeftMarginOrigin = TargetControl.Left
    TopMarginOrigin = TargetControl.Top

    RightMarginOrigin = TargetControl.Parent.Width - _
        TargetControl.Left - TargetControl.Width + RightOffset
    
    BottomMarginOrigin = TargetControl.Parent.Height - _
        TargetControl.Top - TargetControl.Height + BottomOffset
    
    WidthOrigin = TargetControl.Width
    HeightOrigin = TargetControl.Height
    
    ParentWidthOrigin = TargetControl.Parent.Width
    ParentHeightOrigin = TargetControl.Parent.Width

End Sub

Sub Layout()
    Select Case HorizonAnchor
    Case HorizonAnchorType.haRight
        TargetControl.Left = _
            TargetControl.Parent.Width - _
            RightMarginOrigin - TargetControl.Width
    Case HorizonAnchorType.haStretch
        TargetControl.Width = _
            MaxValue(1, _
                TargetControl.Parent.Width - _
                RightMarginOrigin - TargetControl.Left)
    Case HorizonAnchorType.haFloat
        TargetControl.Left = _
            (LeftMarginOrigin * TargetControl.Parent.Width) / ParentWidthOrigin
        TargetControl.Width = _
            (WidthOrigin * TargetControl.Parent.Width) / ParentWidthOrigin
    End Select
    
    Select Case VerticalAnchor
    Case VerticalAnchorType.vaBottom
        TargetControl.Top = _
            TargetControl.Parent.Height - _
            BottomMarginOrigin - TargetControl.Height
    Case VerticalAnchorType.vaStretch
        TargetControl.Height = _
            MaxValue(1, _
                TargetControl.Parent.Height - _
                BottomMarginOrigin - TargetControl.Top)
    Case VerticalAnchorType.vaFloat
        TargetControl.Top = _
            (TopMarginOrigin * TargetControl.Parent.Height) / ParentHeightOrigin
        TargetControl.Height = _
            (HeightOrigin * TargetControl.Parent.Height) / ParentHeightOrigin
    End Select
End Sub


