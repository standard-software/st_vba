'--------------------------------------------------
'st_vba
'--------------------------------------------------
'ModuleName:    ControlSplitter Class
'ObjectName:    st_vba_ControlSplitter
'--------------------------------------------------
'Version:       2015/07/24
'--------------------------------------------------
Option Explicit

Enum SplitterType
    Vertical
    Horizon
End Enum

Private WithEvents FSplitterImageControl As MSForms.Image
Private FSplitterType As SplitterType

Private FSplitterMouseDownFlag As Boolean
Private FLeftTopControls() As MSForms.Control
Private FRightBottomControls() As MSForms.Control
Private FLeftTopMinSize As Double
Private FRightBottomMinSize As Double

Public Sub AddControlLeftTop(ByRef Control As MSForms.Control)
    Call ArrayAdd(FLeftTopControls, Control)
End Sub

Public Sub AddControlRightBottom(ByRef Control As MSForms.Control)
    Call ArrayAdd(FRightBottomControls, Control)
End Sub

Public Sub Initialize( _
ByRef SplitterImageControl As MSForms.Image, _
ByRef SplitterTypeValue As SplitterType, _
ByVal LeftTopMinSize As Double, _
ByVal RightBottomMinSize As Double)
    Set FSplitterImageControl = SplitterImageControl
    FSplitterImageControl.BorderStyle = fmBorderStyleNone
    FSplitterType = SplitterTypeValue
    Select Case FSplitterType
    Case SplitterType.Vertical
        FSplitterImageControl.MousePointer = fmMousePointerSizeWE
    Case SplitterType.Horizon
        FSplitterImageControl.MousePointer = fmMousePointerSizeNS
    End Select
    
    FLeftTopMinSize = LeftTopMinSize
    FRightBottomMinSize = RightBottomMinSize
    
    FSplitterMouseDownFlag = False
End Sub

Private Sub FSplitterImageControl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FSplitterMouseDownFlag = True
    FSplitterImageControl.BackColor = &H80000010
End Sub

Private Sub FSplitterImageControl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FSplitterMouseDownFlag = False
    FSplitterImageControl.BackColor = &H8000000F
End Sub

Private Sub FSplitterImageControl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If FSplitterMouseDownFlag Then
        If CanLayoutUpdate( _
            FSplitterImageControl.Left + X, _
            FSplitterImageControl.Top + Y) Then
                    
            Call LayoutUpdate( _
                FSplitterImageControl.Left + X, _
                FSplitterImageControl.Top + Y)
        End If
    End If
End Sub

Public Function CanLayoutUpdate( _
ByVal NewSplitterLeft As Long, _
ByVal NewSplitterTop As Long) As Boolean

    Dim Result As Boolean: Result = True
    
    Dim NewControlLeft As Long
    Dim NewControlWidth As Long
    Dim RightBuffer As Long
    Dim NewControlTop As Long
    Dim NewControlHeight As Long
    Dim BottomBuffer As Long
    
    Dim I As Long
    Select Case FSplitterType
    Case SplitterType.Vertical
        For I = 0 To ArrayCount(FLeftTopControls) - 1
            NewControlWidth = NewSplitterLeft - FLeftTopControls(I).Left
            
            If NewControlWidth <= FLeftTopMinSize Then
                Result = False
                Exit For
            End If
        Next
        For I = 0 To ArrayCount(FRightBottomControls) - 1
            RightBuffer = FRightBottomControls(I).Left + FRightBottomControls(I).Width
            NewControlLeft = NewSplitterLeft + FSplitterImageControl.Width
            NewControlWidth = RightBuffer - NewControlLeft

            If NewControlWidth <= FRightBottomMinSize Then
                Result = False
                Exit For
            End If
        Next
        
    Case SplitterType.Horizon
        For I = 0 To ArrayCount(FLeftTopControls) - 1
            NewControlHeight = NewSplitterTop - FLeftTopControls(I).Top
            
            If NewControlHeight <= FLeftTopMinSize Then
                Result = False
                Exit For
            End If
        Next
        For I = 0 To ArrayCount(FRightBottomControls) - 1
            BottomBuffer = FRightBottomControls(I).Top + FRightBottomControls(I).Height
            NewControlTop = NewSplitterTop + FSplitterImageControl.Height
            NewControlHeight = BottomBuffer - NewControlTop
            
            If NewControlHeight <= FRightBottomMinSize Then
                Result = False
                Exit For
            End If
        Next
        
    End Select
    
    CanLayoutUpdate = Result
End Function

Public Sub LayoutUpdate( _
ByVal NewSplitterLeft As Long, _
ByVal NewSplitterTop As Long)

    Dim NewControlLeft As Long
    Dim NewControlWidth As Long
    Dim RightBuffer As Long
    Dim NewControlTop As Long
    Dim NewControlHeight As Long
    Dim BottomBuffer As Long
    
    Dim I As Long
    Select Case FSplitterType
    Case SplitterType.Vertical
        For I = 0 To ArrayCount(FLeftTopControls) - 1
            NewControlWidth = NewSplitterLeft - FLeftTopControls(I).Left
            
            FLeftTopControls(I).Width = NewControlWidth
        Next
        For I = 0 To ArrayCount(FRightBottomControls) - 1
            RightBuffer = FRightBottomControls(I).Left + FRightBottomControls(I).Width
            NewControlLeft = NewSplitterLeft + FSplitterImageControl.Width
            NewControlWidth = RightBuffer - NewControlLeft
            
            FRightBottomControls(I).Left = NewControlLeft
            FRightBottomControls(I).Width = NewControlWidth
        Next
        FSplitterImageControl.Left = NewSplitterLeft
        
    Case SplitterType.Horizon
        For I = 0 To ArrayCount(FLeftTopControls) - 1
            NewControlHeight = NewSplitterTop - FLeftTopControls(I).Top
            
            FLeftTopControls(I).Height = NewControlHeight
        Next
        For I = 0 To ArrayCount(FRightBottomControls) - 1
            BottomBuffer = FRightBottomControls(I).Top + FRightBottomControls(I).Height
            NewControlTop = NewSplitterTop + FSplitterImageControl.Height
            NewControlHeight = BottomBuffer - NewControlTop
            
            FRightBottomControls(I).Top = NewControlTop
            FRightBottomControls(I).Height = NewControlHeight
        Next
        FSplitterImageControl.Top = NewSplitterTop
        
    End Select
End Sub
