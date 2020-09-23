Attribute VB_Name = "modMousePositions"
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'   Programmed by: Nick Smith (age 19)
'   Website: http://www.spliff.wideboys.co.uk
'   PlanetSourceCode ID: Nick Spliff
'   Email: npsmith82@hotmail.com
'
'   You may use this code as freely as you like
'   because it was originally gathered from the
'   public domain.
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

'Place the following in a 1 m/s timer and change values to correspond with your projects.

'If MouseHasLeftObject(Form1, lblEMAIL) = False Then
'    If lblEMAIL.FontUnderline = False Then
'        lblEMAIL.ForeColor = vbBlue
'        lblEMAIL.FontUnderline = True
'    End If
'End If

'If MouseHasLeftObject(Form1, lblEMAIL) = True Then
'    If lblEMAIL.FontUnderline = True Then
'        lblEMAIL.ForeColor = vbBlack
'        lblEMAIL.FontUnderline = False
'     End If
'End If

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long 'Used for getting the cursor position
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Enum WhichCoordinate
    X
    Y
End Enum

Public Function MousePosition(ChosenCoordinate As WhichCoordinate) As Long
Dim CursorPos As POINTAPI
    GetCursorPos CursorPos
    If ChosenCoordinate = X Then MousePosition = CursorPos.X
    If ChosenCoordinate = Y Then MousePosition = CursorPos.Y
End Function

Public Function MouseHasLeftObject(Frm As Form, AnyFormObject As Object) As Boolean
Dim ObjectTop As Integer
Dim ObjectLeft As Integer

    'If the form has a border it completely phucks up the top and left positions
    'of every object displayed on the form, so we must account for the change.
    
    'TIP:   You must try and keep the objects out of containers such as picture boxes
    '       or frames, as they will further phuck up the top and left positions.
    
    Select Case Frm.BorderStyle
        Case 0 'No Border
            ObjectTop = AnyFormObject.Top
            ObjectLeft = AnyFormObject.Left
        Case 1 'Fixed Single
            ObjectTop = AnyFormObject.Top + (23 * Screen.TwipsPerPixelY)
            ObjectLeft = AnyFormObject.Left + (3 * Screen.TwipsPerPixelY)
        Case 2 'Sizable
            ObjectTop = AnyFormObject.Top + (24 * Screen.TwipsPerPixelY)
            ObjectLeft = AnyFormObject.Left + (4 * Screen.TwipsPerPixelY)
        Case 3 'Fixed Dialog
            ObjectTop = AnyFormObject.Top + (23 * Screen.TwipsPerPixelY)
            ObjectLeft = AnyFormObject.Left + (3 * Screen.TwipsPerPixelY)
        Case 4 'Fixed ToolWindow
            ObjectTop = AnyFormObject.Top + (23 * Screen.TwipsPerPixelY)
            ObjectLeft = AnyFormObject.Left + (3 * Screen.TwipsPerPixelY)
        Case 5 'Sizable ToolWindow
            ObjectTop = AnyFormObject.Top + (24 * Screen.TwipsPerPixelY)
            ObjectLeft = AnyFormObject.Left + (4 * Screen.TwipsPerPixelY)
    End Select

    'Check mouse to the LEFT
    If (MousePosition(X) * Screen.TwipsPerPixelX) < (Frm.Left + ObjectLeft) Then
        MouseHasLeftObject = True
        'Call MouseOut
    Else
        MouseHasLeftObject = False
        'Call MouseIn 'Mouse IS within the area.
    End If
    
    If (MousePosition(X) * Screen.TwipsPerPixelX) > (Frm.Left + ObjectLeft + AnyFormObject.Width) Then MouseHasLeftObject = True    'Check mouse to the RIGHT
    If (MousePosition(Y) * Screen.TwipsPerPixelY) < (Frm.Top + ObjectTop) Then MouseHasLeftObject = True                            'Check mouse to the TOP
    If (MousePosition(Y) * Screen.TwipsPerPixelY) > (Frm.Top + ObjectTop + AnyFormObject.Height) Then MouseHasLeftObject = True     'Check mouse to the BOTTOM
End Function
