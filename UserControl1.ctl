VERSION 5.00
Begin VB.UserControl ActiveButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   MaskColor       =   &H00C0C0C0&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Image imDis 
      Height          =   495
      Left            =   1800
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imHot 
      Height          =   495
      Left            =   1200
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imDown 
      Height          =   495
      Left            =   600
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imUp 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "ActiveButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'Default Property Values:
Const m_def_BackStyle = 1
Const m_def_MaskColor = &HC0C0C0
Const m_def_Enabled = True
Const m_def_Style = 0
Const m_def_Value = 0
'Property Variables:
Dim m_MaskColor As OLE_COLOR
Dim m_ImageUp As Picture
Dim m_Enabled As Boolean
Dim m_ImageDown As Picture
Dim m_ImageHot As Picture
Dim m_ImageDisabled As Picture
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."

Public Enum VAL_U_P
    abUnPressed = 0
    abPressed = 1
End Enum
Private vval As VAL_U_P

Public Enum STYLE_B
    abCheckButton = 1
    abStandardButton = 0
End Enum
Private sstyle As STYLE_B

Public Enum BACKSTYLE_TO
    abTransparent = 0
    abOpaque = 1
End Enum
Private m_BackStyle As BACKSTYLE_TO

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    If m_Enabled = False Then
        Set UserControl.Picture = imDis.Picture
        Set UserControl.MaskPicture = imDis.Picture
    Else
        If vval = abUnPressed Then
            Set UserControl.Picture = imUp.Picture
            Set UserControl.MaskPicture = imUp.Picture
        ElseIf vval = abPressed Then
            Set UserControl.Picture = imDown.Picture
            Set UserControl.MaskPicture = imDown.Picture
        End If
    End If
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ImageDown() As Picture
    Set ImageDown = m_ImageDown
End Property

Public Property Set ImageDown(ByVal New_ImageDown As Picture)
    Set m_ImageDown = New_ImageDown
    Set imDown.Picture = New_ImageDown
    PropertyChanged "ImageDown"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ImageHot() As Picture
    Set ImageHot = m_ImageHot
End Property

Public Property Set ImageHot(ByVal New_ImageHot As Picture)
    Set m_ImageHot = New_ImageHot
    Set imHot.Picture = New_ImageHot
    PropertyChanged "ImageHot"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ImageDisabled() As Picture
    Set ImageDisabled = m_ImageDisabled
End Property

Public Property Set ImageDisabled(ByVal New_ImageDisabled As Picture)
    Set m_ImageDisabled = New_ImageDisabled
    Set imDis.Picture = New_ImageDisabled
    PropertyChanged "ImageDisabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Style() As STYLE_B
    Style = sstyle
End Property

Public Property Let Style(ByVal New_Style As STYLE_B)
    sstyle = New_Style
    PropertyChanged "Style"
    Set UserControl.Picture = imUp.Picture
    Set UserControl.MaskPicture = imUp.Picture
    vval = abUnPressed
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As VAL_U_P
    Value = vval
End Property

Public Property Let Value(ByVal New_Value As VAL_U_P)
    vval = New_Value
    PropertyChanged "Value"
    If vval = abPressed Then
        Set UserControl.Picture = imDown.Picture
        Set UserControl.MaskPicture = imDown.Picture
    ElseIf vval = abUnPressed Then
        Set UserControl.Picture = imUp.Picture
        Set UserControl.MaskPicture = imUp.Picture
    End If
    UserControl.Refresh
    
End Property

Private Sub UserControl_Click()
RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_Enabled = False Then Exit Sub
RaiseEvent MouseDown(Button, Shift, X, Y)
If Button = 1 Then
    If vval = abUnPressed Then
        Set UserControl.Picture = imDown.Picture
        Set UserControl.MaskPicture = imDown.Picture
    Else
        Set UserControl.Picture = imHot.Picture
        Set UserControl.MaskPicture = imHot.Picture
    End If
    UserControl.Refresh
End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this is most important part of control
'It detemines when mouse is over the control
' and when it goes out. This is made by using
' mouse capturing. Whenever mouse is captured
' to a window, that window will get all the
' mouse input.
Dim lp As POINTAPI
Dim ret As Long
Dim wn As Long

If m_Enabled = False Then Exit Sub
RaiseEvent MouseMove(Button, Shift, X, Y)

If Button = 1 Then
    wn = GetCapture()
    If (wn <> UserControl.hWnd) And (wn <> 0) Then
        'this suppose to happen when button is pressed and mouse is moved
        'over another active button control, so it will not change capture.
        If vval = abUnPressed Then
            Set UserControl.Picture = imUp.Picture
            Set UserControl.MaskPicture = imUp.Picture
        Else
            Set UserControl.Picture = imDown.Picture
            Set UserControl.MaskPicture = imDown.Picture
        End If
        UserControl.Refresh
        Exit Sub
    End If
    SetCapture UserControl.hWnd
    ret = GetCursorPos(lp)
    ScreenToClient UserControl.hWnd, lp
    lp.X = lp.X * Screen.TwipsPerPixelX
    lp.Y = lp.Y * Screen.TwipsPerPixelY
    If lp.X > 0 And lp.X < UserControl.Width Then
        If lp.Y > 0 And lp.Y < UserControl.Height Then
            If vval = abUnPressed Then
                Set UserControl.Picture = imDown.Picture
                Set UserControl.MaskPicture = imDown.Picture
            Else
                Set UserControl.Picture = imHot.Picture
                Set UserControl.MaskPicture = imHot.Picture
            End If
        Else
            If vval = abUnPressed Then
                Set UserControl.Picture = imUp.Picture
                Set UserControl.MaskPicture = imUp.Picture
            Else
                Set UserControl.Picture = imDown.Picture
                Set UserControl.MaskPicture = imDown.Picture
            End If
        End If
    Else
        If vval = abUnPressed Then
            Set UserControl.Picture = imUp.Picture
            Set UserControl.MaskPicture = imUp.Picture
        Else
            Set UserControl.Picture = imDown.Picture
            Set UserControl.MaskPicture = imDown.Picture
        End If
    End If
ElseIf Button = 0 Then
    If GetCapture() <> UserControl.hWnd Then SetCapture UserControl.hWnd
    ret = GetCursorPos(lp)
    ScreenToClient UserControl.hWnd, lp
    lp.X = lp.X * Screen.TwipsPerPixelX
    lp.Y = lp.Y * Screen.TwipsPerPixelY
    If lp.X > 0 And lp.X < UserControl.Width Then
        If lp.Y > 0 And lp.Y < UserControl.Height Then
            If vval = abUnPressed Then
                Set UserControl.Picture = imHot.Picture
                Set UserControl.MaskPicture = imHot.Picture
            End If
        Else
            ReleaseCapture
            If vval = abUnPressed Then
                Set UserControl.Picture = imUp.Picture
                Set UserControl.MaskPicture = imUp.Picture
            Else
                Set UserControl.Picture = imDown.Picture
                Set UserControl.MaskPicture = imDown.Picture
            End If
        End If
    Else
        ReleaseCapture
        If vval = abUnPressed Then
            Set UserControl.Picture = imUp.Picture
            Set UserControl.MaskPicture = imUp.Picture
        Else
            Set UserControl.Picture = imDown.Picture
            Set UserControl.MaskPicture = imDown.Picture
        End If
    End If
End If
UserControl.Refresh

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_Enabled = False Then Exit Sub
RaiseEvent MouseUp(Button, Shift, X, Y)
If Button = 1 Then
    'make sure we released mouse button inside the control
    If X > 0 And X < Extender.Width Then
        If Y > 0 And Y < Extender.Height Then
            If sstyle = abCheckButton Then
                If vval = abPressed Then vval = abUnPressed Else vval = abPressed
                If vval = abUnPressed Then
                    Set UserControl.Picture = imHot.Picture
                    Set UserControl.MaskPicture = imHot.Picture
                Else
                    Set UserControl.Picture = imDown.Picture
                    Set UserControl.MaskPicture = imDown.Picture
                End If
            Else
                Set UserControl.Picture = imHot.Picture
                Set UserControl.MaskPicture = imHot.Picture
            End If
        Else
            If vval = abUnPressed Then
                Set UserControl.Picture = imUp.Picture
                Set UserControl.MaskPicture = imUp.Picture
            Else
                Set UserControl.Picture = imDown.Picture
                Set UserControl.MaskPicture = imDown.Picture
            End If
        End If
    Else
        If vval = abUnPressed Then
            Set UserControl.Picture = imUp.Picture
            Set UserControl.MaskPicture = imUp.Picture
        Else
            Set UserControl.Picture = imDown.Picture
            Set UserControl.MaskPicture = imDown.Picture
        End If
    End If
End If
ReleaseCapture

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    Set m_ImageDown = LoadPicture("")
    Set m_ImageHot = LoadPicture("")
    Set m_ImageDisabled = LoadPicture("")
    Set m_ImageUp = LoadPicture("")
    sstyle = m_def_Style
    vval = m_def_Value
    m_MaskColor = m_def_MaskColor
    m_BackStyle = m_def_BackStyle
    Set UserControl.MaskPicture = LoadPicture("")
    UserControl.BackStyle = m_BackStyle
    UserControl.MaskColor = m_MaskColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_ImageDown = PropBag.ReadProperty("ImageDown", Nothing)
'    Set Picture = PropBag.ReadProperty("ImageUp", Nothing)
    Set m_ImageHot = PropBag.ReadProperty("ImageHot", Nothing)
    Set m_ImageDisabled = PropBag.ReadProperty("ImageDisabled", Nothing)
    sstyle = PropBag.ReadProperty("Style", m_def_Style)
    vval = PropBag.ReadProperty("Value", m_def_Value)
    Set m_ImageUp = PropBag.ReadProperty("ImageUp", Nothing)
    m_MaskColor = PropBag.ReadProperty("MaskColor", m_def_MaskColor)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    
    UserControl.BackStyle = m_BackStyle
    Set imUp.Picture = m_ImageUp
    Set imDown.Picture = m_ImageDown
    Set imHot.Picture = m_ImageHot
    Set imDis.Picture = m_ImageDisabled
    If m_Enabled = True Then
        If vval = abPressed Then
            Set UserControl.Picture = imDown.Picture
            Set UserControl.MaskPicture = imDown.Picture
        ElseIf vval = abUnPressed Then
            Set UserControl.Picture = imUp.Picture
            Set UserControl.MaskPicture = imUp.Picture
            Set imUp.Picture = m_ImageUp
        End If
    Else
        Set UserControl.Picture = imDis.Picture
        Set UserControl.MaskPicture = imDis.Picture
    End If
'Call UserControl_Resize

End Sub

Private Sub UserControl_Resize()
UserControl.Width = imUp.Width
UserControl.Height = imUp.Height

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("ImageDown", m_ImageDown, Nothing)
    Call PropBag.WriteProperty("ImageHot", m_ImageHot, Nothing)
    Call PropBag.WriteProperty("ImageDisabled", m_ImageDisabled, Nothing)
    Call PropBag.WriteProperty("ImageUp", m_ImageUp, Nothing)
    Call PropBag.WriteProperty("Style", sstyle, m_def_Style)
    Call PropBag.WriteProperty("Value", vval, m_def_Value)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ImageUp() As Picture
Attribute ImageUp.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set ImageUp = m_ImageUp
End Property

Public Property Set ImageUp(ByVal New_ImageUp As Picture)
    Set m_ImageUp = New_ImageUp
    Set imUp.Picture = New_ImageUp
    PropertyChanged "ImageUp"
    UserControl.BackStyle = 1
    Set UserControl.Picture = imUp.Picture
    Set UserControl.MaskPicture = imUp.Picture
    DoEvents
    Extender.Width = imUp.Width
    Extender.Height = imUp.Height
    UserControl.BackStyle = m_BackStyle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00C0C0C0&
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the Picture."
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
    UserControl.MaskColor = m_MaskColor
    Set UserControl.MaskPicture = m_ImageUp
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As BACKSTYLE_TO
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BACKSTYLE_TO)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
    UserControl.BackStyle = m_BackStyle
    Set UserControl.MaskPicture = UserControl.Picture
    UserControl.MaskColor = m_MaskColor
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

