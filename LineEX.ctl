VERSION 5.00
Begin VB.UserControl LineEX 
   ClientHeight    =   180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   HasDC           =   0   'False
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   12
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   Windowless      =   -1  'True
End
Attribute VB_Name = "LineEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Enum lineOrient
   lo_horizontal = 0
   lo_vertical = 1
End Enum

Enum lineStyle
   ls_embossedline = 0
   ls_raisedline = 1
End Enum

Enum lineType
   lt_single = 1
   lt_double = 3
   lt_triple = 5
End Enum

'Default Property Values:
Const m_def_line_type = 1
Const m_def_line_orientation = 0
Const m_def_line_style = 0

'Property Variables:
Dim m_line_type As lineType
Dim m_line_orientation As lineOrient
Dim m_line_style As lineStyle


  
Private Sub Paint()
 
 Dim x1&, x2&, y1&, y2&, lcnt&
 Dim clr_dark&, clr_light&
 
 On Error Resume Next
 
 If m_line_style = ls_embossedline Then
     clr_dark = RGB(170, 170, 180)
     clr_light = vbWhite
 ElseIf m_line_style = ls_raisedline Then
     clr_dark = vbWhite
     clr_light = RGB(170, 170, 180)
 End If
 
 For lcnt = 1 To m_line_type Step 2
 
   If m_line_orientation = lo_horizontal Then
      UserControl.Line (0, lcnt - 1)-(ScaleWidth, lcnt - 1), clr_dark
      UserControl.Line (0, lcnt)-(ScaleWidth, lcnt), clr_light
      
   ElseIf m_line_orientation = lo_vertical Then
      UserControl.Line (lcnt - 1, 0)-(lcnt - 1, ScaleHeight), clr_light
      UserControl.Line (lcnt, 0)-(lcnt, ScaleHeight), clr_dark
       
   End If
 
 Next lcnt
 
   
End Sub

Private Sub UserControl_Paint()
  
  Call Paint
  
End Sub

Private Sub UserControl_Resize()
 
 Dim lval&
  
 If m_line_type = lt_single Then
     lval = 30
 ElseIf m_line_type = lt_double Then
     lval = 60
 ElseIf m_line_type = lt_triple Then
     lval = 90
 End If
 
 If line_orientation = lo_horizontal Then
     Height = lval
 ElseIf line_orientation = lo_vertical Then
     Width = lval
 End If
 
 Cls
 Call Paint
  
End Sub

Private Sub UserControl_Show()
  
  ScaleMode = vbPixels
  
End Sub

'line_orientation
Public Property Get line_orientation() As lineOrient
    line_orientation = m_line_orientation
End Property
Public Property Let line_orientation(ByVal New_line_orientation As lineOrient)
  
  'make sure the line orentation has changed
  If New_line_orientation <> m_line_orientation Then
     If New_line_orientation = lo_horizontal Then
        Width = 1500
     ElseIf New_line_orientation = lo_vertical Then
        Height = 1500
     End If
     
     m_line_orientation = New_line_orientation
     PropertyChanged "line_orientation"
     Call UserControl_Resize
  End If
  
End Property
'line_style
Public Property Get line_style() As lineStyle
    line_style = m_line_style
End Property
Public Property Let line_style(ByVal New_line_style As lineStyle)
    m_line_style = New_line_style
    PropertyChanged "line_style"
    Call UserControl_Resize
End Property
'line_type
Public Property Get line_type() As lineType
    line_type = m_line_type
End Property
Public Property Let line_type(ByVal New_line_type As lineType)
    m_line_type = New_line_type
    PropertyChanged "line_type"
    Call UserControl_Resize
End Property
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_line_orientation = m_def_line_orientation
    m_line_style = m_def_line_style
    m_line_type = m_def_line_type
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_line_orientation = PropBag.ReadProperty("line_orientation", m_def_line_orientation)
    m_line_style = PropBag.ReadProperty("line_style", m_def_line_style)
    m_line_type = PropBag.ReadProperty("line_type", m_def_line_type)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("line_orientation", m_line_orientation, m_def_line_orientation)
    Call PropBag.WriteProperty("line_style", m_line_style, m_def_line_style)
    Call PropBag.WriteProperty("line_type", m_line_type, m_def_line_type)
End Sub

 

