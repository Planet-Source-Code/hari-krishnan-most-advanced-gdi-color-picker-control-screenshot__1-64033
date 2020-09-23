VERSION 5.00
Begin VB.UserControl AdvColorPickerDialog 
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   945
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   900
   ScaleWidth      =   945
   ToolboxBitmap   =   "AdvColorPickerDialog.ctx":0000
   Begin VB.Image imgIcon 
      Height          =   690
      Left            =   90
      Picture         =   "AdvColorPickerDialog.ctx":0312
      Top             =   90
      Width           =   690
   End
End
Attribute VB_Name = "AdvColorPickerDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : AdvColorPickerDialog
' DateTime  : 1/24/2005 18:12
' Author    : Hari Krishnan
' Purpose   : The Advanced Color Picker control
'---------------------------------------------------------------------------------------

Option Explicit
'Default Property Values:
Const m_def_Color = 0

Public Sub ShowDialog()
    Load frmClrPick
    frmClrPick.Show vbModal, UserControl.Parent
End Sub

Private Sub UserControl_Initialize()
    m_Color = 0
End Sub

Private Sub UserControl_Resize()
    imgIcon.Move 0, 0
    If UserControl.Width <> imgIcon.Width Or UserControl.height <> imgIcon.height Then
        UserControl.Width = imgIcon.Width
        UserControl.height = imgIcon.height
    End If
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Color() As Long
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As Long)
    m_Color = New_Color
    PropertyChanged "Color"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Color = m_def_Color
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Color = PropBag.ReadProperty("Color", m_def_Color)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
End Sub

