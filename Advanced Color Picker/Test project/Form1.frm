VERSION 5.00
Object = "*\A..\AdvClrPick.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Show Advanced Color Dialog"
      Default         =   -1  'True
      Height          =   510
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2625
   End
   Begin AdvClrPick.AdvColorPickerDialog AdvColorPickerDialog1 
      Left            =   3105
      Top             =   225
      _ExtentX        =   1217
      _ExtentY        =   1217
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    AdvColorPickerDialog1.color = Me.BackColor
    AdvColorPickerDialog1.ShowDialog
    If AdvColorPickerDialog1.color <> &HAA00D9 Then Me.BackColor = AdvColorPickerDialog1.color
End Sub

