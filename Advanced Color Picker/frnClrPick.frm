VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmClrPick 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adv Color Picker"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4020
   Icon            =   "frnClrPick.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTS 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   3750
      Index           =   0
      Left            =   135
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   10
      Tag             =   "9,27,250,250"
      Top             =   405
      Width           =   3750
      Begin VB.PictureBox picSV 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   75
         ScaleHeight     =   205
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   240
         TabIndex        =   12
         Top             =   600
         Width           =   3600
         Begin VB.Line lnSVH 
            DrawMode        =   6  'Mask Pen Not
            Tag             =   "Saturation"
            X1              =   69
            X2              =   21
            Y1              =   57
            Y2              =   57
         End
         Begin VB.Line lnSVV 
            DrawMode        =   6  'Mask Pen Not
            Tag             =   "Value"
            X1              =   42
            X2              =   42
            Y1              =   27
            Y2              =   87
         End
      End
      Begin VB.PictureBox picHue 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   75
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   240
         TabIndex        =   11
         Top             =   75
         Width           =   3600
         Begin VB.Line lnHue 
            BorderWidth     =   2
            DrawMode        =   6  'Mask Pen Not
            X1              =   10
            X2              =   10
            Y1              =   -1
            Y2              =   31
         End
      End
   End
   Begin ComctlLib.TabStrip tsCLR 
      Height          =   4200
      Left            =   45
      TabIndex        =   0
      Tag             =   "3,3,262,280"
      Top             =   45
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   7408
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&HSV Cone"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&RGB Cube"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Palettes"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picAbout 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   1350
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   51
      Top             =   6705
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   90
      TabIndex        =   50
      Top             =   6750
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1935
      TabIndex        =   49
      Top             =   6750
      Width           =   960
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2970
      TabIndex        =   48
      Top             =   6750
      Width           =   960
   End
   Begin VB.PictureBox picTS 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   3750
      Index           =   1
      Left            =   135
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   1
      Tag             =   "9,27,250,250"
      Top             =   360
      Width           =   3750
      Begin VB.PictureBox picRB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1725
         Left            =   75
         ScaleHeight     =   115
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   115
         TabIndex        =   4
         Top             =   1875
         Width           =   1725
         Begin VB.Line lnRBH 
            DrawMode        =   6  'Mask Pen Not
            X1              =   75
            X2              =   24
            Y1              =   57
            Y2              =   57
         End
         Begin VB.Line lnRBV 
            DrawMode        =   6  'Mask Pen Not
            X1              =   48
            X2              =   48
            Y1              =   24
            Y2              =   84
         End
      End
      Begin VB.PictureBox picGB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1725
         Left            =   1875
         ScaleHeight     =   115
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   115
         TabIndex        =   3
         Top             =   75
         Width           =   1725
         Begin VB.Line lnGBH 
            DrawMode        =   6  'Mask Pen Not
            X1              =   39
            X2              =   84
            Y1              =   66
            Y2              =   66
         End
         Begin VB.Line lnGBV 
            DrawMode        =   6  'Mask Pen Not
            X1              =   63
            X2              =   63
            Y1              =   33
            Y2              =   93
         End
      End
      Begin VB.PictureBox picRG 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1725
         Left            =   75
         ScaleHeight     =   115
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   115
         TabIndex        =   2
         Top             =   75
         Width           =   1725
         Begin VB.Line lnRGH 
            DrawMode        =   6  'Mask Pen Not
            X1              =   69
            X2              =   21
            Y1              =   57
            Y2              =   57
         End
         Begin VB.Line lnRGV 
            DrawMode        =   6  'Mask Pen Not
            X1              =   42
            X2              =   42
            Y1              =   27
            Y2              =   87
         End
      End
      Begin VB.Label lblclrRGBhex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   2700
         TabIndex        =   9
         Top             =   2190
         Width           =   90
      End
      Begin VB.Label lblClrRGB 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   1950
         TabIndex        =   8
         Top             =   2025
         Width           =   615
      End
      Begin VB.Label lblBlue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blue : "
         Height          =   195
         Left            =   2100
         TabIndex        =   7
         Top             =   3330
         Width           =   450
      End
      Begin VB.Label lblGreen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Green : "
         Height          =   195
         Left            =   1980
         TabIndex        =   6
         Top             =   3060
         Width           =   570
      End
      Begin VB.Label lblRed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Red : "
         Height          =   195
         Left            =   2115
         TabIndex        =   5
         Top             =   2790
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   90
      TabIndex        =   43
      Top             =   5940
      Width           =   3840
      Begin VB.PictureBox picColorSelPrev 
         BackColor       =   &H00000000&
         Height          =   420
         Left            =   945
         ScaleHeight     =   360
         ScaleWidth      =   855
         TabIndex        =   45
         Top             =   180
         Width           =   915
      End
      Begin VB.PictureBox picColorCur 
         BackColor       =   &H00000000&
         Height          =   420
         Left            =   2700
         ScaleHeight     =   360
         ScaleWidth      =   855
         TabIndex        =   44
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Previous :"
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Current :"
         Height          =   195
         Left            =   2025
         TabIndex        =   46
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   1140
      Index           =   0
      Left            =   135
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   14
      Tag             =   "9,312,250,76"
      Top             =   4680
      Width           =   3750
      Begin VB.CommandButton cmdCustAdd 
         Caption         =   "Add to Custom"
         Height          =   420
         Left            =   2340
         TabIndex        =   34
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox txtHex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2610
         TabIndex        =   30
         Text            =   "128"
         Top             =   135
         Width           =   1050
      End
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1485
         TabIndex        =   28
         Text            =   "128"
         Top             =   765
         Width           =   735
      End
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1485
         TabIndex        =   27
         Text            =   "128"
         Top             =   450
         Width           =   735
      End
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   1485
         TabIndex        =   26
         Text            =   "128"
         Top             =   135
         Width           =   735
      End
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   405
         TabIndex        =   25
         Text            =   "128"
         Top             =   765
         Width           =   735
      End
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   405
         TabIndex        =   24
         Text            =   "128"
         Top             =   450
         Width           =   735
      End
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   405
         TabIndex        =   23
         Text            =   "128"
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# :"
         Height          =   195
         Left            =   2340
         TabIndex        =   29
         Top             =   180
         Width           =   195
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R : "
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G : "
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B : "
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   810
         Width           =   240
      End
      Begin VB.Label lblVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V :"
         Height          =   195
         Left            =   1260
         TabIndex        =   19
         Top             =   810
         Width           =   195
      End
      Begin VB.Label lblSat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S :"
         Height          =   195
         Left            =   1260
         TabIndex        =   18
         Top             =   495
         Width           =   195
      End
      Begin VB.Label lblhue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H :"
         Height          =   195
         Left            =   1245
         TabIndex        =   17
         Top             =   180
         Width           =   210
      End
   End
   Begin ComctlLib.TabStrip tsInfo 
      Height          =   1590
      Left            =   45
      TabIndex        =   13
      Tag             =   "3,288,262,106"
      Top             =   4320
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   2805
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Color &Info"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "C&ustom Palette"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Misc"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   1140
      Index           =   1
      Left            =   2790
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   16
      Top             =   4770
      Width           =   3750
      Begin VB.PictureBox picCust 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   495
         ScaleHeight     =   55
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   145
         TabIndex        =   32
         Top             =   90
         Width           =   2175
         Begin VB.Label lblCust 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   33
            Top             =   75
            Width           =   195
         End
         Begin VB.Shape shpCust 
            BorderColor     =   &H8000000D&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   240
            Left            =   765
            Top             =   135
            Width           =   285
         End
      End
      Begin VB.HScrollBar sbCust 
         Height          =   240
         Left            =   45
         TabIndex        =   31
         Top             =   855
         Width           =   2850
      End
   End
   Begin VB.PictureBox picTS 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   3750
      Index           =   2
      Left            =   3510
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   35
      Top             =   765
      Width           =   3750
      Begin VB.ComboBox cmbPal 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "frnClrPick.frx":1042
         Left            =   135
         List            =   "frnClrPick.frx":104F
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   180
         Width           =   3075
      End
      Begin VB.Label lblStdClr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   42
         Top             =   585
         Width           =   195
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   1140
      Index           =   2
      Left            =   2925
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   15
      Top             =   4590
      Width           =   3750
      Begin VB.PictureBox picScrColor 
         BackColor       =   &H00FFFFFF&
         Height          =   960
         Left            =   1080
         ScaleHeight     =   900
         ScaleWidth      =   675
         TabIndex        =   37
         Top             =   60
         Width           =   735
         Begin VB.Label lblScrB 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B : 0"
            Height          =   195
            Left            =   60
            TabIndex        =   40
            Top             =   675
            Width           =   330
         End
         Begin VB.Label lblScrG 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "G : 0"
            Height          =   195
            Left            =   45
            TabIndex        =   39
            Top             =   360
            Width           =   345
         End
         Begin VB.Label lblScrR 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "R : 0"
            Height          =   195
            Left            =   45
            TabIndex        =   38
            Top             =   45
            Width           =   345
         End
      End
      Begin VB.PictureBox picScrPick 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000014&
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   90
         MouseIcon       =   "frnClrPick.frx":1072
         MousePointer    =   99  'Custom
         ScaleHeight     =   62
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   36
         Top             =   60
         Width           =   960
         Begin VB.Shape shpClrPick 
            Height          =   120
            Left            =   420
            Top             =   420
            Width           =   120
         End
         Begin VB.Line lnPickV 
            DrawMode        =   6  'Mask Pen Not
            Index           =   1
            X1              =   32
            X2              =   32
            Y1              =   36
            Y2              =   66
         End
         Begin VB.Line lnPickH 
            DrawMode        =   6  'Mask Pen Not
            Index           =   1
            X1              =   36
            X2              =   66
            Y1              =   32
            Y2              =   32
         End
         Begin VB.Line lnPickH 
            DrawMode        =   6  'Mask Pen Not
            Index           =   0
            X1              =   -2
            X2              =   28
            Y1              =   32
            Y2              =   32
         End
         Begin VB.Line lnPickV 
            DrawMode        =   6  'Mask Pen Not
            Index           =   0
            X1              =   32
            X2              =   32
            Y1              =   -2
            Y2              =   28
         End
      End
   End
End
Attribute VB_Name = "frmClrPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmClrPick
' DateTime  : 1/24/2005 10:13
' Author    : Hari Krishnan
' Purpose   : The Color picker Dialog Interface functions
'---------------------------------------------------------------------------------------

Option Explicit

Dim giTok&, GSI As GdiplusStartupInput

Dim mCustCurSelected&, mScreenDC As Long




Private Sub cmdAbout_Click()
    ' The about box
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    m_Color = picColorCur.BackColor
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i&, s$, X&
    If KeyCode = vbKeyS And (Shift And vbCtrlMask) = vbCtrlMask Then
        Open App.path & "/Cust" & CStr(Round(Rnd() * 1000, 0)) & ".pal" For Output As #1
        Print #1, "AdvPal32bitFormat ColorList"
        X = 0
        For i = 0 To 15
            s = ""
            For X = 0 To 15
                s = s & CStr(lblCust(i * 16 + X).BackColor)
                If i = 15 And X = 15 Then
                Else
                    s = s & ","
                End If
            Next X
            Print #1, s
        Next i
        Close #1
    ElseIf KeyCode = vbKeyD And (Shift And vbCtrlMask) = vbCtrlMask Then
        cmdCustAdd_Click
    End If
End Sub

Private Sub Form_Load()
    Dim h&, s&, v&
    GSI.GdiplusVersion = 1&
    GdiplusStartup giTok, GSI
    
    Me.Caption = "Adv Color Picker v-" & App.Major & "." & App.Minor & "." & App.Revision
    
    SetPanels
    
    OleTranslateColor m_Color, 0&, m_Color
    toRGB m_Color, h, s, v
    picColorSelPrev.BackColor = m_Color
    
    MoveRGBLines h, s, v
    DrawRGBCube h, s, v
    
    UpdateColorInfo_RGB h, s, v
    
    ConvRGBtoHSL h, s, v, h, s, v
    SetHSV h, s, v
    
    PrepareCustPalette
    PrepareStdPalette
    
    SelectRGBColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GdiplusShutdown giTok
End Sub

Private Sub SetPanels()
'    On Error Resume Next
    Dim i&, s
    s = Split(tsCLR.Tag, ",")
    tsCLR.Move s(0), s(1), s(2), s(3)
    
    tsInfo.ZOrder (0)
    tsCLR.ZOrder (0)
    
    s = Split(picTS(0).Tag, ",")
    For i = 0 To picTS.count - 1
        picTS(i).Move s(0), s(1), s(2), s(3)
        If (tsCLR.SelectedItem.Index <> (i + 1)) Then
            picTS(i).Visible = False
        Else
            picTS(i).Visible = True
            picTS(i).ZOrder (0)
        End If
    Next i
    
    s = Split(tsInfo.Tag, ",")
    tsInfo.Move s(0), s(1), s(2), s(3)
    
    s = Split(picInfo(0).Tag, ",")
    For i = 0 To picInfo.count - 1
        picInfo(i).Move s(0), s(1), s(2), s(3)
        If (tsInfo.SelectedItem.Index <> (i + 1)) Then
            picInfo(i).Visible = False
        Else
            picInfo(i).Visible = True
            picInfo(i).ZOrder (0)
        End If
    Next i
End Sub

Private Sub tsCLR_Click()
    SetPanels
End Sub
Private Sub tsInfo_Click()
    SetPanels
End Sub



'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Function Collection :  RGB Functions
'---------------------------------------------------------------------------------------
'
Public Function MoveRGBLines(ByVal Rval&, ByVal Gval&, ByVal Bval&)
    Dim mh&, ms&, mv&, rv$, gv$, bv$
    ClipColorValues Rval, Gval, Bval
    If picRG.Tag <> "1" Then
        ' Horizontal line, origin at Top Right
        lnRGH.x1 = 0
        lnRGH.x2 = picRG.ScaleWidth
        lnRGH.y1 = Rval * picRG.ScaleHeight / 255
        lnRGH.y2 = lnRGH.y1
        ' Vertical line, origin at Top Right
        lnRGV.x1 = Gval * picRG.ScaleWidth / 255
        lnRGV.x2 = lnRGV.x1
        lnRGV.y1 = 0
        lnRGV.y2 = picRG.ScaleHeight
    End If
    If picGB.Tag <> "1" Then
        ' Horizontal line, origin at Top Left
        lnGBH.x1 = 0
        lnGBH.x2 = picGB.ScaleWidth
        lnGBH.y1 = Bval * picGB.ScaleHeight / 255
        lnGBH.y2 = lnGBH.y1
        ' Vertical line, origin at Top Left
        lnGBV.x1 = (255 - Gval) * picGB.ScaleWidth / 255
        lnGBV.x2 = lnGBV.x1
        lnGBV.y1 = 0
        lnGBV.y2 = picGB.ScaleHeight
    End If
    If picRB.Tag <> "1" Then
        ' Horizontal line, origin at Bottom Right
        lnRBH.x1 = 0
        lnRBH.x2 = picRB.ScaleWidth
        lnRBH.y1 = (255 - Rval) * picRB.ScaleHeight / 255
        lnRBH.y2 = lnRBH.y1
        ' Vertical line, origin at Bottom Right
        lnRBV.x1 = Bval * picRB.ScaleWidth / 255
        lnRBV.x2 = lnRBV.x1
        lnRBV.y1 = 0
        lnRBV.y2 = picRB.ScaleHeight
    End If
    lblRed.Caption = "Red : " & Rval & " ( " & Hex(Rval) & " )"
    lblGreen.Caption = "Green : " & Gval & " ( " & Hex(Gval) & " )"
    lblBlue.Caption = "Blue : " & Bval & " ( " & Hex(Bval) & " )"
    lblClrRGB.BackColor = RGB(Rval, Gval, Bval)
    rv = Hex(Rval)
    gv = Hex(Gval)
    bv = Hex(Bval)
    rv = IIf(Len(rv) < 2, "0" & rv, rv)
    gv = IIf(Len(gv) < 2, "0" & gv, gv)
    bv = IIf(Len(bv) < 2, "0" & bv, bv)
    lblclrRGBhex.Caption = "#" & rv & gv & bv
End Function

'---------------------------------------------------------------------------------------
' Procedure : DrawRGBCube
' DateTime  : 1/24/2005 10:14
' Author    : Hari Krishnan
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub DrawRGBCube(ByVal Rval&, ByVal Gval&, ByVal Bval&)
'    On Error GoTo DrawRGBCube_Error
    Dim gfx&
    Dim ptr(4) As POINTL, pclr(4) As Long, pbr&
    
    ClipColorValues Rval, Gval, Bval
    
    ptr(0).X = 0: ptr(0).Y = 0
    ptr(1).X = picRG.ScaleWidth: ptr(1).Y = 0
    ptr(2).X = picRG.ScaleWidth: ptr(2).Y = picRG.ScaleHeight
    ptr(3).X = 0: ptr(3).Y = picRG.ScaleHeight
    GdipCreatePathGradientI ptr(0), 4, WrapModeTile, pbr
    
    ' RG face of the color cube
    GdipCreateFromHDC picRG.hdc, gfx
    GdipSetInterpolationMode gfx, InterpolationModeHighQualityBicubic
    pclr(0) = ColorARGB(255, 0, 0, 0)
    pclr(1) = ColorARGB(255, 0, Gval, 0)
    pclr(2) = ColorARGB(255, Rval, Gval, 0)
    pclr(3) = ColorARGB(255, Rval, 0, 0)
    GdipSetPathGradientCenterPointI pbr, ptr(0)
    GdipSetPathGradientSurroundColorsWithCount pbr, pclr(0), 4
    GdipFillRectangleI gfx, pbr, 0, 0, ptr(2).X, ptr(2).Y
    picRG.Refresh
    GdipDeleteGraphics gfx
    
    ' GB face of the color cube
    GdipCreateFromHDC picGB.hdc, gfx
    GdipSetInterpolationMode gfx, InterpolationModeHighQualityBicubic
    pclr(0) = ColorARGB(255, 0, Gval, 0)
    pclr(1) = ColorARGB(255, 0, 0, 0)
    pclr(2) = ColorARGB(255, 0, 0, Bval)
    pclr(3) = ColorARGB(255, 0, Gval, Bval)
    GdipSetPathGradientCenterPointI pbr, ptr(1)
    GdipSetPathGradientSurroundColorsWithCount pbr, pclr(0), 4
    GdipFillRectangleI gfx, pbr, 0, 0, ptr(2).X, ptr(2).Y
    picGB.Refresh
    GdipDeleteGraphics gfx
    
    ' RB face of the color cube
    GdipCreateFromHDC picRB.hdc, gfx
    GdipSetInterpolationMode gfx, InterpolationModeHighQualityBicubic
    pclr(0) = ColorARGB(255, Rval, 0, 0)
    pclr(1) = ColorARGB(255, Rval, 0, Bval)
    pclr(2) = ColorARGB(255, 0, 0, Bval)
    pclr(3) = ColorARGB(255, 0, 0, 0)
    GdipSetPathGradientCenterPointI pbr, ptr(3)
    GdipSetPathGradientSurroundColorsWithCount pbr, pclr(0), 4
    GdipFillRectangleI gfx, pbr, 0, 0, ptr(2).X, ptr(2).Y
    picRB.Refresh
    GdipDeleteGraphics gfx
    
    GdipDeleteBrush pbr
    On Error GoTo 0
    Exit Sub
DrawRGBCube_Error:
    GdipDeleteGraphics gfx
    GdipDeleteBrush pbr
     MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure DrawRGBCube of Form frmClrPick"
End Sub

Private Sub picRG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r&, g&, b&
    If Button = vbLeftButton Then
        picRG.Tag = ""
        r = Y * 255 / picRG.ScaleHeight
        g = X * 255 / picRG.ScaleWidth
        b = lnGBH.y1 * 255 / picGB.ScaleHeight
        MoveRGBLines r, g, b
        DrawRGBCube r, g, b
        SelectRGBColor
    End If
End Sub
Private Sub picRG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picRG_MouseMove Button, Shift, X, Y
End Sub

Private Sub picGB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r&, g&, b&
    If Button = vbLeftButton Then
        picRG.Tag = ""
        r = lnRGH.y1 * 255 / picGB.ScaleHeight
        g = 255 - (X * 255 / picGB.ScaleWidth)
        b = Y * 255 / picGB.ScaleHeight
        MoveRGBLines r, g, b
        DrawRGBCube r, g, b
        SelectRGBColor
    End If
End Sub
Private Sub picGB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picGB_MouseMove Button, Shift, X, Y
End Sub

Private Sub picRB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r&, g&, b&
    If Button = vbLeftButton Then
        picRG.Tag = ""
        r = 255 - (Y * 255 / picRB.ScaleHeight)
        g = lnRGV.x1 * 255 / picRG.ScaleHeight
        b = X * 255 / picRB.ScaleWidth
        MoveRGBLines r, g, b
        DrawRGBCube r, g, b
        SelectRGBColor
    End If
End Sub
Private Sub picRB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picRB_MouseMove Button, Shift, X, Y
End Sub

Public Sub SetRGB(ByVal Rval&, ByVal Gval&, ByVal Bval&)
    MoveRGBLines Rval, Gval, Bval
    DrawRGBCube Rval, Gval, Bval
End Sub

Public Sub SelectRGBColor()
    Dim r&, g&, b&, h&, s&, v&
    r = lnRGH.y1 * 255 / picGB.ScaleHeight
    g = lnRGV.x1 * 255 / picRG.ScaleHeight
    b = lnGBH.y1 * 255 / picGB.ScaleHeight
    ConvRGBtoHSL r, g, b, h, s, v
    SetHSV h, s, v
    UpdateColorInfo_RGB r, g, b
End Sub

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Function Collection :  HSV Functions
'---------------------------------------------------------------------------------------
'
Public Sub DrawHue()
    Dim gfx As Long
    Dim pbr As Long, pts(7) As POINTL
    Dim w As Long, i As Long, clrs(7) As Long
    
    On Error GoTo DrawHue_Error

    w = picHue.ScaleWidth / 6#
    For i = 0 To 6
        pts(i).X = CLng(w * i)
        pts(i).Y = 0
    Next i
    pts(6).X = picHue.ScaleWidth
    clrs(0) = ColorARGB(255, 255, 0, 0)
    clrs(1) = ColorARGB(255, 255, 255, 0)
    clrs(2) = ColorARGB(255, 0, 255, 0)
    clrs(3) = ColorARGB(255, 0, 255, 255)
    clrs(4) = ColorARGB(255, 0, 0, 255)
    clrs(5) = ColorARGB(255, 255, 0, 255)
    clrs(6) = ColorARGB(255, 255, 0, 0)
    
    GdipCreateFromHDC picHue.hdc, gfx
    
    For i = 0 To 5
        GdipCreateLineBrushI pts(i), pts(i + 1), clrs(i), clrs(i + 1), WrapModeTileFlipX, pbr
        GdipFillRectangleI gfx, pbr, pts(i).X, pts(i).Y, pts(i + 1).X, picHue.ScaleHeight
        GdipDeleteBrush pbr
    Next i
    
    picHue.Refresh
    GdipDeleteGraphics gfx

    On Error GoTo 0
    Exit Sub

DrawHue_Error:
     MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure DrawHue of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DrawSV
' DateTime  : 1/24/2005 12:38
' Author    : Hari Krishnan
' Purpose   : To draw the Saturation-Value distribution space for given hue
'---------------------------------------------------------------------------------------
'
Public Function DrawSV(ByVal Hueval As Long)
    Dim ptr(4) As POINTL, clr(4) As Long, r&, g&, b&
    Dim gfx&, pbr&
    
    Hueval = IIf(Hueval < 0, 0, IIf(Hueval > 255, 255, Hueval))
    
    ptr(0).X = 0: ptr(0).Y = 0
    ptr(1).X = picSV.ScaleWidth: ptr(1).Y = 0
    ptr(2).X = picSV.ScaleWidth: ptr(2).Y = picSV.ScaleHeight
    ptr(3).X = 0: ptr(3).Y = picSV.ScaleHeight
    GdipCreatePathGradientI ptr(0), 4, WrapModeTile, pbr
    
    ConvHSLtoRGB Hueval, 0, 0, r, g, b
    clr(0) = ColorARGB(255, r, g, b)
    ConvHSLtoRGB Hueval, 255, 0, r, g, b
    clr(1) = ColorARGB(255, r, g, b)
    ConvHSLtoRGB Hueval, 255, 255, r, g, b
    clr(2) = ColorARGB(255, r, g, b)
    ConvHSLtoRGB Hueval, 0, 255, r, g, b
    clr(3) = ColorARGB(255, r, g, b)
    
    GdipCreateFromHDC picSV.hdc, gfx
    
    GdipSetPathGradientCenterPointI pbr, ptr(0)
    GdipSetPathGradientSurroundColorsWithCount pbr, clr(0), 4
    GdipFillRectangleI gfx, pbr, 0, 0, ptr(2).X, ptr(2).Y
    
    picSV.Refresh
    GdipDeleteGraphics gfx
    GdipDeleteBrush pbr
End Function

Public Sub MoveSV(ByVal Sval&, ByVal Vval&)
    Sval = Sval * picSV.ScaleWidth / 255
    Vval = Vval * picSV.ScaleHeight / 255
    lnSVH.x1 = 0: lnSVH.x2 = picSV.ScaleWidth
    lnSVH.y1 = Vval: lnSVH.y2 = lnSVH.y1
    lnSVV.x1 = Sval: lnSVV.x2 = lnSVV.x1
    lnSVV.y1 = 0: lnSVV.y2 = picSV.ScaleHeight
End Sub

Private Sub picHue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        lnHue.x1 = X
        lnHue.x2 = X
        DrawSV X * 255 / picHue.ScaleWidth
        SelectHSVColor
    End If
End Sub
Private Sub picHue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picHue_MouseMove Button, Shift, X, Y
End Sub

Private Sub picSV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Sval, Vval
    If Button = vbLeftButton Then
        Sval = X * 255 / picSV.ScaleWidth
        Vval = Y * 255 / picSV.ScaleHeight
        MoveSV Sval, Vval
        SelectHSVColor
    End If
End Sub
Private Sub picSV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSV_MouseMove Button, Shift, X, Y
End Sub

Private Sub SetHSV(ByVal Hval&, ByVal Sval&, ByVal Vval&)
    DrawHue
    DrawSV Hval
    lnHue.x1 = Hval * picHue.ScaleWidth / 255
    lnHue.x2 = lnHue.x1
    MoveSV Sval, Vval
End Sub

Public Sub SelectHSVColor()
    Dim r&, g&, b&, h&, s&, v&
    h = lnHue.x1 * 255 / picHue.ScaleWidth
    s = lnSVV.x1 * 255 / picSV.ScaleWidth
    v = lnSVH.y1 * 255 / picSV.ScaleHeight
    ConvHSLtoRGB h, s, v, r, g, b
    SetRGB r, g, b
    UpdateColorInfo_HSV h, s, v
End Sub




'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Function Collection :  Color Info Functions
'---------------------------------------------------------------------------------------
'
Public Function UpdateColorInfo_RGB(ByVal Rval&, ByVal Gval&, ByVal Bval&)
    Dim mh&, ms&, mv&, rv$, gv$, bv$, i&
    ClipColorValues Rval, Gval, Bval
    ConvRGBtoHSL Rval, Gval, Bval, mh, ms, mv
    txtR(0).Text = Rval
    txtR(1).Text = Gval
    txtR(2).Text = Bval
    txtR(3).Text = mh
    txtR(4).Text = ms
    txtR(5).Text = mv
    If (m_ShowLong = True) Then
        txtHex.Text = CStr(RGB(Rval, Gval, Bval))
    Else
        rv = Hex(Rval)
        gv = Hex(Gval)
        bv = Hex(Bval)
        rv = IIf(Len(rv) < 2, "0" & rv, rv)
        gv = IIf(Len(gv) < 2, "0" & gv, gv)
        bv = IIf(Len(bv) < 2, "0" & bv, bv)
        txtHex.Text = "#" & rv & gv & bv
    End If
    For i = 0 To 5
        txtR(i).Refresh
    Next i
    txtHex.Refresh
    picColorCur.BackColor = RGB(Rval, Gval, Bval)
End Function

Public Function UpdateColorInfo_HSV(ByVal Hval&, ByVal Sval&, ByVal Vval&)
    Dim Rval&, Gval&, Bval&, rv$, gv$, bv$, i&
    ClipColorValues Hval, Sval, Vval
    ConvHSLtoRGB Hval, Sval, Vval, Rval, Gval, Bval
    txtR(0).Text = Rval
    txtR(1).Text = Gval
    txtR(2).Text = Bval
    txtR(3).Text = Hval
    txtR(4).Text = Sval
    txtR(5).Text = Vval
    If (m_ShowLong = True) Then
        txtHex.Text = CStr(RGB(Rval, Gval, Bval))
    Else
        rv = Hex(Rval)
        gv = Hex(Gval)
        bv = Hex(Bval)
        rv = IIf(Len(rv) < 2, "0" & rv, rv)
        gv = IIf(Len(gv) < 2, "0" & gv, gv)
        bv = IIf(Len(bv) < 2, "0" & bv, bv)
        txtHex.Text = "#" & rv & gv & bv
    End If
    For i = 0 To 5
        txtR(i).Refresh
    Next i
    txtHex.Refresh
    picColorCur.BackColor = RGB(Rval, Gval, Bval)
End Function

Private Sub txtR_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtR_LostFocus Index
    End If
End Sub

Private Sub txtR_LostFocus(Index As Integer)
    txtR(Index).Text = CLng(Val(txtR(Index)))
    SetColorText Index
End Sub

Public Function SetColorText(Optional ByVal Index& = 0)
    Dim r&, g&, b&
    If Index < 3 Then
        r = CLng(txtR(0).Text)
        g = CLng(txtR(1).Text)
        b = CLng(txtR(2).Text)
        SetRGB r, g, b
        UpdateColorInfo_RGB r, g, b
        picColorCur.BackColor = RGB(r, g, b)
        ConvRGBtoHSL r, g, b, r, g, b
        SetHSV r, g, b
    Else
        r = CLng(txtR(3).Text)
        g = CLng(txtR(4).Text)
        b = CLng(txtR(5).Text)
        SetHSV r, g, b
        UpdateColorInfo_HSV r, g, b
        ConvHSLtoRGB r, g, b, r, g, b
        SetRGB r, g, b
        picColorCur.BackColor = RGB(r, g, b)
    End If
End Function



'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Function Collection :  Custom Palette Functions
'---------------------------------------------------------------------------------------
'
Private Sub PrepareCustPalette()
    On Local Error Resume Next
    Dim i&, X&, Y&
    X = 0: Y = 0
    For i = 0 To 255
        Load lblCust(i)
        lblCust(i).Move 5 + X * (lblCust(0).Width + 3), 5 + Y * (lblCust(0).height + 3)
        lblCust(i).BackColor = 0&
        lblCust(i).Visible = True
        lblCust(i).ZOrder (0)
        Y = Y + 1
        If (Y > 2) Then Y = 0: X = X + 1
    Next i
    picCust.Move 0, 0, 10 + (X + 1) * (lblCust(0).Width + 3), picInfo(1).ScaleHeight - sbCust.height
    sbCust.Move 0, picCust.height, picInfo(1).ScaleWidth
    sbCust.value = 0: sbCust.Min = 0
    sbCust.Max = picCust.Width - picInfo(1).ScaleWidth
    With lblCust(0)
        shpCust.Move .Left - 2, .Top - 2, .Width + 4, .height + 4
    End With
End Sub

Private Sub sbCust_Change()
    picCust.Left = -sbCust.value
End Sub
Private Sub sbCust_Scroll()
    sbCust_Change
End Sub

Private Sub SelectCust(ByVal idx&)
    Dim r&, g&, b&
    If idx < 0 Or idx >= lblCust.count Then Exit Sub
    mCustCurSelected = idx
    With lblCust(idx)
        shpCust.Move .Left - 2, .Top - 2, .Width + 4, .height + 4
        
        If .Tag = "" Then Exit Sub
        toRGB .BackColor, r, g, b
        SetRGB r, g, b
        picColorCur.BackColor = RGB(r, g, b)
        UpdateColorInfo_RGB r, g, b
        ConvRGBtoHSL r, g, b, r, g, b
        SetHSV r, g, b
    End With
End Sub

Private Sub lblCust_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        SelectCust Index
    ElseIf Button = vbRightButton Then
        With lblCust(Index)
            shpCust.Move .Left - 2, .Top - 2, .Width + 4, .height + 4
        End With
        mCustCurSelected = Index
        cmdCustAdd_Click
    End If
End Sub

Private Sub cmdCustAdd_Click()
    On Error Resume Next
    lblCust(mCustCurSelected).BackColor = picColorCur.BackColor
    lblCust(mCustCurSelected).Tag = "1"
    mCustCurSelected = mCustCurSelected + 1
    If mCustCurSelected >= lblCust.count Then
        mCustCurSelected = 0
        sbCust.value = 0
    End If
    With lblCust(mCustCurSelected)
        shpCust.Move .Left - 2, .Top - 2, .Width + 4, .height + 4
    End With
End Sub




'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Function Collection :  Standard Palette Functions
'---------------------------------------------------------------------------------------
'
Public Function PrepareStdPalette()
    Const ClrWid& = 17
    On Local Error Resume Next
    Dim i&, X&, Y&, KeyClrList(9) As Long, ClrList(20) As Long
    Dim r&, g&, b&

    cmbPal.Move 5, 5, picTS(2).ScaleWidth - 10

    X = 0: Y = 0

    For i = 0 To 255
        Load lblStdClr(i)
        lblStdClr(i).Move 5 + X * (lblStdClr(i).Width + 1), cmbPal.Top + cmbPal.height + 5 + Y * (lblStdClr(i).height + 1), 13, 12
        lblStdClr(i).Visible = True
        lblStdClr(i).ZOrder (0)
        lblStdClr(i).BackColor = &HFFFFFF
        X = X + 1
        If (X > 16) Then
            Y = Y + 1
            X = 0
        End If
    Next i

'    KeyClrList(0) = RGB(128, 128, 128)
'    KeyClrList(1) = RGB(255, 0, 0)
'    KeyClrList(2) = RGB(255, 128, 0)
'    KeyClrList(3) = RGB(255, 255, 0)
'    KeyClrList(4) = RGB(0, 255, 0)
'    KeyClrList(5) = RGB(0, 255, 255)
'    KeyClrList(6) = RGB(0, 140, 255)
'    KeyClrList(7) = RGB(0, 0, 255)
'    KeyClrList(8) = RGB(255, 0, 255)
'
'
'
'    For Y = 0 To 8
'        GetColorVariations KeyClrList(Y), ClrList, ClrWid + 1
'        For X = 0 To ClrWid
'            If Y = 0 And (X = 1 Or X = ClrWid) Then
'                If X = 1 Then
'                    lblStdClr(Y * ClrWid + X).BackColor = RGB(255, 255, 255)
'                ElseIf X = ClrWid Then
'                    lblStdClr(Y * ClrWid + X).BackColor = RGB(0, 0, 0)
'                End If
'            Else
'                lblStdClr(Y * ClrWid + X).BackColor = ClrList(X)
'            End If
'            toRGB lblStdClr(Y * ClrWid + X).BackColor, r, g, b
'            lblStdClr(Y * ClrWid + X).ToolTipText = "R:" & r & " | G:" & g & " | B:" & b
'            lblStdClr(Y * ClrWid + X).Refresh
'        Next X
'    Next Y

    SelectStdPal 0
    cmbPal.ListIndex = 0
End Function

Private Function GetColorVariations(ByVal KeyColor As Long, CList() As Long, ByVal ClrCount As Long)
    Dim i&, r#, g#, b#, rs#, gs#, bs#
    CList(0) = KeyColor
    ClrCount = ClrCount - 1
    
    toRGB KeyColor, r, g, b
    rs = -r / ((ClrCount \ 2) + 1)
    gs = -g / ((ClrCount \ 2) + 1)
    bs = -b / ((ClrCount \ 2) + 1)
    For i = 1 To (ClrCount \ 2)
        r = r + rs: g = g + gs: b = b + bs
        CList((ClrCount \ 2) + i) = RGB(r, g, b)
    Next i
    
    toRGB KeyColor, r, g, b
    rs = -(255 - r) / (ClrCount \ 2)
    gs = -(255 - g) / (ClrCount \ 2)
    bs = -(255 - b) / (ClrCount \ 2)
    r = 255: g = 255: b = 255
    For i = 1 To (ClrCount \ 2)
        r = r + rs: g = g + gs: b = b + bs
        CList(i) = RGB(r, g, b)
    Next i
End Function

Private Sub SelectStdPal(ByVal idx As Long, Optional ByVal mode& = 0)
    Dim r&, g&, b&
    If idx < 0 Or idx > (lblStdClr.count - 1) Then Exit Sub
    With lblStdClr(idx)
        If mode = 0 Then Exit Sub
        
        toRGB lblStdClr(idx).BackColor, r, g, b
        SetRGB r, g, b
        UpdateColorInfo_RGB r, g, b
        ConvRGBtoHSL r, g, b, r, g, b
        SetHSV r, g, b
        picColorCur.BackColor = lblStdClr(idx).BackColor
    End With
End Sub

Private Sub lblStdClr_Click(Index As Integer)
    SelectStdPal Index, 1
End Sub



Private Sub picScrPick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then
        mScreenDC = GetDC(GetDesktopWindow())
        picScrPick_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub picScrPick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    Dim r&, g&, b&
    Dim pt As POINTAPI
    If Button Then
        pt.X = X
        pt.Y = Y
        ClientToScreen picScrPick.hWnd, pt
        StretchBlt picScrPick.hdc, 0, 0, picScrPick.ScaleWidth - 1, picScrPick.ScaleHeight - 1, _
            mScreenDC, pt.X - 4, pt.Y - 4, 8, 8, vbSrcCopy
            
        picScrColor.BackColor = picScrPick.Point(31, 31) 'GetPixel(mScreenDC, pt.x, pt.y)
        toRGB picScrColor.BackColor, r, g, b
        lblScrR.ForeColor = RGB(255 - r, 255 - g, 255 - b)
        lblScrG.ForeColor = lblScrR.ForeColor
        lblScrB.ForeColor = lblScrR.ForeColor
        lblScrR.Caption = "R : " & r
        lblScrG.Caption = "G : " & g
        lblScrB.Caption = "B : " & b
        picScrPick.Refresh
    End If
End Sub

Private Sub picScrPick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r&, g&, b&
    toRGB picScrColor.BackColor, r, g, b
    SetRGB r, g, b
    UpdateColorInfo_RGB r, g, b
    ConvRGBtoHSL r, g, b, r, g, b
    SetHSV r, g, b
    picColorCur.BackColor = picScrColor.BackColor
End Sub

Private Sub cmbPal_Click()
    Dim clst&(256), i&, r&, g&, b&
    Select Case cmbPal.list(cmbPal.ListIndex)
        Case "Standard"
            GetPalette_Standard clst
        Case "XP Colors"
            GetPalette_XPColors clst
        Case "Web Safe"
            GetPalette_WebSafe216 clst
    End Select
    
    For i = 0 To 255
        lblStdClr(i).BackColor = clst(i)
        toRGB lblStdClr(i).BackColor, r, g, b
        lblStdClr(i).ToolTipText = "R:" & r & " | G:" & g & " | B:" & b
        lblStdClr(i).Refresh
    Next i
End Sub
