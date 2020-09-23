Attribute VB_Name = "mdl_Main"
'---------------------------------------------------------------------------------------
' Module    : mdl_Main
' DateTime  : 1/24/2005 18:12
' Author    : Hari Krishnan
' Purpose   : The General Routines for the Advanced Color Picker Control.
'---------------------------------------------------------------------------------------

Option Explicit

Public m_Color As Long
Public m_ShowLong As Boolean


Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, _
    lpPoint As POINTAPI) As Long
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X _
    As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal _
    hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As _
    Long, ByVal Y As Long) As Long

Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor _
    As Long, ByVal lHPalette As Long, lColorRef As Long) As Long


'---------------------------------------------------------------------------------------
' Procedure : ClipColorValues
' DateTime  : 1/24/2005 11:16
' Author    : Hari Krishnan
' Purpose   : To clip color values within permissible range.
'---------------------------------------------------------------------------------------
'
Public Sub ClipColorValues(ByRef r As Long, ByRef g As Long, ByRef b As Long)
    r = IIf(r < 0, 0, IIf(r > 255, 255, r))
    g = IIf(g < 0, 0, IIf(g > 255, 255, g))
    b = IIf(b < 0, 0, IIf(b > 255, 255, b))
End Sub

Public Function Max(ByVal a, ByVal b, ByVal c)
    a = IIf(a > b, a, b)
    a = IIf(a > c, a, c)
    Max = a
End Function

Public Function Min(ByVal a, ByVal b, ByVal c)
    a = IIf(a < b, a, b)
    a = IIf(a < c, a, c)
    Min = a
End Function

'---------------------------------------------------------------------------------------
' Procedure : ConvRGBtoHSL
' DateTime  : 1/24/2005 12:57
' Author    : Hari Krishnan
' Purpose   : To convert between RGB colorspace and HSV colorspace
'---------------------------------------------------------------------------------------
'
Public Function ConvRGBtoHSL(ByVal r&, ByVal g&, ByVal b&, ByRef Hval&, ByRef _
    Sval&, ByRef Vval&)
    Dim var_R#, var_G#, var_B#, var_Max#, var_Min#, del_Max#
    Dim del_R#, del_G#, del_B#, h#, s#, v#
    var_R = (r / 255)                       'RGB values = From 0 to 255
    var_G = (g / 255)
    var_B = (b / 255)
    
    var_Min = Min(var_R, var_G, var_B)      'Min. value of RGB
    var_Max = Max(var_R, var_G, var_B)      'Max. value of RGB
    del_Max = var_Max - var_Min             'Delta RGB value
    
    v = var_Max
    
    If (del_Max = 0) Then                   'This is a gray, no chroma...
        h = 0                               'HSV results = From 0 to 1
        s = 0
    Else                                    'Chromatic data...
        s = del_Max / var_Max
        
        del_R = (((var_Max - var_R) / 6) + (del_Max / 2)) / del_Max
        del_G = (((var_Max - var_G) / 6) + (del_Max / 2)) / del_Max
        del_B = (((var_Max - var_B) / 6) + (del_Max / 2)) / del_Max
        
        If (var_R = var_Max) Then
            h = del_B - del_G
        ElseIf (var_G = var_Max) Then
            h = (1 / 3) + del_R - del_B
        ElseIf (var_B = var_Max) Then
            h = (2 / 3) + del_G - del_R
        End If
        
        If (h < 0) Then h = h + 1
        If (h > 1) Then h = h - 1
    End If
    Hval = h * 255
    Sval = s * 255
    Vval = v * 255
End Function

'---------------------------------------------------------------------------------------
' Procedure : ConvHSLtoRGB
' DateTime  : 1/24/2005 12:57
' Author    : Hari Krishnan
' Purpose   : To convert between RGB colorspace and HSV colorspace
'---------------------------------------------------------------------------------------
'
Public Function ConvHSLtoRGB(ByVal Hval&, ByVal Sval&, ByVal Vval&, ByRef Rval&, _
    ByRef Gval&, ByRef Bval&)
    Dim r#, g#, b#, var_R#, var_G#, var_B#, var_H#, var_I#, var_1#, var_2#, _
        var_3#
    Dim h#, s#, v#
    h = Hval / 255#
    s = Sval / 255#
    v = Vval / 255#
    If (s = 0) Then                        'HSV values = From 0 to 1
        r = v * 255                      'RGB results = From 0 to 255
        g = v * 255
        b = v * 255
    Else
        var_H = h * 6
        var_I = CInt(var_H - 0.5)            'Or ... var_i = floor( var_h )
        var_1 = v * (1 - s)
        var_2 = v * (1 - s * (var_H - var_I))
        var_3 = v * (1 - s * (1 - (var_H - var_I)))
        
        ' A little tweek needed here when converting HSV(1,1,1) to RGB
        If h = 1 Then var_2 = 0
        
        If (var_I = 0) Then
            var_R = v: var_G = var_3: var_B = var_1
        ElseIf (var_I = 1) Then
            var_R = var_2: var_G = v: var_B = var_1
        ElseIf (var_I = 2) Then
            var_R = var_1: var_G = v: var_B = var_3
        ElseIf (var_I = 3) Then
            var_R = var_1: var_G = var_2: var_B = v
        ElseIf (var_I = 4) Then
            var_R = var_3: var_G = var_1: var_B = v
        Else
            var_R = v: var_G = var_1: var_B = var_2
        End If
        
        r = var_R * 255                  'RGB results = From 0 to 255
        g = var_G * 255
        b = var_B * 255
    End If
    Rval = r
    Gval = g
    Bval = b
End Function


Public Sub GetPalette_Standard(clst() As Long)
    On Local Error Resume Next
    Dim i&, s
    ReDim clst(256) As Long
    s = "8421504,16777215,14671839,13619151,12632256,11579568,10526880,9474192,8421504,"
    s = s & "7500402,6579300,5592405,4671303,3750201,2829099,1842204,921102,255,14671871,"
    s = s & "12566527,10461183,8421631,6316287,4210943,2105599,255,227,198,170,142,113,85,57,"
    s = s & "28,33023,14675967,12574719,10473471,8438015,6336767,4235519,2134271,33023,29411,"
    s = s & "25798,21930,18318,14705,11093,7225,3612,65535,14680063,12582911,10485759,"
    s = s & "8454143,6356991,4259839,2162687,65535,58339,50886,43690,36494,29041,21845,14649,"
    s = s & "7196,65280,14680031,12582847,10485663,8454016,6356832,4259648,2162464,65280,"
    s = s & "58112,50688,43520,36352,28928,21760,14592,7168,16776960,16777183,16777151,"
    s = s & "16777119,16777088,16777056,16777024,16776992,16776960,14934784,13026816,"
    s = s & "11184640,9342464,7434496,5592320,3750144,1842176,16747520,16773599,16769727,"
    s = s & "16766111,16762496,16758624,16755008,16751136,16747520,14908416,13004032,"
    s = s & "11164928,9326080,7421440,5582592,3743488,1839104,16711680,16768991,16760767,"
    s = s & "16752543,16744576,16736352,16728128,16719904,16711680,14876672,12976128,"
    s = s & "11141120,9306112,7405568,5570560,3735552,1835008,16711935,16769023,16760831,16752639,16744703,16736511,16728319,"
    s = s & "16720127,16711935,14876899,12976326,11141290,9306254,7405681,5570645,3735609,"
    s = s & "1835036,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
    s = Split(s, ",")
    For i = 0 To 255
        clst(i) = CLng(s(i))
    Next i
End Sub
Public Sub GetPalette_XPColors(clst() As Long)
    On Local Error Resume Next
    Dim i&, s
    ReDim clst(256) As Long
    s = "26266,39372,52479,131071,10092542,10345471,6737151,3447295,4880895,13311,154,6632243,10040064,13329920,14057984,16750848,"
    s = s & "14653758,16764313,16769717,16776927,16764158,16764108,16751002,13395559,13474201,10118758,26112,39168,3394407,6749850,13434829,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
    s = Split(s, ",")
    For i = 0 To 255
        clst(i) = CLng(s(i))
    Next i
End Sub
Public Sub GetPalette_WebSafe216(clst() As Long)
    On Local Error Resume Next
    Dim i&, s
    ReDim clst(256) As Long
    s = "AdvPal32bitFormat ColorList"
    s = s & "0,3342336,6684672,10027008,13369344,16711680,51,3342387,6684723,10027059,13369395,16711731,102,3342438,6684774,10027110,"
    s = s & "13369446,16711782,153,3342489,6684825,10027161,13369497,16711833,204,3342540,6684876,10027212,13369548,16711884,255,3342591,"
    s = s & "6684927,10027263,13369599,16711935,13056,3355392,6697728,10040064,13382400,16724736,13107,3355443,6697779,10040115,13382451,16724787,"
    s = s & "13158,3355494,6697830,10040166,13382502,16724838,13209,3355545,6697881,10040217,13382553,16724889,13260,3355596,3355596,10040268,"
    s = s & "13382604,16724940,13311,3355647,6697983,10040319,13382655,16724991,26112,3368448,6710784,10053120,13395456,16737792,26163,3368499,"
    s = s & "6710835,10053171,13395507,16737843,26214,3368550,6710886,10053222,13395558,16737894,26265,3368601,6710937,10053273,13395609,16737945,"
    s = s & "26316,3368652,6710988,10053324,13395660,16737996,26367,3368703,6711039,10053375,13395711,16738047,39168,3381504,6723840,10066176,"
    s = s & "13408512,16750848,39219,3381555,6723891,10066227,13408563,16750899,39270,3381606,6723942,10066278,13408614,16750950,39321,3381657,"
    s = s & "6723993,10066329,13408665,16751001,39372,3381708,6724044,10066380,13408716,16751052,39423,3381759,6724095,10066431,13408767,16751103,"
    s = s & "52224,3394560,6736896,10079232,13421568,16763904,52275,3394611,6736947,10079283,13421619,16763955,52326,3394662,6736998,10079334,"
    s = s & "13421670,16764006,52377,3394713,6737049,10079385,13421721,16764057,52428,3394764,6737100,10079436,13421772,16764108,52479,3394815,"
    s = s & "6737151,10079487,13421823,16764159,65280,3407616,6749952,10092288,13434624,16776960,65331,3407667,6750003,10092339,13434675,16777011,"
    s = s & "65382,3407718,6750054,10092390,13434726,16777062,65433,3407769,6750105,10092441,13434777,16777113,65484,3407820,6750156,10092492,"
    s = s & "13434828,16777164,65535,3407871,6750207,10092543,13434879,16777215,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
    s = s & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"

    s = Split(s, ",")
    For i = 0 To 255
        clst(i) = CLng(s(i))
    Next i
End Sub
