VERSION 5.00
Begin VB.Form Wanhuachi 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10995
   ClientLeft      =   540
   ClientTop       =   540
   ClientWidth     =   7800
   FillColor       =   &H00FF80FF&
   ForeColor       =   &H00FF80FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   733
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   10560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Settings"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   8160
      Width           =   7600
      Begin VB.HScrollBar steps 
         Height          =   255
         Left            =   5520
         Max             =   10000
         Min             =   2000
         TabIndex        =   15
         Top             =   2040
         Value           =   2000
         Width           =   1935
      End
      Begin VB.HScrollBar Intvl 
         Height          =   255
         Left            =   1440
         Max             =   200
         Min             =   10
         TabIndex        =   13
         Top             =   2040
         Value           =   20
         Width           =   1815
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   255
         Left            =   1200
         Max             =   20000
         TabIndex        =   10
         Top             =   1680
         Value           =   10000
         Width           =   6195
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   1200
         Max             =   20000
         TabIndex        =   8
         Top             =   1320
         Value           =   10000
         Width           =   6195
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   1200
         Max             =   2000
         TabIndex        =   5
         Top             =   960
         Value           =   1000
         Width           =   6195
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   1200
         Max             =   20000
         TabIndex        =   4
         Top             =   600
         Value           =   6000
         Width           =   6195
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   1200
         Max             =   20000
         TabIndex        =   1
         Top             =   240
         Value           =   10000
         Width           =   6195
      End
      Begin VB.Label Label7 
         Caption         =   "StepsPerInterval=2000"
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Interval=20ms"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Fai2=1.57"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fai1=1.57"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "d=0.5"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "f2=600Hz"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "f1=1000Hz"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Line Line1 
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      Index           =   2
      X1              =   520
      X2              =   520
      Y1              =   0
      Y2              =   680
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   677.267
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   540
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   544
      Y1              =   678
      Y2              =   678
   End
End
Attribute VB_Name = "Wanhuachi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function CreatePen _
                Lib "gdi32" (ByVal nPenStyle As Long, _
                             ByVal nWidth As Long, _
                             ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
            (ByVal hdc As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
            ByVal nXEnd As Long, ByVal nYEnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
            ByVal x As Long, ByVal y As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long, _
            ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
            ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
            ByVal x As Long, ByVal y As Long, _
            ByRef lpPoint As POINTAPI) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Const vbSolid = 0
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
 Const WM_SYSCOMMAND = &H112
 Const SC_MOVE = &HF010&
 Const HTCAPTION = 2
Dim stop0 As Boolean
Dim paint0 As New rcircle, pi As Double

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Command3.Enabled = True
stop0 = fasle
'Me.Width = Me.Height / 1.3
'Me.Scale (-1.01, 1.01)-(1.01, -1.8)
Dim omiga1 As Double, omiga2 As Double, fai1 As Double, fai2 As Double, v As Long
Dim tmpmap As Long, tmpdc As Long
Dim myWidth As Long

Dim t As Double, X0 As Double, Y0 As Double, d As Double, py As Double
Dim op As POINTAPI

Dim brush As Long, pen As Long, blackbrush As Long, blackpen As Long
Dim ob As Long, opn As Long, x1Dc As Long, x1map As Long ', x2Dc As Long, x2map As Long

myWidth = 520



x1Dc = CreateCompatibleDC(Me.hdc)
x1map = CreateCompatibleBitmap(Me.hdc, myWidth, myWidth)
SelectObject x1Dc, x1map

pen = CreatePen(vbSolid, 1, RGB(170, 170, 170))
opn = SelectObject(x1Dc, pen)
brush = CreateSolidBrush(RGB(170, 170, 170))
ob = SelectObject(x1Dc, brush)
Rectangle x1Dc, 0, 0, myWidth, myWidth
DeleteObject brush
DeleteObject pen
'''''''''''''''''''''''''''


tmpdc = CreateCompatibleDC(Me.hdc)
tmpmap = CreateCompatibleBitmap(Me.hdc, myWidth, myWidth)
SelectObject tmpdc, tmpmap

pen = CreatePen(vbSolid, 1, RGB(255, 122, 255))
opn = SelectObject(tmpdc, pen)
brush = CreateSolidBrush(RGB(255, 122, 255))
ob = SelectObject(tmpdc, brush)

blackbrush = CreateSolidBrush(RGB(0, 0, 0))
blackpen = CreatePen(vbSolid, 1, 0)

'opn = SelectObject(Me.hdc, pen)


'Rectangle Me.hdc, 100, 100, 200, 200



omiga1 = 2 * pi * HScroll1.Value / 10
omiga2 = 2 * pi * HScroll2.Value / 10
fai1 = HScroll4.Value / HScroll4.Max * pi
'HScroll4.Value = HScroll4.Max * fai1 / pi
fai2 = HScroll5.Value / HScroll5.Max * pi
'HScroll5.Value = HScroll5.Max * fai2 / pi
t = 0
d = HScroll3.Value / HScroll3.Max
v = 0
py = 0


paint0.initialize omiga1, omiga2, d, Intvl.Value / 1000# / steps.Value, fai1, fai2, py
paint0.getnext X0, Y0
Dim pwid As Long, phei As Long
pwid = (X0 + 1) / 2 * myWidth
phei = (Y0 + 1) / 2 * myWidth
v = MoveToEx(tmpdc, pwid, phei, op)

Do
    v = v + 1
    'Print Str(v)
    'Print t
    Dim dt As Long
    dt = steps.Value
    For i = 0 To dt
        paint0.getnext X0, Y0
        
        'Me.PSet (X0, Y0)
        'Me.PSet (X0, 2# * i / 10000# - 1), RGB(155, 0, 0)
        pwid = (X0 + 1) / 2 * myWidth
        phei = (Y0 + 1) / 2 * myWidth
        LineTo tmpdc, pwid, phei
        'Me.PSet (2# * i / 10000# - 1, Y0), RGB(0, 155, 0)
        'DoEvents
        'If i Mod 500 = 0 Then
            'Me.Line (-1, Y0)-(1, Y0), RGB(255, 44, 44)
            'Me.Line (X0, -1)-(X0, 1), RGB(44, 255, 44)
            'DoEvents
            'Sleep 1
        'End If
    Next i
    
    'Me.Refresh
    'Me.Cls
    
    BitBlt Me.hdc, 0, 0, myWidth, myWidth, tmpdc, 0, 0, vbSrcCopy
    
    BitBlt tmpdc, 0, 0, myWidth, myWidth, x1Dc, 0, 0, vbSrcAnd
    BitBlt x1Dc, 0, 0, myWidth, myWidth, 0, 0, 0, vbDstInvert
    'SelectObject tmpdc, blackbrush
    'SelectObject tmpdc, blackpen
    'SelectObject Me.hdc, opn
    'Rectangle tmpdc, 0, 0, myWidth, myWidth
    'SelectObject tmpdc, brush
    'SelectObject tmpdc, pen
    'opn = SelectObject(tmpdc, pen)
    DoEvents
    Sleep Intvl.Value
    
   'Me.Width = Me.Height / 1.35
    'Me.Scale (-1.01, 1.01)-(1, -1.7)
    
    
Loop Until stop0
DeleteDC tmpdc
DeleteDC x1Dc
DeleteObject x1map
DeleteObject brush
DeleteObject pen
DeleteObject blackbrush
DeleteObject blackpen

DeleteObject tmpmap
'DeleteObject pen

'End
End Sub

Private Sub Command3_Click()
Command3.Enabled = fasle
Command2.Enabled = True
stop0 = True
End Sub

Private Sub Command4_Click()
Dim brush As Long, pen As Long
Dim ob As Long, op As Long

brush = CreateSolidBrush(RGB(100, 200, 200))
ob = SelectObject(Me.hdc, brush)

pen = CreatePen(vbDot, 0, RGB(0, 0, 255))

op = SelectObject(Me.hdc, pen)


v = Rectangle(Me.hdc, 100, 100, 200, 200)
SelectObject Me.hdc, ob
SelectObject Me.hdc, op
     
DeleteObject brush
DeleteObject pen
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
stop0 = True
'End
End Sub

Private Sub Form_Load()
'Print Cos(3.1415)
Command3.Enabled = fasle
stop0 = False
pi = 3.14159265358979
'Me.Width = Me.Height / 1.3
'Me.Scale (-1.01, 1.01)-(1.01, -1.8)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
    Call SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, ByVal 0&)
End Sub

Private Sub HScroll1_Change()
Label1.Caption = "f1=" + Str(HScroll1.Value / 10) + "Hz"
paint0.setomiga1 (2 * pi * HScroll1.Value / 10)
End Sub
Private Sub HScroll2_Change()
Label2.Caption = "f2=" + Str(HScroll2.Value / 10) + "Hz"
paint0.setomiga2 (2 * pi * HScroll2.Value / 10)
End Sub

Private Sub HScroll3_Change()
Label3.Caption = "d=" + Str(HScroll3.Value / HScroll3.Max)
paint0.setd (HScroll3.Value / HScroll3.Max)
End Sub

Private Sub HScroll4_Change()
Label4.Caption = "Fai1=" + Str(HScroll4.Value / HScroll4.Max * pi)
paint0.setfai1 (HScroll4.Value / HScroll4.Max * pi)
End Sub
Private Sub HScroll5_Change()
Label5.Caption = "Fai2=" + Str(HScroll5.Value / HScroll5.Max * pi)
paint0.setfai2 (HScroll5.Value / HScroll5.Max * pi)
End Sub

Private Sub Intvl_Change()
paint0.interval = Intvl.Value / 1000# / steps.Value
Label6.Caption = "Interval=" + Str(Intvl.Value) + "ms"
End Sub

Private Sub steps_Change()
paint0.interval = Intvl.Value / 1000# / steps.Value
Label7.Caption = "StepsPerInterval=" + Str(steps.Value)
End Sub
