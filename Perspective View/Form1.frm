VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perspective View                                    - By Sandeep.G"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13905
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "S E T T I N G S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1815
      Left            =   11280
      TabIndex        =   1
      Top             =   5400
      Width           =   2535
      Begin VB.Image ONVR 
         Height          =   510
         Left            =   1800
         ToolTipText     =   "Hot Key : O"
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Visual Rays"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   435
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   840
         ToolTipText     =   "Hot Key: R"
         Top             =   1200
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "C O N T R O L S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   3495
      Left            =   11280
      TabIndex        =   0
      Top             =   1800
      Width           =   2535
      Begin VB.Image ImUPcCP1 
         Height          =   555
         Left            =   960
         ToolTipText     =   "Hot Key : UP Arrow"
         Top             =   840
         Width           =   630
      End
      Begin VB.Image ImLT 
         Height          =   555
         Left            =   600
         ToolTipText     =   "Hot Key : Left Arrow"
         Top             =   2760
         Width           =   630
      End
      Begin VB.Image ImDNcCP1 
         Height          =   555
         Left            =   960
         ToolTipText     =   "Hot Key : Down Arrow"
         Top             =   1560
         Width           =   630
      End
      Begin VB.Image ImRT 
         Height          =   555
         Left            =   1320
         ToolTipText     =   "Hot Key : Right Arrow"
         Top             =   2760
         Width           =   630
      End
      Begin VB.Image ImUPcCP 
         Height          =   555
         Left            =   120
         ToolTipText     =   "Hot Key : W"
         Top             =   840
         Width           =   630
      End
      Begin VB.Image ImDNcCP 
         Height          =   555
         Left            =   120
         ToolTipText     =   "Hot Key : S"
         Top             =   1560
         Width           =   630
      End
      Begin VB.Image ImDNGL 
         Height          =   555
         Left            =   1800
         ToolTipText     =   "Hot Key : D"
         Top             =   1560
         Width           =   630
      End
      Begin VB.Image ImUpGL 
         Height          =   555
         Left            =   1800
         ToolTipText     =   "Hot Key : E"
         Top             =   840
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   270
         Left            =   990
         Picture         =   "Form1.frx":030A
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   270
         Left            =   1875
         Picture         =   "Form1.frx":0A0C
         Top             =   360
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   270
         Left            =   240
         Picture         =   "Form1.frx":1156
         Top             =   360
         Width           =   450
      End
      Begin VB.Image Image6 
         Height          =   300
         Left            =   990
         Picture         =   "Form1.frx":1810
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Image ONVR_UP 
      Height          =   450
      Left            =   6960
      Picture         =   "Form1.frx":20C2
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ONVR_OR 
      Height          =   450
      Left            =   7560
      Picture         =   "Form1.frx":2BCE
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ONVR_DN 
      Height          =   450
      Left            =   8160
      Picture         =   "Form1.frx":36DA
      Top             =   4080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image1_OR 
      Height          =   375
      Left            =   8760
      Picture         =   "Form1.frx":41E6
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1_DN 
      Height          =   255
      Left            =   8760
      Picture         =   "Form1.frx":6552
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1_UP 
      Height          =   375
      Left            =   8760
      Picture         =   "Form1.frx":88BE
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line30 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9960
      X2              =   9960
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line29 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   10080
      X2              =   10080
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   10200
      X2              =   10200
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   10320
      X2              =   10320
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   3240
      Y2              =   3860
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9120
      X2              =   9120
      Y1              =   3240
      Y2              =   3860
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9240
      X2              =   9900
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9240
      X2              =   9885
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   9840
      X2              =   9840
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   9720
      X2              =   9720
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   9600
      X2              =   9600
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   9480
      X2              =   9480
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   9360
      X2              =   9360
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   9240
      X2              =   9240
      Y1              =   1560
      Y2              =   3600
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00C0FFC0&
      X1              =   4080
      X2              =   5880
      Y1              =   6040
      Y2              =   8040
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00C0FFC0&
      X1              =   4080
      X2              =   6960
      Y1              =   6040
      Y2              =   8040
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C0FFC0&
      X1              =   4080
      X2              =   5880
      Y1              =   6040
      Y2              =   6960
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00C0FFC0&
      X1              =   4080
      X2              =   6960
      Y1              =   6045
      Y2              =   6965
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C0FFC0&
      X1              =   6960
      X2              =   4080
      Y1              =   1320
      Y2              =   2760
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C0FFC0&
      X1              =   5880
      X2              =   4080
      Y1              =   1320
      Y2              =   2760
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C0FFC0&
      X1              =   6960
      X2              =   4080
      Y1              =   240
      Y2              =   2760
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C0FFC0&
      X1              =   5880
      X2              =   4080
      Y1              =   240
      Y2              =   2760
   End
   Begin VB.Line LineSP 
      BorderColor     =   &H00C0FFFF&
      BorderStyle     =   4  'Dash-Dot
      X1              =   4080
      X2              =   4080
      Y1              =   1320
      Y2              =   8040
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   6960
      X2              =   6960
      Y1              =   240
      Y2              =   1320
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   5880
      X2              =   6960
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   5880
      X2              =   5880
      Y1              =   240
      Y2              =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   6960
      Y2              =   8040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   5880
      X2              =   6960
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   5880
      X2              =   6960
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   5880
      X2              =   5880
      Y1              =   6960
      Y2              =   8040
   End
   Begin VB.Line LineGL 
      BorderColor     =   &H008080FF&
      X1              =   240
      X2              =   10500
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line LinePP 
      BorderColor     =   &H008080FF&
      X1              =   240
      X2              =   10500
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Image UP_DN 
      Height          =   495
      Left            =   8040
      Picture         =   "Form1.frx":AC2A
      Top             =   3480
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image UP_OR 
      Height          =   495
      Left            =   7440
      Picture         =   "Form1.frx":BB62
      Top             =   3480
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image UP_UP 
      Height          =   495
      Left            =   6840
      Picture         =   "Form1.frx":CA9A
      Top             =   3480
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image RT_UP 
      Height          =   495
      Left            =   6840
      Picture         =   "Form1.frx":D9D2
      Top             =   2760
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image RT_OR 
      Height          =   495
      Left            =   7440
      Picture         =   "Form1.frx":E90A
      Top             =   2760
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image RT_DN 
      Height          =   495
      Left            =   8040
      Picture         =   "Form1.frx":F842
      Top             =   2760
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image LT_OR 
      Height          =   495
      Left            =   7440
      Picture         =   "Form1.frx":1077A
      Top             =   2160
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image LT_UP 
      Height          =   495
      Left            =   6840
      Picture         =   "Form1.frx":116B2
      Top             =   2160
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image LT_DN 
      Height          =   495
      Left            =   8040
      Picture         =   "Form1.frx":125EA
      Top             =   2160
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Dn_UP 
      Height          =   495
      Left            =   6840
      Picture         =   "Form1.frx":13522
      Top             =   1560
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Dn_OR 
      Height          =   495
      Left            =   7440
      Picture         =   "Form1.frx":1445A
      Top             =   1560
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image DN_DN 
      Height          =   495
      Left            =   8040
      Picture         =   "Form1.frx":15392
      Top             =   1560
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Line Line32 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   11160
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line32 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line31 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   11160
      X2              =   11160
      Y1              =   120
      Y2              =   8400
   End
   Begin VB.Line Line31 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   8400
   End
   Begin VB.Line LineCP1 
      BorderColor     =   &H0080FF80&
      X1              =   3960
      X2              =   4200
      Y1              =   6045
      Y2              =   6045
   End
   Begin VB.Line LineCP 
      BorderColor     =   &H0080FF80&
      X1              =   3960
      X2              =   4200
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   5880
      X2              =   6960
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image ImSP 
      Height          =   270
      Left            =   3960
      Picture         =   "Form1.frx":162CA
      Top             =   960
      Width           =   480
   End
   Begin VB.Image ImGL 
      Height          =   270
      Left            =   10560
      Picture         =   "Form1.frx":16797
      Top             =   7920
      Width           =   495
   End
   Begin VB.Image ImCP1 
      Height          =   300
      Left            =   10560
      Picture         =   "Form1.frx":16C4D
      Top             =   5880
      Width           =   540
   End
   Begin VB.Image ImCP 
      Height          =   270
      Left            =   10560
      Picture         =   "Form1.frx":17164
      Top             =   2640
      Width           =   450
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   10560
      Picture         =   "Form1.frx":17640
      Top             =   1200
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cPP As Pnt, cGL As Pnt, cSP As Pnt, cCP As Pnt, cCP1 As Pnt
Dim TempAr() As Single, Tempcount As Single
Function ReView()
    Line13.X1 = Line2.X2
    Line13.Y1 = Line2.Y2
    Line13.X2 = cCP1.X1
    Line13.Y2 = cCP1.Y2
    
    Line14.X1 = Line1.X1
    Line14.Y1 = Line1.Y1
    Line14.X2 = cCP1.X1
    Line14.Y2 = cCP1.Y2
    
    Line15.X1 = Line4.X2
    Line15.Y1 = Line4.Y2
    Line15.X2 = cCP1.X1
    Line15.Y2 = cCP1.Y2
    
    Line16.X1 = Line1.X2
    Line16.Y1 = Line1.Y2
    Line16.X2 = cCP1.X1
    Line16.Y2 = cCP1.Y2
    
    Line9.X1 = Line6.X1
    Line9.Y1 = Line6.Y1
    Line9.X2 = cCP.X1
    Line9.Y2 = cCP.Y2
    
    Line10.X1 = Line6.X2
    Line10.Y1 = Line6.Y2
    Line10.X2 = cCP.X1
    Line10.Y2 = cCP.Y2
    
    Line11.X1 = Line7.X1
    Line11.Y1 = Line7.Y1
    Line11.X2 = cCP.X1
    Line11.Y2 = cCP.Y2
    
    Line12.X1 = Line7.X2
    Line12.Y1 = Line7.Y2
    Line12.X2 = cCP.X1
    Line12.Y2 = cCP.Y2
    
    With Line23
        .X1 = Line17.X1
        .X2 = Line18.X1
        .Y1 = Line17.Y2
        .Y2 = Line18.Y2
    End With
    
    With Line24
        .X1 = Line21.X1
        .X2 = Line22.X1
        .Y1 = Line21.Y2
        .Y2 = Line22.Y2
    End With
    
    With Line25
        .X1 = Line17.X1
        .X2 = Line17.X1
        .Y1 = Line17.Y2
        .Y2 = Line21.Y2
    End With
    
    With Line26
        .X1 = Line18.X1
        .X2 = Line18.X1
        .Y1 = Line18.Y2
        .Y2 = Line22.Y2
    End With
    
    With Line27
        .X1 = Line17.X1
        .X2 = Line19.X1
        .Y1 = Line17.Y2
        .Y2 = Line19.Y2
    End With
    
    With Line28
        .X1 = Line18.X1
        .X2 = Line20.X1
        .Y1 = Line18.Y2
        .Y2 = Line20.Y2
    End With
    
    With Line29
        .X1 = Line22.X1
        .X2 = Line4.X1
        .Y1 = Line22.Y2
        .Y2 = Line4.Y2
    End With
    
    With Line30
        .X1 = Line21.X1
        .X2 = Line1.X1
        .Y1 = Line21.Y2
        .Y2 = Line1.Y2
    End With
    
    LineCP1.Y1 = cCP1.Y1
    LineCP1.Y2 = cCP1.Y2
    LineCP.Y1 = cCP.Y1
    LineCP.Y2 = cCP.Y2
    
    LineCP1.X1 = cCP1.X1 - 120
    LineCP1.X2 = cCP1.X2 + 120
    LineCP.X1 = cCP.X1 - 120
    LineCP.X2 = cCP.X2 + 120
    
    ImGL.Top = LineGL.Y1
    ImCP.Top = LineCP.Y1
    ImCP1.Top = LineCP1.Y1
    ImSP.Left = LineSP.X1 - 100
    
    DropLines
    DrawEdges
End Function
Function NewPic()
    cPP.X1 = 240
    cPP.X2 = 10500
    cPP.Y1 = 1320
    cPP.Y2 = 1320
    
    cGL.X1 = 240
    cGL.X2 = 10500
    cGL.Y1 = 8040
    cGL.Y2 = 8040
    
    cSP.X1 = 4080
    cSP.X2 = 4080
    cSP.Y1 = 1320
    cSP.Y2 = 8040
    
    cCP.X1 = 4080
    cCP.X2 = 4080
    cCP.Y1 = 2760
    cCP.Y2 = 2760
    
    cCP1.X1 = 4080
    cCP1.X2 = 4080
    cCP1.Y1 = 6040
    cCP1.Y2 = 6040
    
    With LineSP
        .X1 = 4080
        .X2 = 4080
        .Y1 = 1320
        .Y2 = 8040
    End With
    LineGL.Y1 = 8040
    LineGL.Y2 = 8040
    
    With Line1
        .X1 = 5880
        .X2 = 5880
        .Y1 = 6960
        .Y2 = 8040
    End With
    
    With Line2
        .X1 = 5880
        .X2 = 6960
        .Y1 = 6960
        .Y2 = 6960
    End With
    
    With Line3
        .X1 = 5880
        .X2 = 6960
        .Y1 = 8040
        .Y2 = 8040
    End With
    
    With Line4
        .X1 = 6960
        .X2 = 6960
        .Y1 = 6960
        .Y2 = 8040
    End With
    
    ReView
    
End Function

Private Sub Command1_Click()
Form1.Cls
End Sub

Private Sub Command2_Click()
Form1.Line (4, 5)-(789, 7889), vbWhite
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        Call ImUPcCP1_Click
    End If
    If KeyCode = 87 Then
        Call ImUPcCP_Click
    End If
    If KeyCode = 40 Then
        Call ImDNcCP1_Click
    End If
    If KeyCode = 83 Then
        Call ImDNcCP_Click
    End If
    If KeyCode = 37 Then
       Call ImLT_Click
    End If
    If KeyCode = 39 Then
       Call ImRT_Click
    End If
    If KeyCode = 69 Then
       Call ImUpgl_Click
    End If
    If KeyCode = 68 Then
       Call ImDNGL_Click
    End If
    If KeyCode = 82 Then
        Call Image1_Click
    End If
    If KeyCode = 79 Then
        Call ONVR_Click
    End If
End Sub

Private Sub Form_Load()
    Call ONVR_Click
    ImUPcCP.Picture = UP_UP.Picture
    ImUPcCP1.Picture = UP_UP.Picture
    ImUpGL.Picture = UP_UP.Picture
    
    ImDNcCP.Picture = Dn_UP.Picture
    ImDNcCP1.Picture = Dn_UP.Picture
    ImDNGL.Picture = Dn_UP.Picture
    
    ImLT.Picture = LT_UP.Picture
    ImRT.Picture = RT_UP.Picture
    
    Image1.Picture = Image1_UP.Picture
    ONVR.Picture = ONVR_UP.Picture
    Call Image1_Click
End Sub

Function DropLines()
    Dim C As Single
    
If (Line9.X2 - Line9.X1) <> 0 Then
    Line17.Y1 = LinePP.Y1
        C = Abs(cCP.Y1 - LinePP.Y1) / Abs(LinePP.Y1 - Line6.Y1)
    Line17.X1 = (cCP.X1 + (C * Line6.X1)) / (C + 1)
    Line17.X2 = Line17.X1
        C = Abs(cCP1.X1 - Line17.X1) / Abs(Line17.X1 - Line2.X1)
    Line17.Y2 = (cCP1.Y1 + (C * Line2.Y1)) / (C + 1)
End If

If (Line10.X2 - Line10.X1) <> 0 Then
    Line18.Y1 = LinePP.Y1
        C = Abs(cCP.Y1 - LinePP.Y1) / Abs(LinePP.Y1 - Line6.Y2)
    Line18.X1 = (cCP.X1 + (C * Line6.X2)) / (C + 1)
    Line18.X2 = Line18.X1
        C = Abs(cCP1.X1 - Line18.X1) / Abs(Line18.X1 - Line2.X2)
    Line18.Y2 = (cCP1.Y1 + (C * Line2.Y2)) / (C + 1)
End If


    With Line19
        .X1 = Line7.X1
        .X2 = Line2.X1
        .Y1 = LinePP.Y1
        .Y2 = Line2.Y1
    End With
    
    With Line20
        .X1 = Line7.X2
        .X2 = Line2.X2
        .Y1 = LinePP.Y1
        .Y2 = Line2.Y2
    End With
    
If (Line9.X2 - Line9.X1) <> 0 Then
    Line21.Y1 = LinePP.Y1
        C = Abs(cCP.Y1 - LinePP.Y1) / Abs(LinePP.Y1 - Line6.Y1)
    Line21.X1 = (cCP.X1 + (C * Line6.X1)) / (C + 1)
    Line21.X2 = Line21.X1
        C = Abs(cCP1.X1 - Line21.X1) / Abs(Line21.X1 - Line3.X1)
    Line21.Y2 = (cCP1.Y1 + (C * Line3.Y1)) / (C + 1)
End If

If (Line10.X2 - Line10.X1) <> 0 Then
    Line22.Y1 = LinePP.Y1
        C = Abs(cCP.Y1 - LinePP.Y1) / Abs(LinePP.Y1 - Line6.Y2)
    Line22.X1 = (cCP.X1 + (C * Line6.X2)) / (C + 1)
    Line22.X2 = Line22.X1
        C = Abs(cCP1.X1 - Line22.X1) / Abs(Line22.X1 - Line3.X2)
    Line22.Y2 = (cCP1.Y1 + (C * Line3.Y2)) / (C + 1)
End If
    
End Function

Function DrawEdges()
    
    With Line23
        .X1 = Line17.X1
        .X2 = Line18.X1
        .Y1 = Line17.Y2
        .Y2 = Line18.Y2
    End With
    
    With Line24
        .X1 = Line21.X1
        .X2 = Line22.X1
        .Y1 = Line21.Y2
        .Y2 = Line22.Y2
    End With
    
    With Line25
        .X1 = Line17.X1
        .X2 = Line21.X1
        .Y1 = Line17.Y2
        .Y2 = Line21.Y2
    End With
    
    With Line26
        .X1 = Line18.X1
        .X2 = Line22.X1
        .Y1 = Line18.Y2
        .Y2 = Line22.Y2
    End With
    
    With Line27
        .X1 = Line17.X1
        .X2 = Line19.X1
        .Y1 = Line17.Y2
        .Y2 = Line19.Y2
    End With
    
    With Line28
        .X1 = Line18.X1
        .X2 = Line20.X2
        .Y1 = Line18.Y2
        .Y2 = Line20.Y2
    End With
    
    With Line29
        .X1 = Line22.X1
        .X2 = Line3.X2
        .Y1 = Line22.Y2
        .Y2 = Line3.Y2
    End With
    
    With Line30
        .X1 = Line21.X1
        .X2 = Line3.X1
        .Y1 = Line21.Y2
        .Y2 = Line3.Y1
    End With
    
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImUPcCP.Picture = UP_UP.Picture
    ImUPcCP1.Picture = UP_UP.Picture
    ImUpGL.Picture = UP_UP.Picture
    
    ImDNcCP.Picture = Dn_UP.Picture
    ImDNcCP1.Picture = Dn_UP.Picture
    ImDNGL.Picture = Dn_UP.Picture
    
    ImLT.Picture = LT_UP.Picture
    ImRT.Picture = RT_UP.Picture
    
    Image1.Picture = Image1_UP.Picture
    ONVR.Picture = ONVR_UP.Picture
    Image1.Picture = Image1_UP.Picture
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImUPcCP.Picture = UP_UP.Picture
    ImUPcCP1.Picture = UP_UP.Picture
    ImUpGL.Picture = UP_UP.Picture
    
    ImDNcCP.Picture = Dn_UP.Picture
    ImDNcCP1.Picture = Dn_UP.Picture
    ImDNGL.Picture = Dn_UP.Picture
    
    ImLT.Picture = LT_UP.Picture
    ImRT.Picture = RT_UP.Picture
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ONVR.Picture = ONVR_UP.Picture
    Image1.Picture = Image1_UP.Picture

End Sub

Private Sub Image1_Click()
    NewPic
    DropLines
    DrawEdges
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Picture = Image1_DN.Picture
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Picture = Image1_OR.Picture
    If Button = 1 Then
        Image1.Picture = Image1_DN.Picture
    End If
End Sub
Private Sub ImDNcCP_Click()
    If LineCP.Y1 < LineGL.Y1 Then
        cCP.Y1 = cCP.Y1 + 15
        cCP.Y2 = cCP.Y2 + 15
        ReView
    End If
End Sub

Private Sub ImDNcCP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImDNcCP.Picture = DN_DN.Picture
End Sub

Private Sub ImDNcCP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImDNcCP.Picture = Dn_OR.Picture
    If Button = 1 Then
        ImDNcCP.Picture = DN_DN.Picture
    End If
End Sub


Private Sub ImDNcCP1_Click()
    If LineCP1.Y1 < LineGL.Y1 Then
        cCP1.Y1 = cCP1.Y1 + 15
        cCP1.Y2 = cCP1.Y2 + 15
        ReView
    End If
End Sub

Private Sub ImDNcCP1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImDNcCP1.Picture = DN_DN.Picture
End Sub

Private Sub ImDNcCP1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImDNcCP1.Picture = Dn_OR.Picture
    If Button = 1 Then
        ImDNcCP1.Picture = DN_DN.Picture
    End If
End Sub


Private Sub ImDNGL_Click()
If LineGL.Y1 < 8040 Then
    LineGL.Y1 = LineGL.Y1 + 15
    LineGL.Y2 = LineGL.Y2 + 15
    Line1.Y1 = Line1.Y1 + 15
    Line1.Y2 = Line1.Y2 + 15
    Line2.Y1 = Line2.Y1 + 15
    Line2.Y2 = Line2.Y2 + 15
    Line3.Y1 = Line3.Y1 + 15
    Line3.Y2 = Line3.Y2 + 15
    Line4.Y1 = Line4.Y1 + 15
    Line4.Y2 = Line4.Y2 + 15
    ReView
End If
End Sub

Private Sub ImDNGL_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImDNGL.Picture = DN_DN.Picture
End Sub

Private Sub ImDNGL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImDNGL.Picture = Dn_OR.Picture
    If Button = 1 Then
        ImDNGL.Picture = DN_DN.Picture
    End If
End Sub


Private Sub ImLT_Click()
    If LineSP.X1 > LinePP.X1 Then
        LineSP.X1 = LineSP.X1 - 15
        LineSP.X2 = LineSP.X2 - 15
        cCP1.X1 = cCP1.X1 - 15
        cCP1.X2 = cCP1.X2 - 15
        cCP.X1 = cCP.X1 - 15
        cCP.X2 = cCP.X2 - 15
        ReView
    End If
End Sub

Private Sub ImLT_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImLT.Picture = LT_DN.Picture
End Sub

Private Sub ImLT_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImLT.Picture = LT_OR.Picture
    If Button = 1 Then
        ImLT.Picture = LT_DN.Picture
    End If
End Sub


Private Sub ImRT_Click()
    If LineSP.X1 < LinePP.X2 Then
        LineSP.X1 = LineSP.X1 + 15
        LineSP.X2 = LineSP.X2 + 15
        cCP1.X1 = cCP1.X1 + 15
        cCP1.X2 = cCP1.X2 + 15
        cCP.X1 = cCP.X1 + 15
        cCP.X2 = cCP.X2 + 15
        ReView
    End If
End Sub

Private Sub ImRT_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImRT.Picture = RT_DN.Picture
End Sub

Private Sub ImRT_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImRT.Picture = RT_OR.Picture
    If Button = 1 Then
        ImRT.Picture = RT_DN.Picture
    End If
End Sub


Private Sub ImUPcCP_Click()
    If LineCP.Y1 > LinePP.Y1 Then
        cCP.Y1 = cCP.Y1 - 15
        cCP.Y2 = cCP.Y2 - 15
        ReView
    End If
End Sub

Private Sub ImUPcCP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImUPcCP.Picture = UP_DN.Picture
End Sub

Private Sub ImUPcCP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImUPcCP.Picture = UP_OR.Picture
    If Button = 1 Then
        ImUPcCP.Picture = UP_DN.Picture
    End If
End Sub

Private Sub ImUPcCP1_Click()
    If LineCP1.Y1 > LinePP.Y1 Then
        cCP1.Y1 = cCP1.Y1 - 15
        cCP1.Y2 = cCP1.Y2 - 15
        ReView
    End If
End Sub

Private Sub ImUPcCP1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImUPcCP1.Picture = UP_DN.Picture
End Sub


Private Sub ImUPcCP1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImUPcCP1.Picture = UP_OR.Picture
    If Button = 1 Then
        ImUPcCP1.Picture = UP_DN.Picture
    End If
End Sub


Private Sub ImUpgl_Click()
If LineGL.Y1 > LinePP.Y1 Then
    LineGL.Y1 = LineGL.Y1 - 15
    LineGL.Y2 = LineGL.Y2 - 15
    Line1.Y1 = Line1.Y1 - 15
    Line1.Y2 = Line1.Y2 - 15
    Line2.Y1 = Line2.Y1 - 15
    Line2.Y2 = Line2.Y2 - 15
    Line3.Y1 = Line3.Y1 - 15
    Line3.Y2 = Line3.Y2 - 15
    Line4.Y1 = Line4.Y1 - 15
    Line4.Y2 = Line4.Y2 - 15
    ReView
End If
End Sub

Private Sub ImUpGL_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImUpGL.Picture = UP_OR.Picture
End Sub

Private Sub ImUpGL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ImUpGL.Picture = UP_OR.Picture
    If Button = 1 Then
        ImUpGL.Picture = UP_DN.Picture
    End If
End Sub

Private Sub ONVR_Click()
Label1.FontStrikethru = Not Label1.FontStrikethru
If LinePP.BorderStyle = 1 Then
LinePP.BorderStyle = 0
LineGL.BorderStyle = 0
LineSP.BorderStyle = 0
LineCP.BorderStyle = 0
LineCP1.BorderStyle = 0
Line9.BorderStyle = 0
Line10.BorderStyle = 0
Line11.BorderStyle = 0
Line12.BorderStyle = 0
Line13.BorderStyle = 0
Line14.BorderStyle = 0
Line15.BorderStyle = 0
Line16.BorderStyle = 0
Line17.BorderStyle = 0
Line18.BorderStyle = 0
Line19.BorderStyle = 0
Line20.BorderStyle = 0
Line21.BorderStyle = 0
Line22.BorderStyle = 0
Line5.BorderStyle = 0
Line6.BorderStyle = 0
Line7.BorderStyle = 0
Line8.BorderStyle = 0
ImSP.Visible = False
Image4.Visible = False
ImCP.Visible = False
ImCP1.Visible = False
ImGL.Visible = False
Else
LinePP.BorderStyle = 1
LineGL.BorderStyle = 1
LineSP.BorderStyle = 4
LineCP.BorderStyle = 1
LineCP1.BorderStyle = 1
Line9.BorderStyle = 1
Line10.BorderStyle = 1
Line11.BorderStyle = 1
Line12.BorderStyle = 1
Line13.BorderStyle = 1
Line14.BorderStyle = 1
Line15.BorderStyle = 1
Line16.BorderStyle = 1
Line17.BorderStyle = 3
Line18.BorderStyle = 3
Line19.BorderStyle = 3
Line20.BorderStyle = 3
Line21.BorderStyle = 3
Line22.BorderStyle = 3
Line5.BorderStyle = 1
Line6.BorderStyle = 1
Line7.BorderStyle = 1
Line8.BorderStyle = 1
ImSP.Visible = True
Image4.Visible = True
ImCP.Visible = True
ImCP1.Visible = True
ImGL.Visible = True
End If
End Sub

Private Sub ONVR_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ONVR.Picture = ONVR_DN.Picture
End Sub

Private Sub ONVR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ONVR.Picture = ONVR_OR.Picture
    If Button = 1 Then
        ONVR.Picture = ONVR_DN.Picture
    End If
End Sub



