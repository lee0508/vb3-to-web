VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm사업장정보 
   BorderStyle     =   0  '없음
   Caption         =   "사업장정보"
   ClientHeight    =   10095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15405
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "사업장정보.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10100
   ScaleMode       =   0  '사용자
   ScaleWidth      =   15405
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   48
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "사업장정보.frx":0CCA
         Style           =   1  '그래픽
         TabIndex        =   57
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "사업장정보.frx":162D
         Style           =   1  '그래픽
         TabIndex        =   54
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "사업장정보.frx":1FD2
         Style           =   1  '그래픽
         TabIndex        =   52
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "사업장정보.frx":2920
         Style           =   1  '그래픽
         TabIndex        =   51
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "사업장정보.frx":32A4
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "사업장정보.frx":3B2B
         Style           =   1  '그래픽
         TabIndex        =   50
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "사 업 장 정 보 관 리"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   18
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   165
         TabIndex        =   49
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   4935
      Left            =   60
      TabIndex        =   47
      Top             =   5100
      Width           =   15195
      _cx             =   26802
      _cy             =   8705
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   60
      TabIndex        =   33
      Top             =   630
      Width           =   15195
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   30
         Left            =   11520
         MaxLength       =   1
         TabIndex        =   31
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   29
         Left            =   6030
         MaxLength       =   1
         TabIndex        =   30
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   28
         Left            =   14160
         MaxLength       =   3
         TabIndex        =   29
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   27
         Left            =   11505
         MaxLength       =   3
         TabIndex        =   28
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   26
         Left            =   8820
         MaxLength       =   3
         TabIndex        =   27
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   25
         Left            =   6030
         MaxLength       =   3
         TabIndex        =   26
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   24
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   25
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   23
         Left            =   11505
         MaxLength       =   6
         TabIndex        =   24
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   22
         Left            =   8820
         MaxLength       =   6
         TabIndex        =   23
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   21
         Left            =   6030
         MaxLength       =   6
         TabIndex        =   22
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   20
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   21
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   19
         Left            =   1275
         TabIndex        =   20
         Top             =   2520
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   18
         Left            =   9300
         TabIndex        =   19
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   17
         Left            =   1275
         TabIndex        =   18
         Top             =   2160
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   16
         Left            =   12720
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1305
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   15
         Left            =   12720
         MaxLength       =   1
         TabIndex        =   16
         Top             =   945
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   14
         Left            =   12720
         MaxLength       =   7
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpOpenDate 
         Height          =   270
         Left            =   12720
         TabIndex        =   14
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   19857409
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   4
         Left            =   5910
         MaxLength       =   20
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   13
         Left            =   9300
         TabIndex        =   13
         Top             =   1665
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   11
         Left            =   9315
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1305
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   12
         Left            =   1275
         TabIndex        =   12
         Top             =   1665
         Width           =   6945
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   10
         Left            =   9300
         MaxLength       =   1
         TabIndex        =   10
         Top             =   945
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   9
         Left            =   9300
         MaxLength       =   14
         TabIndex        =   9
         Top             =   585
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   8
         Left            =   9300
         MaxLength       =   20
         TabIndex        =   8
         Top             =   233
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   7
         Left            =   5910
         TabIndex        =   7
         Top             =   1305
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   6
         Left            =   5910
         TabIndex        =   6
         Top             =   945
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   5
         Left            =   5910
         MaxLength       =   14
         TabIndex        =   5
         Top             =   585
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   3
         Left            =   1275
         MaxLength       =   14
         TabIndex        =   3
         Top             =   1305
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   2
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   2
         Top             =   945
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   1
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   1
         Top             =   585
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   1275
         MaxLength       =   4
         TabIndex        =   0
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   5
         Left            =   15000
         TabIndex        =   86
         Top             =   4020
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "최종 매출단가 자동변경"
         Height          =   240
         Index           =   38
         Left            =   9240
         TabIndex        =   85
         Top             =   4020
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "최종 매입단가 자동변경"
         Height          =   240
         Index           =   37
         Left            =   3720
         TabIndex        =   84
         Top             =   4020
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   83
         Top             =   4020
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "단가자동변경"
         Height          =   240
         Index           =   36
         Left            =   75
         TabIndex        =   82
         ToolTipText     =   "거래명세서/세금계산서"
         Top             =   4020
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   120
         X2              =   15015
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   3
         Left            =   14985
         TabIndex        =   81
         Top             =   3555
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   2
         Left            =   12360
         TabIndex        =   80
         Top             =   3060
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "출력타입구분"
         Height          =   240
         Index           =   35
         Left            =   1320
         TabIndex        =   79
         Top             =   3555
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   78
         Top             =   3555
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   77
         Top             =   3060
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "기초마감년월"
         Height          =   240
         Index           =   34
         Left            =   75
         TabIndex        =   76
         Top             =   3060
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "세금계산서왼쪽마진"
         Height          =   240
         Index           =   33
         Left            =   12360
         TabIndex        =   75
         Top             =   3555
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "세금계산서상단마진"
         Height          =   240
         Index           =   32
         Left            =   9720
         TabIndex        =   74
         Top             =   3555
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "거래명세서왼쪽마진"
         Height          =   240
         Index           =   31
         Left            =   6900
         TabIndex        =   73
         Top             =   3555
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "거래명세서상단마진"
         Height          =   240
         Index           =   30
         Left            =   4200
         TabIndex        =   72
         Top             =   3555
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "거래/계산서"
         Height          =   240
         Index           =   29
         Left            =   75
         TabIndex        =   71
         ToolTipText     =   "거래명세서/세금계산서"
         Top             =   3550
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   120
         X2              =   15015
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "회계 기초마감년월"
         Height          =   240
         Index           =   28
         Left            =   9840
         TabIndex        =   70
         Top             =   3060
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "미수금 기초마감년월"
         Height          =   240
         Index           =   27
         Left            =   6900
         TabIndex        =   69
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "미지급금 기초마감년월"
         Height          =   240
         Index           =   26
         Left            =   3960
         TabIndex        =   68
         Top             =   3060
         Width           =   1935
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   120
         X2              =   15015
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "자재 기초마감년월"
         Height          =   240
         Index           =   25
         Left            =   1320
         TabIndex        =   67
         Top             =   3060
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "백업폴더"
         Height          =   240
         Index           =   24
         Left            =   75
         TabIndex        =   66
         ToolTipText     =   "서버의 백업폴더입니다."
         Top             =   2600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "홈페이지주소"
         Height          =   240
         Index           =   23
         Left            =   8100
         TabIndex        =   65
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "이메일주소"
         Height          =   240
         Index           =   22
         Left            =   75
         TabIndex        =   64
         Top             =   2200
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   15015
         Y1              =   2055
         Y2              =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "(1.전표, 2.계산서)"
         Height          =   240
         Index           =   21
         Left            =   13440
         TabIndex        =   63
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "(1.전표, 2.계산서)"
         Height          =   240
         Index           =   20
         Left            =   13440
         TabIndex        =   62
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "미수금발생"
         Height          =   240
         Index           =   19
         Left            =   11475
         TabIndex        =   61
         ToolTipText     =   "(1.전표, 2.계산서)"
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "미지급금발생"
         Height          =   240
         Index           =   17
         Left            =   11475
         TabIndex        =   60
         ToolTipText     =   "(1.전표, 2.계산서)"
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "(0.정상)"
         Height          =   240
         Index           =   18
         Left            =   10200
         TabIndex        =   59
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부가세율(%)"
         Height          =   240
         Index           =   16
         Left            =   11475
         TabIndex        =   58
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   10440
         TabIndex        =   56
         ToolTipText     =   "300"
         Top             =   1365
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "개업일자"
         Height          =   240
         Index           =   14
         Left            =   11475
         TabIndex        =   55
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사용구분"
         Height          =   240
         Index           =   9
         Left            =   8100
         TabIndex        =   53
         ToolTipText     =   "0.정상, 기타.시용불가"
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "번지"
         Height          =   240
         Index           =   13
         Left            =   8100
         TabIndex        =   46
         Top             =   1725
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "팩스번호"
         Height          =   240
         Index           =   12
         Left            =   8100
         TabIndex        =   45
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "우편번호"
         Height          =   240
         Index           =   11
         Left            =   8100
         TabIndex        =   44
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "주소"
         Height          =   240
         Index           =   10
         Left            =   75
         TabIndex        =   43
         Top             =   1725
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "전화번호"
         Height          =   240
         Index           =   8
         Left            =   8100
         TabIndex        =   42
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "업종"
         Height          =   240
         Index           =   7
         Left            =   4710
         TabIndex        =   41
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "업태"
         Height          =   240
         Index           =   6
         Left            =   4710
         TabIndex        =   40
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "주민번호"
         Height          =   240
         Index           =   5
         Left            =   4710
         TabIndex        =   39
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "대표자명"
         Height          =   240
         Index           =   4
         Left            =   4710
         TabIndex        =   38
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "법인번호"
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   37
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사업자번호"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   36
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사업장명"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   35
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사업장코드"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   34
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm사업장정보"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 사업장정보
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 사업장
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 22  '그리드 한 페이지 당 행수(FixedRows 포함)

'+--------------------------------+
'/// LOAD FORM ( 한번만 실행 ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
End Sub

'+-------------------------------------------+
'/// ACTIVATE FORM 활성화 ( 한번만 실행 ) ///
'+-------------------------------------------+
Private Sub Form_Activate()
Dim SQL     As String
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       Subvsfg1_INIT
       Subvsfg1_FILL
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '인쇄, 조회
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '추가, 저장, 인쇄, 조회
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       dtpOpenDate.Value = Format("19000101", "0000-00-00")
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 정보(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+-------------------+
'/// Text1(index) ///
'+-------------------+
Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim strExeFile As String
Dim varRetVal  As Variant
    If (Index = 11 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then
       If Len(Trim(Text1(Index).Text)) = 6 Then
          Text1(Index).Text = Format(Trim(Text1(Index).Text), "###-###")
       End If
       PB_strPostCode = Trim(Text1(Index).Text)
       PB_strPostName = Trim(Text1(Index + 1).Text)
       frm우편번호검색.Show vbModal
       If (Len(PB_strPostCode) + Len(PB_strPostName)) = 0 Then '검색에서 취소(ESC)
       Else
          Text1(Index).Text = PB_strPostCode
          Text1(Index + 1).Text = PB_strPostName
       End If
       If PB_strPostCode <> "" Then
          Text1(Index + 2).SetFocus
       Else
          Text1(Index + 1).SetFocus
       End If
       Exit Sub
    End If
    If KeyCode = vbKeyReturn Then
       Select Case Index
       End Select
       SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = Format(Val(Trim(.Text)), "00")
                     If Trim(Text1(Index).Text) = "00" Then
                        Text1(Index).Text = ""
                     End If
                     If Text1(Index).Enabled = True Then
                        P_adoRec.CursorLocation = adUseClient
                        strSQL = "SELECT * FROM 사업장 " _
                                & "WHERE 사업장코드 = '" & Trim(.Text) & "' "
                        On Error GoTo ERROR_TABLE_SELECT
                        P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                        If P_adoRec.RecordCount <> 0 Then
                           P_adoRec.Close
                           .Text = ""
                           .SetFocus
                           Exit Sub
                        End If
                        P_adoRec.Close
                     End If
                Case 2 '사업자등록번호
                     If Len(Trim(.Text)) = 10 Then
                        .Text = Format(Trim(.Text), "###-##-#####")
                     End If
                Case 5 '주민등록번호
                     If Len(Trim(.Text)) = 13 Then
                        .Text = Format(Trim(.Text), "######-#######")
                     End If
                Case 11 '우편번호
                     If Len(Trim(.Text)) = 6 Then
                        .Text = Format(Trim(.Text), "###-###")
                     End If
                Case 14 '부가세율
                     .Text = Format(Vals(Trim(.Text)), "#00.00")
                Case 15, 16 '15.미지급금발생, 16.미수금발생
                     If Len(.Text) = 0 Or Trim(.Text) = "0" Then .Text = "2"
                     .Text = Format(Val(Trim(.Text)), "0")
                Case 24     '24.출력타입구분
                     .Text = Fix(Val(Trim(.Text)))
                     If Len(.Text) = 0 Or Trim(.Text) = "0" Then .Text = "1"
                     .Text = Fix(Val(Trim(.Text)))
                Case 25 To 28   '상단, 왼쪽
                     .Text = Fix(Val(Trim(.Text)))
                Case 29 To 30   '단가자동변경
                     If Val(Trim(.Text)) <> 1 Then
                        .Text = "0"
                     End If
                     .Text = Fix(Val(Trim(.Text)))
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub dtpOpenDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

'+-----------+
'/// Grid ///
'+-----------+
Private Sub vsfg1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         .ToolTipText = ""
         If .MouseRow < .FixedRows Or .MouseCol < 0 Then
            Exit Sub
         End If
         .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub
Private Sub vsfg1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    P_intButton = Button
End Sub
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack '.ForeColorSel
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 0))
            'If P_intButton = 1 Then
            '   .Sort = flexSortGenericAscending
            'Else
            '   .Sort = flexSortGenericDescending
            'End If
            'If .FindRow(strData, , 0) > 0 Then
            '   .Row = .FindRow(strData, , 0)
            'End If
            'If PC_intRowCnt < .Rows Then
            '   .TopRow = .Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
         If .Row >= .FixedRows And OldRow <> NewRow Then
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case Is <= 10
                            Text1(lngC).Text = .TextMatrix(.Row, lngC)
                       Case 12
                            If Len(.TextMatrix(.Row, lngC)) = 10 Then
                               dtpOpenDate.Value = .TextMatrix(.Row, lngC)
                            Else
                               dtpOpenDate.Value = Format("19000101", "0000-00-00")
                            End If
                       Case 13 To 15
                            Text1(lngC - 2).Text = .TextMatrix(.Row, lngC)
                       Case 17
                            Text1(14).Text = Format(.ValueMatrix(.Row, lngC), "#00.00")
                       Case 18, 19
                            Text1(lngC - 3).Text = Format(.ValueMatrix(.Row, lngC), "0")
                       Case 20 To 26
                            Text1(lngC - 3).Text = .TextMatrix(.Row, lngC)
                       Case 27 To 33
                            Text1(lngC - 3).Text = .TextMatrix(.Row, lngC)
                       Case Else
                End Select
            Next lngC
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         If .Row = 0 Then
            Text1(Text1.LBound).Enabled = True
         Else
            Text1(Text1.LBound).Enabled = False
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         'If KeyCode = vbKeyInsert Then
         '   SubClearText
         '   .Row = 0
         '   Text1(Text1.LBound).SetFocus
         '   Exit Sub
         'End If
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub

End Sub

'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
    SubClearText
    vsfg1.Row = 0
    Text1(Text1.LBound).Enabled = True
    Text1(Text1.LBound).SetFocus
End Sub
'+-----------+
'/// 조회 ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    SubClearText
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL    As String
Dim lngR      As Long
Dim lngC      As Long
Dim blnOK     As Boolean
Dim intRetVal As Integer
    '입력내역 검사
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       Text1(lngC).SetFocus
       Exit Sub
    End If
    If Text1(Text1.LBound).Enabled = True Then
       intRetVal = MsgBox("입력된 자료를 추가하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 추가")
    Else
       intRetVal = MsgBox("수정된 자료를 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 저장")
    End If
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    cmdSave.Enabled = False
    With vsfg1
         Screen.MousePointer = vbHourglass
         P_adoRec.CursorLocation = adUseClient
         If Text1(Text1.LBound).Enabled = True Then '사업장 추가면 검색
            strSQL = "SELECT * FROM 사업장 " _
                    & "WHERE 사업장코드 = '" & Trim(Text1(Text1.LBound).Text) & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount <> 0 Then
               P_adoRec.Close
               Text1(Text1.LBound).SetFocus
               Screen.MousePointer = vbDefault
               cmdSave.Enabled = True
               Exit Sub
            End If
            P_adoRec.Close
         End If
         PB_adoCnnSQL.BeginTrans
         If Text1(Text1.LBound).Enabled = True Then '사업장 추가
            strSQL = "INSERT INTO 사업장(사업장코드, 사업장명, 사업자번호," _
                                       & "법인번호, 대표자명, 대표자주민번호," _
                                       & "개업일자," _
                                       & "우편번호, 주소, 번지," _
                                       & "업태, 업종, 전화번호," _
                                       & "팩스번호, 부가세율, 미지급금발생구분, 미수금발생구분, " _
                                       & "사용구분, 수정일자, 사용자코드, " _
                                       & "이메일주소, 홈페이지주소, 백업폴더, " _
                                       & "자재기초마감년월,미지급금기초마감년월,미수금기초마감년월, 회계기초마감년월, " _
                                       & "출력타입구분, 거래명세서상단마진, 거래명세서왼쪽마진, " _
                                       & "세금계산서상단마진, 세금계산서왼쪽마진, " _
                                       & "최종입고단가자동갱신구분, 최종입고단가자동갱신구분 ) VALUES( " _
                    & "'" & Trim(Text1(0).Text) & "','" & Trim(Text1(1).Text) & "','" & Trim(Text1(2).Text) & "', " _
                    & "'" & Trim(Text1(3).Text) & "','" & Trim(Text1(4).Text) & "','" & Trim(Text1(5).Text) & "', " _
                    & "'" & IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)) & "', " _
                    & "'" & Trim(Text1(11).Text) & "','" & Trim(Text1(12).Text) & "','" & Trim(Text1(13).Text) & "', " _
                    & "'" & Trim(Text1(6).Text) & "','" & Trim(Text1(7).Text) & "','" & Trim(Text1(8).Text) & "', " _
                    & "'" & Trim(Text1(9).Text) & "', " & Vals(Text1(14).Text) & ", " & Vals(Text1(15).Text) & ", " & Vals(Text1(16).Text) & ", " _
                    & "" & Val(Text1(10).Text) & ", '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "', " _
                    & "'" & Trim(Text1(17).Text) & "', '" & Trim(Text1(18).Text) & "', '" & Trim(Text1(19).Text) & "', " _
                    & "'" & Trim(Text1(20).Text) & "', '" & Trim(Text1(21).Text) & "', " _
                    & "'" & Trim(Text1(22).Text) & "', '" & Trim(Text1(23).Text) & "', " _
                    & "" & Vals(Trim(Text1(24).Text)) & ", " & Vals(Trim(Text1(25).Text)) & ", " & Vals(Trim(Text1(26).Text)) & ", " _
                    & "" & Vals(Trim(Text1(27).Text)) & ", " & Vals(Trim(Text1(28).Text)) & ", " _
                    & "" & Vals(Trim(Text1(29).Text)) & ", " & Vals(Trim(Text1(30).Text)) & " ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            '자재원장생성(자재->자재원장)
            strSQL = "INSERT INTO 자재원장 " _
                        & "SELECT '" & Trim(Text1(0).Text) & "' AS 사업장코드, " _
                               & "ISNULL(T1.분류코드 , '') AS 분류코드, ISNULL(T1.세부코드 , '') AS 세부코드, " _
                               & "0 AS 적정재고, 0 AS 최저재고, '' AS 최종입고일자, '' AS 최종출고일자, " _
                               & "0 AS 사용구분, " _
                               & "'" & PB_regUserinfoU.UserServerDate & "' AS 수정일자, '" & PB_regUserinfoU.UserCode & "' AS 사용자코드, " _
                               & "'' AS 비고란, '' AS 주매입처코드, " _
                               & "ISNULL(T2.입고단가1, 0) AS 입고단가1, ISNULL(T2.입고단가2, 0) AS 입고단가2, " _
                               & "ISNULL(T2.입고단가3, 0) AS 입고단가3, ISNULL(T2.출고단가1, 0) AS 출고단가1, " _
                               & "ISNULL(T2.출고단가3, 0) AS 출고단가2, ISNULL(T2.출고단가3, 0) AS 출고단가3 " _
                          & "FROM 자재 T1 " _
                         & "INNER JOIN 자재원장 T2 ON T2.사업장코드 = '01' AND T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드 " _
                         & "ORDER BY T1.분류코드, T1.세부코드 "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            .AddItem .Rows
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case Is <= 9
                            .TextMatrix(.Rows - 1, lngC) = Text1(lngC).Text
                            If lngC = 0 Then .Cell(flexcpData, .Rows - 1, lngC, .Rows - 1, lngC) = Text1(lngC).Text
                       Case 10
                            .TextMatrix(.Rows - 1, lngC) = Val(Trim(Text1(lngC).Text))
                            Select Case Val(Text1(lngC).Text)
                                   Case 0
                                        .TextMatrix(.Rows - 1, lngC + 1) = "정상"
                                   Case 9
                                        .TextMatrix(.Rows - 1, lngC + 1) = "사용불가"
                                   Case Else
                                        .TextMatrix(.Rows - 1, lngC + 1) = "구분오류"
                            End Select
                       Case 11 To 13
                            .TextMatrix(.Rows - 1, lngC + 2) = Text1(lngC).Text
                       Case 14
                            .TextMatrix(.Rows - 1, 17) = Vals(Text1(14).Text)
                       Case 15, 16
                            .TextMatrix(.Rows - 1, lngC + 3) = Val(Text1(lngC).Text)
                       Case 17 To 23, 24 To 28
                            .TextMatrix(.Rows - 1, lngC + 3) = Trim(Text1(lngC).Text)
                       Case 29 To 30
                            .TextMatrix(.Rows - 1, lngC + 3) = Trim(Text1(lngC).Text)
                       Case Else
                End Select
            Next lngC
            .TextMatrix(.Rows - 1, 12) = Format(IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)), "0000-00-00")
            .TextMatrix(.Rows - 1, 16) = Trim(.TextMatrix(.Rows - 1, 14)) & Space(1) & Trim(.TextMatrix(.Rows - 1, 15))
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(Text1.LBound).Enabled = False
            .Row = .Rows - 1          '자동으로 vsfg1_EnterCell Event 발생
         Else                                          '사업장 변경
            strSQL = "UPDATE 사업장 SET " _
                          & "사업장명 = '" & Trim(Text1(1).Text) & "', " _
                          & "사업자번호 = '" & Trim(Text1(2).Text) & "', " _
                          & "법인번호 = '" & Trim(Text1(3).Text) & "', " _
                          & "대표자명 = '" & Trim(Text1(4).Text) & "', " _
                          & "대표자주민번호 = '" & Trim(Text1(5).Text) & "', " _
                          & "개업일자 = '" & IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)) & "', " _
                          & "우편번호 = '" & Trim(Text1(11).Text) & "', " _
                          & "주소 = '" & Trim(Text1(12).Text) & "', 번지 = '" & Trim(Text1(13).Text) & "', " _
                          & "업태 = '" & Trim(Text1(6).Text) & "', 업종 = '" & Trim(Text1(7).Text) & "', " _
                          & "전화번호 = '" & Trim(Text1(8).Text) & "', 팩스번호 = '" & Trim(Text1(9).Text) & "', " _
                          & "부가세율 = " & Vals(Text1(14).Text) & ", " _
                          & "미지급금발생구분 = " & Val(Text1(15).Text) & ", 미수금발생구분 = " & Val(Text1(16).Text) & ", " _
                          & "사용구분 = " & Val(Text1(10).Text) & ", " _
                          & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', 사용자코드 = '" & PB_regUserinfoU.UserCode & "', " _
                          & "이메일주소 = '" & Trim(Text1(17).Text) & "', 홈페이지주소 = '" & Trim(Text1(18).Text) & "', " _
                          & "백업폴더 = '" & Trim(Text1(19).Text) & "', 자재기초마감년월 = '" & Trim(Text1(20).Text) & "', " _
                          & "미지급금기초마감년월 = '" & Trim(Text1(21).Text) & "', " _
                          & "미수금기초마감년월 = '" & Trim(Text1(22).Text) & "', 회계기초마감년월 = '" & Trim(Text1(23).Text) & "', " _
                          & "출력타입구분 = " & Vals(Trim(Text1(24).Text)) & ", " _
                          & "거래명세서상단마진 = " & Vals(Trim(Text1(25).Text)) & ",거래명세서왼쪽마진 = " & Vals(Trim(Text1(26).Text)) & ",  " _
                          & "세금계산서상단마진 = " & Vals(Trim(Text1(27).Text)) & ",세금계산서왼쪽마진 = " & Vals(Trim(Text1(28).Text)) & ", " _
                          & "최종입고단가자동갱신구분 = " & Vals(Trim(Text1(29).Text)) & ", " _
                          & "최종출고단가자동갱신구분 = " & Vals(Trim(Text1(30).Text)) & " " _
                    & "WHERE 사업장코드 = '" & Trim(Text1(Text1.LBound).Text) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            PB_adoCnnSQL.Execute strSQL
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case Is <= 9
                            .TextMatrix(.Row, lngC) = Text1(lngC).Text
                       Case 10
                            .TextMatrix(.Row, lngC) = Val(Trim(Text1(lngC).Text))
                            Select Case Val(Text1(lngC).Text)
                                   Case 0
                                        .TextMatrix(.Row, lngC + 1) = "정상"
                                   Case 9
                                        .TextMatrix(.Row, lngC + 1) = "사용불가"
                                   Case Else
                                        .TextMatrix(.Row, lngC + 1) = "구분오류"
                            End Select
                       Case 11 To 13
                            .TextMatrix(.Row, lngC + 2) = Text1(lngC).Text
                       Case 14
                            .TextMatrix(.Row, 17) = Vals(Trim(Text1(lngC).Text))
                       Case 15, 16
                            .TextMatrix(.Row, lngC + 3) = Val(Trim(Text1(lngC).Text))
                       Case 17 To 23, 24 To 30
                            .TextMatrix(.Row, lngC + 3) = Trim(Text1(lngC).Text)
                       Case Else
                End Select
            Next lngC
            .TextMatrix(.Row, 12) = Format(IIf(DTOS(dtpOpenDate.Value) = "19000101", "", DTOS(dtpOpenDate.Value)), "0000-00-00")
            .TextMatrix(.Row, 16) = Trim(.TextMatrix(.Row, 14)) & Space(1) & Trim(.TextMatrix(.Row, 15))
         End If
         PB_adoCnnSQL.CommitTrans
         
         '+--------+
         ' 전역변수
         '+--------+
         '(부가세율 : 현재사업장이고 부가세율이 변경되면)
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            Vals(PB_curVatRate) <> Vals(Text1(14).Text) Then
            PB_curVatRate = Vals(Text1(14).Text) / 100
         End If
         '(최종입고단가자동갱신구분 : 현재사업장이고 최종입고단가자동갱신구분이 변경되면)
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            Vals(PB_intIAutoPriceGbn) <> Vals(Text1(29).Text) Then
            PB_intIAutoPriceGbn = Vals(Trim(Text1(29).Text))
         End If
         '(최종출고단가자동갱신구분 : 현재사업장이고 최종출고단가자동갱신구분이 변경되면)
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            Vals(PB_intOAutoPriceGbn) <> Vals(Text1(30).Text) Then
            PB_intOAutoPriceGbn = Vals(Trim(Text1(30).Text))
         End If
         
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strEnterNo <> Trim(Text1(2).Text) Then
            PB_strEnterNo = Trim(Text1(2).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strEnterName <> Trim(Text1(1).Text) Then
            PB_strEnterName = Trim(Text1(1).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strRepName <> Trim(Text1(4).Text) Then
            PB_strRepName = Trim(Text1(4).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            PB_strEnterAddress <> (Trim(Text1(12).Text) + Space(1) + Trim(Text1(13).Text)) Then
            PB_strEnterAddress = Trim(Text1(12).Text) + Space(1) + Trim(Text1(13).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strUptae <> Trim(Text1(6).Text) Then
            PB_strUptae = Trim(Text1(6).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And PB_strUpjong <> Trim(Text1(7).Text) Then
            PB_strUpjong = Trim(Text1(7).Text)
         End If
                  
         '(출력타입구분 : 현재사업장이고 출력타입이 변경되면)
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intPrtTypeGbn) <> Vals(Text1(24).Text) Then
            PB_intPrtTypeGbn = Vals(Text1(24).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intDTopMargin) <> Vals(Text1(25).Text) Then
            PB_intDTopMargin = Vals(Text1(25).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intDLeftMargin) <> Vals(Text1(26).Text) Then
            PB_intDLeftMargin = Vals(Text1(26).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intTTopMargin) <> Vals(Text1(27).Text) Then
            PB_intTTopMargin = Vals(Text1(27).Text)
         End If
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And Vals(PB_intTLeftMargin) <> Vals(Text1(28).Text) Then
            PB_intTLeftMargin = Vals(Text1(28).Text)
         End If
         
         '+----------+
         ' 레지스트리
         '+----------+
         '사업장이름이 바뀐 경우 레지스트리 변경
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            PB_regUserinfoU.UserBranchName <> Trim(Text1(1).Text) Then
            frmMain.Caption = PB_strSystemName & " - " & Trim(Text1(1).Text)
            PB_regUserinfoU.UserBranchName = Trim(Text1(1).Text)
            UserinfoU_Save PB_regUserinfoU
         End If
         '미지급금발생구분이 바뀐 경우 레지스트리 변경
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            PB_regUserinfoU.UserMJGbn <> Trim(Text1(15).Text) Then
            PB_regUserinfoU.UserMJGbn = Trim(Text1(15).Text)
            UserinfoU_Save PB_regUserinfoU
         End If
         '미수금발생구분이 바뀐 경우 레지스트리 변경
         If PB_regUserinfoU.UserBranchCode = Trim(Text1(0).Text) And _
            PB_regUserinfoU.UserMSGbn <> Trim(Text1(16).Text) Then
            PB_regUserinfoU.UserMSGbn = Trim(Text1(16).Text)
            UserinfoU_Save PB_regUserinfoU
         End If
         .SetFocus
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 변경 실패"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// 삭제 ///
'+-----------+
Private Sub cmdDelete_Click()
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         If .Row >= .FixedRows Then
            intRetVal = MsgBox("등록된 자료를 삭제하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton2, "자료 삭제")
            If intRetVal = vbYes Then
               Screen.MousePointer = vbHourglass
               cmdDelete.Enabled = False
               '삭제전 관련테이블 검사
               'P_adoRec.CursorLocation = adUseClient
               'strSQL = "SELECT Count(*) AS 해당건수 FROM TableName " _
               '        & "WHERE 사업장구분 = " & .TextMatrix(.Row, 0) & " "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("해당건수")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "건)가 있으므로 삭제할 수 없습니다.", vbCritical, "사업장 삭제 불가"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "DELETE FROM 자재원장 WHERE 사업장코드 = " & .TextMatrix(.Row, 0) & " "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               strSQL = "DELETE FROM 사업장 WHERE 사업장코드 = " & .TextMatrix(.Row, 0) & " "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               Text1(Text1.LBound).Enabled = False
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarHorizontal
               End If
               Screen.MousePointer = vbDefault
               If .Rows = 1 Then
                  SubClearText
                  .Row = 0
                  Text1(Text1.LBound).Enabled = True
                  Text1(Text1.LBound).SetFocus
                  Exit Sub
               End If
               cmdDelete.Enabled = True
               vsfg1_EnterCell
               vsfg1.SetFocus
               vsfg1_AfterRowColChange 0, 0, 1, 1
            End If
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 삭제 실패"
    cmdDelete.Enabled = True
    vsfg1.SetFocus
    'Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

'+-----------+
'/// 종료 ///
'+-----------+
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    If P_adoRec.State <> adStateClosed Then
       P_adoRec.Close
    End If
    Set P_adoRec = Nothing
    Set frm사업장정보 = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfg1) 초기화 ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    Text1(Text1.LBound).Enabled = False                '사업장코드 FLASE
    With vsfg1              'Rows 0, Cols 34, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 2
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 34
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 850    '사업장코드
         .ColWidth(1) = 3000   '사업장명
         .ColWidth(2) = 1600   '사업자번호
         .ColWidth(3) = 1600   '법인번호
         .ColWidth(4) = 1000   '대표자명
         .ColWidth(5) = 1500   '대표자주민번호
         .ColWidth(6) = 2500   '업태
         .ColWidth(7) = 2500   '업종
         .ColWidth(8) = 2000   '전화번호
         .ColWidth(9) = 1600   '팩스번호
         .ColWidth(10) = 1     '사용구분
         .ColWidth(11) = 1000  '사용구분
         
         .ColWidth(12) = 1000  '개업일자
         .ColWidth(13) = 1000  '우편번호
         .ColWidth(14) = 1     '사업장주소
         .ColWidth(15) = 1     '사업장번지
         .ColWidth(16) = 7000  '사업장주소(주소+번지)
         .ColWidth(17) = 1000  '부가세율
         .ColWidth(18) = 1000  '미지급금발생구분
         .ColWidth(19) = 1000  '미수금발생구분
         
         .ColWidth(20) = 4600  '이메일주소
         .ColWidth(21) = 4600  '홈페이지주소
         .ColWidth(22) = 4600  '백업폴더
         .ColWidth(23) = 2000  '자재기초마감년월
         .ColWidth(24) = 2000  '미지급금기초마감년월
         .ColWidth(25) = 2000  '미수금기초마감년월
         .ColWidth(26) = 2000  '회계기초마감년월
         .ColWidth(27) = 2000  '출력타입구분
         .ColWidth(28) = 2000  '거래명세서상단마진
         .ColWidth(29) = 2000  '거래명세서왼쪽마진
         .ColWidth(30) = 2000  '세금계산서상단마진
         .ColWidth(31) = 2000  '세금계산서왼쪽마진
         .ColWidth(32) = 2000  '최종입고단가자동갱신구분
         .ColWidth(33) = 2000  '최종출고단가자동갱신구분
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "코드"
         .TextMatrix(0, 1) = "사업장명"
         .TextMatrix(0, 2) = "사업자번호"
         .TextMatrix(0, 3) = "법인번호"
         .TextMatrix(0, 4) = "대표자명"
         .TextMatrix(0, 5) = "주민번호"
         .TextMatrix(0, 6) = "업태"
         .TextMatrix(0, 7) = "업종"
         .TextMatrix(0, 8) = "전화번호"
         .TextMatrix(0, 9) = "팩스번호"
         .TextMatrix(0, 10) = "사용구분"
         .TextMatrix(0, 11) = "사용구분"
         .TextMatrix(0, 12) = "개업일자"
         .TextMatrix(0, 13) = "우편번호"
         .TextMatrix(0, 14) = "사업장주소"
         .TextMatrix(0, 15) = "사업장번지"
         .TextMatrix(0, 16) = "사업장주소" '주소+번지
         .TextMatrix(0, 17) = "부가세율"
         .TextMatrix(0, 18) = "미지급금"
         .TextMatrix(0, 19) = "미수금"
         .TextMatrix(0, 20) = "이메일주소"
         .TextMatrix(0, 21) = "홈페이지주소"
         .TextMatrix(0, 22) = "백업폴더"
         .TextMatrix(0, 23) = "자재기초마감년월"
         .TextMatrix(0, 24) = "미지급금기초마감년월"
         .TextMatrix(0, 25) = "미수금기초마감년월"
         .TextMatrix(0, 26) = "회계기초마감년월"
         .TextMatrix(0, 27) = "출력타입"
         .TextMatrix(0, 28) = "거래명세서상단마진"
         .TextMatrix(0, 29) = "거래명세서왼쪽마진"
         .TextMatrix(0, 30) = "세금계산서상단마진"
         .TextMatrix(0, 31) = "세금계산서왼쪽마진"
         .TextMatrix(0, 32) = "매입단가갱신"
         .TextMatrix(0, 33) = "매출단가갱신"
         .ColHidden(10) = True: .ColHidden(14) = True: .ColHidden(15) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1 To 9, 16
                        .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 11 To 13, 17 To 19, 23 To 26, 27, 32, 33
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 17
                        .ColFormat(17) = "#0.00"
             End Select
         Next lngC
         
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfg1) 채우기///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim SQL        As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
    P_adoRec.CursorLocation = adUseClient
    SQL = "SELECT * " _
          & "FROM 사업장 T1 " _
         & "ORDER BY T1.사업장코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Text1(Text1.LBound).Enabled = True
       Text1(Text1.LBound).SetFocus
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("사업장코드")
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow 사용을 위해
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("사업장명")), "", P_adoRec("사업장명"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("사업자번호")), "", P_adoRec("사업자번호"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("법인번호")), "", P_adoRec("법인번호"))
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("대표자명")), "", P_adoRec("대표자명"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("대표자주민번호")), "", P_adoRec("대표자주민번호"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("업태")), "", P_adoRec("업태"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("업종")), "", P_adoRec("업종"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("전화번호")), "", P_adoRec("전화번호"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("팩스번호")), "", P_adoRec("팩스번호"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("사용구분")), "", P_adoRec("사용구분"))
               Select Case .ValueMatrix(lngR, 10)
                      Case 0
                           .TextMatrix(lngR, 11) = "정상"
                      Case 9
                           .TextMatrix(lngR, 11) = "사용불가"
                      Case Else
                           .TextMatrix(lngR, 11) = "구분오류"
               End Select
               .TextMatrix(lngR, 12) = IIf(IsNull(P_adoRec("개업일자")), "", Format(P_adoRec("개업일자"), "0000-00-00"))
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("우편번호")), "", P_adoRec("우편번호"))
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("주소")), "", P_adoRec("주소"))
               .TextMatrix(lngR, 15) = IIf(IsNull(P_adoRec("번지")), "", P_adoRec("번지"))
               .TextMatrix(lngR, 16) = Trim(.TextMatrix(lngR, 14)) & Space(1) & Trim(.TextMatrix(lngR, 13))
               .TextMatrix(lngR, 17) = P_adoRec("부가세율")
               .TextMatrix(lngR, 18) = IIf(IsNull(P_adoRec("미지급금발생구분")), "", P_adoRec("미지급금발생구분"))
               .TextMatrix(lngR, 19) = IIf(IsNull(P_adoRec("미수금발생구분")), "", P_adoRec("미수금발생구분"))
               .TextMatrix(lngR, 20) = IIf(IsNull(P_adoRec("이메일주소")), "", P_adoRec("이메일주소"))
               .TextMatrix(lngR, 21) = IIf(IsNull(P_adoRec("홈페이지주소")), "", P_adoRec("홈페이지주소"))
               .TextMatrix(lngR, 22) = IIf(IsNull(P_adoRec("백업폴더")), "", P_adoRec("백업폴더"))
               .TextMatrix(lngR, 23) = IIf(IsNull(P_adoRec("자재기초마감년월")), "", P_adoRec("자재기초마감년월"))
               .TextMatrix(lngR, 24) = IIf(IsNull(P_adoRec("미지급금기초마감년월")), "", P_adoRec("미지급금기초마감년월"))
               .TextMatrix(lngR, 25) = IIf(IsNull(P_adoRec("미수금기초마감년월")), "", P_adoRec("미수금기초마감년월"))
               .TextMatrix(lngR, 26) = IIf(IsNull(P_adoRec("회계기초마감년월")), "", P_adoRec("회계기초마감년월"))
               
               .TextMatrix(lngR, 27) = P_adoRec("출력타입구분")
               .TextMatrix(lngR, 28) = P_adoRec("거래명세서상단마진")
               .TextMatrix(lngR, 29) = P_adoRec("거래명세서왼쪽마진")
               .TextMatrix(lngR, 30) = P_adoRec("세금계산서상단마진")
               .TextMatrix(lngR, 31) = P_adoRec("세금계산서왼쪽마진")
               .TextMatrix(lngR, 32) = P_adoRec("최종입고단가자동갱신구분")
               .TextMatrix(lngR, 33) = P_adoRec("최종출고단가자동갱신구분")
               
               If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
                  lngRR = lngR
               End If
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Rows - PC_intRowCnt + .FixedRows
               End If
               Text1(Text1.LBound).Enabled = True
               Text1(Text1.LBound).SetFocus
               Exit Sub
            Else
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Row
               End If
            End If
            vsfg1_EnterCell       'vsfg1_EnterCell 자동실행(만약 한건 일때도 강제로 자동실행)
            .SetFocus
            If lngRR <> 0 Then
               vsfg1_AfterRowColChange 0, 0, 1, 1
            End If
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------------------+
'/// Clear text1(index) ///
'+-------------------------+
Private Sub SubClearText()
Dim lngC As Long
    For lngC = Text1.LBound To Text1.UBound
        Text1(lngC).Text = ""
    Next lngC
    dtpOpenDate.Value = Format("19000101", "0000-00-00")
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '사업장코드
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "00")
                    If Not (Text1(lngC).Text >= "01" And Text1(lngC).Text <= "99") Then
                       Exit Function
                    End If
               Case 5  '주민번호
                    If Len(Trim(Text1(lngC).Text)) > 14 Then
                       Exit Function
                    End If
               Case 10  '사용구분
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0")
                    If Not (Text1(lngC).Text >= "0" And Text1(lngC).Text <= "9") Then
                       Exit Function
                    End If
               Case 15, 16  '미지급금발생, 미수금발생
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0")
                    If Not (Text1(lngC).Text >= "1" And Text1(lngC).Text <= "2") Then
                       Exit Function
                    End If
               Case 17, 19  '17.이메일주소, 18.홈페이지주소, 19.백업폴더
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 50) Then
                       Exit Function
                    End If
               Case 20 To 23  '기초마감년월
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) = 6) Then
                       Exit Function
                    End If
               Case 24 To 28  '출력타입구분
                    Text1(lngC).Text = Val(Trim(Text1(lngC).Text))
                    If Len(Trim(Text1(lngC).Text)) > 2 Then
                       Exit Function
                    End If
               Case 29 To 30  '단가변경
                    Text1(lngC).Text = Val(Trim(Text1(lngC).Text))
                    If Val(Trim(Text1(lngC).Text)) > 1 Then
                       Exit Function
                    End If
               Case Else
        End Select
    Next lngC
    blnOK = True
End Function

