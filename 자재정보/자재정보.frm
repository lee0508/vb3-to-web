VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm자재정보 
   BorderStyle     =   0  '없음
   Caption         =   "자재정보"
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
      TabIndex        =   23
      Top             =   0
      Width           =   15195
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4920
         Top             =   200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowLeft      =   0
         WindowTop       =   0
         WindowWidth     =   15405
         WindowHeight    =   11100
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "자재정보.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   45
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "자재정보.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   27
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "자재정보.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "자재정보.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "자재정보.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "자재정보.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   38
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "자 재 코 드 관 리"
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
         TabIndex        =   24
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   6450
      Left            =   60
      TabIndex        =   19
      Top             =   3585
      Width           =   15195
      _cx             =   26802
      _cy             =   11377
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
      Height          =   2925
      Left            =   60
      TabIndex        =   20
      Top             =   630
      Width           =   15195
      Begin VB.CheckBox chkCodeException 
         Caption         =   "CODE 품목 제외"
         Height          =   180
         Left            =   10860
         TabIndex        =   67
         Top             =   240
         Value           =   1  '확인
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   8
         Left            =   1155
         TabIndex        =   11
         Top             =   2535
         Width           =   9465
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   9
         Left            =   11985
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1755
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   10
         Left            =   11985
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2115
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   11
         Left            =   11985
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2475
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   12
         Left            =   13410
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1755
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   13
         Left            =   13410
         MaxLength       =   20
         TabIndex        =   16
         Top             =   2115
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   14
         Left            =   13410
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2475
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   6
         Left            =   7920
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1440
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   7
         Left            =   7920
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1800
         Width           =   2685
      End
      Begin VB.TextBox txtSebuCodeRe 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   3795
         MaxLength       =   16
         TabIndex        =   55
         Top             =   570
         Width           =   2055
      End
      Begin VB.ComboBox cboMtGpRe 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1155
         Style           =   2  '드롭다운 목록
         TabIndex        =   53
         Top             =   570
         Width           =   1575
      End
      Begin VB.TextBox txtFindBarCode 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   13400
         MaxLength       =   13
         TabIndex        =   36
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox txtFindCD 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   10860
         MaxLength       =   18
         TabIndex        =   33
         Top             =   585
         Width           =   1605
      End
      Begin VB.TextBox txtFindSZ 
         Appearance      =   0  '평면
         Height          =   285
         Left            =   10860
         MaxLength       =   30
         TabIndex        =   35
         Top             =   945
         Width           =   1605
      End
      Begin VB.TextBox txtFindNM 
         Appearance      =   0  '평면
         Height          =   285
         Left            =   7920
         MaxLength       =   30
         TabIndex        =   34
         Top             =   945
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   2
         Left            =   4500
         MaxLength       =   13
         TabIndex        =   3
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   5
         Left            =   1155
         MaxLength       =   8
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   1  '입력 상태 설정
         Index           =   4
         Left            =   1155
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   1
         Left            =   4500
         Style           =   2  '드롭다운 목록
         TabIndex        =   7
         Top             =   1440
         Width           =   1350
      End
      Begin VB.ComboBox cboTaxGbn 
         Height          =   300
         Left            =   4515
         Style           =   2  '드롭다운 목록
         TabIndex        =   8
         Top             =   1800
         Width           =   1350
      End
      Begin VB.ComboBox cboMtGp 
         Height          =   300
         Index           =   1
         Left            =   7920
         Style           =   2  '드롭다운 목록
         TabIndex        =   32
         Top             =   555
         Width           =   1850
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   0
         Left            =   13650
         Style           =   2  '드롭다운 목록
         TabIndex        =   31
         Top             =   555
         Width           =   1100
      End
      Begin VB.ComboBox cboMtGp 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1155
         Style           =   2  '드롭다운 목록
         TabIndex        =   0
         Top             =   200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   3
         Left            =   1155
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1440
         Width           =   2400
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         Index           =   1
         Left            =   1140
         MaxLength       =   30
         TabIndex        =   2
         Top             =   945
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   3795
         MaxLength       =   16
         TabIndex        =   1
         Top             =   225
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처명"
         Height          =   240
         Index           =   27
         Left            =   6520
         TabIndex        =   66
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   " 매입단가 "
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   12105
         TabIndex        =   65
         ToolTipText     =   "자재원장"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   " 매출단가 "
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   13545
         TabIndex        =   64
         ToolTipText     =   "자재원장"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "1."
         Height          =   240
         Index           =   20
         Left            =   11640
         TabIndex        =   63
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "2."
         Height          =   240
         Index           =   21
         Left            =   11640
         TabIndex        =   62
         Top             =   2175
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "3."
         Height          =   240
         Index           =   22
         Left            =   11640
         TabIndex        =   61
         Top             =   2535
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처코드"
         Height          =   240
         Index           =   17
         Left            =   6520
         TabIndex        =   60
         ToolTipText     =   "자재원장"
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "[Home]"
         Height          =   240
         Index           =   16
         Left            =   9225
         TabIndex        =   59
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "비고란"
         Height          =   240
         Index           =   23
         Left            =   195
         TabIndex        =   58
         ToolTipText     =   "자재원장"
         Top             =   2595
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   ")"
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   15
         Left            =   5880
         TabIndex        =   57
         ToolTipText     =   "코드변경"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "[Home]"
         Height          =   240
         Index           =   14
         Left            =   6120
         TabIndex        =   56
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "세부코드"
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   13
         Left            =   2835
         TabIndex        =   54
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "분류"
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   12
         Left            =   550
         TabIndex        =   52
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "=>("
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   51
         ToolTipText     =   "코드변경"
         Top             =   600
         Width           =   375
      End
      Begin VB.Line Line2 
         X1              =   7080
         X2              =   7080
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "바코드"
         Height          =   240
         Index           =   9
         Left            =   12480
         TabIndex        =   50
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "[Home]"
         Height          =   240
         Index           =   8
         Left            =   6120
         TabIndex        =   49
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "세부코드"
         Height          =   240
         Index           =   7
         Left            =   9975
         TabIndex        =   48
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "규격"
         Height          =   240
         Index           =   5
         Left            =   10200
         TabIndex        =   47
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품명"
         Height          =   240
         Index           =   11
         Left            =   7140
         TabIndex        =   46
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "바코드"
         Height          =   240
         Index           =   10
         Left            =   3660
         TabIndex        =   44
         Top             =   1005
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   15000
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "폐기율"
         Height          =   240
         Index           =   6
         Left            =   195
         TabIndex        =   43
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "과세구분"
         Height          =   240
         Index           =   3
         Left            =   3660
         TabIndex        =   42
         Top             =   1845
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "단위"
         Height          =   240
         Index           =   2
         Left            =   195
         TabIndex        =   41
         Top             =   1845
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품명"
         Height          =   240
         Index           =   36
         Left            =   300
         TabIndex        =   40
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "(검색조건)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   35
         Left            =   9600
         TabIndex        =   39
         ToolTipText     =   "300"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "분류"
         Height          =   240
         Index           =   34
         Left            =   7005
         TabIndex        =   37
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "규격"
         Height          =   240
         Index           =   26
         Left            =   195
         TabIndex        =   30
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사용구분"
         Height          =   240
         Index           =   25
         Left            =   12800
         TabIndex        =   29
         Top             =   615
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사용구분"
         Height          =   240
         Index           =   24
         Left            =   3660
         TabIndex        =   28
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "세부코드"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   2835
         TabIndex        =   22
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "분류"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   21
         Top             =   255
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm자재정보"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 자재정보
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 자재, (자재분류)
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private P_strFindString1   As String
Private P_strFindString2   As String
Private P_strFindString3   As String
Private P_strFindString4   As String
Private Const PC_intRowCnt As Integer = 20  '그리드 한 페이지 당 행수(FixedRows 포함)

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
       SubOther_FILL
       frmMain.SBar.Panels(4).Text = ""
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재 (서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+-----------------------+
'/// cboMtGp(index) ///
'+-----------------------+
Private Sub cboMtGp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Select Case Index
       End Select
       SendKeys "{tab}"
    End If
End Sub
'+------------------------------+
'/// cboMtGpRe(분류코드변경) ///
'+------------------------------+
Private Sub cboMtGpRe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtSebuCodeRe.SetFocus
    End If
End Sub
Private Sub txtSebuCodeRe_GotFocus()
    With txtSebuCodeRe
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtSebuCodeRe_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strExeFile As String
Dim varRetVal  As Variant
Dim inti       As Integer
    If (Len(Trim(txtSebuCodeRe.Text)) > 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '품목코드 변경시에만
       PB_strCallFormName = "frm자재정보"
       PB_strMaterialsCode = Trim(txtSebuCodeRe.Text) 'Mid(cboMtGp(0).Text, 1, 2)
       PB_strMaterialsName = "" 'Trim(Text1(1).Text)
       frm자재검색.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '검색에서 취소(ESC)
       Else
          For inti = 0 To cboMtGpRe.ListCount - 1
              cboMtGpRe.ListIndex = inti
              If Mid(cboMtGpRe.Text, 1, 2) = Mid(PB_strMaterialsCode, 1, 2) Then
                 Exit For
              End If
          Next inti
          txtSebuCodeRe.Text = Mid(PB_strMaterialsCode, 3) '세부코드
       End If
       'If PB_strMaterialsCode = "" Then
       '   PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
       'Else
          PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
       'End If
       Text1(1).SetFocus
    Else
       If KeyCode = vbKeyReturn Then
          Text1(1).SetFocus
       End If
    End If
End Sub

Private Sub txtSebuCodeRe_LostFocus()
    txtSebuCodeRe.Text = UPPER(Trim(txtSebuCodeRe.Text))
End Sub
'+-----------------------+
'/// cboState(index) ///
'+-----------------------+
Private Sub cboState_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Select Case Index
       End Select
       SendKeys "{tab}"
    End If
End Sub
'+----------------+
'/// cboTaxGbn ///
'+----------------+
Private Sub cboTaxGbn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
'+-------------------+
'/// dtpInputDate ///
'+-------------------+
Private Sub dtpInputDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
'+--------------------+
'/// dtpOutputDate ///
'+--------------------+
Private Sub dtpOutputDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
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
Dim inti       As Integer
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '자재정보 추가시에만
       PB_strCallFormName = "frm자재정보"
       PB_strMaterialsCode = (Text1(0).Text) 'Mid(cboMtGp(0).Text, 1, 2)
       PB_strMaterialsName = "" 'Trim(Text1(1).Text)
       frm자재검색.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '검색에서 취소(ESC)
       Else
          For inti = 0 To cboMtGp(0).ListCount - 1
              cboMtGp(0).ListIndex = inti
              If Mid(cboMtGp(0).Text, 1, 2) = Mid(PB_strMaterialsCode, 1, 2) Then
                 Exit For
              End If
          Next inti
          Text1(0).Text = Mid(PB_strMaterialsCode, 3) '세부코드
          Text1(1).Text = PB_strMaterialsName         '품명
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    ElseIf _
       (Index = 6 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '매입처검색
       PB_strFMCCallFormName = "frm자재정보"
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
       frm매입처검색.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '검색에서 취소(ESC)
       Else
          Text1(Index).Text = PB_strSupplierCode
          Text1(Index + 1).Text = PB_strSupplierName
       End If
       If PB_strSupplierCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strSupplierCode = "": PB_strSupplierName = ""
    Else
       If KeyCode = vbKeyReturn Then
          Select Case Index
          End Select
          SendKeys "{tab}"
       End If
    End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
Dim lngR   As Long
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = UPPER(Trim(.Text))
                     If Len(Text1(Index).Text) < 1 Then
                        Text1(Index).Text = ""
                        Exit Sub
                     End If
                     'If Text1(Index).Enabled = True Then
                     '   P_adoRec.CursorLocation = adUseClient
                     '   strSQL = "SELECT * FROM 자재 " _
                     '           & "WHERE 분류코드 = '" & Left(cboMtGp(0).List(lngR), 2) & "' " _
                     '             & "AND 세부코드 = '" & Trim(.Text) & "' "
                     '   On Error GoTo ERROR_TABLE_SELECT
                     '   P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                     '   If P_adoRec.RecordCount <> 0 Then
                     '      P_adoRec.Close
                     '      .Text = ""
                     '      .SetFocus
                     '      Exit Sub
                     '   End If
                     '   P_adoRec.Close
                     'End If
                Case 6
                     .Text = UPPER(Trim(.Text))
                Case 5, 9 To 14
                     .Text = Format(Vals(Trim(.Text)), "#,0.00")
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+--------------+
'/// txtFind ///
'+--------------+
'+-------------------------------------+
'/// chkCodeException(코드품목제외) ///
'+-------------------------------------+
Private Sub chkCodeException_Click()
    cboMtGp(1).SetFocus
End Sub
Private Sub txtFindCD_GotFocus()
    With txtFindCD
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindCD_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       txtFindNM.SetFocus
    End If
End Sub
Private Sub txtFindCD_LostFocus()
    With txtFindCD
         .Text = UPPER(Trim(.Text))
    End With
End Sub
'+--------------+
'/// txtFind ///
'+--------------+
Private Sub txtFindNM_GotFocus()
    With txtFindNM
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindNM_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       txtFindSZ.SetFocus
    End If
End Sub
'+--------------+
'/// txtFind ///
'+--------------+
Private Sub txtFindSZ_GotFocus()
    With txtFindSZ
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindSZ_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       txtFindBarCode.SetFocus
    End If
End Sub
'+--------------+
'/// txtFind ///
'+--------------+
Private Sub txtFindBarCode_GotFocus()
    With txtFindBarCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       cmdFind_Click
    End If
End Sub

'+-----------+
'/// Grid ///
'+-----------+
Private Sub vsfg1_BeforeSort(ByVal Col As Long, Order As Integer)
    With vsfg1
         'P_strFindString2 = Trim(.Cell(flexcpData, .Row, 0)) 'Not Used
    End With
End Sub
Private Sub vsfg1_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfg1
         'If .FindRow(P_strFindString2, , 0) > 0 Then
         '   .Row = .FindRow(P_strFindString2, , 0) 'Not Used
         'End If
         'If PC_intRowCnt < .Rows Then
         '   .TopRow = .Row
         'End If
    End With
End Sub
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
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfg1
         If .Row >= .FixedRows Then
            If KeyCode = vbKeyF1 Then '자재시세검색
               'PB_strFMCCallFormName = "frm자재정보"
               'PB_strMaterialsCode = .TextMatrix(.Row, 4)
               'PB_strMaterialsName = .TextMatrix(.Row, 3)
               'PB_strSupplierCode = ""
               'frm자재시세검색.Show vbModal
            ElseIf _
               KeyCode = vbKeyF2 Then '자재원장
               'PB_strFMWCallFormName = "frm자재정보"
               'PB_strMaterialsCode = .TextMatrix(.Row, 4)
               'PB_strMaterialsName = .TextMatrix(.Row, 3)
               'frm자재시세.Show vbModal
               'MsgBox "자재원장"
            End If
         End If
    End With
End Sub
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = Trim(.Cell(flexcpData, .Row, 4))
            Select Case .MouseCol
                   Case 0, 2
                        .ColSel = 2
                        .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(1) = flexSortNone
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 1
                        .ColSel = 2
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 4) > 0 Then
               .Row = .FindRow(strData, , 4)
            End If
            If PC_intRowCnt < .Rows Then
               .TopRow = .Row
            End If
         End If
    End With
End Sub
Private Sub vsfg1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         cboMtGp(cboMtGp.LBound).Enabled = False
         Text1(Text1.LBound).Enabled = False
         cboMtGpRe.Enabled = True
         txtSebuCodeRe.Enabled = True
         txtSebuCodeRe.Text = ""
         If .Row >= .FixedRows And OldRow <> NewRow Then
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case 0
                            cboMtGp(0).Text = .TextMatrix(.Row, 0) + ". " + .TextMatrix(.Row, 1)
                       Case 2 To 3
                            Text1(lngC - 2).Text = .TextMatrix(.Row, lngC) '세부코드
                       Case 16 '바코드
                            Text1(2).Text = .TextMatrix(.Row, lngC)
                       Case 5 To 6 '규격, 단위
                            Text1(lngC - 2).Text = .TextMatrix(.Row, lngC)
                       Case 17 '폐기율
                            Text1(5).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
                       Case 22 '사용구분 listindex
                            cboState(1).ListIndex = .ValueMatrix(.Row, lngC)
                       Case 19 '과세구분 listindex
                            cboTaxGbn.ListIndex = .ValueMatrix(.Row, lngC)
                       Case 7 To 8 '매입처
                            Text1(lngC - 1).Text = .TextMatrix(.Row, lngC)
                       Case 15 '비고란
                            Text1(8).Text = .TextMatrix(.Row, lngC)
                       Case 9 To 14
                            Text1(lngC).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
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
         'cboMtGp(cboMtGp.LBound).Enabled = False
         'Text1(Text1.LBound).Enabled = False
         'cboMtGpRe.Enabled = True
         'txtSebuCodeRe.Enabled = True
         'txtSebuCodeRe.Text = ""
    End With
End Sub

'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    SubClearText
    vsfg1.Row = 0
    cboMtGp(cboMtGp.LBound).Enabled = True
    Text1(Text1.LBound).Enabled = True  'Log Counter사용시에는 False
    cboMtGpRe.Enabled = False
    txtSebuCodeRe.Enabled = False
    cboMtGp(cboMtGp.LBound).SetFocus
End Sub
'+-----------+
'/// 조회 ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    P_strFindString1 = Trim(txtFindCD.Text)       '조회할 경우 검색할 자재코드 보관
    P_strFindString2 = Trim(txtFindNM.Text)       '조회할 경우 검색할 자재명 보관
    P_strFindString3 = Trim(txtFindSZ.Text)       '조회할 경우 검색할 규격 보관
    P_strFindString4 = Trim(txtFindBarCode.Text)  '조회할 경우 검색할 바코드 보관
    SubClearText
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL        As String
Dim lngR          As Long
Dim lngC          As Long
Dim blnOK         As Boolean
Dim intRetVal     As Integer
Dim strServerDate As String
Dim strServerTime As String
    '추가이면 이미있는 품목인지 검사
    If Text1(0).Enabled = True Then
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT 분류코드, 세부코드 FROM 자재 " _
               & "WHERE 분류코드 = '" & Left(cboMtGp(0).Text, 2) & "' " _
                 & "AND 세부코드 = '" & Trim(Text1(0).Text) & "' "
       On Error GoTo ERROR_TABLE_SELECT
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       If P_adoRec.RecordCount <> 0 Then
          P_adoRec.Close
          MsgBox "이미 등록이 완료된 품목입니다. 확인후 다시 입력하여 주세요.", vbCritical, "품목 중복 등록"
          Text1(0).SetFocus
          Exit Sub
       End If
       P_adoRec.Close
    End If
    If cboMtGp(cboMtGp.LBound).Enabled = False Then  '갱신이면
       If Len(Trim(Text1(Text1.LBound))) = 0 Then
          Text1(1).SetFocus
          Exit Sub
       End If
       If Len(Trim(txtSebuCodeRe.Text)) > 0 Then     '코드변경이면
          P_adoRec.CursorLocation = adUseClient
          strSQL = "SELECT 분류코드, 세부코드 FROM 자재 " _
                  & "WHERE 분류코드 = '" & Left(cboMtGpRe.Text, 2) & "' " _
                    & "AND 세부코드 = '" & Trim(txtSebuCodeRe.Text) & "' "
          On Error GoTo ERROR_TABLE_SELECT
          P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          If P_adoRec.RecordCount <> 0 Then
             P_adoRec.Close
             MsgBox "이미 등록이 완료된 품목입니다. 확인후 다시 입력하여 주세요.", vbCritical, "품목코드 변경 불가"
             txtSebuCodeRe.SetFocus
             Exit Sub
          End If
          P_adoRec.Close
       End If
    End If
    '바코드 중복체크
    If Len(Trim(Text1(2).Text)) > 0 Then
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT 분류코드, 세부코드, 바코드 FROM 자재 " _
               & "WHERE NOT (분류코드 = '" & Left(cboMtGp(0).Text, 2) & "' " _
                 & "AND 세부코드 = '" & Trim(Text1(0).Text) & "') " _
                 & "AND 바코드 = '" & Trim(Text1(2).Text) & "' "
       On Error GoTo ERROR_TABLE_SELECT
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       If P_adoRec.RecordCount <> 0 Then
          P_adoRec.Close
          MsgBox "이미 등록이 완료된 바코드입니다. 확인후 다시 입력하여 주세요.", vbCritical, "품목 바코드 중복 등록"
          Text1(2).SetFocus
          Exit Sub
       End If
       P_adoRec.Close
    End If
    '입력내역 검사
    blnOK = False
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       'If lngC = 0 Then
       '   Text1(lngC).Enabled = True
       'End If
       Text1(lngC).SetFocus
       Exit Sub
    End If
    If cboMtGp(cboMtGp.LBound).Enabled = True Then
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
         PB_adoCnnSQL.BeginTrans
         If cboMtGp(cboMtGp.LBound).Enabled = True Then '자재 추가면 검색  '로그
            'strSQL = "SELECT * FROM 자재 " _
            '        & "WHERE 자재코드 = '" & Trim(Text1(Text1.LBound).Text) & "' "
            'On Error GoTo ERROR_TABLE_SELECT
            'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            'If P_adoRec.RecordCount <> 0 Then
            '   P_adoRec.Close
            '   Text1(Text1.LBound).SetFocus
            '   Screen.MousePointer = vbDefault
            '   cmdSave.Enabled = True
            '   Exit Sub
            'End If
            'P_adoRec.Close
            '// Log Counter
            'P_adoRec.CursorLocation = adUseClient
            'strSQL = "spLogCounter '자재', '" & Left(cboMtGp(0).Text, 2) & "', 0, 0, '" & PB_regUserinfoU.UserCode & "','' "
            'On Error GoTo ERROR_STORED_PROCEDURE
            'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            'Text1(0).Text = Format(P_adoRec(0), "0000")
            'P_adoRec.Close
         End If
         If cboMtGp(cboMtGp.LBound).Enabled = True Then '자재 추가
            strSQL = "INSERT INTO 자재(분류코드, 세부코드, " _
                                    & "자재명, 바코드, 규격," _
                                    & "단위, 폐기율, " _
                                    & "과세구분, 적요, 사용구분, " _
                                    & "수정일자, 사용자코드) VALUES( " _
                    & "'" & Left(Trim(cboMtGp(0).Text), 2) & "', '" & Trim(Text1(0).Text) & "', " _
                    & "'" & Trim(Text1(1).Text) & "', '" & Trim(Text1(2).Text) & "', '" & Trim(Text1(3).Text) & "', " _
                    & "'" & Trim(Text1(4).Text) & "'," & Vals(Trim(Text1(5).Text)) & "," _
                    & "" & Vals(Left(cboTaxGbn.Text, 1)) & ", '', " & Vals(Left(cboState(1).Text, 1)) & ", " _
                    & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            '자재원장추가
            strSQL = "INSERT INTO 자재원장 " _
                   & "SELECT '" & PB_regUserinfoU.UserBranchCode & "' AS 사업장코드, " _
                         & "ISNULL(T1.분류코드 , '') AS 분류코드, ISNULL(T1.세부코드 , '') AS 세부코드, " _
                         & "0 AS 적정재고, 0 AS 최저재고, '' AS 최종입고일자, '' AS 최종출고일자, " _
                         & "T1.사용구분 AS 사용구분, T1.수정일자 AS 수정일자, T1.사용자코드 AS 사용자코드, " _
                         & "'" & Trim(Text1(8).Text) & "' AS 비고란, '" & Trim(Text1(6).Text) & "' AS 주매입처코드, " _
                         & "" & Vals(Trim(Text1(9).Text)) & " AS 입고단가1, " & Vals(Trim(Text1(10).Text)) & " AS 입고단가2, " _
                         & "" & Vals(Trim(Text1(11).Text)) & " AS 입고단가3, " & Vals(Trim(Text1(12).Text)) & " AS 출고단가1, " _
                         & "" & Vals(Trim(Text1(13).Text)) & " AS 출고단가2, " & Vals(Trim(Text1(14).Text)) & " AS 출고단가3 " _
                     & "FROM 자재 T1 " _
                    & "WHERE T1.분류코드 = '" & Left(Trim(cboMtGp(0).Text), 2) & "' AND T1.세부코드 = '" & Trim(Text1(0).Text) & "' "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            P_adoRec.CursorLocation = adUseClient
            strSQL = "SELECT 사업장코드 FROM 사업장 " _
                    & "WHERE 사업장코드 <> '" & PB_regUserinfoU.UserBranchCode & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount = 0 Then
               P_adoRec.Close
            Else
               Do Until P_adoRec.EOF
                  strSQL = "INSERT INTO 자재원장 " _
                         & "SELECT '" & P_adoRec("사업장코드") & "' AS 사업장코드, " _
                                & "ISNULL(T1.분류코드 , '') AS 분류코드, ISNULL(T1.세부코드 , '') AS 세부코드, " _
                                & "0 AS 적정재고, 0 AS 최저재고, '' AS 최종입고일자, '' AS 최종출고일자, " _
                                & "T1.사용구분 AS 사용구분, T1.수정일자 AS 수정일자, T1.사용자코드 AS 사용자코드, " _
                                & "'' AS 비고란, '' AS 주매입처코드, " _
                                & "T1.입고단가1 AS 입고단가1, T1.입고단가2 AS 입고단가2, T1.입고단가3 AS 입고단가3, " _
                                & "T1.출고단가1 AS 출고단가1, T1.출고단가2 AS 출고단가2, T1.출고단가3 AS 출고단가3 " _
                           & "FROM 자재 T1 " _
                          & "WHERE T1.분류코드 = '" & Left(Trim(cboMtGp(0).Text), 2) & "' AND T1.세부코드 = '" & Trim(Text1(0).Text) & "' "
                  On Error GoTo ERROR_TABLE_INSERT
                  PB_adoCnnSQL.Execute strSQL
                  P_adoRec.MoveNext
               Loop
               P_adoRec.Close
            End If
            .AddItem .Rows
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 0
                            .TextMatrix(.Rows - 1, 0) = Left(cboMtGp(0).Text, 2)
                            .TextMatrix(.Rows - 1, 1) = Right(Trim(cboMtGp(0).Text), Len(Trim(cboMtGp(0).Text)) - 4)
                            .TextMatrix(.Rows - 1, 2) = Trim(Text1(0).Text)
                            .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 0) + .TextMatrix(.Rows - 1, 2)
                            .Cell(flexcpData, .Rows - 1, 4, .Rows - 1, 4) = .TextMatrix(.Rows - 1, 4)
                            .TextMatrix(.Rows - 1, 18) = Vals(Left(cboTaxGbn.Text, 1)) '과세구분
                            .TextMatrix(.Rows - 1, 19) = cboTaxGbn.ListIndex
                            .TextMatrix(.Rows - 1, 20) = Right(Trim(cboTaxGbn.Text), Len(Trim(cboTaxGbn.Text)) - 3)
                            .TextMatrix(.Rows - 1, 21) = Vals(Left(cboState(1).Text, 1)) '사용구분
                            .TextMatrix(.Rows - 1, 22) = cboState(1).ListIndex
                            .TextMatrix(.Rows - 1, 23) = Right(Trim(cboState(1).Text), Len(Trim(cboState(1).Text)) - 3)
                       Case 1 '1.품명
                            .TextMatrix(.Rows - 1, 3) = Trim(Text1(1).Text)
                       Case 2 '2.바코드
                            .TextMatrix(.Rows - 1, 16) = Trim(Text1(lngC).Text)
                       Case 3 To 4 '3.규격, 4.단위
                            .TextMatrix(.Rows - 1, lngC + 2) = Trim(Text1(lngC).Text)
                       Case 5 '5.폐기율
                            .TextMatrix(.Rows - 1, 17) = Vals(Trim(Text1(lngC).Text))
                       Case 6 To 7 '6.주매입처코드, 7.주매입처명
                            .TextMatrix(.Rows - 1, lngC + 1) = Trim(Text1(lngC).Text)
                       Case 8 '8.비고란
                            .TextMatrix(.Rows - 1, 15) = Trim(Text1(lngC).Text)
                       Case 9 To 14
                            .TextMatrix(.Rows - 1, lngC) = Vals(Trim(Text1(lngC).Text))
                       Case Else
                End Select
            Next lngC
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            cboMtGp(0).Enabled = False
            .Row = .Rows - 1          '자동으로 vsfg1_EnterCell Event 발생
         Else                                          '자재 변경
            strSQL = "UPDATE 자재 SET " _
                          & "자재명 = '" & Trim(Text1(1).Text) & "', 바코드 = '" & Trim(Text1(2).Text) & "', " _
                          & "규격 = '" & Trim(Text1(3).Text) & "', " _
                          & "단위 = '" & Trim(Text1(4).Text) & "', 폐기율 = " & Vals(Trim(Text1(5).Text)) & ", " _
                          & "과세구분 = " & Vals(Left(cboTaxGbn.Text, 1)) & ", " _
                          & "사용구분 = " & Vals(Left(cboState(1).Text, 1)) & ", " _
                          & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE 분류코드 = '" & Left(Trim(cboMtGp(0).Text), 2) & "' " _
                      & "AND 세부코드 = '" & Trim(Text1(0).Text) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            PB_adoCnnSQL.Execute strSQL
            strSQL = "UPDATE 자재원장 SET " _
                          & "주매입처코드 = '" & Trim(Text1(6).Text) & "', 입고단가1 = " & Vals(Trim(Text1(9).Text)) & ", " _
                          & "입고단가2 = " & Vals(Trim(Text1(10).Text)) & ", 입고단가3 = " & Vals(Trim(Text1(11).Text)) & ", " _
                          & "출고단가1 = " & Vals(Trim(Text1(12).Text)) & ", 출고단가2 = " & Vals(Trim(Text1(13).Text)) & ", " _
                          & "출고단가3 = " & Vals(Trim(Text1(14).Text)) & ", 비고란 = '" & Trim(Text1(8).Text) & "', " _
                          & "사용구분 = " & Vals(Left(cboState(1).Text, 1)) & ", " _
                          & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND 분류코드 = '" & Left(Trim(cboMtGp(0).Text), 2) & "' " _
                      & "AND 세부코드 = '" & Trim(Text1(0).Text) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            PB_adoCnnSQL.Execute strSQL
            If Len(txtSebuCodeRe.Text) > 0 Then '코드변경이면
               '자재코드변경 테이블 추가
               intRetVal = MsgBox(Left(cboMtGp(0).Text, 2) + Trim(Text1(0).Text) + " 코드를 " _
                         & Left(cboMtGpRe.Text, 2) + Trim(txtSebuCodeRe.Text) + " 로 변경하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton2, "품목코드 변경")
               If intRetVal = vbYes Then
                  '서버시간을 가져옴
                  P_adoRec.CursorLocation = adUseClient
                  strSQL = "SELECT CONVERT(VARCHAR(8),GETDATE(), 112) AS 서버일자, " _
                          & "RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
                  On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
                  P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                  strServerDate = P_adoRec("서버일자")
                  strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) _
                                + Mid(P_adoRec("서버시간"), 7, 2) + Mid(P_adoRec("서버시간"), 10)
                  P_adoRec.Close
                  strSQL = "INSERT INTO 자재코드변경(변경전분류코드, 변경전세부코드, 변경일시, " _
                                                  & "변경후분류코드, 변경후세부코드, 사용자코드) VALUES(" _
                                                  & "'" & Left(Trim(cboMtGp(0).Text), 2) & "', " _
                                                  & "'" & Trim(Text1(0).Text) & "', " _
                                                  & "'" & strServerDate + strServerTime & "', " _
                                                  & "'" & Left(Trim(cboMtGpRe.Text), 2) & "', " _
                                                  & "'" & Trim(txtSebuCodeRe.Text) & "', " _
                                                  & "'" & PB_regUserinfoU.UserCode & "' )"
                  On Error GoTo ERROR_TABLE_UPDATE
                  PB_adoCnnSQL.Execute strSQL
                  '자재코드변경(전체변경)
                  strSQL = "UPDATE 자재 SET " _
                                & "분류코드 = '" & Left(Trim(cboMtGpRe.Text), 2) & "', " _
                                & "세부코드 = '" & Trim(txtSebuCodeRe.Text) & "' " _
                          & "WHERE 분류코드 = '" & Left(Trim(cboMtGp(0).Text), 2) & "' " _
                            & "AND 세부코드 = '" & Trim(Text1(0).Text) & "' "
                  On Error GoTo ERROR_TABLE_UPDATE
                  PB_adoCnnSQL.Execute strSQL
                  '발주내역 자재코드 변경
                  strSQL = "UPDATE 발주내역 SET " _
                                & "자재코드 = '" & Left(Trim(cboMtGpRe.Text), 2) + Trim(txtSebuCodeRe.Text) & "' " _
                          & "WHERE 자재코드 = '" & Left(Trim(cboMtGp(0).Text), 2) + Trim(Text1(0).Text) & "' "
                  On Error GoTo ERROR_TABLE_UPDATE
                  PB_adoCnnSQL.Execute strSQL
                  '견적내역 자재코드 변경
                  strSQL = "UPDATE 견적내역 SET " _
                                & "자재코드 = '" & Left(Trim(cboMtGpRe.Text), 2) + Trim(txtSebuCodeRe.Text) & "' " _
                          & "WHERE 자재코드 = '" & Left(Trim(cboMtGp(0).Text), 2) + Trim(Text1(0).Text) & "' "
                  On Error GoTo ERROR_TABLE_UPDATE
                  PB_adoCnnSQL.Execute strSQL
                  '견적내역 자재코드 변경
               End If
            End If
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 0
                            If Len(txtSebuCodeRe.Text) > 0 And intRetVal = vbYes Then '코드변경이면
                               cboMtGp(0).ListIndex = cboMtGpRe.ListIndex
                               .TextMatrix(.Row, 1) = Mid(Trim(cboMtGpRe.Text), 5)
                               Text1(0).Text = Trim(txtSebuCodeRe.Text)
                               txtSebuCodeRe.Text = ""
                            End If
                            .TextMatrix(.Row, 0) = Left(cboMtGp(0).Text, 2)
                            .TextMatrix(.Row, 1) = Mid(Trim(cboMtGp(0).Text), 5)
                            .TextMatrix(.Row, 2) = Trim(Text1(0).Text)
                            .TextMatrix(.Row, 4) = .TextMatrix(.Row, 0) + .TextMatrix(.Row, 2)
                            .Cell(flexcpData, .Row, 4, .Row, 4) = .TextMatrix(.Row, 4)
                            .TextMatrix(.Row, 18) = Vals(Left(cboTaxGbn.Text, 1)) '과세구분
                            .TextMatrix(.Row, 19) = cboTaxGbn.ListIndex
                            .TextMatrix(.Row, 20) = Right(Trim(cboTaxGbn.Text), Len(Trim(cboTaxGbn.Text)) - 3)
                            .TextMatrix(.Row, 21) = Vals(Left(cboState(1).Text, 1)) '사용구분
                            .TextMatrix(.Row, 22) = cboState(1).ListIndex
                            .TextMatrix(.Row, 23) = Right(Trim(cboState(1).Text), Len(Trim(cboState(1).Text)) - 3)
                       Case 1 '1.품명
                            .TextMatrix(.Row, 3) = Trim(Text1(1).Text)
                       Case 2 '2.바코드
                            .TextMatrix(.Row, 16) = Trim(Text1(lngC).Text)
                       Case 3 To 4 '3.규격, 4.단위
                            .TextMatrix(.Row, lngC + 2) = Trim(Text1(lngC).Text)
                       Case 5 '5.폐기율
                            .TextMatrix(.Row, 17) = Vals(Trim(Text1(lngC).Text))
                       Case 6 To 7 '6.주매입처코드, 7.주매입처명
                            .TextMatrix(.Row, lngC + 1) = Trim(Text1(lngC).Text)
                       Case 8 '8.비고란
                            .TextMatrix(.Row, 15) = Trim(Text1(lngC).Text)
                       Case 9 To 14
                            .TextMatrix(.Row, lngC) = Vals(Trim(Text1(lngC).Text))
                       Case Else
                End Select
            Next lngC
         End If
         PB_adoCnnSQL.CommitTrans
         .SetFocus
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재 읽기 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재 변경 실패"
    Unload Me
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    PB_adoCnnSQL.RollbackTrans
    MsgBox PB_varErrCode & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "판매 관리 시스템 (서버와의 연결 실패)"
    Unload frmLogin
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "로그 변경 실패"
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
               cmdDelete.Enabled = False
               Screen.MousePointer = vbHourglass
               '삭제전 관련테이블 검사
               'P_adoRec.CursorLocation = adUseClient
               'strSQL = "SELECT Count(*) AS 해당건수 FROM 자재시세 " _
               '        & "WHERE 분류코드 = '" & Left(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
               '          & "AND 세부코드 = '" & Trim(.TextMatrix(.Row, 2)) & "' " _
               '          & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("해당건수")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "자재시세(" & Format(lngCnt, "#,#") & "건)가 있으므로 삭제할 수 없습니다.", vbCritical, "자재 삭제 불가"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "DELETE FROM 자재원장 " _
                       & "WHERE 분류코드 = '" & Left(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
                         & "AND 세부코드 = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               strSQL = "DELETE FROM 자재 " _
                       & "WHERE 분류코드 = '" & Left(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
                         & "AND 세부코드 = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               cboMtGp(cboMtGp.LBound).Enabled = False
               cboMtGpRe.Enabled = True
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarHorizontal
               End If
               Screen.MousePointer = vbDefault
               SubClearText
               .Row = 0
               vsfg1_EnterCell
               cmdDelete.Enabled = True
            End If
            .SetFocus
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재시세 읽기 실패"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재 삭제 실패"
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
    Set frm자재정보 = Nothing
    frmMain.SBar.Panels(4).Text = ""
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
    Text1(Text1.LBound).Enabled = False '자재코드 FLASE
    With vsfg1                 'Rows 1, Cols 24, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .Ellipsis = flexEllipsisEnd
         '.ExplorerBar = flexExSortShow
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 5
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 24
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '분류코드
         .ColWidth(1) = 1150   '분류명
         .ColWidth(2) = 2000   '세부코드
         .ColWidth(3) = 2500   '자재명
         .ColWidth(4) = 3000   '분류코드 + 세부코드
         .ColWidth(5) = 2500   '규격
         .ColWidth(6) = 800    '단위
         .ColWidth(7) = 1200   '매입처코드
         .ColWidth(8) = 3000   '매입처명
         .ColWidth(9) = 1350   '입고단가1
         .ColWidth(10) = 1350  '입고단가2
         .ColWidth(11) = 1350  '입고단가3
         .ColWidth(12) = 1350  '출고단가1
         .ColWidth(13) = 1350  '출고단가2
         .ColWidth(14) = 1350  '출고단가3
         .ColWidth(15) = 9400  '비고란
         .ColWidth(16) = 3000  '바코드
         .ColWidth(17) = 900   '폐기율
         .ColWidth(18) = 1     '과세구분
         .ColWidth(19) = 1     '과세구분ListIndex
         .ColWidth(20) = 1000  '과세구분
         .ColWidth(21) = 1     '사용구분
         .ColWidth(22) = 1     '사용구분ListIndex
         .ColWidth(23) = 1000  '사용구분
         
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "분류코드"                'H
         .TextMatrix(0, 1) = "분류명"
         .TextMatrix(0, 2) = "코드"
         .TextMatrix(0, 3) = "품명"
         .TextMatrix(0, 4) = "(분류코드+세부코드)코드"  'H
         .TextMatrix(0, 5) = "규격"
         .TextMatrix(0, 6) = "단위"
         .TextMatrix(0, 7) = "매입처코드"
         .TextMatrix(0, 8) = "매입처명"
         .TextMatrix(0, 9) = "매입단가1"
         .TextMatrix(0, 10) = "매입단가2"
         .TextMatrix(0, 11) = "매입단가3"
         .TextMatrix(0, 12) = "매출단가1"
         .TextMatrix(0, 13) = "매출단가2"
         .TextMatrix(0, 14) = "매출단가3"
         .TextMatrix(0, 15) = "비고란"
         .TextMatrix(0, 16) = "바코드"
         .TextMatrix(0, 17) = "폐기율"
         .TextMatrix(0, 18) = "과세구분"       'H
         .TextMatrix(0, 19) = "과세구분"       'H
         .TextMatrix(0, 20) = "과세구분"
         .TextMatrix(0, 21) = "사용구분"       'H
         .TextMatrix(0, 22) = "사용구분"       'H
         .TextMatrix(0, 23) = "사용구분"
         
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 4, 18, 19, 21, 22
                         .ColHidden(lngC) = True
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 9 To 14, 17
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 1, 2, 3, 5, 6, 7, 8, 15, 16
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 4, 18 To 23
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictColumns
         For lngC = 0 To 1
             .MergeCol(lngC) = True
         Next lngC
    End With
End Sub

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    cboMtGp(cboMtGp.LBound).Enabled = False
    Text1(Text1.LBound).Enabled = False
    cboMtGpRe.Enabled = True
    txtSebuCodeRe.Enabled = True
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.분류코드 AS 분류코드, ISNULL(T1.분류명,'') AS 분류명 " _
             & "FROM 자재분류 T1 " _
            & "ORDER BY T1.분류코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboMtGp(0).ListIndex = -1
       cboMtGp(1).ListIndex = -1
       cboMtGpRe.ListIndex = -1
       cmdClear.Enabled = False: cmdFind.Enabled = False
       cmdSave.Enabled = False: cmdDelete.Enabled = False
       Exit Sub
    Else
       cboMtGp(1).AddItem "00. " & "전체"
       Do Until P_adoRec.EOF
          cboMtGp(0).AddItem P_adoRec("분류코드") & ". " & P_adoRec("분류명")
          cboMtGp(1).AddItem P_adoRec("분류코드") & ". " & P_adoRec("분류명")
          cboMtGpRe.AddItem P_adoRec("분류코드") & ". " & P_adoRec("분류명")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       cboMtGp(0).ListIndex = 0
       cboMtGp(1).ListIndex = 0
       cboMtGpRe.ListIndex = 0
    End If
    With cboState(0)
         .AddItem "전    체"
         .AddItem "정    상"
         .AddItem "사용불가"
         .AddItem "기    타"
         .ListIndex = 1
    End With
    With cboTaxGbn
         .AddItem "0. 비 과 세"
         .AddItem "1. 과    세"
         .ListIndex = 1
    End With
    With cboState(1)
         .AddItem "0. 정    상"
         .AddItem "9. 사용불가"
         .ListIndex = 0
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfg1) 채우기///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL     As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
     
    'If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3) + Len(P_strFindString4)) = 0 Then
    '   txtFindCD.SetFocus
    '   Exit Sub
    'End If
    If Left(cboMtGp(1).Text, 2) <> "00" And (cboMtGp(1).ListCount > 0) Then
       cboMtGp(0).ListIndex = (cboMtGp(1).ListIndex) - 1
    End If
    '검색조건 자재분류
    Select Case Left(Trim(cboMtGp(1).Text), 2)
           Case "00" '전체
                strWhere = ""
           Case Else
                strWhere = "WHERE T1.분류코드 = '" & Mid(Trim(cboMtGp(1).Text), 1, 2) & "' "
    End Select
    '검색조건 사용구분
    Select Case cboState(0).ListIndex
           Case 0 '전체
                strWhere = strWhere
           Case 1 '정상
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.사용구분 = 0 "
           Case 2 '사용불가
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.사용구분 = 9 "
           Case Else
                strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "NOT(T1.사용구분 = 0 OR T1.사용구분 = 9) "
    End Select
    If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3) + Len(P_strFindString4)) = 0 Then   '정상적인 조회
       strOrderBy = "ORDER BY T1.자재명 "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + "T1.세부코드 LIKE '%" & P_strFindString1 & "%' " _
                & "AND T1.자재명 LIKE '%" & P_strFindString2 & "%' AND T1.규격 LIKE '%" & P_strFindString3 & "%' " _
                & "AND T1.바코드 LIKE '%" & P_strFindString4 & "%' "
       strOrderBy = "ORDER BY T1.분류코드, T1.세부코드 "
    End If
    '??CODE????? 로된 품목 제외
    If chkCodeException.Value = 1 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                + "NOT (DATALENGTH(T1.세부코드) = 9 AND UPPER(SUBSTRING(T1.세부코드, 1, 4)) = 'CODE' " _
                + "AND T1.세부코드 LIKE 'CODE_____' " _
                + "AND ISNUMERIC(SUBSTRING(T1.세부코드, 5, 5)) = 1) "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT ISNULL(T1.분류코드,'') AS 분류코드, ISNULL(T2.분류명,'') AS 분류명, " _
                  & "ISNULL(T1.세부코드,'') AS 세부코드, T1.자재명 AS 자재명, " _
                  & "T1.바코드 AS 바코드, T1.규격 AS 규격, T1.단위 AS 단위, " _
                  & "T1.폐기율 AS 폐기율, T1.과세구분 AS 과세구분, " _
                  & "T1.사용구분 AS 사용구분, T3.주매입처코드 AS 주매입처코드, ISNULL(T4.매입처명, '') AS 주매입처명, " _
                  & "ISNULL(T3.입고단가1,0) AS 입고단가1, ISNULL(T3.입고단가2,0) AS 입고단가2, ISNULL(T3.입고단가3,0) AS 입고단가3, " _
                  & "ISNULL(T3.출고단가1,0) AS 출고단가1, ISNULL(T3.출고단가2,0) AS 출고단가2, ISNULL(T3.출고단가3,0) AS 출고단가3, " _
                  & "T3.비고란 AS 비고란 " _
             & "FROM 자재 T1 " _
             & "LEFT JOIN 자재분류 T2 " _
                    & "ON T2.분류코드 = T1.분류코드 " _
             & "LEFT JOIN 자재원장 T3 ON T3.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                   & " AND T3.분류코드 = T1.분류코드 AND T3.세부코드 = T1.세부코드 " _
             & "LEFT JOIN 매입처 T4 ON T4.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                   & " AND T4.매입처코드 = T3.주매입처코드 " _
               & "" & strWhere & " " _
               & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + vsfg1.FixedRows
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = IIf(IsNull(P_adoRec("분류코드")), "", P_adoRec("분류코드"))
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("분류명")), "", P_adoRec("분류명"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("세부코드")), "", P_adoRec("세부코드"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("자재명")), "", P_adoRec("자재명"))
               'FindRow 사용을 위해
               .TextMatrix(lngR, 4) = .TextMatrix(lngR, 0) & P_adoRec("세부코드")
               .Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4)
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("규격")), "", P_adoRec("규격"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("단위")), "", P_adoRec("단위"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("주매입처코드")), "", P_adoRec("주매입처코드"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("주매입처명")), "", P_adoRec("주매입처명"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("입고단가1")), 0, P_adoRec("입고단가1"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("입고단가2")), 0, P_adoRec("입고단가2"))
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("입고단가3")), 0, P_adoRec("입고단가3"))
               .TextMatrix(lngR, 12) = IIf(IsNull(P_adoRec("출고단가1")), 0, P_adoRec("출고단가1"))
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("출고단가2")), 0, P_adoRec("출고단가2"))
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("출고단가3")), 0, P_adoRec("출고단가3"))
               .TextMatrix(lngR, 15) = IIf(IsNull(P_adoRec("비고란")), "", P_adoRec("비고란"))
               .TextMatrix(lngR, 16) = IIf(IsNull(P_adoRec("바코드")), "", P_adoRec("바코드"))
               .TextMatrix(lngR, 17) = IIf(IsNull(P_adoRec("폐기율")), 0, P_adoRec("폐기율"))
               .TextMatrix(lngR, 18) = IIf(IsNull(P_adoRec("과세구분")), 0, P_adoRec("과세구분"))
               'ListIndex
               For lngRRR = 0 To cboTaxGbn.ListCount - 1
                   If .ValueMatrix(lngR, 18) = Vals(Left(cboTaxGbn.List(lngRRR), 1)) Then
                      .TextMatrix(lngR, 19) = lngRRR
                      .TextMatrix(lngR, 20) = Right(Trim(cboTaxGbn.List(lngRRR)), Len(Trim(cboTaxGbn.List(lngRRR))) - 3)
                      Exit For
                   End If
               Next lngRRR
               .TextMatrix(lngR, 21) = IIf(IsNull(P_adoRec("사용구분")), 0, P_adoRec("사용구분"))
               'ListIndex
               For lngRRR = 0 To cboState(1).ListCount - 1
                   If .ValueMatrix(lngR, 21) = Vals(Left(cboState(1).List(lngRRR), 1)) Then
                      .TextMatrix(lngR, 22) = lngRRR
                      .TextMatrix(lngR, 23) = Right(Trim(cboState(1).List(lngRRR)), Len(Trim(cboState(1).List(lngRRR))) - 3)
                      Exit For
                   End If
               Next lngRRR
               'If .TextMatrix(lngR, 3) = P_strFindString2 Then
               '   lngRR = lngR
               'End If
               If P_adoRec.RecordCount = 1 Then
                  lngRR = 1
               End If
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt Then
                  '.TopRow = .Rows - PC_intRowCnt + .FixedRows
                  .TopRow = 1
               End If
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재 읽기 실패"
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
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '자재코드
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If LenH(Text1(lngC).Text) < 1 Or LenH(Text1(lngC).Text) > 16 Then
                       Exit Function
                    End If
               Case 1  '자재명
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) > 0 And LenH(Text1(lngC).Text) <= 30) Then
                       Exit Function
                    End If
               Case 2  '바코드
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 13) Then
                       Exit Function
                    End If
               Case 3  '규격
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 30) Then
                       Exit Function
                    End If
               Case 4  '단위
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 20) Then
                       Exit Function
                    End If
               Case 8  '비고란
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 100) Then
                       Exit Function
                    End If
        End Select
    Next lngC
    blnOK = True
End Function

'+---------------------------+
'/// 크리스탈 리포터 출력 ///
'+---------------------------+
Private Sub cmdPrint_Click()
Dim strSQL                 As String
Dim strWhere               As String
Dim strOrderBy             As String

Dim varRetVal              As Variant '리포터 파일
Dim strExeFile             As String
Dim strExeMode             As String
Dim intRetCHK              As Integer '실행여부

Dim lngR                   As Long
Dim lngC                   As Long

Dim strForAppDate          As String  '실행일자       (Formula)
Dim strForBranchName       As String  '사업장명       (Formula)
Dim strForPrtDateTime      As String  '출력일시       (Formula)
Dim strParGroupCode        As Integer '식품군소분류   (Parameter)
Dim intParStateCode        As Integer '사용구분       (Parameter)

    Screen.MousePointer = vbHourglass
    '서버일시(출력일시)
    strSQL = "SELECT CONVERT(VARCHAR(19),GETDATE(), 120) AS 서버시간 "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strForPrtDateTime = Format(PB_regUserinfoU.UserServerDate, "0000-00-00") & Space(1) _
                      & Format(Right(P_adoRec("서버시간"), 8), "hh:mm:ss")
    P_adoRec.Close
    
    intRetCHK = 99
    With CrystalReport1
         If PB_Test = 0 Then
            strExeFile = App.Path & ".\자재정보.rpt"
         Else
            strExeFile = App.Path & ".\자재정보T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForAppDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'"
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"
            '--- Parameter Fields ---
            .StoredProcParam(0) = Mid(cboMtGp(1).Text, 1, 2)    '자재분류(분류코드)
            If cboState(0).ListIndex < 2 Then                   '사용구분(0.전체, 1.정상, 2.삭제, 3.오 류)
               .StoredProcParam(1) = 0
            Else
               .StoredProcParam(1) = 9
            End If
         End If
         If intRetCHK = 99 Then
            .Connect = PB_adoCnnSQL.ConnectionString
            .Destination = crptToWindow
            .DiscardSavedData = True
            .ProgressDialog = True
            .ReportSource = crptReport
            .WindowAllowDrillDown = False
            .WindowShowProgressCtls = True
            .WindowShowCloseBtn = True
            .WindowShowExportBtn = False
            .WindowShowGroupTree = False
            .WindowShowPrintSetupBtn = True
            .WindowTop = 0: .WindowTop = 0: .WindowHeight = 11100: .WindowWidth = 15405
            .WindowState = crptMaximized
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "자재정보"
            .Action = 1
            .Reset
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_CRYSTAL_REPORTS:
    MsgBox Err.Number & Space(1) & Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
 
