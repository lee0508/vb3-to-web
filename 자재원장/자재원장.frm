VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm자재원장 
   BorderStyle     =   0  '없음
   Caption         =   "자재원장"
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
      TabIndex        =   28
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdZero 
         Caption         =   "현재재고 0 으로 조정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   67
         Top             =   195
         Width           =   2535
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "자재원장.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   59
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "자재원장.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   33
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "자재원장.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   31
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "자재원장.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   32
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "자재원장.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "자재원장.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   30
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "자 재 원 장"
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
         TabIndex        =   29
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   6317
      Left            =   60
      TabIndex        =   26
      Top             =   3645
      Width           =   15195
      _cx             =   26802
      _cy             =   11142
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
      Height          =   2949
      Left            =   60
      TabIndex        =   27
      Top             =   630
      Width           =   15195
      Begin VB.CheckBox chkCodeException 
         Caption         =   "CODE 품목 제외"
         Height          =   180
         Left            =   6360
         TabIndex        =   38
         Top             =   240
         Value           =   1  '확인
         Width           =   1575
      End
      Begin VB.TextBox txtFindBarCode 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   13400
         MaxLength       =   14
         TabIndex        =   44
         Top             =   600
         Width           =   1350
      End
      Begin VB.TextBox txtBarCode 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Left            =   4740
         MaxLength       =   14
         TabIndex        =   74
         Top             =   585
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   20
         Left            =   1515
         TabIndex        =   24
         Top             =   2550
         Width           =   9465
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   19
         Left            =   13420
         MaxLength       =   20
         TabIndex        =   23
         Top             =   2160
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   18
         Left            =   13420
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1800
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   17
         Left            =   13420
         MaxLength       =   20
         TabIndex        =   21
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   16
         Left            =   12000
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2160
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   15
         Left            =   12000
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1800
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   14
         Left            =   12000
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   10
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1800
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   9
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1440
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   13
         Left            =   8925
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2160
         Width           =   2610
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   12
         Left            =   10150
         MaxLength       =   8
         TabIndex        =   16
         Top             =   1800
         Width           =   825
      End
      Begin VB.TextBox txtFindCD 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   11355
         MaxLength       =   18
         TabIndex        =   41
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox txtFindSZ 
         Appearance      =   0  '평면
         Height          =   285
         Left            =   10725
         MaxLength       =   30
         TabIndex        =   43
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox txtFindNM 
         Appearance      =   0  '평면
         Height          =   285
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   42
         Top             =   600
         Width           =   2685
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   8
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1080
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   7
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2160
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   4
         Left            =   1515
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   5
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1440
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   11
         Left            =   7320
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2160
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   1  '입력 상태 설정
         Index           =   3
         Left            =   1515
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   1
         Left            =   4740
         Style           =   2  '드롭다운 목록
         TabIndex        =   6
         Top             =   1080
         Width           =   1350
      End
      Begin VB.ComboBox cboTaxGbn 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1515
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   2160
         Width           =   1350
      End
      Begin VB.ComboBox cboMtGp 
         Height          =   300
         Left            =   8640
         Style           =   2  '드롭다운 목록
         TabIndex        =   40
         Top             =   200
         Width           =   1695
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   0
         Left            =   13680
         Style           =   2  '드롭다운 목록
         TabIndex        =   39
         Top             =   200
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   6
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1800
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   1  '입력 상태 설정
         Index           =   2
         Left            =   1515
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   1
         Left            =   1515
         MaxLength       =   30
         TabIndex        =   1
         Top             =   585
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   1515
         MaxLength       =   18
         TabIndex        =   0
         Top             =   225
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpInputDate 
         Height          =   270
         Left            =   10150
         TabIndex        =   14
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpOutputDate 
         Height          =   270
         Left            =   10150
         TabIndex        =   15
         Top             =   1440
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "바코드"
         Height          =   240
         Index           =   28
         Left            =   12670
         TabIndex        =   76
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "바코드"
         Height          =   240
         Index           =   27
         Left            =   3645
         TabIndex        =   75
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "비고란"
         Height          =   240
         Index           =   23
         Left            =   315
         TabIndex        =   73
         Top             =   2600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "3."
         Height          =   240
         Index           =   22
         Left            =   11650
         TabIndex        =   72
         Top             =   2220
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "2."
         Height          =   240
         Index           =   21
         Left            =   11650
         TabIndex        =   71
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "1."
         Height          =   240
         Index           =   20
         Left            =   11650
         TabIndex        =   70
         Top             =   1485
         Width           =   255
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
         Left            =   13560
         TabIndex        =   69
         Top             =   1125
         Width           =   1095
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
         Left            =   12120
         TabIndex        =   68
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "재고이동"
         Height          =   240
         Index           =   17
         Left            =   6405
         TabIndex        =   66
         Top             =   1845
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "재고조정"
         Height          =   240
         Index           =   16
         Left            =   6405
         TabIndex        =   65
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "[Home]"
         Height          =   240
         Index           =   14
         Left            =   11025
         TabIndex        =   64
         Top             =   1860
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처코드"
         Height          =   240
         Index           =   12
         Left            =   8805
         TabIndex        =   63
         Top             =   1845
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "세부코드"
         Height          =   240
         Index           =   11
         Left            =   10440
         TabIndex        =   62
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "규격"
         Height          =   240
         Index           =   10
         Left            =   10080
         TabIndex        =   61
         Top             =   645
         Width           =   495
      End
      Begin VB.Line Line2 
         X1              =   6360
         X2              =   6360
         Y1              =   480
         Y2              =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품명"
         Height          =   240
         Index           =   13
         Left            =   6405
         TabIndex        =   60
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "[Home]"
         Height          =   240
         Index           =   1
         Left            =   3840
         TabIndex        =   58
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품목코드"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   57
         Top             =   280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매출수량"
         Height          =   240
         Index           =   15
         Left            =   6405
         TabIndex        =   56
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입수량"
         Height          =   240
         Index           =   9
         Left            =   3645
         TabIndex        =   55
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "최종출고일자"
         Height          =   240
         Index           =   8
         Left            =   8805
         TabIndex        =   54
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "최종입고일자"
         Height          =   240
         Index           =   7
         Left            =   8805
         TabIndex        =   53
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   15000
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "폐기율"
         Height          =   240
         Index           =   6
         Left            =   315
         TabIndex        =   52
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "적정재고"
         Height          =   240
         Index           =   5
         Left            =   3660
         TabIndex        =   51
         Top             =   1485
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "현재재고"
         Height          =   240
         Index           =   4
         Left            =   5925
         TabIndex        =   50
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "과세구분"
         Height          =   240
         Index           =   3
         Left            =   315
         TabIndex        =   49
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "단위"
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   48
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품명"
         Height          =   240
         Index           =   36
         Left            =   315
         TabIndex        =   47
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "검색조건)"
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
         Left            =   5280
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "분류"
         Height          =   240
         Index           =   34
         Left            =   7725
         TabIndex        =   45
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "이월재고"
         Height          =   240
         Index           =   31
         Left            =   3660
         TabIndex        =   37
         Top             =   1845
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "규격"
         Height          =   240
         Index           =   26
         Left            =   315
         TabIndex        =   36
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사용구분"
         Height          =   240
         Index           =   25
         Left            =   12720
         TabIndex        =   35
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사용구분"
         Height          =   240
         Index           =   24
         Left            =   3645
         TabIndex        =   34
         Top             =   1125
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm자재원장"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 자재원장
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 자재원장, 자재, 자재분류
'                  자재원장마감, 자재입출내역
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_adoRec1          As New ADODB.Recordset
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

    frmMain.SBar.Panels(4).Text = ""
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
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 (서버와의 연결 실패)"
    Unload Me
    Exit Sub
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
'+-------------------+
'/// dtpFirstDate ///
'+-------------------+
Private Sub dtpFirstDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
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
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '자재검색
       PB_strCallFormName = "frm자재원장"
       PB_strMaterialsCode = Trim(Text1(0).Text)
       PB_strMaterialsName = Trim(Text1(1).Text)
       frm자재검색.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '검색에서 취소(ESC)
       Else
          Text1(0).Text = PB_strMaterialsCode
          Text1(1).Text = PB_strMaterialsName
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    ElseIf _
       (Index = 12 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then '매입처검색
       PB_strFMCCallFormName = "frm자재원장"
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
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 읽기 실패"
    Unload Me
    Exit Sub
End Sub
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
Dim lngR   As Long
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = Trim(.Text)
                     If Len(Trim(Text1(Index).Text)) = 0 Then
                        Text1(1).Text = "": txtBarCode.Text = "": Text1(2).Text = "": Text1(3).Text = "": Text1(4).Text = ""
                        Exit Sub
                     End If
                Case 4
                     .Text = Format(Vals(Trim(.Text)), "#,0.00")
                Case 5 To 11
                     .Text = Format(Fix(Vals(Trim(.Text))), "#,0")
                Case 12
                     If Len(Trim(.Text)) = 0 Then
                        Text1(Index + 1).Text = ""
                     End If
                Case 14 To 19
                     If Vals(Trim(.Text)) < 0 Then
                        .Text = Vals(Trim(.Text)) * -1
                     End If
                     .Text = Format(Vals(Trim(.Text)), "#,0.00")
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// 검색조건 ///
'+---------------+
'+-------------------------------------------+
'/// chkCodeException(CODE 품목제외 검색) ///
'+-------------------------------------------+
Private Sub chkCodeException_KeyDown(KeyCode As Integer, Shift As Integer)
    With chkCodeException
         If KeyCode = vbKeyReturn Then cboMtGp.SetFocus
    End With
End Sub
'+-----------------------+
'/// cboMtGp(index) ///
'+-----------------------+
Private Sub cboMtGp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtFindCD.SetFocus
    End If
End Sub
'+----------------------------+
'/// txtFindCode(코드검색) ///
'+----------------------------+
Private Sub txtFindCD_GotFocus()
    With txtFindCD
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindCD_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
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

'+----------------------------+
'/// txtFindNM(자재명검색) ///
'+----------------------------+
Private Sub txtFindNM_GotFocus()
    With txtFindNM
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindNM_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       txtFindSZ.SetFocus
    End If
End Sub

'+--------------------------+
'/// txtFindSZ(규격검색) ///
'+--------------------------+
Private Sub txtFindSZ_GotFocus()
    With txtFindSZ
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindSZ_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       txtFindBarCode.SetFocus
    End If
End Sub

'+---------------------------------+
'/// txtFindBarCode(바코드검색) ///
'+---------------------------------+
Private Sub txtFindBarCode_GotFocus()
    With txtFindBarCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtFindBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
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
         'P_strFindString1 = Trim(.Cell(flexcpData, .Row, 0)) 'Not Used
    End With
End Sub

Private Sub vsfg1_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfg1
         'If .FindRow(P_strFindString1, , 0) > 0 Then
         '   .Row = .FindRow(P_strFindString1, , 0) 'Not Used
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
               'PB_strFMCCallFormName = "frm자재원장"
               'PB_strMaterialsCode = .TextMatrix(.Row, 4)
               'PB_strMaterialsName = .TextMatrix(.Row, 3)
               'PB_strSupplierCode = ""
               'frm자재시세검색.Show vbModal
            End If
         End If
    End With
End Sub
Private Sub vsfg1_DblClick()
    With vsfg1
         If .Row >= .FixedRows Then
             vsfg1_KeyDown vbKeyF1, 0  '자재시세검색
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
         Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
         If .Row >= .FixedRows And OldRow <> NewRow Then
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case 0
                            Text1(0).Text = .TextMatrix(.Row, 4) '자재코드
                       Case 3
                            Text1(1).Text = .TextMatrix(.Row, 3)
                       Case 5 To 6
                            Text1(lngC - 3).Text = .TextMatrix(.Row, lngC)
                       Case 7 To 13
                            Text1(lngC - 2).Text = Format(.ValueMatrix(.Row, lngC), "#,0")
                       Case 14
                            If Len(.TextMatrix(.Row, lngC)) = 10 Then
                               dtpInputDate.Value = .TextMatrix(.Row, lngC)
                            Else
                               dtpInputDate.Value = Format("19000101", "0000-00-00")
                            End If
                       Case 15
                            If Len(.TextMatrix(.Row, lngC)) = 10 Then
                               dtpOutputDate.Value = .TextMatrix(.Row, lngC)
                            Else
                               dtpOutputDate.Value = Format("19000101", "0000-00-00")
                            End If
                       Case 16, 17 '매입처
                            Text1(lngC - 4).Text = .TextMatrix(.Row, lngC)
                       Case 18 To 23  '단가
                            Text1(lngC - 4).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
                       Case 24  '비고란
                            Text1(20).Text = .TextMatrix(.Row, lngC)
                       Case 25  '바코드
                            txtBarCode.Text = .TextMatrix(.Row, lngC)
                       Case 26  '폐기율
                            Text1(4).Text = Format(.ValueMatrix(.Row, lngC), "#,0.00")
                       Case 27  '과세구분
                            cboTaxGbn.ListIndex = IIf(.TextMatrix(.Row, lngC) = "비 과 세", 0, 1)
                       Case 29  '사용구분
                            cboState(1).ListIndex = .ValueMatrix(.Row, lngC)
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
         'Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
    End With
End Sub

'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
Dim strSQL As String
    SubClearText
    vsfg1.Row = 0
    Text1(Text1.LBound).Enabled = True
    Text1(Text1.LBound + 1).Enabled = False
    Text1(Text1.LBound).SetFocus
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
Dim strSQL         As String
Dim lngR           As Long
Dim lngC           As Long
Dim blnOK          As Boolean
Dim intRetVal      As Integer
Dim curAdjustAmt   As Currency
Dim lngLogCnt      As Long
Dim strServerTime  As String
Dim strTime        As String
Dim CurInputMny    As Currency '입고단가
Dim CurInputVat    As Currency '입고부가
Dim CurOutPutMny   As Currency '출고단가
Dim CurOutPutVat   As Currency '출고부가

    '입력내역 검사
    blnOK = False
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       If lngC = 0 Then
          Text1(lngC).Enabled = True
       End If
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
    '서버시간 구하기
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) + Mid(P_adoRec("서버시간"), 7, 2) _
                  + Mid(P_adoRec("서버시간"), 10)
    P_adoRec.Close
    strTime = strServerTime
    If Text1(Text1.LBound).Enabled = True Then '자재원장 추가면 검색
       curAdjustAmt = Vals(Trim(Text1(11).Text))
    Else
       curAdjustAmt = Vals(Trim(Text1(11).Text)) - vsfg1.ValueMatrix(vsfg1.Row, 13)
    End If
    '단가
    CurInputMny = Vals(Trim(Text1(14).Text)): CurInputVat = Fix(Vals(Trim(Text1(14).Text)) * (PB_curVatRate))
    CurOutPutMny = Vals(Trim(Text1(17).Text)): CurOutPutVat = Fix(Vals(Trim(Text1(17).Text)) * (PB_curVatRate))
    With vsfg1
         Screen.MousePointer = vbHourglass
         P_adoRec.CursorLocation = adUseClient
         If Text1(Text1.LBound).Enabled = True Then '자재원장 추가면 검색
            strSQL = "SELECT * FROM 자재원장 T1 " _
                    & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND T1.분류코드 = '" & Mid(Text1(0).Text, 1, 2) & "' AND T1.세부코드 = '" & Mid(Text1(0).Text, 3) & "' "
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
         If curAdjustAmt > 0 Then '입고(+)
            '거래번호 구하기
            P_adoRec.CursorLocation = adUseClient
            strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "5" & "', " _
                                 & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
            On Error GoTo ERROR_STORED_PROCEDURE
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            lngLogCnt = P_adoRec(0)
            P_adoRec.Close
            '자재입출내역
            strSQL = "INSERT INTO 자재입출내역(사업장코드, 분류코드, " _
                                            & "세부코드, 입출고구분, " _
                                            & "입출고일자, 입출고시간, " _
                                            & "입고수량, 입고단가, " _
                                            & "입고부가, 출고수량, " _
                                            & "출고단가, 출고부가, " _
                                            & "매입처코드, 매출처코드, " _
                                            & "원래입출고일자, 직송구분, " _
                                            & "발견일자, 발견번호, 거래일자, 거래번호, " _
                                            & "계산서발행여부, 현금구분, 감가구분, 적요, 작성년도, 책번호, 일련번호, " _
                                            & "사용구분, 수정일자, " _
                                            & "사용자코드, 재고이동사업장코드) VALUES( " _
                      & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(Trim(Text1(0).Text), 1, 2) & "', " _
                      & "'" & Mid(Trim(Text1(0).Text), 3) & "', 5, " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                      & "" & curAdjustAmt & ", " & CurInputMny & ", " _
                      & "" & CurInputVat & ", 0, " _
                      & "0, 0, " _
                      & "'', '', " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                      & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                      & "0, 0, 0, '', '', 0, 0, " _
                      & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                      & "'" & PB_regUserinfoU.UserCode & "', '' ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
         ElseIf _
            curAdjustAmt < 0 Then '출고(-)
            '거래번호 구하기
            P_adoRec.CursorLocation = adUseClient
            strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "6" & "', " _
                                 & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
            On Error GoTo ERROR_STORED_PROCEDURE
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            lngLogCnt = P_adoRec(0)
            P_adoRec.Close
            strSQL = "INSERT INTO 자재입출내역(사업장코드, 분류코드, " _
                                            & "세부코드, 입출고구분, " _
                                            & "입출고일자, 입출고시간, " _
                                            & "입고수량, 입고단가, " _
                                            & "입고부가, 출고수량, " _
                                            & "출고단가, 출고부가, " _
                                            & "매입처코드, 매출처코드, " _
                                            & "원래입출고일자, 직송구분, " _
                                            & "발견일자, 발견번호, 거래일자, 거래번호, " _
                                            & "계산서발행여부, 현금구분, 감가구분, 적요, 작성년도, 책번호, 일련번호, " _
                                            & "사용구분, 수정일자, " _
                                            & "사용자코드, 재고이동사업장코드) VALUES( " _
                      & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(Trim(Text1(0).Text), 1, 2) & "', " _
                      & "'" & Mid(Trim(Text1(0).Text), 3) & "', 6, " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                      & "0, 0, " _
                      & "0, " & (curAdjustAmt * -1) & ", " _
                      & "" & CurOutPutMny & ", " & CurOutPutVat & ", " _
                      & "'', '', " _
                      & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                      & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                      & "0, 0, 0, '', '', 0, 0, " _
                      & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                      & "'" & PB_regUserinfoU.UserCode & "', '' ) "
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
         End If
    End With
    With vsfg1
         If Text1(Text1.LBound).Enabled = True Then '자재원장 추가
            strSQL = "INSERT INTO 자재원장(사업장코드, 분류코드, " _
                                        & "세부코드, 적정재고, " _
                                        & "최종입고일자, 최종출고일자, " _
                                        & "사용구분, 수정일자, " _
                                        & "사용자코드, 비고란, 주매입처코드, " _
                                        & "입고단가1, 입고단가2, 입고단가3, " _
                                        & "출고단가1, 출고단가2, 출고단가3 ) Values( " _
                    & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Mid(Trim(Text1(0).Text), 1, 2) & "', " _
                    & "'" & Mid(Trim(Text1(0).Text), 3) & "', " & Vals(Trim(Text1(5).Text)) & ", " _
                    & "'" & IIf(DTOS(dtpInputDate.Value) = "19000101", "", DTOS(dtpInputDate.Value)) & "', " _
                    & "'" & IIf(DTOS(dtpOutputDate.Value) = "19000101", "", DTOS(dtpOutputDate.Value)) & "', " _
                    & "" & Vals(Left(cboState(1).Text, 1)) & ",'" & PB_regUserinfoU.UserServerDate & "', " _
                    & "'" & PB_regUserinfoU.UserCode & "', '" & Trim(Text1(20).Text) & "', '" & Trim(Text1(12).Text) & "', " _
                    & "" & Vals(Trim(Text1(14).Text)) & ", " & Vals(Trim(Text1(15).Text)) & ", " & Vals(Trim(Text1(15).Text)) & ", " _
                    & "" & Vals(Trim(Text1(17).Text)) & ", " & Vals(Trim(Text1(18).Text)) & ", " & Vals(Trim(Text1(19).Text)) & " ) "
            On Error GoTo ERROR_TABLE_INSERT
            .AddItem .Rows
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 0
                            .TextMatrix(.Rows - 1, 0) = Left(Text1(0).Text, 2)
                            For lngR = 0 To cboMtGp.ListCount - 1
                                If Left(Text1(0).Text, 2) = Left(cboMtGp.List(lngR), 2) Then
                                   .TextMatrix(.Rows - 1, 1) = Trim(Mid(cboMtGp.List(lngR), 5))
                                   Exit For
                                End If
                            Next lngR
                            .TextMatrix(.Rows - 1, 2) = Mid(Text1(0).Text, 3)
                            .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 0) + .TextMatrix(.Rows - 1, 2)
                            .Cell(flexcpData, .Rows - 1, 4, .Rows - 1, 4) = .TextMatrix(.Rows - 1, 4)
                       Case 1      '1.품명
                            .TextMatrix(.Rows - 1, 3) = Trim(Text1(1).Text)
                       Case 2, 3   '2.규격, 3.단위
                            .TextMatrix(.Rows - 1, lngC + 3) = Trim(Text1(lngC).Text)
                       Case 4      '4,과세구분
                            .TextMatrix(.Rows - 1, 14) = Format(DTOS(dtpInputDate.Value), "0000-00-00")   '최종입고일자
                            .TextMatrix(.Rows - 1, 15) = Format(DTOS(dtpOutputDate.Value), "0000-00-00")  '최종출고일자
                            .TextMatrix(.Rows - 1, 25) = Trim(txtBarCode.Text)
                            .TextMatrix(.Rows - 1, 26) = Vals(Trim(Text1(lngC).Text))                     '폐기율
                            .TextMatrix(.Rows - 1, 27) = Right(Trim(cboTaxGbn.Text), Len(Trim(cboTaxGbn.Text)) - 3) '과세구분
                            .TextMatrix(.Rows - 1, 28) = Vals(Left(cboState(1).Text, 1))                  '사용구분
                            .TextMatrix(.Rows - 1, 29) = cboState(1).ListIndex
                            .TextMatrix(.Rows - 1, 30) = Right(Trim(cboState(1).Text), Len(Trim(cboState(1).Text)) - 3)
                       Case 5      '5.적정재고
                            .TextMatrix(.Rows - 1, 7) = Vals(Trim(Text1(lngC).Text))
                       Case 9      '9.재고조정수량
                            .TextMatrix(.Rows - 1, 11) = Vals(Trim(Text1(lngC).Text)) + curAdjustAmt
                            Text1(lngC).Text = Format(Vals(Trim(Text1(lngC).Text)) + curAdjustAmt, "#,0")
                       Case 11     '11.현재재고
                            .TextMatrix(.Rows - 1, 13) = Vals(Trim(Text1(lngC).Text))
                       Case 12     '12.주매입처코드
                            .TextMatrix(.Rows - 1, 16) = Trim(Text1(lngC).Text)
                       Case 13     '13.주매입처명
                            .TextMatrix(.Rows - 1, 17) = Trim(Text1(lngC).Text)
                       Case 14 To 19  '단가
                            .TextMatrix(.Rows - 1, lngC + 4) = Vals(Trim(Text1(lngC).Text))
                       Case 20     '20.비고란
                            .TextMatrix(.Rows - 1, 24) = Trim(Text1(lngC).Text)
                       Case Else
                End Select
            Next lngC
            '미달상품표시
            If .ValueMatrix(.Rows - 1, 13) < .ValueMatrix(.Rows - 1, 7) Then
               .Cell(flexcpForeColor, .Rows - 1, 13, .Rows - 1, 13) = vbRed
               .Cell(flexcpFontBold, .Rows - 1, 13, .Rows - 1, 13) = True
            Else
               .Cell(flexcpForeColor, .Rows - 1, 13, .Rows - 1, 13) = vbBlack
               .Cell(flexcpFontBold, .Rows - 1, 13, .Rows - 1, 13) = False
            End If
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(0).Enabled = False: Text1(1).Enabled = True
            .Row = .Rows - 1          '자동으로 vsfg1_EnterCell Event 발생
         Else
            strSQL = "UPDATE 자재원장 SET " _
                          & "적정재고 = " & Vals(Trim(Text1(5).Text)) & ", " _
                          & "최종입고일자 = '" & IIf(DTOS(dtpInputDate.Value) = "19000101", "", DTOS(dtpInputDate.Value)) & "'," _
                          & "최종출고일자 = '" & IIf(DTOS(dtpOutputDate.Value) = "19000101", "", DTOS(dtpOutputDate.Value)) & "', " _
                          & "사용구분 = " & Vals(Left(cboState(1).Text, 1)) & ", " _
                          & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "사용자코드 = '" & PB_regUserinfoU.UserCode & "', " _
                          & "비고란 = '" & Trim(Text1(20).Text) & "', " _
                          & "주매입처코드 = '" & Trim(Text1(12).Text) & "', " _
                          & "입고단가1 = " & Vals(Trim(Text1(14).Text)) & ", " _
                          & "입고단가2 = " & Vals(Trim(Text1(15).Text)) & ", " _
                          & "입고단가3 = " & Vals(Trim(Text1(16).Text)) & ", " _
                          & "출고단가1 = " & Vals(Trim(Text1(17).Text)) & ", " _
                          & "출고단가2 = " & Vals(Trim(Text1(18).Text)) & ", " _
                          & "출고단가3 = " & Vals(Trim(Text1(19).Text)) & " " _
                    & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND 분류코드 = '" & Mid(Text1(0).Text, 1, 2) & "' " _
                      & "AND 세부코드 = '" & Mid(Text1(0).Text, 3) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 0
                            .TextMatrix(.Row, 0) = Left(Text1(0).Text, 2)                         '분류코드
                            .TextMatrix(.Row, 2) = Mid(Text1(0).Text, 3)                          '세부코드
                            .TextMatrix(.Row, 4) = .TextMatrix(.Row, 0) + .TextMatrix(.Row, 2)    '분류코드 + 세부코드
                            .Cell(flexcpData, .Row, 4, .Row, 4) = .TextMatrix(.Row, 4)
                       Case 1     '품명
                            .TextMatrix(.Row, 3) = Trim(Text1(1).Text)
                       Case 2, 3  '규격, 단위
                            .TextMatrix(.Row, lngC + 3) = Trim(Text1(lngC).Text)
                       Case 4     '과세구분
                            .TextMatrix(.Row, 14) = Format(DTOS(dtpInputDate.Value), "0000-00-00")   '최종입고일자
                            .TextMatrix(.Row, 15) = Format(DTOS(dtpOutputDate.Value), "0000-00-00")  '최종출고일자
                            .TextMatrix(.Row, 25) = Trim(txtBarCode.Text)
                            .TextMatrix(.Row, 26) = Vals(Trim(Text1(lngC).Text))                     '폐기율
                            .TextMatrix(.Row, 27) = Right(Trim(cboTaxGbn.Text), Len(Trim(cboTaxGbn.Text)) - 3) '과세구분
                            .TextMatrix(.Row, 28) = Vals(Left(cboState(1).Text, 1))                  '사용구분
                            .TextMatrix(.Row, 29) = cboState(1).ListIndex
                            .TextMatrix(.Row, 30) = Right(Trim(cboState(1).Text), Len(Trim(cboState(1).Text)) - 3)
                       Case 5     '적정재고
                            .TextMatrix(.Row, 7) = Vals(Trim(Text1(lngC).Text))
                       Case 9     '재고조정수량
                            .TextMatrix(.Row, 11) = Vals(Trim(Text1(lngC).Text)) + curAdjustAmt
                            Text1(lngC).Text = Format(Vals(Trim(Text1(lngC).Text)) + curAdjustAmt, "#,0")
                       Case 11    '현재재고
                            .TextMatrix(.Row, 13) = Vals(Trim(Text1(lngC).Text))
                       Case 12    '주매입처코드
                            .TextMatrix(.Row, 16) = Trim(Text1(lngC).Text)
                       Case 13    '주매입처명
                            .TextMatrix(.Row, 17) = Trim(Text1(lngC).Text)
                       Case 14 To 19  '단가
                            .TextMatrix(.Row, lngC + 4) = Vals(Trim(Text1(lngC).Text))
                       Case 20    '비고란
                            .TextMatrix(.Row, 24) = Trim(Text1(lngC).Text)
                       Case Else
                End Select
            Next lngC
            '미달상품표시
            If .ValueMatrix(.Row, 13) < .ValueMatrix(.Row, 7) Then
               .Cell(flexcpForeColor, .Row, 13, .Row, 13) = vbRed
               .Cell(flexcpFontBold, .Row, 13, .Row, 13) = True
            Else
               .Cell(flexcpForeColor, .Row, 13, .Row, 13) = vbBlack
               .Cell(flexcpFontBold, .Row, 13, .Row, 13) = False
            End If
         End If
         PB_adoCnnSQL.Execute strSQL
         PB_adoCnnSQL.CommitTrans
         .SetFocus
         Screen.MousePointer = vbDefault
         
    End With
    cmdSave.Enabled = True
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 읽기 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 변경 실패"
    Unload Me
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
                       & "WHERE 대분류코드 = '" & Left(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
                         & "AND 중분류코드 = '" & Mid(Trim(.TextMatrix(.Row, 0)), 3, 2) & "' " _
                         & "AND 소분류코드 = '" & Right(Trim(.TextMatrix(.Row, 0)), 2) & "' " _
                         & "AND 세분류코드 = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("해당건수")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "자재시세(" & Format(lngCnt, "#,#") & "건)가 있으므로 삭제할 수 없습니다.", vbCritical, "자재원장 삭제 불가"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               'strSQL = "DELETE FROM 자재원장 " _
                       & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND 분류코드 = '" & Mid(Trim(.TextMatrix(.Row, 0)), 1, 2) & "' " _
                         & "AND 세부코드 = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               strSQL = "UPDATE 자재원장 SET 사용구분 = 9 " _
                       & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND 분류코드 = '" & Mid(Trim(.TextMatrix(.Row, 0)), 1, 2) & "' " _
                         & "AND 세부코드 = '" & Trim(.TextMatrix(.Row, 2)) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               Text1(Text1.LBound).Enabled = False: Text1(1).Enabled = True
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 읽기 실패"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 삭제 실패"
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
    Set frm자재원장 = Nothing
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
    Text1(Text1.LBound).Enabled = False      '자재코드 FLASE
    Text1(Text1.LBound + 1).Enabled = True   '자재명 FLASE
    With vsfg1                 'Rows 1, Cols 31, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .Ellipsis = flexEllipsisEnd
         '.ExplorerBar = flexExSortShow
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 6
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 31
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '자재분류(분류코드) 'H
         .ColWidth(1) = 1000   '분류명
         .ColWidth(2) = 1550   '세부코드
         .ColWidth(3) = 2700   '품명
         .ColWidth(4) = 1900   '분류코드+세부코드  'H
         .ColWidth(5) = 2000   '규격
         .ColWidth(6) = 600    '단위
         .ColWidth(7) = 1000   '적정재고
         .ColWidth(8) = 1000   '이월재고
         .ColWidth(9) = 1000   '입고수량
         .ColWidth(10) = 1000  '출고수량
         .ColWidth(11) = 1000  '재고조정수량
         .ColWidth(12) = 1000  '재고이동수량
         .ColWidth(13) = 1000  '현재재고
         .ColWidth(14) = 1400  '최종입고일자
         .ColWidth(15) = 1400  '최종출고일자
         .ColWidth(16) = 1200  '매입처코드
         .ColWidth(17) = 3000  '매입처명
         .ColWidth(18) = 1350  '입고단가1
         .ColWidth(19) = 1350  '입고단가2
         .ColWidth(20) = 1350  '입고단가3
         .ColWidth(21) = 1350  '출고단가1
         .ColWidth(22) = 1350  '출고단가2
         .ColWidth(23) = 1350  '출고단가3
         .ColWidth(24) = 9400  '비고란
         .ColWidth(25) = 3000  '바코드
         .ColWidth(26) = 1000  '폐기율
         .ColWidth(27) = 1000  '과세구분
         .ColWidth(28) = 1     '사용구분
         .ColWidth(29) = 1     '사용구분ListIndex
         .ColWidth(30) = 1000  '사용구분
         
         .Cell(flexcpFontBold, 0, 0, .FixedRows - 1, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "분류코드"         'H
         .TextMatrix(0, 1) = "분류명"           'H
         .TextMatrix(0, 2) = "코드"             'H
         .TextMatrix(0, 3) = "품명"
         .TextMatrix(0, 4) = "코드"             'H(분류코드+세부코드)
         .TextMatrix(0, 5) = "규격"
         .TextMatrix(0, 6) = "단위"
         .TextMatrix(0, 7) = "적정재고"
         .TextMatrix(0, 8) = "이월재고"
         .TextMatrix(0, 9) = "매입수량"
         .TextMatrix(0, 10) = "매출수량"
         .TextMatrix(0, 11) = "재고조정"
         .TextMatrix(0, 12) = "재고이동"
         .TextMatrix(0, 13) = "현재재고"
         .TextMatrix(0, 14) = "최종매입일자"
         .TextMatrix(0, 15) = "최종매출일자"
         .TextMatrix(0, 16) = "매입처코드"
         .TextMatrix(0, 17) = "매입처명"
         .TextMatrix(0, 18) = "매입단가1"
         .TextMatrix(0, 19) = "매입단가2"
         .TextMatrix(0, 20) = "매입단가3"
         .TextMatrix(0, 21) = "매출단가1"
         .TextMatrix(0, 22) = "매출단가2"
         .TextMatrix(0, 23) = "매출단가3"
         .TextMatrix(0, 24) = "비고란"
         .TextMatrix(0, 25) = "바코드"
         .TextMatrix(0, 26) = "폐기율"
         .TextMatrix(0, 27) = "과세구분"
         .TextMatrix(0, 28) = "사용구분"        'H
         .TextMatrix(0, 29) = "사용구분"        'H
         .TextMatrix(0, 30) = "사용구분"
         For lngC = 7 To 13
             .ColFormat(lngC) = "#,#"
         Next lngC
         For lngC = 18 To 23
             .ColFormat(lngC) = "#,#.00"
         Next lngC
         .ColFormat(26) = "#,#.00"
         .ColHidden(0) = True: .ColHidden(4) = True
         .ColHidden(28) = True: .ColHidden(29) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 1, 2, 3, 4, 5, 6, 16, 17, 24, 25
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 14, 15, 27, 28, 29, 30
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
    Text1(Text1.LBound).Enabled = False: Text1(Text1.LBound + 1).Enabled = True
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.분류코드 AS 분류코드, " _
                  & "ISNULL(T1.분류명,'') AS 분류명 " _
             & "FROM 자재분류 T1 " _
            & "ORDER BY T1.분류코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cboMtGp.ListIndex = -1
       cmdClear.Enabled = False: cmdFind.Enabled = False
       cmdSave.Enabled = False: cmdDelete.Enabled = False
       Exit Sub
    Else
       cboMtGp.AddItem "00. " & "전체"
       Do Until P_adoRec.EOF
          cboMtGp.AddItem P_adoRec("분류코드") & ". " & P_adoRec("분류명")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       cboMtGp.ListIndex = 0
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
    'dtpInputDate.Value = Format(Left(PB_regUserinfoU.UserClientDate, 4) + "0101", "0000-00-00")
    dtpInputDate.Value = Format("19000101", "0000-00-00")
    dtpOutputDate.Value = Format("19000101", "0000-00-00")
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
Dim strGroupBy As String
Dim strHaving  As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
Dim lngRRR     As Long

    'If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3)) = 0 Then
    '   txtFindCD.SetFocus
    '   Exit Sub
    'End If
    With vsfg1
         '검색조건 자재분류
         Select Case Left(Trim(cboMtGp.Text), 2)
                Case "00" '전체
                     strWhere = ""
                Case Else
                     strWhere = "WHERE T1.분류코드 = '" & Mid(Trim(cboMtGp.Text), 1, 2) & "' "
         End Select
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
    End With
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
    '정상적인 조회
    If (Len(P_strFindString1) + Len(P_strFindString2) + Len(P_strFindString3) + Len(P_strFindString4)) = 0 Then
       strOrderBy = "ORDER BY T1.사업장코드, T1.분류코드, T1.세부코드 "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") + " " _
                     & "T1.세부코드 LIKE '%" & P_strFindString1 & "%' AND T3.자재명 LIKE '%" & P_strFindString2 & "%' " _
                & " AND T3.규격 LIKE '%" & P_strFindString3 & "%' AND T3.바코드 LIKE '%" & P_strFindString4 & "%' "
       strOrderBy = "ORDER BY T1.사업장코드, T1.분류코드, T1.세부코드 "
    End If
    '??CODE????? 로된 품목 제외
    If chkCodeException.Value = 1 Then
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                + "NOT (DATALENGTH(T1.세부코드) = 9 AND UPPER(SUBSTRING(T1.세부코드, 1, 4)) = 'CODE' " _
                + "AND T1.세부코드 LIKE 'CODE_____' " _
                + "AND ISNUMERIC(SUBSTRING(T1.세부코드, 5, 5)) = 1) "
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                  & "ISNULL(T1.분류코드,'') AS 분류코드, ISNULL(T4.분류명,'') AS 분류명, " _
                  & "ISNULL(T1.세부코드,'') AS 세부코드, T3.자재명 AS 자재명, " _
                  & "T3.규격 AS 규격, T3.단위 AS 단위, T3.바코드 AS 바코드, T3.폐기율 AS 폐기율, T3.과세구분 AS 과세구분, " _
                  & "T1.사용구분 AS 사용구분, T1.적정재고 AS 적정재고, " _
                  & "ISNULL(T1.주매입처코드,'') AS 주매입처코드, ISNULL(T5.매입처명, '') AS 주매입처명, " _
                  & "ISNULL(T1.최종입고일자,'') AS 최종입고일자, ISNULL(T1.최종출고일자,'') AS 최종출고일자, " _
                  & "ISNULL(T1.입고단가1,0) AS 입고단가1, ISNULL(T1.입고단가2,0) AS 입고단가2, ISNULL(T1.입고단가3,0) AS 입고단가3," _
                  & "ISNULL(T1.출고단가1,0) AS 출고단가1, ISNULL(T1.출고단가2,0) AS 출고단가2, ISNULL(T1.출고단가3,0) AS 출고단가3," _
                  & "ISNULL(T1.비고란, '') AS 비고란, " _
                  & "(SELECT ISNULL(SUM(입고누계수량-출고누계수량), 0) " _
                     & "FROM 자재원장마감 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 " _
                      & "AND 마감년월 >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND 마감년월 < '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS 이월재고, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(입고수량), 0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND (입출고구분 = 1) " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 입고수량, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(출고수량), 0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND (입출고구분 = 2) " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 출고수량, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(입고수량 - 출고수량), 0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND (입출고구분 = 5 OR 입출고구분 = 6) " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 재고조정수량, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(입고수량 - 출고수량), 0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND (입출고구분 = 11 OR 입출고구분 = 12) " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 재고이동수량 "
    strSQL = strSQL _
             & "FROM 자재원장 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 자재 T3 " _
                    & "ON T3.분류코드 = T1.분류코드 AND T3.세부코드 = T1.세부코드 " _
             & "LEFT JOIN 자재분류 T4 ON T4.분류코드 = T1.분류코드 " _
             & "LEFT JOIN 매입처 T5 ON T5.사업장코드 = T1.사업장코드 AND T5.매입처코드 = T1.주매입처코드 "
    strSQL = strSQL _
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
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("적정재고")), 0, P_adoRec("적정재고"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("이월재고")), 0, P_adoRec("이월재고"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("입고수량")), 0, P_adoRec("입고수량"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("출고수량")), 0, P_adoRec("출고수량"))
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("재고조정수량")), 0, P_adoRec("재고조정수량"))
               .TextMatrix(lngR, 12) = IIf(IsNull(P_adoRec("재고이동수량")), 0, P_adoRec("재고이동수량"))
               .TextMatrix(lngR, 13) = .ValueMatrix(lngR, 8) + .ValueMatrix(lngR, 9) - .ValueMatrix(lngR, 10) _
                                     + .ValueMatrix(lngR, 11) + .ValueMatrix(lngR, 12) '현재재고
               If .ValueMatrix(lngR, 7) <> 0 And .ValueMatrix(lngR, 13) < .ValueMatrix(lngR, 7) Then
                  .Cell(flexcpForeColor, lngR, 13, lngR, 13) = vbRed
                  .Cell(flexcpFontBold, lngR, 13, lngR, 13) = True
               End If
               If Len(P_adoRec("최종입고일자")) = 8 Then
                  .TextMatrix(lngR, 14) = Format(P_adoRec("최종입고일자"), "0000-00-00")
               End If
               If Len(P_adoRec("최종출고일자")) = 8 Then
                  .TextMatrix(lngR, 15) = Format(P_adoRec("최종출고일자"), "0000-00-00")
               End If
               .TextMatrix(lngR, 16) = P_adoRec("주매입처코드")
               .TextMatrix(lngR, 17) = P_adoRec("주매입처명")
               .TextMatrix(lngR, 18) = IIf(IsNull(P_adoRec("입고단가1")), 0, P_adoRec("입고단가1"))
               .TextMatrix(lngR, 19) = IIf(IsNull(P_adoRec("입고단가2")), 0, P_adoRec("입고단가2"))
               .TextMatrix(lngR, 20) = IIf(IsNull(P_adoRec("입고단가3")), 0, P_adoRec("입고단가3"))
               .TextMatrix(lngR, 21) = IIf(IsNull(P_adoRec("출고단가1")), 0, P_adoRec("출고단가1"))
               .TextMatrix(lngR, 22) = IIf(IsNull(P_adoRec("출고단가2")), 0, P_adoRec("출고단가2"))
               .TextMatrix(lngR, 23) = IIf(IsNull(P_adoRec("출고단가3")), 0, P_adoRec("출고단가3"))
               .TextMatrix(lngR, 24) = IIf(IsNull(P_adoRec("비고란")), "", P_adoRec("비고란"))
               .TextMatrix(lngR, 25) = IIf(IsNull(P_adoRec("바코드")), "", P_adoRec("바코드"))
               .TextMatrix(lngR, 26) = IIf(IsNull(P_adoRec("폐기율")), 0, P_adoRec("폐기율"))
               If P_adoRec("과세구분") = 0 Then
                  .TextMatrix(lngR, 27) = "비 과 세"
               Else
                  .TextMatrix(lngR, 27) = "과    세"
               End If
               .TextMatrix(lngR, 28) = IIf(IsNull(P_adoRec("사용구분")), "", P_adoRec("사용구분"))
               'ListIndex
               For lngRRR = 0 To cboState(1).ListCount - 1
                   If .ValueMatrix(lngR, 28) = Vals(Left(cboState(1).List(lngRRR), 1)) Then
                      .TextMatrix(lngR, 29) = lngRRR
                      .TextMatrix(lngR, 30) = Right(Trim(cboState(1).List(lngRRR)), Len(Trim(cboState(1).List(lngRRR))) - 3)
                      Exit For
                   End If
               Next lngRRR
               If .TextMatrix(lngR, 3) = P_strFindString1 Then
                  lngRR = lngR
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------------------+
'/// Clear text1(index) ///
'+-------------------------+
Private Sub SubClearText()
Dim lngC As Long
    txtBarCode.Text = ""
    For lngC = Text1.LBound To Text1.UBound
        Text1(lngC).Text = ""
    Next lngC
    dtpInputDate.Value = Format("19000101", "0000-00-00")
    dtpOutputDate.Value = Format("19000101", "0000-00-00")
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0  '자재코드
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (Len(Text1(lngC).Text) >= 1 And Len(Text1(lngC).Text) <= 18) Then
                       Text1(lngC).Text = ""
                       Exit Function
                    End If
               Case 1  '자재명
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    'If Not Len(Text1(lngC).Text) > 0 Then
                    '   Text1(lngC).Text = ""
                    '   Exit Function
                    'End If
               Case 20  '비고란
                    Text1(lngC).Text = Trim(Text1(lngC).Text)
                    If Not (LenH(Text1(lngC).Text) <= 100) Then
                       Exit Function
                    End If
        End Select
    Next lngC
    blnOK = True
End Function

'+---------------------------+
'/// 현재재고 0 으로 조정 ///
'+---------------------------+
Private Sub cmdZero_Click()
Dim strSQL         As String
Dim strGroupBy     As String
Dim strHaving      As String
Dim strWhere       As String
Dim strOrderBy     As String
Dim intRetVal      As Integer
Dim lngCnt         As Long
Dim curAdjustAmt   As Currency
Dim lngLogCnt      As Long
Dim strServerTime  As String
Dim strTime        As String
Dim CurInputMny    As Currency '입고단가
Dim CurInputVat    As Currency '입고부가
Dim CurOutPutMny   As Currency '출고단가
Dim CurOutPutVat   As Currency '출고부가

    intRetVal = MsgBox("현재재고를 0 으로 자동 재고조정하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton2, "현재재고조정")
    If intRetVal = vbNo Then
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    cmdZero.Enabled = False
    With vsfg1
         strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                  & "T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " 'AND T1.사용구분 = 0 AND T3.사용구분 = 0 " 'T3(자재)
    End With
    strOrderBy = "ORDER BY T1.사업장코드, T1.분류코드, T1.세부코드  "
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.사업장코드 AS 사업장코드, ISNULL(T1.분류코드,'') AS 분류코드, ISNULL(T1.세부코드,'') AS 세부코드, " _
                  & "T1.입고단가1 AS 입고단가1, T1.입고단가2 AS 입고단가2, T1.입고단가1 AS 입고단가3, " _
                  & "T1.입고단가1 AS 출고단가1, T1.출고단가2 AS 출고단가2, T1.출고단가3 AS 출고단가3, " _
                  & "(SELECT ISNULL(SUM(입고누계수량 - 출고누계수량), 0) " _
                     & "FROM 자재원장마감 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 " _
                      & "AND 마감년월 >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND 마감년월 < '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS 이월재고, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(입고수량), 0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND (입출고구분 = 1 OR 입출고구분 = 3) " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 입고수량, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(출고수량), 0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND (입출고구분 = 2 OR 입출고구분 = 4) " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 출고수량, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(입고수량 - 출고수량), 0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND (입출고구분 = 5 OR 입출고구분 = 6) " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 재고조정수량, "
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(입고수량 - 출고수량), 0) " _
                     & "FROM 자재입출내역 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 AND 사용구분 = 0 " _
                      & "AND (입출고구분 = 11 OR 입출고구분 = 12) " _
                      & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                         & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 재고이동수량 "
    strSQL = strSQL _
             & "FROM 자재원장 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 자재 T3 " _
                    & "ON T3.분류코드 = T1.분류코드 AND T3.세부코드 = T1.세부코드 " _
             & "LEFT JOIN 자재분류 T4 ON T4.분류코드 = T1.분류코드 " _
             & "LEFT JOIN 매입처 T5 ON T5.사업장코드 = T1.사업장코드 AND T5.매입처코드 = T1.주매입처코드 "
    strSQL = strSQL _
           & "" & strWhere & " " _
           & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdZero.Enabled = False
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       PB_adoCnnSQL.BeginTrans
       Do Until P_adoRec.EOF
          curAdjustAmt = (P_adoRec("이월재고") + P_adoRec("입고수량") - P_adoRec("출고수량") _
                        + P_adoRec("재고조정수량") + P_adoRec("재고이동수량")) * -1
          If curAdjustAmt <> 0 Then
             '서버시간 구하기
             Screen.MousePointer = vbHourglass
             P_adoRec1.CursorLocation = adUseClient
             strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
             On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
             P_adoRec1.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
             strServerTime = Mid(P_adoRec1("서버시간"), 1, 2) + Mid(P_adoRec1("서버시간"), 4, 2) + Mid(P_adoRec1("서버시간"), 7, 2) _
                           + Mid(P_adoRec1("서버시간"), 10)
             P_adoRec1.Close
             strTime = strServerTime
          End If
          If curAdjustAmt <> 0 Then
             CurInputMny = P_adoRec("입고단가1"): CurInputVat = Fix(P_adoRec("입고단가1") * (PB_curVatRate))
             CurOutPutMny = P_adoRec("출고단가1"): CurOutPutVat = Fix(P_adoRec("출고단가1") * (PB_curVatRate))
          End If
          If curAdjustAmt > 0 Then '입고(+)
             '거래번호 구하기
             P_adoRec1.CursorLocation = adUseClient
             strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "5" & "', " _
                                  & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
             On Error GoTo ERROR_STORED_PROCEDURE
             P_adoRec1.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
             lngLogCnt = P_adoRec1(0)
             P_adoRec1.Close
             '자재입출내역
             strSQL = "INSERT INTO 자재입출내역(사업장코드, 분류코드, " _
                                             & "세부코드, 입출고구분, " _
                                             & "입출고일자, 입출고시간, " _
                                             & "입고수량, 입고단가, " _
                                             & "입고부가, 출고수량, " _
                                             & "출고단가, 출고부가, " _
                                             & "매입처코드, 매출처코드, " _
                                             & "원래입출고일자, 직송구분, " _
                                             & "발견일자, 발견번호, 거래일자, 거래번호, " _
                                             & "계산서발행여부, 현금구분, 감가구분, 적요, 작성년도, 책번호, 일련번호, " _
                                             & "사용구분, 수정일자, " _
                                             & "사용자코드, 재고이동사업장코드) VALUES( " _
                       & "'" & PB_regUserinfoU.UserBranchCode & "', '" & P_adoRec("분류코드") & "', " _
                       & "'" & P_adoRec("세부코드") & "', 5, " _
                       & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                       & "" & curAdjustAmt & ", " & CurInputMny & ", " _
                       & "" & CurInputVat & ", 0, " _
                       & "0, 0, " _
                       & "'', '', " _
                       & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                       & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                       & "0, 0, 0, '', '', 0, 0, " _
                       & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                       & "'" & PB_regUserinfoU.UserCode & "', '' ) "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
          ElseIf _
             curAdjustAmt < 0 Then '출고(-)
             '거래번호 구하기
             P_adoRec1.CursorLocation = adUseClient
             strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "6" & "', " _
                                  & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
             On Error GoTo ERROR_STORED_PROCEDURE
             P_adoRec1.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
             lngLogCnt = P_adoRec1(0)
             P_adoRec1.Close
             strSQL = "INSERT INTO 자재입출내역(사업장코드, 분류코드, " _
                                             & "세부코드, 입출고구분, " _
                                             & "입출고일자, 입출고시간, " _
                                             & "입고수량, 입고단가, " _
                                             & "입고부가, 출고수량, " _
                                             & "출고단가, 출고부가, " _
                                             & "매입처코드, 매출처코드, " _
                                             & "원래입출고일자, 직송구분, " _
                                             & "발견일자, 발견번호, 거래일자, 거래번호, " _
                                             & "계산서발행여부, 현금구분, 감가구분, 적요, 작성년도, 책번호, 일련번호, " _
                                             & "사용구분, 수정일자, " _
                                             & "사용자코드, 재고이동사업장코드) VALUES( " _
                       & "'" & PB_regUserinfoU.UserBranchCode & "', '" & P_adoRec("분류코드") & "', " _
                       & "'" & P_adoRec("세부코드") & "', 6, " _
                       & "'" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                       & "0, 0, " _
                       & "0, " & (curAdjustAmt * -1) & ", " _
                       & "" & CurOutPutMny & ", " & CurOutPutVat & ", " _
                       & "'', '', " _
                       & "'" & PB_regUserinfoU.UserClientDate & "', 0, " _
                       & "'', 0, '" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                       & "0, 0, 0, '', '', 0, 0, " _
                       & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                       & "'" & PB_regUserinfoU.UserCode & "', '' ) "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
          End If
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
       PB_adoCnnSQL.CommitTrans
    End If
    cmdZero.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재원장 읽기 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자동재고조정 실패"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "로그 변경 실패"
    Unload Me
    Exit Sub
End Sub

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
Dim strForPrtDateTime      As String  '출력일시           (Formula)

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
            strExeFile = App.Path & ".\자재원장.rpt"
         Else
            strExeFile = App.Path & ".\자재원장T.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForAppPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                       '사업장명
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                                  '출력일시
            .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'"   '적용일자
            'DECLARE @ParAppPgDate VarChar(8), @ParMtGroupCode VarChar(6),  @ParMtName VarChar(20), @ParStateCode int
            '--- Parameter Fields ---
            '프로그램실행일자
            .StoredProcParam(0) = PB_regUserinfoU.UserClientDate
            '자재분류(분류코드)
            If Mid(cboMtGp.Text, 1, 2) = "00" Then
               .StoredProcParam(1) = " "
            Else
               .StoredProcParam(1) = Mid(cboMtGp.Text, 1, 2)
            End If
            '자재명
            If Len(txtFindNM.Text) = 0 Then
               .StoredProcParam(2) = " "
            Else
               .StoredProcParam(2) = Trim(txtFindNM.Text)
            End If
            '사용구분(0.전체, 1.정상, 2.삭제, 3.오 류)
            If cboState(0).ListIndex < 2 Then
               .StoredProcParam(3) = 0
            Else
               .StoredProcParam(3) = 9
            End If
            .StoredProcParam(4) = PB_regUserinfoU.UserBranchCode
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
            .WindowShowGroupTree = True
            .WindowShowPrintSetupBtn = True
            .WindowTop = 0: .WindowTop = 0: .WindowHeight = 11100: .WindowWidth = 15405
            .WindowState = crptMaximized
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "자재원장"
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

