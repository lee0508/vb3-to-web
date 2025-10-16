VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm매출세금계산서 
   BorderStyle     =   0  '없음
   Caption         =   "매출세금계산서"
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
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   4920
         Style           =   2  '드롭다운 목록
         TabIndex        =   49
         Top             =   240
         Width           =   2235
      End
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "건별"
         Height          =   255
         Left            =   7200
         TabIndex        =   32
         Top             =   150
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "전체"
         Height          =   255
         Left            =   7200
         TabIndex        =   31
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "매출세금계산서.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   27
         Top             =   195
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   0
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "매출세금계산서.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "매출세금계산서.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "매출세금계산서.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "매출세금계산서.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "매출세금계산서.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   0
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "매출처미수금 장부 조회및수정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
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
      Height          =   7876
      Left            =   60
      TabIndex        =   18
      Top             =   2055
      Width           =   15195
      _cx             =   26802
      _cy             =   13892
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
      Height          =   1395
      Left            =   60
      TabIndex        =   22
      Top             =   630
      Width           =   15195
      Begin VB.TextBox txtT_No2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   13440
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "50"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtT_No1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   12600
         MaxLength       =   4
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optNo 
         Caption         =   "번호"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   54
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton optDate 
         Caption         =   "일자"
         Height          =   375
         Left            =   480
         TabIndex        =   53
         Top             =   600
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox txtF_No2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   10200
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "1"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtF_No1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   9360
         MaxLength       =   4
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtpF_Year 
         Height          =   255
         Left            =   8280
         TabIndex        =   5
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy"
         Format          =   57540611
         UpDown          =   -1  'True
         CurrentDate     =   38268
      End
      Begin VB.ComboBox cboUsage 
         Height          =   300
         Left            =   13440
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboCredit 
         Height          =   300
         Left            =   11520
         Style           =   2  '드롭다운 목록
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox cboRS 
         Height          =   300
         Left            =   9720
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboMny 
         Height          =   300
         Left            =   7920
         Style           =   2  '드롭다운 목록
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboMake 
         Height          =   300
         Left            =   6240
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboPrint 
         Height          =   300
         Left            =   4200
         Style           =   2  '드롭다운 목록
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.ListBox lstPort 
         Height          =   240
         Left            =   8280
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "저장후 인쇄"
         Enabled         =   0   'False
         Height          =   375
         Left            =   13440
         TabIndex        =   38
         Top             =   195
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   1
         Left            =   5050
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   3700
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   3480
         TabIndex        =   3
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   5640
         TabIndex        =   4
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Year 
         Height          =   255
         Left            =   11520
         TabIndex        =   8
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy"
         Format          =   57540611
         UpDown          =   -1  'True
         CurrentDate     =   38268
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   25
         Left            =   1680
         TabIndex        =   60
         Top             =   690
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   24
         Left            =   240
         TabIndex        =   59
         Top             =   690
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   23
         Left            =   14640
         TabIndex        =   58
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   22
         Left            =   14040
         TabIndex        =   57
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         Height          =   240
         Index           =   21
         Left            =   13200
         TabIndex        =   56
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         Height          =   240
         Index           =   20
         Left            =   12360
         TabIndex        =   55
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   19
         Left            =   10800
         TabIndex        =   52
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         Height          =   240
         Index           =   18
         Left            =   9960
         TabIndex        =   51
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         Height          =   240
         Index           =   17
         Left            =   9120
         TabIndex        =   50
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   16
         Left            =   14640
         TabIndex        =   48
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "사용 :"
         Height          =   240
         Index           =   14
         Left            =   12720
         TabIndex        =   47
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "미수 :"
         Height          =   240
         Index           =   13
         Left            =   10800
         TabIndex        =   46
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "영청 :"
         Height          =   240
         Index           =   12
         Left            =   9000
         TabIndex        =   45
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   11
         Left            =   3000
         TabIndex        =   44
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "구분"
         Height          =   240
         Index           =   10
         Left            =   2160
         TabIndex        =   43
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "금액 :"
         Height          =   240
         Index           =   9
         Left            =   7200
         TabIndex        =   42
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "작성 :"
         Height          =   240
         Index           =   8
         Left            =   5520
         TabIndex        =   41
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "발행 :"
         Height          =   240
         Index           =   7
         Left            =   3480
         TabIndex        =   40
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   6
         Left            =   7080
         TabIndex        =   37
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   5
         Left            =   4920
         TabIndex        =   36
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "/"
         Height          =   240
         Index           =   4
         Left            =   7995
         TabIndex        =   35
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   3
         Left            =   3000
         TabIndex        =   34
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "발행일자"
         Height          =   240
         Index           =   2
         Left            =   1680
         TabIndex        =   33
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "전체금액"
         Height          =   240
         Index           =   1
         Left            =   8880
         TabIndex        =   30
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label lblTotMny 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "0.00"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10200
         TabIndex        =   29
         Top             =   285
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   3000
         TabIndex        =   28
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처코드"
         Height          =   240
         Index           =   0
         Left            =   1680
         TabIndex        =   26
         Top             =   285
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
         Left            =   480
         TabIndex        =   25
         Top             =   285
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm매출세금계산서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 매출세금계산서
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 사업장, 매출처, 자재입출내역, 매출세금계산서장부
' 업  무  설  명 : 매출세금계산서장부 내역 조회/수정/삭제
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_adoRecW            As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private Const PC_intRowCnt1  As Integer = 25   '그리드1의 한 페이지 당 행수(FixedRows 포함)

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
Dim strSQL             As String
Dim inti               As Integer

Dim p                  As Printer
Dim strDefaultPrinter  As String
Dim aryPrinter()       As String
Dim strBuffer          As String

    frmMain.SBar.Panels(4).Text = "매출세금계산서장부를 수정/삭제하여도 세금계산서에는 반영되지 않으며, 미수금 수금내역에도 반영되지 않습니다. "
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       
       'Printer Setting and Seaching(API)
       strBuffer = Space(1024)
       inti = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
       aryPrinter = Split(strBuffer, ",")
       strDefaultPrinter = Trim(Trim(aryPrinter(0)))
       For Each p In Printers
           cboPrinter.AddItem Trim(p.DeviceName)
           lstPort.AddItem p.Port
       Next
       For inti = 0 To cboPrinter.ListCount - 1
           cboPrinter.ListIndex = inti
           If UCase(Trim(cboPrinter.Text)) = UCase(Trim(strDefaultPrinter)) Then
              Exit For
           End If
       Next inti
       '---
       Subvsfg1_INIT  '세금계산서
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '인쇄, 조회
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '추가, 저장, 인쇄, 조회
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = False
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       SubOther_FILL
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매출세금계산서(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+--------------------+
'--- Select Printer ---
'+--------------------+
Private Sub cboPrinter_Click()
    lstPort.ListIndex = cboPrinter.ListIndex
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
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '매출처검색
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
       frm매출처검색.Show vbModal
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매출처 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------+
'/// 매출처 ///
'+-------------+
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
Dim lngR   As Long
    With Text1(Index)
         Select Case Index
                Case 0
                     .Text = UPPER(Trim(.Text))
                     If Len(.Text) < 1 Then
                        Text1(Index).Text = ""
                        Text1(Index + 1).Text = ""
                        Exit Sub
                     End If
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "TABLE 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+--------------+
'/// 기준조회 ///
'+--------------+
Private Sub optDate_Click()
    If optDate.Value = True Then
       dtpF_Date.Enabled = True: dtpT_Date.Enabled = True
       dtpF_Year.Enabled = False: txtF_No1.Enabled = False: txtF_No2.Enabled = False
       dtpT_Year.Enabled = False: txtT_No1.Enabled = False: txtT_No2.Enabled = False
    End If
    dtpF_Date.SetFocus
End Sub
Private Sub optDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub optNo_Click()
    If optNo.Value = True Then
       dtpF_Date.Enabled = False: dtpT_Date.Enabled = False
       dtpF_Year.Enabled = True: txtF_No1.Enabled = True: txtF_No2.Enabled = True
       dtpT_Year.Enabled = True: txtT_No1.Enabled = True: txtT_No2.Enabled = True
    End If
    dtpF_Year.SetFocus
End Sub
Private Sub opNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
'+--------------+
'/// 발행일자 ///
'+--------------+
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
'+---------------------+
'/// 세금계산서번호 ///
'+---------------------+
Private Sub dtpF_Year_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txtF_No1_GotFocus()
    With txtF_No1
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtF_No1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txtF_No2_GotFocus()
    With txtF_No2
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtF_No2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Year_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txtT_No1_GotFocus()
    With txtT_No1
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtT_No1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub txtT_No2_GotFocus()
    With txtT_No2
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtT_No2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
'+---------------+
'/// 구분선택 ///
'+---------------+
Private Sub cboPrint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboMake_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboMny_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboRS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboCredit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cboUsage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If cmdFind.Enabled = True Then
          cmdFind_Click
       Else
          cmdFind.SetFocus
       End If
    End If
End Sub

'+------------+
'/// vsfg1 ///
'+------------+
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
    With vsfg1
         P_intButton = Button
         If (.MouseRow < .FixedRows) Then
            Exit Sub
         End If
         If (.MouseRow >= .FixedRows And _
            .TextMatrix(.MouseRow, 20) = "정상") Then
            If cmdSave.Enabled = False Then Exit Sub
            If (.MouseCol = 8) Then      '공급가액
                If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            'ElseIf _
            '   (.MouseCol = 9) Then      '세액
            '   If Button = vbLeftButton Then
            '      .Select .MouseRow, .MouseCol
            '      .EditCell
            '    End If
            ElseIf _
               (.MouseCol = 11) Then     '품목및규격
               If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            ElseIf _
                (.MouseCol = 13) Then    '수량(종)
                If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            End If
         End If
    End With
End Sub
Sub vsfg1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim blnModify As Boolean
Dim curTmpMny As Currency
    With vsfg1
         If Row >= .FixedRows Then
            If (Col = 8) Then   '공급가액
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then '소숫점이하 사용불가
                     'IsNumeric(Right(.EditText, 1)) = False) Then                                            '소숫점이하 사용가
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 8)
                     .TextMatrix(Row, 8) = Vals(.EditText)
                     .TextMatrix(Row, 9) = Fix(Vals(.EditText) * (PB_curVatRate))  '부가세
                     .TextMatrix(Row, 10) = Vals(.EditText) + .ValueMatrix(Row, 9)
                  End If
               End If
            ElseIf _
               (Col = 9) Then  '세액
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 8)
                     .TextMatrix(Row, 10) = .ValueMatrix(Row, 8) + Vals(.EditText)
                  End If
               End If
            ElseIf _
               (Col = 11) Then  '품목및규격
               If .TextMatrix(Row, Col) <> .EditText Then
                  If .Cell(flexcpChecked, Row, 19) = flexChecked Then '미수이면(거래내역 있으면)
                     If .ValueMatrix(Row, 13) = 0 Then   '수량
                        If Not ((LenH(Trim(.EditText))) <= 50) Then
                           Beep
                           Cancel = True
                        Else
                           blnModify = True
                        End If
                     Else
                        If Not ((LenH(Trim(.EditText)) + LenH(" 외 ") + Len(.TextMatrix(Row, 13)) + LenH("종")) <= 50) Then
                           Beep
                           Cancel = True
                        Else
                           blnModify = True
                        End If
                     End If
                  Else                                                '미수아니면(거래내역 없으면)
                     If Not (LenH(Trim(.EditText)) <= 50) Then
                        Beep
                        Cancel = True
                     Else
                        blnModify = True
                     End If
                  End If
               End If
            ElseIf _
               (Col = 13) Then  '수량(종)
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     If .Cell(flexcpChecked, Row, 19) = flexChecked Then '미수이면(거래내역 있으면)
                        If Val(.EditText) = 0 Then '수량
                           If Not (LenH(.TextMatrix(Row, 11)) <= 50) Then
                              Beep
                              Cancel = True
                           Else
                              blnModify = True
                              curTmpMny = .ValueMatrix(Row, 8)
                           End If
                        Else
                           If Not ((LenH(Trim(.TextMatrix(Row, 11))) + LenH(" 외 ") + Len(.EditText) + LenH("종")) <= 50) Then
                              Beep
                              Cancel = True
                           Else
                              blnModify = True
                              curTmpMny = .ValueMatrix(Row, 8)
                           End If
                        End If
                     Else                                                '미수아니면(거래내역 없으면)
                        blnModify = True
                        curTmpMny = .ValueMatrix(Row, 8)
                     End If
                  End If
               End If
            'ElseIf _
            '   (Col = 18) Then  '계산서발행여부
            '   If (Len(.TextMatrix(Row, 9)) = 0) Then '매출처가 없음
            '      .Cell(flexcpChecked, Row, 18, Row, 18) = flexUnchecked
            '      Beep
            '      Cancel = True
            '      Exit Sub
            '   End If
            '   If .Cell(flexcpChecked, Row, Col) <> .EditText Then
            '      blnModify = True
            '   End If
            End If
            '변경표시 + 금액재계산
            If blnModify = True Then
               If .TextMatrix(Row, 21) = "" Then
                  .TextMatrix(Row, 21) = "U"
               End If
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
               Select Case Col
                      Case 8, 9, 13
                           lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - curTmpMny + .ValueMatrix(Row, 8), "#,#.00")
                      Case Else
               End Select
            End If
         End If
    End With
End Sub
Private Sub vsfg1_Click()
Dim strData As String
Dim lngR1   As Long
Dim lngRH1  As Long
Dim lngR2   As Long
Dim lngRR2  As Long
Dim lngC    As Long
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = Trim(.Cell(flexcpData, .Row, 4))
            Select Case .MouseCol
                   Case 4
                        '.ColSel = 4
                        .Select 0, 0, 0, 4
                        .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 5
                        .ColSel = 5
                        .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(5) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 4) > 0 Then
               .Row = .FindRow(strData, , 4)
            End If
            If PC_intRowCnt1 < .Rows Then
               .TopRow = .Row
            End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR1  As Long
Dim lngR2  As Long
Dim lngC   As Long
Dim lngCnt As Long
    With vsfg1
         If .Row >= .FixedRows Then
         End If
    End With
End Sub
Private Sub vsfg1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Dim lngR1  As Long
Dim lngR2  As Long
Dim lngC   As Long
Dim lngCnt As Long
    With vsfg1
         If NewRow <> OldRow Then
            'For lngR2 = 1 To vsfg2.Rows - 1
            '    vsfg2.RowHidden(lngR2) = True
            'Next lngR2
            'If NewRow > 0 Then 'Add 20041002
            '   For lngR1 = .ValueMatrix(.Row, 14) To .ValueMatrix(.Row, 15)
            '       vsfg2.RowHidden(lngR1) = False
            '       lngCnt = lngCnt + 1
            '   Next lngR1
            'End If
            'If PC_intRowCnt2 < lngCnt Then
            '   vsfg2.TopRow = vsfg2.Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lngR     As Long
Dim blnDupOK As Boolean
    With vsfg1
         If .Row >= .FixedRows Then
         End If
    End With
End Sub

'+-----------+
'/// 출력 ///
'+-----------+
Private Sub cmdPrint_Click()
Dim p              As Printer
Dim strSQL         As String
Dim lngR           As Long
Dim lngLogCnt      As Long
Dim strMakeYear    As String
Dim lngLogCnt1     As Long
Dim lngLogCnt2     As Long
Dim strServerTime  As String
Dim strTime        As String
    
    For Each p In Printers
        If Trim(p.DeviceName) = Trim(cboPrinter.Text) And lstPort.List(cboPrinter.ListIndex) = p.Port Then
           Set Printer = p
           Exit For
        End If
    Next
    
    If (vsfg1.Rows = 1) Then
       Exit Sub
    End If
    '서버시간 구하기
    cmdSave.Enabled = False
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) + Mid(P_adoRec("서버시간"), 7, 2) _
                  + Mid(P_adoRec("서버시간"), 10)
    P_adoRec.Close
    strTime = strServerTime
    If optPrtChk0.Value = True Then '세금계산서인쇄(건별)
       If (vsfg1.Row < 1) Then
          Screen.MousePointer = vbDefault
          Exit Sub
       End If
       With vsfg1
            If .TextMatrix(.Row, 20) = "정상" Then
               If .Cell(flexcpChecked, .Row, 15) = flexUnchecked Then
                  PB_adoCnnSQL.BeginTrans
                  strSQL = "UPDATE 매출세금계산서장부 SET " _
                                & "발행여부 = 1," _
                                & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                                & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                          & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 0) & "' " _
                            & "AND 작성일자 = '" & DTOS(.TextMatrix(.Row, 5)) & "' AND 작성시간 = '" & .TextMatrix(.Row, 22) & "' " _
                            & "AND 매출처코드 = '" & .TextMatrix(.Row, 6) & "' "
                   On Error GoTo ERROR_TABLE_UPDATE
                   PB_adoCnnSQL.Execute strSQL
                   '세금계산서번호가 있으면(Not Used)
                   'If Len(.TextMatrix(.Row, 1)) > 0 Then
                   '   strSQL = "UPDATE 세금계산서 SET " _
                   '             & "발행여부 = 1," _
                   '             & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                   '             & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                   '       & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 0) & "' " _
                   '         & "AND 작성년도 = '" & .TextMatrix(.Row, 1) & "' AND 책번호 = " & .ValueMatrix(.Row, 2) & " " _
                   '         & "AND 일련번호 = '" & .ValueMatrix(.Row, 3) & "' "
                   '   On Error GoTo ERROR_TABLE_UPDATE
                   '   PB_adoCnnSQL.Execute strSQL
                   'End If
                   PB_adoCnnSQL.CommitTrans
                   .Cell(flexcpChecked, .Row, 15) = flexChecked
                   .Cell(flexcpText, .Row, 15) = "발행"
               End If
               '사업장코드, 작성일자, 작성시간, 매출처코드
               SubPubPrint_TaxBill p, PB_intPrtTypeGbn, .TextMatrix(.Row, 0), "", 0, 0, _
                           1, DTOS(.TextMatrix(.Row, 5)), .TextMatrix(.Row, 22), .TextMatrix(.Row, 6)
            End If
       End With
    End If
    If optPrtChk1.Value = True Then '세금계산서인쇄(전체)
       PB_adoCnnSQL.BeginTrans
       With vsfg1
            For lngR = 1 To .Rows - 1
                If .TextMatrix(lngR, 20) = "정상" Then
                   If .Cell(flexcpChecked, lngR, 15) = flexUnchecked Then '미발행
                      strSQL = "UPDATE 매출세금계산서장부 SET " _
                                    & "발행여부 = 1," _
                                    & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                                    & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                              & "WHERE 사업장코드 = '" & .TextMatrix(lngR, 0) & "' " _
                                & "AND 작성일자 = '" & DTOS(.TextMatrix(lngR, 5)) & "' AND 작성시간 = '" & .TextMatrix(lngR, 22) & "' " _
                                & "AND 매출처코드 = '" & .TextMatrix(lngR, 6) & "' "
                       On Error GoTo ERROR_TABLE_UPDATE
                       PB_adoCnnSQL.Execute strSQL
                       '세금계산서번호가 있으면(Not Used)
                       'If Len(.TextMatrix(.Row, 1)) > 0 Then
                       '   strSQL = "UPDATE 세금계산서 SET " _
                       '                 & "발행여부 = 1," _
                       '                 & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                       '                 & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                       '           & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 0) & "' " _
                       '           & "AND 작성년도 = '" & .TextMatrix(lngR, 1) & "' AND 책번호 = " & .ValueMatrix(lngR, 2) & " " _
                       '           & "AND 일련번호 = '" & .ValueMatrix(lngR, 3) & "' "
                       '   On Error GoTo ERROR_TABLE_UPDATE
                       '   PB_adoCnnSQL.Execute strSQL
                       'End If
                       .Cell(flexcpChecked, lngR, 15) = flexChecked
                       .Cell(flexcpText, lngR, 15) = "발행"
                   End If
                   '사업장코드, 작성일자, 작성시간, 매출처코드
                   SubPubPrint_TaxBill p, PB_intPrtTypeGbn, .TextMatrix(lngR, 0), "", 0, 0, _
                           1, DTOS(.TextMatrix(lngR, 5)), .TextMatrix(lngR, 22), .TextMatrix(lngR, 6)
               End If
            Next lngR
       End With
       PB_adoCnnSQL.CommitTrans
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "변경 실패"
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
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
    '
End Sub
'+-----------+
'/// 조회 ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    lblTotMny.Caption = "0.00"
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
Dim p              As Printer
Dim blnSaveOK      As Boolean
Dim strSQL         As String
Dim lngR           As Long
Dim lngRR          As Long
Dim lngRRR         As Long
Dim lngC           As Long
Dim blnOK          As Boolean
Dim intRetVal      As Integer
Dim lngChkCnt      As Long
Dim lngDelCntS     As Long
Dim lngDelCntE     As Long
Dim lngDelCnt      As Long
Dim lngLogCnt      As Long
Dim intAddTax      As Integer '세금계산서(0.무작성, 1.변경, 2.추가, 3.삭제후추가)
Dim intCreateTax   As Integer '세금계산서발행여부(1.발행)
Dim strOldMakeYear As String
Dim lngOldLogCnt1  As Long
Dim lngOldLogCnt2  As Long
Dim strNewMakeYear As String
Dim lngNewLogCnt1  As Long
Dim lngNewLogCnt2  As Long
Dim curInputMoney  As Long
Dim curOutputMoney As Long
Dim strServerTime  As String
Dim strTime        As String
Dim strJukyo       As String  '자재입출내역의 적요

    For Each p In Printers
        If Trim(p.DeviceName) = Trim(cboPrinter.Text) And lstPort.List(cboPrinter.ListIndex) = p.Port Then
           Set Printer = p
           Exit For
        End If
    Next
    
    If vsfg1.Row >= vsfg1.FixedRows Then
       With vsfg1
            If (.TextMatrix(.Row, 21) = "U") Then
               blnSaveOK = True
            End If
            If blnSaveOK = False Then '저장할(변경된) 것이없으면
               Exit Sub
            End If
       End With
       intRetVal = MsgBox("수정된 매출세금계산서장부를 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "매출세금계산서장부 저장")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       '서버시간 구하기
       cmdSave.Enabled = False
       Screen.MousePointer = vbHourglass
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) + Mid(P_adoRec("서버시간"), 7, 2) _
                     + Mid(P_adoRec("서버시간"), 10)
       P_adoRec.Close
       strTime = strServerTime
       PB_adoCnnSQL.BeginTrans
       P_adoRec.CursorLocation = adUseClient
       With vsfg1
            strSQL = "UPDATE 매출세금계산서장부 SET " _
                          & "품목및규격 = '" & .TextMatrix(.Row, 11) & "', " _
                          & "수량 = " & .ValueMatrix(.Row, 13) & ", " _
                          & "공급가액 = " & .ValueMatrix(.Row, 8) & "," _
                          & "세액 = " & .ValueMatrix(.Row, 9) & ", " _
                          & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 0) & "' " _
                      & "AND 작성일자 = '" & DTOS(.TextMatrix(.Row, 5)) & "' " _
                      & "AND 작성시간 = " & .ValueMatrix(.Row, 22) & " "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
             '(색상 원위치)
             .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = vbWhite
             .Cell(flexcpForeColor, .Row, .FixedCols, .Row, .Cols - 1) = vbBlack
             .Cell(flexcpText, .Row, 21, .Row, 21) = ""  'SQL구분 지움
       End With
       PB_adoCnnSQL.CommitTrans
       If (chkPrint.Value = 1) And (cboPrinter.ListIndex >= 0) Then                        '세금계산서 출력
          With vsfg1
               SubPubPrint_TaxBill p, PB_intPrtTypeGbn, .TextMatrix(.Row, 0), "", 0, 0, _
                           1, DTOS(.TextMatrix(.Row, 5)), .TextMatrix(.Row, 22), .TextMatrix(.Row, 6)
          End With
       End If
       Screen.MousePointer = vbDefault
    End If
    cmdSave.Enabled = True
    vsfg1.SetFocus
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "검색 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "저장 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "삭제 실패"
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
Dim strSQL     As String
Dim lngR       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
Dim lngC       As Long
Dim blnOK      As Boolean
Dim intRetVal  As Integer
Dim lngChkCnt  As Long
Dim lngDelCntS As Long
Dim lngDelCntE As Long
Dim lngLogCnt  As Long
    If (vsfg1.Row >= vsfg1.FixedRows And vsfg1.TextMatrix(vsfg1.Row, 20) = "정상") Then
       intRetVal = MsgBox("매출세금계산서장부에서 삭제하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton2, "매출세금계산서장부 삭제")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       cmdDelete.Enabled = False
       Screen.MousePointer = vbHourglass
       With vsfg1
            lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 8), "#,#.00") '전체금액에서 제외
            PB_adoCnnSQL.BeginTrans
            If intRetVal = vbYes Then  '매출세금계산서장부 삭제
               strSQL = "UPDATE 매출세금계산서장부 SET " _
                             & "사용구분 = 9, " _
                             & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 0) & "' AND 작성일자 = '" & DTOS(.TextMatrix(.Row, 5)) & "' " _
                         & "AND 작성시간 = '" & .TextMatrix(.Row, 22) & "' "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
            End If
            PB_adoCnnSQL.CommitTrans
            lblTotMny.Caption = Format(Vals(lblTotMny.Caption) + .ValueMatrix(.Row, 8), "#,#.00")
           .RemoveItem .Row
           .Row = 0
       End With
       cmdFind.SetFocus
       Screen.MousePointer = vbDefault
    End If
    cmdDelete.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "검색 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "저장 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "삭제 실패"
    Unload Me
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
    Set frm매출세금계산서 = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    Text1(0).Text = "": Text1(1).Text = ""
    dtpF_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00") 'Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
    dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpF_Year.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpT_Year.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    cboPrint.AddItem "전  체"
    cboPrint.AddItem "미발행"
    cboPrint.AddItem "발  행"
    cboPrint.ListIndex = 0
    cboMake.AddItem "전체"
    cboMake.AddItem "거래"
    cboMake.AddItem "임의"
    cboMake.AddItem "일괄"
    cboMake.ListIndex = 0
    cboMny.AddItem "전체"
    cboMny.AddItem "현금"
    cboMny.AddItem "수표"
    cboMny.AddItem "어음"
    cboMny.AddItem "외상"
    cboMny.ListIndex = 0
    cboRS.AddItem "전체"
    cboRS.AddItem "영수"
    cboRS.AddItem "청구"
    cboRS.ListIndex = 0
    cboCredit.AddItem "전체"
    cboCredit.AddItem "일반"
    cboCredit.AddItem "미수"
    cboCredit.ListIndex = 0
    cboUsage.AddItem "전체"
    cboUsage.AddItem "정상"
    cboUsage.AddItem "삭제"
    cboUsage.ListIndex = 1
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub
End Sub
'+----------------------------------+
'/// VsFlexGrid(vsfg1) 초기화 ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1              'Rows 1, Cols 23, RowHeightMax(Min) 300
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
         .FixedCols = 5
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 23
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1200   '사업장코드 'H
         .ColWidth(1) = 1000   '작성년도   'H
         .ColWidth(2) = 1000   '책번호     'H
         .ColWidth(3) = 1000   '일련번호   'H
         .ColWidth(4) = 1450   'KEY        'H(사업장코드-작성일자-작성시간)
         .ColWidth(5) = 1200   '작성일자   '0000-00-00
         
         .ColWidth(6) = 1000   '매출처코드
         .ColWidth(7) = 2500   '매출처명
         .ColWidth(8) = 1600   '공급가액(단가)
         .ColWidth(9) = 1400   '세액(부가)
         .ColWidth(10) = 1600  '합계       'H
         .ColWidth(11) = 3700  '품목및규격
         .ColWidth(12) = 300   '외
         .ColWidth(13) = 500   '수량
         .ColWidth(14) = 300   '종
         .ColWidth(15) = 650   '발행여부
         .ColWidth(16) = 450   '작성구분
         .ColWidth(17) = 450   '금액구분
         .ColWidth(18) = 450   '영청구분
         .ColWidth(19) = 800   '미수구분
         .ColWidth(20) = 450   '사용구분
         .ColWidth(21) = 450   'SQL구분
         .ColWidth(22) = 1200  '작성시간
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "사업장코드" 'H
         .TextMatrix(0, 1) = "작성년도"   'H
         .TextMatrix(0, 2) = "책번호"     'H
         .TextMatrix(0, 3) = "일련번호"   'H
         .TextMatrix(0, 4) = "KEY"        'H
         .TextMatrix(0, 5) = "작성일자"
         .TextMatrix(0, 6) = "매출처코드" 'H
         .TextMatrix(0, 7) = "매출처명"
         .TextMatrix(0, 8) = "공급가액"
         .TextMatrix(0, 9) = "세액"
         .TextMatrix(0, 10) = "합계금액"  'H
         .TextMatrix(0, 11) = "품목및규격"
         .TextMatrix(0, 12) = "외"
         .TextMatrix(0, 13) = "수량"
         .TextMatrix(0, 14) = "종"
         .TextMatrix(0, 15) = "발행"
         .TextMatrix(0, 16) = "작성"
         .TextMatrix(0, 17) = "구분"
         .TextMatrix(0, 18) = "영청"
         .TextMatrix(0, 19) = "미수"
         .TextMatrix(0, 20) = "사용"
         .TextMatrix(0, 21) = "SQL"
         .TextMatrix(0, 22) = "작성시간"
         
         .ColHidden(0) = True: .ColHidden(1) = True: .ColHidden(2) = True: .ColHidden(3) = True: .ColHidden(4) = True:
         .ColHidden(6) = True: .ColHidden(10) = True: .ColHidden(21) = True: .ColHidden(22) = True
         .ColFormat(8) = "#,#.00": .ColFormat(9) = "#,#.00": .ColFormat(10) = "#,#.00"
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 7, 11, 12, 14
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 2, 3, 4, 5, 6, 15, 16, 17, 18, 19, 20, 21, 22
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         
         .ColComboList(16) = "거래|임의|일괄"
         .ColComboList(17) = "현금|수표|어음|외상"
         .ColComboList(18) = "영수|청구"
         .ColComboList(20) = "정상|삭제"
         
         '.MergeCells = flexMergeRestrictRows  'flexMergeFixedOnly
         '.MergeRow(0) = True
         'For lngC = 0 To 4
         '    .MergeCol(lngC) = True
         'Next lngC
    End With
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
    vsfg1.Rows = 1
    P_adoRec.CursorLocation = adUseClient
    If Len(Text1(0).Text) > 0 Then
       strWhere = "AND T1.매출처코드 = '" & Trim(Text1(0).Text) & "' "
    End If
    If optDate.Value = True Then '발행일자
       strWhere = strWhere + "AND T1.작성일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' "
       strOrderBy = "ORDER BY T1.사업장코드, T1.작성일자, T1.작성시간 "
    End If
    'If optNo.Value = True Then   '세금계산서번호
    '   strWhere = strWhere + "AND (T1.작성년도 BETWEEN '" & Year(dtpF_Year.Value) & "' AND '" & Year(dtpT_Year.Value) & "' ) " _
    '                       & "AND (T1.책번호 BETWEEN " & Vals(Trim(txtF_No1.Text)) & " AND " & Vals(Trim(txtT_No1.Text)) & " ) " _
    '                       & "AND (T1.일련번호 BETWEEN " & Vals(Trim(txtF_No2.Text)) & " AND " & Vals(Trim(txtT_No2.Text)) & ") "
    '   strOrderBy = "ORDER BY T1.사업장코드, T1.작성년도, T1.책번호, T1.일련번호 "
    'End If
    If cboPrint.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.발행여부 = " & (cboPrint.ListIndex - 1) & " "
    End If
    If cboMake.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.작성구분 = " & (cboMake.ListIndex - 1) & " "
    End If
    If cboMny.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.금액구분 = " & (cboMny.ListIndex - 1) & " "
    End If
    If cboRS.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.영청구분 = " & (cboRS.ListIndex - 1) & " "
    End If
    If cboCredit.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.미수구분 = " & (cboCredit.ListIndex - 1) & " "
    End If
    If cboUsage.ListIndex = 0 Then
       vsfg1.ColHidden(20) = False
    Else
       vsfg1.ColHidden(20) = True
    End If
    If cboUsage.ListIndex > 0 Then
       strWhere = strWhere + "AND T1.사용구분 = " & IIf(cboUsage.ListIndex = 1, 0, 9) & " "
    End If
    strSQL = "SELECT T1.사업장코드 AS 사업장코드, T1.작성일자 AS 작성일자, " _
                  & "T1.작성시간 AS 작성시간, T1.매출처코드 AS 매출처코드, T2.매출처명 AS 매출처명, " _
                  & "T1.공급가액 AS 공급가액, T1.세액 AS 세액, " _
                  & "T1.품목및규격 AS 품목및규격, T1.수량 AS 수량, " _
                  & "T1.금액구분 AS 금액구분, T1.영청구분 AS 영청구분, T1.발행여부 AS 발행여부, " _
                  & "T1.작성구분 AS 작성구분, T1.미수구분 AS 미수구분, T1.사용구분 AS 사용구분, " _
                  & "T1.작성년도 AS 작성년도, T1.책번호 AS 책번호, T1.일련번호 AS 일련번호 " _
             & "FROM 매출세금계산서장부 T1 " _
             & "LEFT JOIN 매출처 T2 ON T2.사업장코드 = T1.사업장코드 AND T2.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
              & "" & strWhere & " " _
              & "" & strOrderBy & " "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdExit.SetFocus
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt1 Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = IIf(IsNull(P_adoRec("사업장코드")), "", P_adoRec("사업장코드"))
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("작성년도")), "", P_adoRec("작성년도"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("책번호")), 0, P_adoRec("책번호"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("일련번호")), "", P_adoRec("일련번호"))
               .TextMatrix(lngR, 4) = P_adoRec("사업장코드") & "-" & P_adoRec("작성일자") & "-" & P_adoRec("작성시간")
               .Cell(flexcpData, lngR, 4, lngR, 4) = Trim(.TextMatrix(lngR, 4)) 'FindRow 사용을 위해
               .TextMatrix(lngR, 5) = Format(P_adoRec("작성일자"), "0000-00-00")
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("매출처코드")), "", P_adoRec("매출처코드"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("매출처명")), "", P_adoRec("매출처명"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("공급가액")), 0, P_adoRec("공급가액"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("세액")), 0, P_adoRec("세액"))
               .TextMatrix(lngR, 10) = .ValueMatrix(lngR, 8) + .ValueMatrix(lngR, 9)
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("품목및규격")), 0, P_adoRec("품목및규격"))
               .TextMatrix(lngR, 12) = "외"
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("수량")), 0, P_adoRec("수량"))
               .TextMatrix(lngR, 14) = "종"
               If P_adoRec("발행여부") = 1 Then
                  .Cell(flexcpChecked, lngR, 15) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 15) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 15) = "발행"
               Select Case P_adoRec("작성구분")
                      Case 0: .Cell(flexcpText, lngR, 16) = "거래"
                      Case 1: .Cell(flexcpText, lngR, 16) = "임의"
                      Case 2: .Cell(flexcpText, lngR, 16) = "일괄"
               End Select
               Select Case P_adoRec("금액구분")
                      Case 0: .Cell(flexcpText, lngR, 17) = "현금"
                      Case 1: .Cell(flexcpText, lngR, 17) = "수표"
                      Case 2: .Cell(flexcpText, lngR, 17) = "어음"
                      Case 3: .Cell(flexcpText, lngR, 17) = "외상"
                      Case Else: .Cell(flexcpText, lngR, 17) = "오류"
               End Select
               Select Case P_adoRec("영청구분")
                      Case 0: .Cell(flexcpText, lngR, 18) = "영수"
                      Case 1: .Cell(flexcpText, lngR, 18) = "청구"
               End Select
               If P_adoRec("미수구분") = 1 Then
                  .Cell(flexcpChecked, lngR, 19) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 19) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 19) = "미수"
               Select Case P_adoRec("사용구분")
                      Case 0: .Cell(flexcpText, lngR, 20) = "정상"
                      Case 9: .Cell(flexcpText, lngR, 20) = "삭제"
               End Select
               .TextMatrix(lngR, 22) = P_adoRec("작성시간")
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
               '계산서 합계금액 계산
               lblTotMny.Caption = Format(Vals(lblTotMny.Caption) + .ValueMatrix(lngR, 8), "#,#.00")
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt1 Then
                  '.TopRow = .Rows - PC_intRowCnt1 + .FixedRows
                  .TopRow = 1
               End If
            Else
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt1 Then
                  .TopRow = .Row
               End If
            End If
            'vsfg1_EnterCell       'vsfg1_EnterCell 자동실행(만약 한건 일때도 강제로 자동실행)
            '.SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계산서 읽기 실패"
    Unload Me
    Exit Sub
End Sub

