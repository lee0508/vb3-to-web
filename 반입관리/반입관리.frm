VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm반입관리 
   BorderStyle     =   0  '없음
   Caption         =   "반입관리"
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
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "업체별"
         Height          =   255
         Left            =   6840
         TabIndex        =   41
         Top             =   390
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "일자별"
         Height          =   255
         Left            =   6840
         TabIndex        =   40
         Top             =   150
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "반입관리.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   30
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "반입관리.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   28
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "반입관리.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   27
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "반입관리.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "반입관리.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "반입관리.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   0
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
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "거래명세서 반품처리"
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
      Height          =   7575
      Left            =   60
      TabIndex        =   12
      Top             =   2460
      Width           =   15195
      _cx             =   26802
      _cy             =   13361
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
      Height          =   1755
      Left            =   60
      TabIndex        =   13
      Top             =   630
      Width           =   15195
      Begin VB.CheckBox chkCash 
         Enabled         =   0   'False
         Height          =   255
         Left            =   13680
         TabIndex        =   42
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   1  '입력 상태 설정
         Index           =   4
         Left            =   6510
         TabIndex        =   5
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   1  '입력 상태 설정
         Index           =   6
         Left            =   6510
         MaxLength       =   14
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpAppDate 
         Height          =   270
         Left            =   6510
         TabIndex        =   6
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   9
         Left            =   9030
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1320
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   8
         Left            =   11430
         MaxLength       =   14
         TabIndex        =   10
         Top             =   945
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   7
         Left            =   9030
         MaxLength       =   14
         TabIndex        =   9
         Top             =   945
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   5
         Left            =   6510
         MaxLength       =   14
         TabIndex        =   7
         Top             =   945
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   3
         Left            =   1275
         TabIndex        =   4
         Top             =   1305
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   2
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   3
         Top             =   945
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   1
         Left            =   1275
         TabIndex        =   2
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
         MaxLength       =   8
         TabIndex        =   1
         Top             =   225
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFDate 
         Height          =   270
         Left            =   10800
         TabIndex        =   34
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpTDate 
         Height          =   270
         Left            =   12840
         TabIndex        =   35
         Top             =   240
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
         Caption         =   "현금반품"
         Height          =   240
         Index           =   16
         Left            =   13920
         TabIndex        =   43
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "단위"
         Height          =   240
         Index           =   14
         Left            =   5310
         TabIndex        =   39
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "규격"
         Height          =   240
         Index           =   13
         Left            =   5310
         TabIndex        =   38
         Top             =   285
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   8520
         X2              =   14660
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   12
         Left            =   14160
         TabIndex        =   37
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   11
         Left            =   12120
         TabIndex        =   36
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "반입일자"
         Height          =   240
         Index           =   10
         Left            =   9720
         TabIndex        =   33
         Top             =   285
         Width           =   855
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
         Left            =   8400
         TabIndex        =   32
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   9
         Left            =   3840
         TabIndex        =   31
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   2520
         TabIndex        =   29
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "적요"
         Height          =   240
         Index           =   8
         Left            =   7830
         TabIndex        =   22
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "출고부가"
         Height          =   240
         Index           =   7
         Left            =   10230
         TabIndex        =   21
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "출고단가"
         Height          =   240
         Index           =   6
         Left            =   7830
         TabIndex        =   20
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "반입수량"
         Height          =   240
         Index           =   5
         Left            =   5310
         TabIndex        =   19
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "기존매출일자"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   5310
         TabIndex        =   18
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품명"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   17
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "품목코드"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   16
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매출처명"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   15
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매출처코드"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   14
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm반입관리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 반입관리
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   :
' 업  무  설  명 : 매출처에서 매출한 것을 -수량으로 반입
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

    frmMain.SBar.Panels(4).Text = "기존매출일자를 정확히 입력하여 주세요. "
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
       dtpAppDate.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       dtpFDate.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       dtpTDate.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFDate.Value))
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

'+---------------+
'/// 검색조건 ///
'+---------------+
Private Sub dtpFDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpTDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdFind_Click
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
    ElseIf _
       (Index = 2 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '자재검색
       PB_strCallFormName = "frm반입관리"
       PB_strMaterialsCode = Trim(Text1(Index).Text)
       PB_strMaterialsName = ""
       frm자재검색.Show vbModal
       If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '검색에서 취소(ESC)
       Else
          Text1(2).Text = PB_strMaterialsCode
          Text1(3).Text = PB_strMaterialsName
       End If
       If PB_strMaterialsCode <> "" Then
          SendKeys "{tab}"
       End If
       PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    Else
       If KeyCode = vbKeyReturn Then
          Select Case Index
                 Case Text1.UBound
                      If cmdSave.Enabled = True Then
                         cmdSave.SetFocus
                         Exit Sub
                      End If
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

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         Select Case Index
                Case 0 '매출처
                     .Text = UPPER(Trim(.Text))
                     If Len(Trim(Text1(Index).Text)) < 1 Then
                        Text1(1).Text = ""
                     End If
                Case 2 '자재검색
                     If Len(Trim(Text1(Index).Text)) = 0 Then
                        Text1(3).Text = ""
                     End If
                Case 5
                     If Vals(.Text) > 0 Then
                        .Text = Vals(.Text) * -1
                     End If
                     .Text = Format(Vals(Trim(.Text)), "#,0")
                Case Else
                     .Text = Trim(.Text)
         End Select
    End With
    Exit Sub
End Sub

Private Sub dtpAppDate_KeyDown(KeyCode As Integer, Shift As Integer)
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
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 0))
            'Select Case .MouseCol
            '       Case 1
            '            .ColSel = 2
            '            .ColSort(0) = flexSortNone
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case 2
            '            .ColSel = 2
            '            .ColSort(0) = flexSortNone
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case Else
            '            .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            'End Select
            'If .FindRow(strData, , 0) > 0 Then
            '   .Row = .FindRow(strData, , 0)
            'End If
            'If PC_intRowCnt < .Rows Then
            '   .TopRow = .Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         Text1(0).Enabled = False: Text1(2).Enabled = False
         If .Row >= .FixedRows Then
            Text1(0).Text = .TextMatrix(.Row, 4): Text1(1).Text = .TextMatrix(.Row, 5) '매출처
            Text1(2).Text = .TextMatrix(.Row, 6): Text1(3).Text = .TextMatrix(.Row, 7) '자재
            Text1(4).Text = .TextMatrix(.Row, 8)                      '규격
            dtpAppDate.Value = Format(DTOS(.TextMatrix(.Row, 3)), "0000-00-00")      '출고일자
            Text1(5).Text = Format(.ValueMatrix(.Row, 9), "#,0")      '반입수량
            Text1(6).Text = .TextMatrix(.Row, 10)                     '단위
            Text1(7).Text = Format(.ValueMatrix(.Row, 11), "#,0.00")  '출고단가
            Text1(8).Text = Format(.ValueMatrix(.Row, 12), "#,0")     '출고부가
            Text1(9).Text = .TextMatrix(.Row, 15)                     '적요
            If .Cell(flexcpChecked, .Row, 14) = flexChecked Then
               chkCash.Value = 1
            Else
               chkCash.Value = 0
            End If
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
Dim strSQL As String
    SubClearText
    vsfg1.Row = 0
    Text1(0).Enabled = True
    Text1(2).Enabled = True
    Text1(0).SetFocus
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
Dim strSQL        As String
Dim lngR          As Long
Dim lngC          As Long
Dim lngCnt        As Long
Dim lngLogCnt     As Long
Dim blnOK         As Boolean
Dim intRetVal     As Integer
Dim strServerTime As String
Dim intChkCash    As Integer
    '입력내역 검사
    blnOK = False
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       Select Case lngC
              Case 0, 2
                   If Text1(lngC).Enabled = False Then
                      Text1(0).Enabled = True: Text1(2).Enabled = True
                   End If
       End Select
       Text1(lngC).SetFocus
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    '자재입출내역 검사
    strSQL = "SELECT ISNULL(SUM(T1.출고수량),0) AS 출고수량 " _
             & "FROM 자재입출내역 T1 " _
            & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
              & "AND (T1.분류코드 + T1.세부코드) = '" & Text1(2).Text & "' " _
              & "AND T1.매출처코드 = '" & Text1(0).Text & "' " _
              & "AND ((T1.입출고구분 = 2 AND T1.입출고일자 = '" & DTOS(dtpAppDate.Value) & "') " _
               & "OR (T1.입출고구분 = 2 AND T1.원래입출고일자 = '" & DTOS(dtpAppDate.Value) & "')) "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       dtpAppDate.SetFocus
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       lngCnt = P_adoRec("출고수량")
       P_adoRec.Close
       If lngCnt = 0 Then
          MsgBox "출고내역이 없어 반입처리할 수 없습니다.", vbCritical, "자재 반입 불가"
          Screen.MousePointer = vbDefault
          Exit Sub
       Else
          If (lngCnt < 1) Or (lngCnt < (Vals(Text1(5).Text) * -1)) Then
             MsgBox "반입수량이 출고수량(" & Format(lngCnt, "#,#") & ") 보다 많아서 반입처리할 수 없습니다.", vbCritical, "자재 반입 불가"
             Text1(5).SetFocus
             Screen.MousePointer = vbDefault
             Exit Sub
          End If
       End If
    End If
    Screen.MousePointer = vbDefault
    If Text1(Text1.LBound).Enabled = True Then
       intRetVal = MsgBox("입력된 자료를 추가하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 추가")
    Else
       If vsfg1.ValueMatrix(vsfg1.Row, 19) = 1 Then
          MsgBox "세금계산서 발행분 임으로 변경할 수 없습니다.", vbCritical, "세금계산서 발행분"
          Exit Sub
       End If
       intRetVal = MsgBox("수정된 자료를 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 저장")
    End If
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    cmdSave.Enabled = False
    With vsfg1
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
         intRetVal = MsgBox("현금반입으로 처리하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton2, "현금반입")
         If intRetVal = vbYes Then
            intChkCash = 1
            chkCash.Value = 1
         Else
            intChkCash = 0
            chkCash.Value = 0
         End If
         '자재시세 검색(자재입출내역 테이블에서)
         strSQL = "SELECT TOP 1 ISNULL(T1.출고단가, 0) AS 출고단가, ISNULL(T1.출고부가, 0) AS 출고부가 " _
                  & "FROM 자재입출내역 T1 " _
                 & "WHERE (T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "') " _
                   & "AND (T1.분류코드 + T1.세부코드) = '" & Text1(2).Text & "' " _
                   & "AND (T1.매출처코드 = '" & Text1(0).Text & "') " _
                   & "AND (T1.입출고구분 = 2 AND T1.입출고일자 = '" & DTOS(dtpAppDate.Value) & "' AND 출고수량 > 0) "
         On Error GoTo ERROR_TABLE_SELECT
         P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
         If P_adoRec.RecordCount = 0 Then
            P_adoRec.Close
            dtpAppDate.SetFocus
            Screen.MousePointer = vbDefault
            cmdSave.Enabled = True
            Exit Sub
         Else
            Text1(7).Text = Format(P_adoRec("출고단가"), "#,0.00")
            Text1(8).Text = Format(P_adoRec("출고부가"), "#,0")
         End If
         P_adoRec.Close
         PB_adoCnnSQL.BeginTrans
         If Text1(Text1.LBound).Enabled = True Then '반품 추가
            '거래번호 구하기
            P_adoRec.CursorLocation = adUseClient
            strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + PB_regUserinfoU.UserClientDate + "2" & "', " _
                                   & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
            On Error GoTo ERROR_STORED_PROCEDURE
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            lngLogCnt = P_adoRec(0)
            P_adoRec.Close
            strSQL = "INSERT INTO 자재입출내역(사업장코드, 분류코드, 세부코드, " _
                                            & "입출고구분, 입출고일자, 입출고시간, " _
                                            & "입고수량, 입고단가, 입고부가, " _
                                            & "출고수량, 출고단가, 출고부가, " _
                                            & "매입처코드, 매출처코드, 원래입출고일자, 직송구분, " _
                                            & "발견일자, 발견번호, 거래일자, 거래번호, " _
                                            & "계산서발행여부, 현금구분, 감가구분, 적요, 책번호, 일련번호, " _
                                            & "사용구분, 수정일자, 사용자코드, 재고이동사업장코드) Values( " _
                    & "'" & PB_regUserinfoU.UserBranchCode & "','" & Mid(Text1(2).Text, 1, 2) & "','" & Mid(Text1(2).Text, 3) & "', " _
                    & "2, '" & PB_regUserinfoU.UserClientDate & "', '" & strServerTime & "', " _
                    & "0 ,0 ,0, " _
                    & "" & Vals(Text1(5).Text) & ", " & Vals(Text1(7).Text) & ", " & Vals(Text1(8).Text) & ", " _
                    & "'', '" & Text1(0).Text & "', '" & DTOS(dtpAppDate.Value) & "', 0, '', 0, " _
                    & "'" & PB_regUserinfoU.UserClientDate & "', " & lngLogCnt & ", " _
                    & "0, " & intChkCash & ", 0, '" & Text1(9).Text & "', 0, 0, " _
                    & "0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "', '' )"
            On Error GoTo ERROR_TABLE_INSERT
            PB_adoCnnSQL.Execute strSQL
            .AddItem .Rows
            .TextMatrix(.Rows - 1, 0) = Text1(2).Text & "-" & Text1(0) & "-" & "" & "-" & "2" & "-" _
                                      & PB_regUserinfoU.UserClientDate & "-" & strServerTime
            .Cell(flexcpData, .Rows - 1, 0, .Rows - 1, 0) = Trim(.TextMatrix(.Rows - 1, 0))
            .TextMatrix(.Rows - 1, 1) = Format(PB_regUserinfoU.UserClientDate, "0000-00-00") '입출고일자(반입일자)
            .TextMatrix(.Rows - 1, 2) = Format(Mid(strServerTime, 1, 6), "00:00:00")         '입출고시간
            .TextMatrix(.Rows - 1, 3) = Format(DTOS(dtpAppDate.Value), "0000-00-00")         '출고일자
            .TextMatrix(.Rows - 1, 4) = Text1(0).Text: .TextMatrix(.Rows - 1, 5) = Text1(1).Text '매출처
            .TextMatrix(.Rows - 1, 6) = Text1(2).Text: .TextMatrix(.Rows - 1, 7) = Text1(3).Text '자재
            .TextMatrix(.Rows - 1, 8) = Text1(4).Text                                        '규격
            .TextMatrix(.Rows - 1, 9) = Vals(Text1(5).Text)                                  '반입수량
            .TextMatrix(.Rows - 1, 10) = Text1(6).Text                                       '단위
            .TextMatrix(.Rows - 1, 11) = Vals(Text1(7).Text)                                 '출고단가
            .TextMatrix(.Rows - 1, 12) = Vals(Text1(8).Text)                                 '출고부가
            .TextMatrix(.Rows - 1, 13) = .ValueMatrix(.Rows - 1, 9) * .ValueMatrix(.Rows - 1, 11)  '출고금액
            .TextMatrix(.Rows - 1, 14) = intChkCash                                          '현금구분
            If intChkCash = 1 Then
               .Cell(flexcpChecked, .Rows - 1, 14) = flexChecked   '1
            Else
               .Cell(flexcpChecked, .Rows - 1, 14) = flexUnchecked '2
            End If
            .Cell(flexcpText, .Rows - 1, 14) = "현금반입"
            .TextMatrix(.Rows - 1, 15) = Text1(9).Text                                       '적요
            .TextMatrix(.Rows - 1, 16) = PB_regUserinfoU.UserCode                            '사용자코드
            .TextMatrix(.Rows - 1, 17) = PB_regUserinfoU.UserName                            '사용자명
            .TextMatrix(.Rows - 1, 18) = strServerTime                                       '입출고시간
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(Text1.LBound).Enabled = False
            .Row = .Rows - 1          '자동으로 vsfg1_EnterCell Event 발생
         Else                                          '자재입출내역 변경
            strSQL = "UPDATE 자재입출내역 SET " _
                          & "원래입출고일자 = '" & DTOS(dtpAppDate.Value) & "', " _
                          & "출고수량 = " & Vals(Text1(5).Text) & ", " _
                          & "출고단가 = " & Vals(Text1(7).Text) & ", " _
                          & "출고부가 = " & Vals(Text1(8).Text) & ", " _
                          & "현금구분 = " & intChkCash & ", " _
                          & "적요 = '" & Trim(Text1(9).Text) & "', " _
                          & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                      & "AND 분류코드 = '" & Mid(.TextMatrix(.Row, 6), 1, 2) & "' " _
                      & "AND 세부코드 = '" & Mid(.TextMatrix(.Row, 6), 3) & "' " _
                      & "AND 매입처코드 = '' " _
                      & "AND 매출처코드 = '" & .TextMatrix(.Row, 4) & "' " _
                      & "AND 입출고구분 = 2 " _
                      & "AND 입출고일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                      & "AND 입출고시간 = '" & .TextMatrix(.Row, 18) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            PB_adoCnnSQL.Execute strSQL
            .TextMatrix(.Row, 3) = Format(DTOS(dtpAppDate.Value), "0000-00-00")     '원래입출고일자
            .TextMatrix(.Row, 9) = Vals(Text1(5).Text)                              '반입수량
            .TextMatrix(.Row, 11) = Vals(Text1(7).Text)                             '출고단가
            .TextMatrix(.Row, 12) = Vals(Text1(8).Text)                             '출고부가
            .TextMatrix(.Row, 13) = .ValueMatrix(.Row, 9) * (.ValueMatrix(.Row, 11))  '출고금액
            .TextMatrix(.Row, 14) = intChkCash
            If intChkCash = 1 Then                                                  '현금구분
               .Cell(flexcpChecked, .Row, 14) = flexChecked    '1
            Else
               .Cell(flexcpChecked, .Row, 14) = flexUnchecked  '2
            End If
            .Cell(flexcpText, .Row, 14) = "현금반입"
            .TextMatrix(.Row, 15) = Text1(9).Text                                   '적요
            .TextMatrix(.Row, 16) = PB_regUserinfoU.UserCode                        '사용자코드
            .TextMatrix(.Row, 17) = PB_regUserinfoU.UserName                        '사용자명
            'if x then '세금계산서발행분이면
            'end if
         End If
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "반입내역 읽기 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "반입내역 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "반입내역 갱신 실패"
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

Private Sub cmdDelete_Click()
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         If .Row >= .FixedRows Then
            intRetVal = MsgBox("등록된 자료를 삭제하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton2, "자료 삭제")
            If vsfg1.ValueMatrix(vsfg1.Row, 19) = 1 Then
               MsgBox "세금계산서 발행분 임으로 삭제할 수 없습니다.", vbCritical, "세금계산서 발행분"
               Exit Sub
            End If
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
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "건)가 있으므로 삭제할 수 없습니다.", vbCritical, "반입내역 삭제 불가"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "UPDATE 자재입출내역 SET " _
                             & "사용구분 = 9, " _
                             & "적요 = '" & Trim(Text1(9).Text) & "', " _
                             & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                         & "AND 분류코드 = '" & Mid(.TextMatrix(.Row, 6), 1, 2) & "' " _
                         & "AND 세부코드 = '" & Mid(.TextMatrix(.Row, 6), 3) & "' " _
                         & "AND 매입처코드 = '" & .TextMatrix(.Row, 4) & "' " _
                         & "AND 매출처코드 = '' " _
                         & "AND 입출고구분 = 2 " _
                         & "AND 입출고일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                         & "AND 입출고시간 = '" & .TextMatrix(.Row, 18) & "' "
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
               vsfg1_EnterCell
               cmdDelete.Enabled = True
            End If
            .SetFocus
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "반입내역 삭제 실패"
    Unload Me
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
    Set frm반입관리 = Nothing
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
    Text1(Text1.LBound).Enabled = False                '매출처코드 FLASE
    With vsfg1              'Rows 1, Cols 22, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 3
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 22
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   'H, KEY(자재코드-매입처코드-매출처코드-입출고구분-입출고일자-입출고시간)
         .ColWidth(1) = 1200   '반입일자
         .ColWidth(2) = 1200   '반입시간(시초분)
         .ColWidth(3) = 1200   '출고일자(원래출고일자)
         .ColWidth(4) = 1000   'H, 매출처코드
         .ColWidth(5) = 2000   '매출처명
         .ColWidth(6) = 1900   '자재코드
         .ColWidth(7) = 2500   '자재명
         .ColWidth(8) = 2000   '자재규격
         .ColWidth(9) = 800    '반입수량
         .ColWidth(10) = 600   '자재단위
         .ColWidth(11) = 1500  '출고단가
         .ColWidth(12) = 1300  'H, 출고부가
         .ColWidth(13) = 1600  '출고금액
         .ColWidth(14) = 1200  '매출구분
         .ColWidth(15) = 5000  '적요
         .ColWidth(16) = 1000  '사용자코드
         .ColWidth(17) = 1000  '사용자명
         .ColWidth(18) = 1000  '입출고시간
         
         .ColWidth(19) = 1000  '계산서발행여부
         .ColWidth(20) = 1000  '세금계산서(책번호)
         .ColWidth(21) = 1000  '세금계산서(일련번호)
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "KEY"         'H
         .TextMatrix(0, 1) = "반입일자"
         .TextMatrix(0, 2) = "반입시간"
         .TextMatrix(0, 3) = "매출일자"
         .TextMatrix(0, 4) = "매출처코드"  'H
         .TextMatrix(0, 5) = "매출처명"
         .TextMatrix(0, 6) = "품목코드"
         .TextMatrix(0, 7) = "품명"
         .TextMatrix(0, 8) = "규격"
         .TextMatrix(0, 9) = "수량"
         .TextMatrix(0, 10) = "단위"
         .TextMatrix(0, 11) = "매출단가"
         .TextMatrix(0, 12) = "매출부가"   'H
         .TextMatrix(0, 13) = "매출금액"
         .TextMatrix(0, 14) = "매출구분"
         .TextMatrix(0, 15) = "적요"
         .TextMatrix(0, 16) = "사용자코드" 'H
         .TextMatrix(0, 17) = "사용자명"
         .TextMatrix(0, 18) = "입출고시간" 'H
         .TextMatrix(0, 19) = "계산서발행" 'H
         .TextMatrix(0, 20) = "책번호"     'H
         .TextMatrix(0, 21) = "일련번호"   'H
         .ColHidden(0) = True: .ColHidden(4) = True:  .ColHidden(12) = True
         .ColHidden(16) = True: .ColHidden(18) = True: .ColHidden(19) = True: .ColHidden(20) = True: .ColHidden(21) = True
         .ColFormat(9) = ","
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 5, 6, 7, 8, 10, 14, 15
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 1, 2, 3, 4, 16, 17, 18, 19
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 11, 13
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
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
    Text1(0).Enabled = False: Text1(2).Enabled = False
    If dtpFDate > dtpTDate Then
       dtpFDate.SetFocus
       Exit Sub
    End If
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT (T1.분류코드 + T1.세부코드) AS 자재코드, ISNULL(T2.자재명,'') AS 자재명, " _
                  & "ISNULL(T1.매입처코드,'') AS 매입처코드, ISNULL(T3.매입처명,'') AS 매입처명, " _
                  & "ISNULL(T1.매출처코드,'') AS 매출처코드, ISNULL(T4.매출처명,'') AS 매출처명, " _
                  & "T1.입출고구분,  T1.입출고일자, T1.입출고시간, " _
                  & "T2.규격 AS 자재규격,  T2.단위 AS 자재단위, " _
                  & "T1.입고수량 AS 입고수량, T1.입고단가, 입고부가, " _
                  & "T1.출고수량 AS 출고수량, T1.출고단가, 출고부가, " _
                  & "T1.원래입출고일자, T1.현금구분, T1.적요, T1.사용자코드, ISNULL(T5.사용자명, '') AS 사용자명, " _
                  & "T1.계산서발행여부, T1.책번호, T1.일련번호 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 자재 T2 " _
                    & "ON T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드 " _
             & "LEFT JOIN 매입처 T3 " _
                    & "ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드= T1.매입처코드 " _
             & "LEFT JOIN 매출처 T4 " _
                    & "ON T4.사업장코드 = T1.사업장코드 AND T4.매출처코드= T1.매출처코드 " _
             & "LEFT JOIN 사용자 T5 " _
                    & "ON T5.사업장코드 = T1.사업장코드 AND T5.사용자코드= T1.사용자코드 " _
            & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.사용구분 = 0 " _
              & "AND T1.입출고구분 = 2 AND 출고수량 < 0 " _
              & "AND T1.입출고일자 BETWEEN '" & DTOS(dtpFDate.Value) & "' AND '" & DTOS(dtpTDate.Value) & "' " _
            & "ORDER BY T1.입출고일자, T1.입출고시간, T1.원래입출고일자 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
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
               .TextMatrix(lngR, 0) = P_adoRec("자재코드") & "-" & P_adoRec("매입처코드") & "-" _
                                    & P_adoRec("매출처코드") & "-" & P_adoRec("입출고구분") & "-" _
                                    & P_adoRec("입출고일자") & "-" & P_adoRec("입출고시간") & "-"
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow 사용을 위해
               .TextMatrix(lngR, 1) = Format(P_adoRec("입출고일자"), "0000-00-00")
               .TextMatrix(lngR, 2) = Format(Mid(P_adoRec("입출고시간"), 1, 6), "00:00:00")
               .TextMatrix(lngR, 3) = Format(P_adoRec("원래입출고일자"), "0000-00-00")
               .TextMatrix(lngR, 4) = P_adoRec("매출처코드")
               .TextMatrix(lngR, 5) = P_adoRec("매출처명")
               .TextMatrix(lngR, 6) = P_adoRec("자재코드")
               .TextMatrix(lngR, 7) = P_adoRec("자재명")
               .TextMatrix(lngR, 8) = P_adoRec("자재규격")
               .TextMatrix(lngR, 9) = P_adoRec("출고수량")
               .TextMatrix(lngR, 10) = P_adoRec("자재단위")
               .TextMatrix(lngR, 11) = P_adoRec("출고단가")
               .TextMatrix(lngR, 12) = P_adoRec("출고부가")
               .TextMatrix(lngR, 13) = .ValueMatrix(lngR, 9) * (.ValueMatrix(lngR, 11))
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("현금구분")), 0, P_adoRec("현금구분"))
               If P_adoRec("현금구분") = 1 Then
                  .Cell(flexcpChecked, lngR, 14) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 14) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 14) = "현금반입"
               .TextMatrix(lngR, 15) = P_adoRec("적요")
               .TextMatrix(lngR, 16) = P_adoRec("사용자코드")
               .TextMatrix(lngR, 17) = P_adoRec("사용자명")
               .TextMatrix(lngR, 18) = P_adoRec("입출고시간")
               .TextMatrix(lngR, 19) = P_adoRec("계산서발행여부")
               .TextMatrix(lngR, 20) = P_adoRec("책번호")
               .TextMatrix(lngR, 21) = P_adoRec("일련번호")
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
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
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "반품내역 읽기 실패"
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
               Case 0  '매출처코드
                    If Not (Len(Text1(lngC).Text) > 0) Then
                       Exit Function
                    End If
               Case 1  '매출처명
                    If Len(Trim(Text1(lngC).Text)) = 0 Then
                       lngC = 0
                       Exit Function
                    End If
               Case 2  '자재코드
                    If Not (Len(Text1(lngC).Text) > 0) Then
                       Exit Function
                    End If
               Case 3  '자재명
                    If Len(Trim(Text1(lngC).Text)) = 0 Then
                       lngC = 2
                       Exit Function
                    End If
               Case 5  '반입수량
                    If Not (Vals(Text1(lngC).Text) < 0) Then
                       Exit Function
                    End If
               Case 9  '적요
                    If Not (LenH(Text1(lngC).Text) <= 50) Then
                       Exit Function
                    End If
               Case Else
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
Dim strForPrtDateTime      As String  '출력일시           (Formula)

    If dtpFDate.Value > dtpTDate.Value Then
       dtpFDate.SetFocus
       Exit Sub
    End If
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
            strExeFile = App.Path & IIf(optPrtChk0.Value = True, ".\반입내역(일자별).rpt", ".\반입내역(업체별).rpt")
         Else
            strExeFile = App.Path & IIf(optPrtChk0.Value = True, ".\반입내역(일자별)T.rpt", ".\반입내역(업체별)T.rpt")
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForPgDate = '" & Format(PB_regUserinfoU.UserClientDate, "0000-00-00") & "'" '프로그램실행일자
            .Formulas(1) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                    '사업장명
            .Formulas(2) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                               '출력일시
            .Formulas(3) = "ForAppDate = '기준일자 : ' & '" & Format(DTOS(dtpFDate.Value), "0000-00-00") & "' & ' 부터 ' & '" & Format(DTOS(dtpTDate.Value), "0000-00-00") & "' & ' 까지' "
            '--- Parameter Fields ---
            .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode  '지점코드
            .StoredProcParam(1) = DTOS(dtpFDate.Value)            '기준일자(시작일자)
            .StoredProcParam(2) = DTOS(dtpTDate.Value)            '기준일자(종료일자)
            .StoredProcParam(3) = " "                             '자재명
            .StoredProcParam(4) = " "                             '업체명
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & IIf(optPrtChk0.Value = True, "반입내역(일자별).rpt", "반입내역(업체별)")
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


