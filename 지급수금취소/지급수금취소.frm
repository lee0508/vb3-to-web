VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm지급수금취소 
   BorderStyle     =   0  '없음
   Caption         =   "지급수금취소"
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
      TabIndex        =   10
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "지급수금취소.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   9120
         Picture         =   "지급수금취소.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "지급수금취소.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "지급수금취소.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "지급수금취소.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "지급수금취소.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   7
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
         Caption         =   "지급/수금취소"
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
         TabIndex        =   11
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   8476
      Left            =   60
      TabIndex        =   8
      Top             =   1620
      Width           =   15195
      _cx             =   26802
      _cy             =   14951
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
      Height          =   960
      Left            =   60
      TabIndex        =   9
      Top             =   630
      Width           =   15195
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   6000
         TabIndex        =   4
         Top             =   570
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   6000
         MaxLength       =   8
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "전체 업체"
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Value           =   1  '확인
         Width           =   1215
      End
      Begin VB.OptionButton optJSGbn 
         Caption         =   "미 수 금 입금"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   550
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optJSGbn 
         Caption         =   "미지급금 지급"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   270
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpFDate 
         Height          =   270
         Left            =   10800
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57540609
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label lblTotMny 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   10660
         TabIndex        =   29
         Top             =   630
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "전체금액"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   9720
         TabIndex        =   28
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "업체명"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   4920
         TabIndex        =   27
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   7200
         TabIndex        =   26
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   3
         Left            =   9480
         TabIndex        =   25
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   24
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "업체코드"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   4920
         TabIndex        =   23
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   ")"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   22
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "("
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   21
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   12
         Left            =   14160
         TabIndex        =   20
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   11
         Left            =   12120
         TabIndex        =   19
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "거래일자"
         Height          =   240
         Index           =   10
         Left            =   9720
         TabIndex        =   18
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
         Left            =   120
         TabIndex        =   17
         Top             =   405
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm지급수금취소"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 미지급금지급 취소, 미수금입금 취소
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   :
' 업  무  설  명 : 미지급금내역 또는 미수금내역만 삭제(UPDATE)
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private P_intBeforeOptGbn  As Integer
Private Const PC_intRowCnt As Integer = 28  '그리드 한 페이지 당 행수(FixedRows 포함)

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

    frmMain.SBar.Panels(4).Text = "단순히 미지급금내역(지급)/미수금내역(수금)만 취소(삭제) 또는 수정합니다. "
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       Subvsfg1_INIT
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '인쇄, 조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '추가, 저장, 인쇄, 조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       P_intBeforeOptGbn = 1
       optJSGbn(0).ForeColor = vbBlue: optJSGbn(1).ForeColor = vbRed
       Label1(0).ForeColor = vbRed: Label1(1).ForeColor = vbRed
       dtpFDate.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
       dtpTDate.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFDate.Value))
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "거래내역취소(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// 검색조건 ///
'+---------------+
Private Sub optJSGbn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With optJSGbn
         If KeyCode = vbKeyReturn Then
            chkTotal.SetFocus
         End If
    End With
End Sub
Private Sub optJSGbn_Click(Index As Integer)
    With optJSGbn
         If optJSGbn(0).Value = True Then
            Label1(0).ForeColor = vbBlue: Label1(1).ForeColor = vbBlue
         Else
            Label1(0).ForeColor = vbRed: Label1(1).ForeColor = vbRed
         End If
         If Index <> P_intBeforeOptGbn Then
            P_intBeforeOptGbn = Index
            Text1(0).Text = "": Text1(1).Text = "": lblTotMny.Caption = "0"
            vsfg1.Rows = 1
         End If
    End With
End Sub

Private Sub chkTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    With chkTotal
         If KeyCode = vbKeyReturn Then
            If chkTotal.Value = 1 Then
               dtpFDate.SetFocus
            Else
               Text1(0).SetFocus
            End If
         End If
    End With
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
       PB_strSupplierName = ""  'Trim(Text1(Index + 1).Text)
       If optJSGbn(0).Value = True Then
          frm매입처검색.Show vbModal
       Else
          frm매출처검색.Show vbModal
       End If
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
                 Case Text1.UBound
                      'If chkTotal.Value = 0 And cmdSave.Enabled = True And vsfg1.Rows > 1 Then
                      '   cmdSave.SetFocus
                      '   Exit Sub
                      'End If
           End Select
           SendKeys "{tab}"
       End If
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매입/매출처 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         .Text = Trim(.Text)
         Select Case Index
                Case 0 '업체
                     .Text = UPPER(Trim(.Text))
                     If Len(Trim(.Text)) < 1 Then
                        Text1(1).Text = ""
                     End If
                Case Else
                     .Text = Trim(.Text)
         End Select
    End With
    Exit Sub
End Sub

Private Sub dtpFDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpTDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdFind_Click
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
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            .Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            strData = Trim(.Cell(flexcpData, .Row, 0))
            Select Case .MouseCol
                   Case 1
                        .ColSel = 2
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case 2
                        .ColSel = 2
                        .ColSort(0) = flexSortNone
                        .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                        .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 0) > 0 Then
               .Row = .FindRow(strData, , 0)
            End If
            If PC_intRowCnt < .Rows Then
               .TopRow = .Row
            End If
         End If
    End With
End Sub
Private Sub vsfg1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         P_intButton = Button
         If .MouseRow >= .FixedRows Then
            If (.MouseCol = 8) Then   '거래금액
               If Button = vbLeftButton Then
                  .Select .MouseRow, IIf(.MouseCol = .Col, .MouseCol, .Col)
                  .EditCell
               End If
            ElseIf _
               (.MouseCol = 9 And .TextMatrix(.Row, 7) = "어음") Then    '만기일자, col(7).결제
               If Button = vbLeftButton Then
                  .Select .MouseRow, IIf(.MouseCol = .Col, .MouseCol, .Col)
                  .EditCell
               End If
            ElseIf _
               (.MouseCol = 10 And .TextMatrix(.Row, 7) = "어음") Then   '어음번호, col(7).결제
               If Button = vbLeftButton Then
                  .Select .MouseRow, IIf(.MouseCol = .Col, .MouseCol, .Col)
                  .EditCell
               End If
            ElseIf _
               (.MouseCol = 11) Then   '적요
               If Button = vbLeftButton Then
                  .Select .MouseRow, IIf(.MouseCol = .Col, .MouseCol, .Col)
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
            If (Col = 8) Then         '거래금액
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 8)
                     .TextMatrix(Row, 8) = Vals(.EditText)   '거래금액
                  End If
               End If
            ElseIf _
               (Col = 9) Then '만기일자
               If .TextMatrix(Row, 9) <> .EditText Then
                  .EditText = Format(Replace(.EditText, "-", ""), "0000-00-00")
                  If Not ((Len(Trim(.EditText)) = 10) And IsDate(.EditText) And Val(.EditText) > 2000) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            ElseIf _
               (Col = 10) Then '어음번호
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 20) Then
                     Beep
                     .TextMatrix(Row, Col) = .EditText
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            ElseIf _
               (Col = 11) Then '적요 길이 검사
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 50) Then
                     Beep
                     .TextMatrix(Row, Col) = .EditText
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            End If
            '변경표시 + 금액재계산
            If blnModify = True Then
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
               Select Case Col
                      Case 8
                           lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - curTmpMny + .ValueMatrix(Row, 8), "#,0")
                      Case Else
                      
               End Select
            End If
         End If
    End With
End Sub

Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
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
End Sub
'+-----------+
'/// 조회 ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         If .Row >= .FixedRows Then
            If .Cell(flexcpBackColor, .Row, 8, .Row, 8) = vbRed Or .Cell(flexcpBackColor, .Row, 9, .Row, 9) = vbRed Or _
               .Cell(flexcpBackColor, .Row, 10, .Row, 10) = vbRed Or .Cell(flexcpBackColor, .Row, 11, .Row, 11) = vbRed Then
               intRetVal = MsgBox("변경된 자료를 저장하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton1, "자료 저장")
            Else
               .SetFocus
               Exit Sub
            End If
            If intRetVal = vbYes Then
               Screen.MousePointer = vbHourglass
               cmdSave.Enabled = False
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               If (.ValueMatrix(.Row, 3) Mod 2) = 1 Then '1.지급
                  strSQL = "UPDATE 미지급금내역 SET " _
                                & "미지급금지급금액 = " & .ValueMatrix(.Row, 8) & ", " _
                                & "만기일자 = '" & DTOS(.TextMatrix(.Row, 9)) & "', " _
                                & "어음번호 = '" & .TextMatrix(.Row, 10) & "', " _
                                & "적요 = '" & .TextMatrix(.Row, 10) & "', " _
                                & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                                & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                          & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                            & "AND 매입처코드 = '" & .TextMatrix(.Row, 5) & "' " _
                            & "AND 미지급금지급일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                            & "AND 미지급금지급시간 = '" & .TextMatrix(.Row, 14) & "' "
               Else                                      '2.수금
                  strSQL = "UPDATE 미수금내역 SET " _
                                & "미수금입금금액 = " & .ValueMatrix(.Row, 8) & ", " _
                                & "만기일자 = '" & DTOS(.TextMatrix(.Row, 9)) & "', " _
                                & "어음번호 = '" & .TextMatrix(.Row, 10) & "', " _
                                & "적요 = '" & .TextMatrix(.Row, 10) & "', " _
                                & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                                & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                          & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                            & "AND 매출처코드 = '" & .TextMatrix(.Row, 5) & "' " _
                            & "AND 미수금입금일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                            & "AND 미수금입금시간 = '" & .TextMatrix(.Row, 14) & "' "
               End If
               On Error GoTo ERROR_TABLE_UPDATE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .TextMatrix(.Row, 12) = PB_regUserinfoU.UserCode
               .TextMatrix(.Row, 13) = PB_regUserinfoU.UserName
               'lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 8), "#,0")
               .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = vbWhite
               .Cell(flexcpForeColor, .Row, .FixedCols, .Row, .Cols - 1) = vbBlack
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarVertical
               End If
               Screen.MousePointer = vbDefault
               vsfg1_EnterCell
               cmdSave.Enabled = True
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "읽기 실패"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "변경 실패"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "삭제 실패"
    Unload Me
    Screen.MousePointer = vbDefault
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
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "건)가 있으므로 삭제할 수 없습니다.", vbCritical, "미지급금/수금내역 삭제 불가"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               If (.ValueMatrix(.Row, 3) Mod 2) = 1 Then '1.지급
                  strSQL = "DELETE FROM 미지급금내역 " _
                          & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                            & "AND 매입처코드 = '" & .TextMatrix(.Row, 5) & "' " _
                            & "AND 미지급금지급일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                            & "AND 미지급금지급시간 = '" & .TextMatrix(.Row, 14) & "' "
               Else                                      '2.수금
                  strSQL = "DELETE FROM 미수금내역 " _
                          & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                            & "AND 매출처코드 = '" & .TextMatrix(.Row, 5) & "' " _
                            & "AND 미수금입금일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                            & "AND 미수금입금시간 = '" & .TextMatrix(.Row, 14) & "' "
               End If
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 8), "#,0")
               .RemoveItem .Row
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarHorizontal
               End If
               Screen.MousePointer = vbDefault
               If .Rows = 1 Then
                  .Row = 0
                  cmdFind.SetFocus
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
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "읽기 실패"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "변경 실패"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "삭제 실패"
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
    Set frm지급수금취소 = Nothing
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
    With vsfg1              'Rows 1, Cols 15, RowHeightMax(Min) 300
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
         .Cols = 15
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   'KEY(자재코드-매입처코드-매출처코드-입출고구분-입출고일자-입출고시간)
         .ColWidth(1) = 1000   '거래일자
         .ColWidth(2) = 900    '거래시간
         .ColWidth(3) = 1000   '거래구분
         .ColWidth(4) = 600    '거래구분명
         .ColWidth(5) = 1200   '업체코드
         .ColWidth(6) = 2000   '업체명
         .ColWidth(7) = 600    '결제방법
         .ColWidth(8) = 1300   '거래금액
         .ColWidth(9) = 1000   '만기일자
         .ColWidth(10) = 2000  '어음번호
         .ColWidth(11) = 4500  '적요
         .ColWidth(12) = 1000  '사용자코드
         .ColWidth(13) = 900   '사용자명
         .ColWidth(14) = 1000  '시간(밀리초)
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "KEY"         'H
         .TextMatrix(0, 1) = "거래일자"
         .TextMatrix(0, 2) = "거래시간"
         .TextMatrix(0, 3) = "구분"        'H
         .TextMatrix(0, 4) = "구분"        '1.지급, 2.수금
         .TextMatrix(0, 5) = "업체코드"  'H
         .TextMatrix(0, 6) = "업체명"
         .TextMatrix(0, 7) = "결제"
         .TextMatrix(0, 8) = "거래금액"
         .TextMatrix(0, 9) = "만기일자"
         .TextMatrix(0, 10) = "어음번호"
         .TextMatrix(0, 11) = "적요"
         .TextMatrix(0, 12) = "사용자코드" 'H
         .TextMatrix(0, 13) = "사용자명"
         .TextMatrix(0, 14) = "시간"
         .ColHidden(0) = True: .ColHidden(3) = True: .ColHidden(5) = True
         .ColHidden(12) = True: .ColHidden(14) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 6, 10, 11
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 1, 2, 3, 4, 5, 7, 9, 12, 13, 14
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 8
                         .ColFormat(lngC) = "#,#"
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
    If dtpFDate > dtpTDate Then
       dtpFDate.SetFocus
       Exit Sub
    End If
    lblTotMny.Caption = "0" '전체금액
    vsfg1.Rows = 1
    With vsfg1
         '검색조건 업체
         strWhere = "T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
         If chkTotal.Value = 0 Then '건별 조회
            If Len(Text1(0).Text) > 0 Then
               strWhere = strWhere & "AND " & IIf(optJSGbn(0).Value = True, "T1.매입처코드", "T1.매출처코드") & " = '" & Trim(Text1(0).Text) & "' "
            Else
               Text1(0).SetFocus
               Exit Sub
            End If
         End If
         If optJSGbn(0).Value = True Then
            strWhere = strWhere & "AND T1.미지급금지급일자 BETWEEN '" & DTOS(dtpFDate.Value) & "' AND '" & DTOS(dtpTDate.Value) & "' "
         Else
            strWhere = strWhere & "AND T1.미수금입금일자 BETWEEN '" & DTOS(dtpFDate.Value) & "' AND '" & DTOS(dtpTDate.Value) & "' "
         End If
    End With
    P_adoRec.CursorLocation = adUseClient
    If optJSGbn(0).Value = True Then
       strSQL = "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                     & "T1.매입처코드 AS 업체코드, T3.매입처명 AS 업체명, " _
                     & "T1.미지급금지급일자 AS 거래일자, T1.미지급금지급시간 AS 거래시간, 1 AS 거래구분, " _
                     & "결제방법 = CASE WHEN T1.결제방법 = 0 THEN '현금' " _
                                     & "WHEN T1.결제방법 = 1 THEN '수표' " _
                                     & "WHEN T1.결제방법 = 2 THEN '어음' " _
                                     & "ELSE '오  류' " _
                                 & "END, " _
                     & "ISNULL(T1.미지급금지급금액,0) AS 거래금액, " _
                     & "만기일자 = CASE WHEN T1.결제방법 = 0 THEN '' " _
                                     & "WHEN T1.결제방법 = 2 THEN T1.만기일자 " _
                                & "END, " _
                     & "어음번호 = CASE WHEN T1.결제방법 = 0 THEN '' " _
                                     & "WHEN T1.결제방법 = 2 THEN T1.어음번호 " _
                                & "END, " _
                     & "T1.적요 AS 적요, T1.사용자코드, T4.사용자명 " _
                & "FROM 미지급금내역 T1 " _
                & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
                & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.매입처코드 " _
                & "LEFT JOIN 사용자 T4 " _
                       & "ON T4.사업장코드 = T1.사업장코드 AND T4.사용자코드= T1.사용자코드 " _
               & "WHERE " & strWhere & " " _
               & "ORDER BY T1.사업장코드, T1.거래일자, T1.거래시간, T1.거래구분, T1.결제방법 "
    Else
       strSQL = "SELECT T1.사업장코드 AS 사업장코드, T2.사업장명 AS 사업장명, " _
                     & "T1.매출처코드 AS 업체코드, T3.매출처명 AS 업체명, " _
                     & "T1.미수금입금일자 AS 거래일자, T1.미수금입금시간 AS 거래시간, 2 AS 거래구분, " _
                     & "결제방법 = CASE WHEN T1.결제방법 = 0 THEN '현금' " _
                                     & "WHEN T1.결제방법 = 1 THEN '수표' " _
                                     & "WHEN T1.결제방법 = 2 THEN '어음' " _
                                     & "Else '오  류' " _
                                & "END, " _
                     & "ISNULL(T1.미수금입금금액,0) AS 거래금액, " _
                     & "만기일자 = CASE WHEN T1.결제방법 = 0 THEN '' " _
                                     & "WHEN T1.결제방법 = 2 THEN T1.만기일자 " _
                                & "END, " _
                     & "어음번호 = CASE WHEN T1.결제방법 = 0 THEN '' " _
                                     & "WHEN T1.결제방법 = 2 THEN T1.어음번호 " _
                                & "END, " _
                     & "T1.적요 AS 적요, T1.사용자코드, T4.사용자명 " _
                & "FROM 미수금내역 T1 " _
                & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
                & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
                & "LEFT JOIN 사용자 T4 " _
                       & "ON T4.사업장코드 = T1.사업장코드 AND T4.사용자코드= T1.사용자코드 " _
               & "WHERE " & strWhere & " " _
               & "ORDER BY T1.사업장코드, T1.거래일자, T1.거래시간, T1.거래구분, T1.결제방법 "
    End If
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
               .TextMatrix(lngR, 0) = P_adoRec("사업장코드") & "-" & P_adoRec("업체코드") & "-" _
                                    & P_adoRec("거래구분") & "-" & P_adoRec("결제방법") & "-" _
                                    & P_adoRec("거래일자") & "-" & P_adoRec("거래시간") & "-"
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow 사용을 위해
               .TextMatrix(lngR, 1) = Format(P_adoRec("거래일자"), "0000-00-00")
               .TextMatrix(lngR, 2) = Format(Mid(P_adoRec("거래시간"), 1, 6), "00:00:00")
               .TextMatrix(lngR, 3) = P_adoRec("거래구분")
               If P_adoRec("거래구분") = 1 Then
                  .TextMatrix(lngR, 4) = "지급"
                  .Cell(flexcpForeColor, lngR, 4, lngR, 4) = vbBlue
                  '.Cell(flexcpForeColor, lngR, 8, lngR, 8) = vbBlue
               ElseIf _
                  P_adoRec("거래구분") = 2 Then
                  .Cell(flexcpForeColor, lngR, 4, lngR, 4) = vbRed
                  '.Cell(flexcpForeColor, lngR, 8, lngR, 8) = vbRed
                  .TextMatrix(lngR, 4) = "수금"
               End If
               .TextMatrix(lngR, 5) = P_adoRec("업체코드")
               .TextMatrix(lngR, 6) = P_adoRec("업체명")
               .TextMatrix(lngR, 7) = P_adoRec("결제방법")
               .TextMatrix(lngR, 8) = P_adoRec("거래금액")
               lblTotMny.Caption = Format(Vals(lblTotMny.Caption) + P_adoRec("거래금액"), "#,0")
               If P_adoRec("결제방법") = "어음" Then
                  .TextMatrix(lngR, 9) = Format(P_adoRec("만기일자"), "0000-00-00")
                  .TextMatrix(lngR, 10) = P_adoRec("어음번호")
               End If
               .TextMatrix(lngR, 11) = P_adoRec("적요")
               .TextMatrix(lngR, 12) = P_adoRec("사용자코드")
               .TextMatrix(lngR, 13) = P_adoRec("사용자명")
               .TextMatrix(lngR, 14) = P_adoRec("거래시간")
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "지급/수금 읽기 실패"
    Unload Me
    Exit Sub
End Sub

