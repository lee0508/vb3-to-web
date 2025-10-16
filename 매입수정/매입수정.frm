VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm매입수정 
   BorderStyle     =   0  '없음
   Caption         =   "매입수정"
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
   Begin VSFlex7Ctl.VSFlexGrid vsfg2 
      Height          =   5777
      Left            =   60
      TabIndex        =   6
      Top             =   4248
      Width           =   15195
      _cx             =   26802
      _cy             =   10190
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
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   12
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   7980
         Picture         =   "매입수정.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   16
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
         Picture         =   "매입수정.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "매입수정.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "매입수정.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "매입수정.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "매입수정.frx":2E61
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
         Caption         =   "매입전표 조회및 수정"
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
         TabIndex        =   13
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   2475
      Left            =   60
      TabIndex        =   5
      Top             =   1695
      Width           =   15195
      _cx             =   26802
      _cy             =   4366
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
      Height          =   1035
      Left            =   60
      TabIndex        =   11
      Top             =   630
      Width           =   15195
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   1
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   19791873
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   5160
         TabIndex        =   4
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   19791873
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   4
         Left            =   7520
         TabIndex        =   24
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   3
         Left            =   2520
         TabIndex        =   23
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   6
         Left            =   6600
         TabIndex        =   22
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   5
         Left            =   4440
         TabIndex        =   21
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "거래일자"
         Height          =   240
         Index           =   2
         Left            =   1200
         TabIndex        =   20
         Top             =   650
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "전체금액"
         Height          =   240
         Index           =   1
         Left            =   10800
         TabIndex        =   19
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
         Left            =   12120
         TabIndex        =   18
         Top             =   285
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   2520
         TabIndex        =   17
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "매입처코드"
         Height          =   240
         Index           =   0
         Left            =   1200
         TabIndex        =   15
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
         Left            =   120
         TabIndex        =   14
         Top             =   285
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm매입수정"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 매입수정
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 자재입출내역, 사업장, 매입처, 매출처
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_adoRecW            As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private Const PC_intRowCnt1  As Integer = 8   '그리드1의 한 페이지 당 행수(FixedRows 포함)
Private Const PC_intRowCnt2  As Integer = 18  '그리드2의 한 페이지 당 행수(FixedRows 포함)

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
       Subvsfg1_INIT  '거래합계
       Subvsfg2_INIT  '거래내역
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 20 '인쇄, 조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 40 '추가, 저장, 인쇄, 조회
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
                   cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
                   cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매입수정(서버와의 연결 실패)"
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
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '매입처검색
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "매입처 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------+
'/// 매입처 ///
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

'+-------------------+
'/// 거래일자선택 ///
'+-------------------+
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If cmdFind.Enabled = True Then
          cmdFind_Click
       Else
          SendKeys "{tab}"
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
         If .MouseRow >= .FixedRows Then
            If (.MouseCol = 9) Then      '현금구분
                If Button = vbLeftButton Then
                  .Select .MouseRow, .MouseCol
                  .EditCell
                End If
            ElseIf _
               (.MouseCol = 10) Then    '매입일자
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
    With vsfg1
         If Row >= .FixedRows Then
            If (Col = 9) Then  '현금구분
               If .Cell(flexcpChecked, Row, Col) <> .EditText Then
                  blnModify = True
               End If
            ElseIf _
               (Col = 10) Then '매입일자
               If .TextMatrix(Row, 1) <> .EditText Then
                  .EditText = Format(Replace(.EditText, "-", ""), "0000-00-00")
                  If Not ((Len(Trim(.EditText)) = 10) And IsDate(.EditText) And Val(.EditText) > 2000) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            End If
            '변경표시
            If blnModify = True Then
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
            Else
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbWhite
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbBlack
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
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 3))
            'Select Case .MouseCol
            '       Case 0
            '            .ColSel = 3
            '            .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case 1
            '            .ColSel = 3
            '            .ColSort(0) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(1) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .ColSort(2) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            '            .Sort = flexSortUseColSort
            '       Case Else
            '            .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            'End Select
            'If .FindRow(strData, , 3) > 0 Then
            '   .Row = .FindRow(strData, , 3)
            'End If
            'If PC_intRowCnt1 < .Rows Then
            '   .TopRow = .Row
            'End If
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
         If NewRow < 1 Then Exit Sub
         If NewRow <> OldRow Then
            For lngR2 = 1 To vsfg2.Rows - 1
                vsfg2.RowHidden(lngR2) = True
            Next lngR2
            For lngR1 = .ValueMatrix(.Row, 14) To .ValueMatrix(.Row, 15)
                If vsfg2.TextMatrix(lngR1, 28) = "D" Then
                   vsfg2.RowHidden(lngR1) = True
                Else
                   vsfg2.RowHidden(lngR1) = False
                   lngCnt = lngCnt + 1
                   vsfg2.TextMatrix(lngR1, 0) = lngCnt '순번
                End If
            Next lngR1
            If PC_intRowCnt2 < lngCnt Then
               vsfg2.TopRow = vsfg2.Row
            End If
            vsfg2.Row = 0
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
    End With
    Exit Sub
End Sub

'+------------+
'/// vsfg2 ///
'+------------+
Private Sub vsfg2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg2
         .ToolTipText = ""
         If .MouseRow < .FixedRows Or .MouseCol < 0 Then
            Exit Sub
         End If
         .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub
Private Sub vsfg2_DblClick()
    With vsfg1
         If .Row >= .FixedRows Then
             vsfg2_KeyDown vbKeyF1, 0  '자재시세검색
         End If
    End With
End Sub
Private Sub vsfg2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg2
         P_intButton = Button
         If .Row >= .FixedRows Then
            If (.Col = 16) Then      '수량
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
               (.Col = 18) Then      '직송구분
               If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
                (.Col = 19) Then     '입고단가
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
                (.Col = 20) Then     '입고부가
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            ElseIf _
                (.Col = 27) Then     '적요
                If Button = vbLeftButton Then
                  .Select .Row, .Col
                  .EditCell
                End If
            End If
         End If
    End With
End Sub
Sub vsfg2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim blnModify As Boolean
Dim curTmpMny As Currency
    With vsfg2
         If Row >= .FixedRows Then
            If (Col = 16) Then  '수량
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 22)
                     .TextMatrix(Row, 22) = Vals(.EditText) * .ValueMatrix(Row, 19)
                  End If
               End If
            ElseIf _
               (Col = 18) Then  '직송구분
               If (Len(.TextMatrix(Row, 9)) = 0) Then '매출처가 없음
                  .Cell(flexcpChecked, Row, 18, Row, 18) = flexUnchecked
                  Beep
                  Cancel = True
                  Exit Sub
               End If
               If .Cell(flexcpChecked, Row, Col) <> .EditText Then
                  blnModify = True
               End If
            ElseIf _
               (Col = 19) Then  '입고단가
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     IsNumeric(Right(.EditText, 1)) = False) Then                                            '소숫점이하 사용가
                     'fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then '소숫점이하 사용불가
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     .TextMatrix(Row, 20) = Fix(Vals(.EditText) * (PB_curVatRate))  '부가세
                     curTmpMny = .ValueMatrix(Row, 22)
                     .TextMatrix(Row, 21) = Vals(.EditText) + .ValueMatrix(Row, 20)
                     .TextMatrix(Row, 22) = .ValueMatrix(Row, 16) * Vals(.EditText)
                  End If
               End If
            ElseIf _
               (Col = 20) Then  '입고부가
               If .TextMatrix(Row, Col) <> .EditText Then
                  If (IsNumeric(.EditText) = False Or Vals(.EditText) < 0 Or _
                     Fix(Vals(.EditText)) < Vals(.EditText) Or IsNumeric(Right(.EditText, 1)) = False) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                     curTmpMny = .ValueMatrix(Row, 22)
                     .TextMatrix(Row, 21) = .ValueMatrix(Row, 19) + Vals(.EditText)
                     .TextMatrix(Row, 22) = .ValueMatrix(Row, 16) * .ValueMatrix(Row, 19)
                  End If
               End If
            ElseIf _
               (Col = 27) Then '적요 길이 검사
               If .TextMatrix(Row, Col) <> .EditText Then
                  If Not (LenH(Trim(.EditText)) <= 50) Then
                     Beep
                     Cancel = True
                  Else
                     blnModify = True
                  End If
               End If
            End If
            '변경표시 + 금액재계산
            If blnModify = True Then
               If .TextMatrix(Row, 28) = "" Then
                  .TextMatrix(Row, 28) = "U"
               End If
               .Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
               .Cell(flexcpForeColor, Row, Col, Row, Col) = vbWhite
               Select Case Col
                      Case 16, 19, 20
                           vsfg1.TextMatrix(vsfg1.Row, 6) = vsfg1.ValueMatrix(vsfg1.Row, 6) - curTmpMny + .ValueMatrix(Row, 22)
                           lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - curTmpMny + .ValueMatrix(Row, 22), "#,#.00")
                      Case Else
               End Select
            End If
         End If
    End With
End Sub
Private Sub vsfg2_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg2
         .Editable = flexEDNone
         If .Row >= .FixedRows Then
             Select Case .Col
                    Case 16, 19, 27
                         .Editable = flexEDKbdMouse
                         vsfg2_MouseUp vbLeftButton, 0, 0, 0
             End Select
         End If
    End With
End Sub
Private Sub vsfg2_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim lngR    As Long
    With vsfg2
         If KeyCode = vbKeyReturn Then
            If Col = 16 Then
               .Col = 19
            ElseIf _
               Col = 19 Then
               .Col = 27
            ElseIf _
               Col = 27 Then
               For lngR = Row To vsfg1.ValueMatrix(vsfg1.Row, 15)
                   If .RowHidden(lngR) = False Then
                      Exit For
                   End If
               Next lngR
               If lngR <> vsfg1.ValueMatrix(vsfg1.Row, 15) Then
                  .Col = 16: .LeftCol = 15
                  .Row = .Row + 1
                  If .ValueMatrix(lngR, 0) > PC_intRowCnt2 Then
                     .TopRow = .TopRow + 1
                  End If
               End If
            End If
         End If
    End With
End Sub
Private Sub vsfg2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lngR     As Long
Dim lngPos   As Long
Dim blnDupOK As Boolean
Dim strTime  As String
Dim strHH    As String
Dim strMM    As String
Dim strSS    As String
Dim strMS    As String
Dim intRetVal As Integer
Dim CtrlDown  As Variant
    With vsfg2
         If (.Row >= .FixedRows) Then     '내역시세검색
            If KeyCode = vbKeyF2 And (Len(vsfg1.TextMatrix(vsfg1.Row, 5)) > 0) And _
              (Len(.TextMatrix(.Row, 4)) > 0) Then
               PB_strFMCCallFormName = "frm매입수정"
               PB_strMaterialsCode = .TextMatrix(.Row, 4)
               PB_strMaterialsName = .TextMatrix(.Row, 5)
               PB_strSupplierCode = vsfg1.TextMatrix(vsfg1.Row, 4)
               frm내역시세검색.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
               End If
            End If
         End If
    End With
    With vsfg2
         '내역이 없는 경우 추가
         If .Row = 0 And KeyCode = vbKeyInsert Then
            For lngR = vsfg1.ValueMatrix(vsfg1.Row, 14) To vsfg1.ValueMatrix(vsfg1.Row, 15)
                If .RowHidden(lngR) = False Then
                   lngPos = lngPos + 1
                End If
            Next lngR
            If lngPos = 0 Then .Row = vsfg1.ValueMatrix(vsfg1.Row, 15)
         End If
         If .Row >= .FixedRows Then
            If KeyCode = vbKeyF1 Then  '자재시세검색
               'If (.MouseCol = 5) Then
                  PB_strFMCCallFormName = "frm매입수정"
                  PB_strMaterialsCode = .TextMatrix(.Row, 4)
                  PB_strMaterialsName = .TextMatrix(.Row, 5)
                  PB_strSupplierCode = .TextMatrix(.Row, 13)
                  frm자재시세검색.Show vbModal
                  If Len(PB_strMaterialsCode) <> 0 Then
                     PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  End If
               'ElseIf _
               '   (.Col = 9) Then      '매출처검색
               '   PB_strSupplierCode = .TextMatrix(.Row, 8)
               '   PB_strSupplierName = .TextMatrix(.Row, 9)
               '   frm매출처검색.Show vbModal
               '   If Len(PB_strSupplierCode) <> 0 Then
               '      'For lngR = vsfg1.ValueMatrix(vsfg1.Row, 14) To vsfg1.ValueMatrix(vsfg1.Row, 15) '자재코드(변경후) + 매출처코드
               '      '    If .Row <> lngR And .TextMatrix(lngR, 0) = .TextMatrix(.Row, 0) And _
               '      '       .TextMatrix(lngR, 4) = PB_strSupplierCode Then
               '      '       blnDupOK = True
               '      '       Exit For
               '      '    End If
               '      'Next lngR
               '      If blnDupOK = False Then
               '         If PB_strSupplierCode <> .TextMatrix(.Row, 8) Then
               '            .Cell(flexcpBackColor, .Row, 9, .Row, 9) = vbRed
               '            .Cell(flexcpForeColor, .Row, 9, .Row, 9) = vbWhite
               '         End If
               '         .TextMatrix(.Row, 8) = PB_strSupplierCode
               '         .TextMatrix(.Row, 9) = PB_strSupplierName
               '         If .TextMatrix(.Row, 28) = "" Then
               '            .TextMatrix(.Row, 28) = "U"
               '         End If
               '      End If
               '   End If
               'End If
            ElseIf _
               KeyCode = vbKeyInsert Then '거래내역 추가
               .AddItem "", .Row + 1
               .Row = .Row + 1
               .TopRow = .Row
               .TextMatrix(.Row, 0) = .ValueMatrix(.Row - 1, 0) + 1 '순번
               .TextMatrix(.Row, 1) = .TextMatrix(.Row - 1, 1)   '사업장코드
               .TextMatrix(.Row, 2) = .TextMatrix(.Row - 1, 2)   '거래일자
               .TextMatrix(.Row, 3) = .TextMatrix(.Row - 1, 3)   '거래번호
               .TextMatrix(.Row, 13) = .TextMatrix(.Row - 1, 13) '매입처코드
               .TextMatrix(.Row, 14) = .TextMatrix(.Row - 1, 14) '매입처명
               .Cell(flexcpChecked, .Row, 18) = flexUnchecked    '직송
               .Cell(flexcpText, .Row, 18) = "직 송"
               .Cell(flexcpAlignment, .Row, 18, .Row, 18) = flexAlignLeftCenter
               .TextMatrix(.Row, 28) = "I"                       'SQL구분
               '자재입출시간
               strTime = .TextMatrix(.Row - 1, 29)
               If .Row <= vsfg1.ValueMatrix(vsfg1.Row, 15) Then  '거래번호의 마지막 아니면
                  strTime = Format(Fix((.ValueMatrix(.Row - 1, 29) + .ValueMatrix(.Row + 1, 29)) / 2), "000000000")
                  '추가 가능한지 검사
                  If (strTime = .TextMatrix(.Row - 1, 29)) Or (strTime = .TextMatrix(.Row - 1, 29)) Then
                     MsgBox "이 행에는 더 이상 추가 할 수 없습니다. 다른 행에 추가하세요.", vbCritical + vbDefaultButton1, "추가"
                     .RemoveItem (.Row)
                     Exit Sub
                  End If
               Else                                              '거래번호의 마지막이면
                  strTime = Format((Val(strTime) + 10000), "000000000")
                  strHH = Mid(strTime, 1, 2): strMM = Mid(strTime, 3, 2): strSS = Mid(strTime, 5, 2): strMS = Mid(strTime, 7, 3)
                  If Val(strMS) > 999 Then
                     strMS = Format(0, "000")
                     strSS = Format(Val(strMM) + 1, "00")
                  End If
                  If Val(strSS) > 59 Then
                     strSS = Format(Val(strSS) - 60, "00")
                     strMM = Format(Val(strMM) + 1, "00")
                  End If
                  If Val(strMM) > 59 Then
                     strMM = Format(Val(strMM) - 60, "00")
                     strHH = Format(Val(strHH) + 1, "00")
                  End If
                  strTime = strHH & strMM & strSS & strMS
               End If
               .TextMatrix(.Row, 29) = strTime                   '자재입출시간
               PB_strFMCCallFormName = "frm매입수정"
               PB_strMaterialsCode = .TextMatrix(.Row, 4)
               PB_strMaterialsName = .TextMatrix(.Row, 5)
               PB_strSupplierCode = .TextMatrix(.Row, 13)
               frm자재시세검색.Show vbModal
               If Len(PB_strMaterialsCode) <> 0 Then '정상적인 선택이면
                  PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  '순번
                  For lngR = (.Row + 1) To (vsfg1.ValueMatrix(vsfg1.Row, 15) + 1)
                      If .TextMatrix(lngR, 28) <> "D" Then
                         .TextMatrix(lngR, 0) = .ValueMatrix(lngR, 0) + 1
                      End If
                  Next lngR
                  vsfg1.TextMatrix(vsfg1.Row, 15) = vsfg1.ValueMatrix(vsfg1.Row, 15) + 1
                  For lngR = 1 To vsfg1.Rows - 1
                      If (lngR <> .Row) Then
                         If (vsfg1.ValueMatrix(lngR, 14) >= .Row) Then
                            vsfg1.TextMatrix(lngR, 14) = vsfg1.ValueMatrix(lngR, 14) + 1
                            vsfg1.TextMatrix(lngR, 15) = vsfg1.ValueMatrix(lngR, 15) + 1
                         End If
                      End If
                  Next lngR
               Else
                  .RemoveItem (.Row)
                  .Row = .Row - 1
               End If
            ElseIf _
               KeyCode = vbKeyDelete And .Col = 9 Then  '매출처 지움
               If (Len(.TextMatrix(.Row, 9)) <> 0) Then
                  .TextMatrix(.Row, 8) = "": .TextMatrix(.Row, 9) = ""
                  .Cell(flexcpBackColor, .Row, 9, .Row, 9) = vbRed
                  .Cell(flexcpForeColor, .Row, 9, .Row, 9) = vbWhite
                  If .TextMatrix(.Row, 28) = "" Then
                     .TextMatrix(.Row, 28) = "U"
                  End If
               End If
               If .Cell(flexcpChecked, .Row, 18, .Row, 18) = flexChecked Then
                  .Cell(flexcpChecked, .Row, 18, .Row, 18) = flexUnchecked
                  .Cell(flexcpBackColor, .Row, 18, .Row, 18) = vbRed
                  .Cell(flexcpForeColor, .Row, 18, .Row, 18) = vbWhite
               End If
            ElseIf _
               KeyCode = vbKeyDelete And (.Col <> 9) And (.Row > 0) And .RowHidden(.Row) = False Then
               intRetVal = MsgBox("입력한 거래내역을 삭제하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton2, "거래내역삭제")
               If intRetVal = vbYes Then
                  .TextMatrix(.Row, 28) = "D": .TextMatrix(.Row, 0) = "0"
                  vsfg1.TextMatrix(vsfg1.Row, 6) = vsfg1.ValueMatrix(vsfg1.Row, 6) - .ValueMatrix(.Row, 22)
                  lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 22), "#,#.00")
                  .RowHidden(.Row) = True
                  For lngR = .Row + 1 To vsfg1.ValueMatrix(vsfg1.Row, 15)
                      If .TextMatrix(lngR, 28) <> "D" Then
                         .TextMatrix(lngR, 0) = .ValueMatrix(lngR, 0) - 1
                         If lngPos = 0 Then
                            lngPos = lngR
                         End If
                      End If
                  Next lngR
                  If lngPos = 0 Then
                     For lngR = vsfg1.ValueMatrix(vsfg1.Row, 14) To .Row
                         If .TextMatrix(lngR, 28) <> "D" And lngR < .Row Then
                            lngPos = lngR
                         End If
                     Next lngR
                  End If
                  .Row = lngPos
               End If
            End If
         End If
    End With
End Sub

'+-----------+
'/// 출력 ///
'+-----------+
Private Sub cmdPrint_Click()
    If ((vsfg1.Rows + vsfg2.Rows) = 2) Or (vsfg1.Row < 1) Then
       Exit Sub
    End If
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
    Subvsfg2_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
Dim blnASaveOK     As Boolean
Dim blnBSaveOK     As Boolean
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
Dim lngOrgLogCnt   As Long
Dim curInputMoney  As Long
Dim curOutputMoney As Long
Dim strServerTime  As String
Dim strTime        As String
Dim strHH          As String
Dim strMM          As String
Dim strSS          As String
Dim strMS          As String
Dim intChkCash     As Integer '현금구분(현금매입)
Dim intChgDate     As Integer '매입일자 변경 검사
Dim strOrgDate     As String  '변경전 매입일자
Dim strChgDate     As String  '변경한 매입일자

    If vsfg1.Row >= vsfg1.FixedRows Then
       With vsfg1
            '현금매입 부분 변경, 매입일자 변경
            If (.Cell(flexcpBackColor, .Row, 9, .Row, 9) = vbRed) Or (.Cell(flexcpBackColor, .Row, 10, .Row, 10) = vbRed) Then
               blnASaveOK = True
            End If
            If .Cell(flexcpChecked, .Row, 9) = flexChecked Then          '현금매입
               intChkCash = 1
            End If
            If (.Cell(flexcpBackColor, .Row, 10, .Row, 10) = vbRed) Then '매입일자
               intChgDate = 1
            End If
       End With
       With vsfg2
            'If .RowHidden(lngRR) = False Then
                For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 14) To vsfg1.ValueMatrix(vsfg1.Row, 15)
                    If (.TextMatrix(lngRR, 28) <> "") Then
                       blnBSaveOK = True
                       Exit For
                    End If
                Next lngRR
            'End If
            If (blnASaveOK = False And blnBSaveOK = False) Then '저장할(변경된) 것이없으면
               Exit Sub
            End If
       End With
       intRetVal = MsgBox("수정된 매입자료를 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "자료 저장")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       cmdSave.Enabled = False
       Screen.MousePointer = vbHourglass
       If vsfg1.Cell(flexcpChecked, vsfg1.Row, 9) = flexChecked Then
          intChkCash = 1
       End If
       '서버시간 구하기
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       strServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) + Mid(P_adoRec("서버시간"), 7, 2) _
                     + Mid(P_adoRec("서버시간"), 10)
       P_adoRec.Close
       strTime = strServerTime
       PB_adoCnnSQL.BeginTrans
       '현재 거래일자, 거래번호 보관
       strOrgDate = DTOS(vsfg1.TextMatrix(vsfg1.Row, 1))
       lngOrgLogCnt = vsfg1.ValueMatrix(vsfg1.Row, 2)
       '거래번호 구하기
       If intChgDate = 1 Then '매입일자 변경이면
          strChgDate = DTOS(vsfg1.TextMatrix(vsfg1.Row, 10))
          strSQL = "spLogCounter '자재입출내역', '" & PB_regUserinfoU.UserBranchCode + DTOS(vsfg1.TextMatrix(vsfg1.Row, 10)) + "1" & "', " _
                               & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
          On Error GoTo ERROR_STORED_PROCEDURE
          P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          lngLogCnt = P_adoRec(0)
          P_adoRec.Close
          '자재입출내역 매입일자 변경
          strSQL = "UPDATE 자재입출내역 SET " _
                        & "입출고일자 = '" & strChgDate & "', " _
                        & "거래일자 = '" & strChgDate & "', 거래번호 = " & lngLogCnt & " " _
                  & "WHERE 사업장코드 = '" & vsfg1.TextMatrix(vsfg1.Row, 0) & "' AND 입출고구분 = 1 AND 사용구분 = 0 " _
                    & "AND 거래일자 = '" & strOrgDate & "' " _
                    & "AND 거래번호 = " & lngOrgLogCnt & " "
          On Error GoTo ERROR_TABLE_UPDATE
          PB_adoCnnSQL.Execute strSQL
          '그리드 상,하 거래일자 바꾸기
          '상
          With vsfg1
               .TextMatrix(.Row, 1) = .TextMatrix(.Row, 10)
               .TextMatrix(.Row, 2) = lngLogCnt
               .TextMatrix(.Row, 3) = .TextMatrix(.Row, 0) & "-" & Format(strChgDate, "0000/00/00") & "-" & CStr(lngLogCnt)
               .Cell(flexcpData, .Row, 3, lngR, 3) = Trim(.TextMatrix(.Row, 3)) 'FindRow 사용을 위해
          End With
          '하
          With vsfg2
               For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 14) To vsfg1.ValueMatrix(vsfg1.Row, 15)
                   .TextMatrix(lngRR, 2) = vsfg1.TextMatrix(vsfg1.Row, 10) '거래(매입)일자
                   .TextMatrix(lngRR, 3) = lngLogCnt                       '거래번호
                   .TextMatrix(lngRR, 12) = .TextMatrix(lngRR, 1) & "-" & Format(strChgDate, "0000/00/00") _
                                     & "-" & CStr(lngLogCnt) & "-" & .TextMatrix(lngRR, 29) _
                                     & "-" & .TextMatrix(lngRR, 6) & "-" & .TextMatrix(lngRR, 10)
               Next lngRR
          End With
       End If
       With vsfg2
            '거래내역
            For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 14) To vsfg1.ValueMatrix(vsfg1.Row, 15)
                If (.TextMatrix(lngRR, 28) = "I") Then 'And .ValueMatrix(lngRR, 16) <> 0) Then '거래내역 추가
                   strSQL = "INSERT INTO 자재입출내역(사업장코드, 분류코드, 세부코드, 입출고구분, " _
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
                             & "'" & .TextMatrix(lngRR, 1) & "', '" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 4), 3) & "', 1, " _
                             & "'" & DTOS(.TextMatrix(lngRR, 2)) & "', '" & .TextMatrix(lngRR, 29) & "', " _
                             & "" & .ValueMatrix(lngRR, 16) & ", " & .ValueMatrix(lngRR, 19) & ", " _
                             & "" & .ValueMatrix(lngRR, 20) & ", 0, " _
                             & "" & .ValueMatrix(lngRR, 23) & ", " & .ValueMatrix(lngRR, 24) & ", " _
                             & "'" & .TextMatrix(lngRR, 13) & "' , '" & .TextMatrix(lngRR, 8) & "', " _
                             & "'" & DTOS(.TextMatrix(lngRR, 2)) & "'," & IIf(.Cell(flexcpChecked, lngRR, 18, lngRR, 18) = flexUnchecked, 0, 1) & ", " _
                             & "'', 0, " _
                             & "'" & DTOS(.TextMatrix(lngRR, 2)) & "', " & .ValueMatrix(lngRR, 3) & ", " _
                             & "0, " & intChkCash & ", 0, '" & Trim(.TextMatrix(lngRR, 27)) & "', '', 0, 0, " _
                             & "0, '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "'" & PB_regUserinfoU.UserCode & "', '' ) "
                   On Error GoTo ERROR_TABLE_INSERT
                   PB_adoCnnSQL.Execute strSQL
                   '최종입출고일자갱신(사업장코드, 분류코드, 세부코드, 입출고구분)
                   strSQL = "sp자재최종입출고일자갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 1 "
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                   '자재최종단가갱신(사업장코드, 분류코드, 세부코드, 입출고구분, 업체코드, 단가, 거래일자)
                   If .ValueMatrix(lngRR, 19) > 0 And PB_intIAutoPriceGbn = 1 Then
                      strSQL = "sp자재최종단가갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 1, " _
                             & "'" & .TextMatrix(lngRR, 13) & "', " _
                             & "" & .ValueMatrix(lngRR, 19) & ", '" & DTOS(.TextMatrix(lngRR, 2)) & "' "
                      On Error GoTo ERROR_STORED_PROCEDURE
                      PB_adoCnnSQL.Execute strSQL
                   End If
                ElseIf _
                   (.TextMatrix(lngRR, 28) = "D") Then '거래내역 삭제
                   'strSQL = "UPDATE 자재입출내역 SET " _
                                 & "사용구분 = 9, " _
                                 & "분류코드 = '" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', " _
                                 & "세부코드 = '" & Mid(.TextMatrix(lngRR, 4), 3) & "', " _
                                 & "입고수량 = " & .ValueMatrix(lngRR, 16) & ", " _
                                 & "직송구분 = " & IIf(.Cell(flexcpChecked, lngRR, 18, lngRR, 18) = flexUnchecked, 0, 1) & ", " _
                                 & "매출처코드 = '" & .TextMatrix(lngRR, 8) & "', " _
                                 & "입고단가 = " & .ValueMatrix(lngRR, 19) & ", " _
                                 & "입고부가 = " & .ValueMatrix(lngRR, 20) & ", " _
                                 & "출고단가 = " & .ValueMatrix(lngRR, 23) & "," _
                                 & "출고부가 = " & .ValueMatrix(lngRR, 24) & ", " _
                                 & "적요 = '" & .TextMatrix(lngRR, 27) & "', " _
                                 & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                                 & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' "
                   strSQL = "DELETE FROM 자재입출내역 " _
                           & "WHERE 사업장코드 = '" & .TextMatrix(lngRR, 1) & "' " _
                             & "AND 분류코드 = '" & Mid(.TextMatrix(lngRR, 6), 1, 2) & "' " _
                             & "AND 세부코드 = '" & Mid(.TextMatrix(lngRR, 6), 3) & "' " _
                             & "AND 입출고구분 = 1 " _
                             & "AND 입출고일자 = '" & DTOS(.TextMatrix(lngRR, 2)) & "' " _
                             & "AND 입출고시간 = '" & .TextMatrix(lngRR, 29) & "' " _
                             & "AND 매출처코드 = '" & .TextMatrix(lngRR, 10) & "' " _
                             & "AND 거래번호 = " & .ValueMatrix(lngRR, 3) & " "
                   lngDelCnt = lngDelCnt + 1     '삭제할 Row수 계산
                   On Error GoTo ERROR_TABLE_DELETE
                   PB_adoCnnSQL.Execute strSQL
                   '최종입출고일자갱신(사업장코드, 분류코드, 세부코드, 입출고구분)
                   strSQL = "sp자재최종입출고일자갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 6), 3) & "', 1 " '변경전
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                ElseIf _
                   (.TextMatrix(lngRR, 28) = "U") Then 'And .ValueMatrix(lngRR, 16) <> 0) Then '거래내역 변경
                   strSQL = "UPDATE 자재입출내역 SET " _
                                 & "분류코드 = '" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', " _
                                 & "세부코드 = '" & Mid(.TextMatrix(lngRR, 4), 3) & "', " _
                                 & "입고수량 = " & .ValueMatrix(lngRR, 16) & ", " _
                                 & "직송구분 = " & IIf(.Cell(flexcpChecked, lngRR, 18, lngRR, 18) = flexUnchecked, 0, 1) & ", " _
                                 & "매출처코드 = '" & .TextMatrix(lngRR, 8) & "', " _
                                 & "입고단가 = " & .ValueMatrix(lngRR, 19) & ", " _
                                 & "입고부가 = " & .ValueMatrix(lngRR, 20) & ", " _
                                 & "출고단가 = " & .ValueMatrix(lngRR, 23) & "," _
                                 & "출고부가 = " & .ValueMatrix(lngRR, 24) & ", " _
                                 & "적요 = '" & .TextMatrix(lngRR, 27) & "', " _
                                 & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                                 & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                           & "WHERE 사업장코드 = '" & .TextMatrix(lngRR, 1) & "' " _
                             & "AND 분류코드 = '" & Mid(.TextMatrix(lngRR, 6), 1, 2) & "' " _
                             & "AND 세부코드 = '" & Mid(.TextMatrix(lngRR, 6), 3) & "' " _
                             & "AND 입출고구분 = 1 " _
                             & "AND 입출고일자 = '" & DTOS(.TextMatrix(lngRR, 2)) & "' " _
                             & "AND 입출고시간 = '" & .TextMatrix(lngRR, 29) & "' " _
                             & "AND 매출처코드 = '" & .TextMatrix(lngRR, 10) & "' " _
                             & "AND 거래번호 = " & .ValueMatrix(lngRR, 3) & " "
                   On Error GoTo ERROR_TABLE_UPDATE
                   PB_adoCnnSQL.Execute strSQL
                   '최종입출고일자갱신(사업장코드, 분류코드, 세부코드, 입출고구분)
                   strSQL = "sp자재최종입출고일자갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 1 "    '변경후
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                   If .TextMatrix(lngRR, 4) <> .TextMatrix(lngRR, 6) Then '자재코드 변경이면
                      '최종입출고일자갱신(사업장코드, 분류코드, 세부코드, 입출고구분)
                      strSQL = "sp자재최종입출고일자갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 6), 3) & "', 1 " '변경전
                      On Error GoTo ERROR_STORED_PROCEDURE
                      PB_adoCnnSQL.Execute strSQL
                   End If
                   '자재최종단가갱신(사업장코드, 분류코드, 세부코드, 입출고구분, 업체코드, 단가, 거래일자)
                   If .ValueMatrix(lngRR, 19) > 0 And PB_intIAutoPriceGbn = 1 Then
                      strSQL = "sp자재최종단가갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                             & "'" & Mid(.TextMatrix(lngRR, 4), 1, 2) & "', '" & Mid(.TextMatrix(lngRR, 4), 3) & "', 1, " _
                             & "'" & .TextMatrix(lngRR, 13) & "', " _
                             & "" & .ValueMatrix(lngRR, 19) & ", '" & DTOS(.TextMatrix(lngRR, 2)) & "' "
                      On Error GoTo ERROR_STORED_PROCEDURE
                      PB_adoCnnSQL.Execute strSQL
                   End If
                End If
                If ((.TextMatrix(lngRR, 28) = "I" Or .TextMatrix(lngRR, 28) = "U") And .ValueMatrix(lngRR, 16) <> 0) Then '추가, 변경
                ElseIf _
                   (.TextMatrix(lngRR, 28) = "U" And .ValueMatrix(lngRR, 16) = 0) Then '거래내역 삭제
                End If
            Next lngRR
       End With
       '현금매입
       With vsfg1
            If blnASaveOK = True Then
               strSQL = "UPDATE 자재입출내역 SET 현금구분 = " & intChkCash & " " _
                       & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 0) & "' AND 거래일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                         & "AND 거래번호 = " & .ValueMatrix(.Row, 2) & " "
               On Error GoTo ERROR_TABLE_UPDATE
               PB_adoCnnSQL.Execute strSQL
            End If
       End With
       With vsfg1
            If (.ValueMatrix(.Row, 15) - .ValueMatrix(.Row, 14) + 1) = lngDelCnt Then '거래내역 모두 삭제
               strSQL = "UPDATE 자재입출내역 SET " _
                             & "사용구분 = 9, " _
                             & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                             & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                       & "WHERE 사업장코드 = '" & .TextMatrix(.Row, 0) & "' AND 거래일자 = '" & DTOS(.TextMatrix(.Row, 1)) & "' " _
                         & "AND 거래번호 = " & .ValueMatrix(.Row, 2) & " "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               lngDelCntS = .ValueMatrix(.Row, 14): lngDelCntE = .ValueMatrix(.Row, 15)
               For lngRR = .ValueMatrix(.Row, 15) To .ValueMatrix(.Row, 14) Step -1
                   '최종입출고일자갱신(사업장코드, 분류코드, 세부코드, 입출고구분)
                   strSQL = "sp자재최종입출고일자갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                          & "'" & Mid(vsfg2.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(vsfg2.TextMatrix(lngRR, 6), 3) & "', 1 "    '변경전
                   On Error GoTo ERROR_STORED_PROCEDURE
                   PB_adoCnnSQL.Execute strSQL
                   vsfg2.RemoveItem lngRR
               Next lngRR
               .RemoveItem .Row
               For lngRRR = 1 To .Rows - 1
                   If lngDelCntS < .ValueMatrix(lngRRR, 14) Then
                      .TextMatrix(lngRRR, 14) = .ValueMatrix(lngRRR, 14) - (lngDelCntE - lngDelCntS + 1)
                      .TextMatrix(lngRRR, 15) = .ValueMatrix(lngRRR, 15) - (lngDelCntE - lngDelCntS + 1)
                   End If
               Next lngRRR
               .Row = 0 '현재선택된 거래내역 Row를 해제
            Else
               lngDelCntS = .ValueMatrix(.Row, 14): lngDelCntE = .ValueMatrix(.Row, 15)
               For lngRR = .ValueMatrix(.Row, 15) To .ValueMatrix(.Row, 14) Step -1
                   If vsfg2.TextMatrix(lngRR, 28) = "D" Then
                      vsfg2.RemoveItem lngRR
                   End If
               Next lngRR
               vsfg1.TextMatrix(vsfg1.Row, 15) = vsfg1.ValueMatrix(vsfg1.Row, 15) - lngDelCnt
               For lngR = 1 To vsfg1.Rows - 1
                   If (lngR <> vsfg1.Row) Then
                      If (vsfg1.ValueMatrix(.Row, 15) < vsfg1.ValueMatrix(lngR, 14)) Then
                         vsfg1.TextMatrix(lngR, 14) = vsfg1.ValueMatrix(lngR, 14) - lngDelCnt
                         vsfg1.TextMatrix(lngR, 15) = vsfg1.ValueMatrix(lngR, 15) - lngDelCnt
                      End If
                   End If
               Next lngR
            End If
       End With
       With vsfg2
            '변경후(재정렬)
            If vsfg1.Row > 0 Then
               For lngRR = vsfg1.ValueMatrix(vsfg1.Row, 14) To vsfg1.ValueMatrix(vsfg1.Row, 15)
                   If (.TextMatrix(lngRR, 28) = "I" Or .TextMatrix(lngRR, 28) = "U") Then 'And .ValueMatrix(lngRR, 16) > 0) Then   '거래내역 변경
                      .TextMatrix(lngRR, 6) = .TextMatrix(lngRR, 4)   '자재코드(변경후->변경전)
                      .TextMatrix(lngRR, 7) = .TextMatrix(lngRR, 5)   '자재명(변경후->변경전)
                      .TextMatrix(lngRR, 10) = .TextMatrix(lngRR, 8)  '매출처코드(변경후->변경전)
                      .TextMatrix(lngRR, 11) = .TextMatrix(lngRR, 9)  '매출처명(변경후->변경전)
                   End If
               Next lngRR
            End If
       End With
       With vsfg2
            '거래내역(색상 원위치)
            If vsfg1.Row > 0 Then '현재선택된 거래 Row를 해제
               vsfg1.Cell(flexcpBackColor, vsfg1.Row, vsfg1.FixedCols, vsfg1.Row, vsfg1.Cols - 1) = vbWhite
               vsfg1.Cell(flexcpForeColor, vsfg1.Row, vsfg1.FixedCols, vsfg1.Row, vsfg1.Cols - 1) = vbBlack
               .Cell(flexcpBackColor, vsfg1.ValueMatrix(vsfg1.Row, 14), 0, vsfg1.ValueMatrix(vsfg1.Row, 15), .FixedCols - 1) = _
               .Cell(flexcpBackColor, 0, 0, 0, 0)
               .Cell(flexcpForeColor, vsfg1.ValueMatrix(vsfg1.Row, 14), 0, vsfg1.ValueMatrix(vsfg1.Row, 15), .FixedCols - 1) = _
               .Cell(flexcpForeColor, 0, 0, 0, 0)
               .Cell(flexcpBackColor, vsfg1.ValueMatrix(vsfg1.Row, 14), .FixedCols, vsfg1.ValueMatrix(vsfg1.Row, 15), .Cols - 1) = vbWhite
               .Cell(flexcpForeColor, vsfg1.ValueMatrix(vsfg1.Row, 14), .FixedCols, vsfg1.ValueMatrix(vsfg1.Row, 15), .Cols - 1) = vbBlack
               .Cell(flexcpText, vsfg1.ValueMatrix(vsfg1.Row, 14), 28, vsfg1.ValueMatrix(vsfg1.Row, 15), 28) = "" 'SQL구분 지움
            End If
       End With
       '저장후 삭제 부분 정리
       'With vsfg2
       '     If vsfg1.Row > 0 Then
       '        For lngRRR = 1 To vsfg1.Rows - 1
       '            If lngDelCntS < vsfg1.ValueMatrix(lngRRR, 14) Then
       '               vsfg1.TextMatrix(lngRRR, 14) = vsfg1.ValueMatrix(lngRRR, 14) - (lngDelCntE - lngDelCntS + 1)
       '               vsfg1.TextMatrix(lngRRR, 15) = vsfg1.ValueMatrix(lngRRR, 15) - (lngDelCntE - lngDelCntS + 1)
       '            End If
       '        Next lngRRR
       '     End If
       'End With
       PB_adoCnnSQL.CommitTrans
       'vsfg1.Row = 0: vsfg2.Row = 0
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "서버 연결 실패"
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
    If vsfg1.Row >= vsfg1.FixedRows Then
       intRetVal = MsgBox("매입처리된 자료를 삭제하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton2, "자료 삭제")
       If intRetVal = vbNo Then
          vsfg1.SetFocus
          Exit Sub
       End If
       cmdDelete.Enabled = False
       Screen.MousePointer = vbHourglass
       PB_adoCnnSQL.BeginTrans
       P_adoRec.CursorLocation = adUseClient
       With vsfg1
            lngDelCntS = .ValueMatrix(.Row, 14): lngDelCntE = .ValueMatrix(.Row, 15)
            '거래내역
            For lngRR = .ValueMatrix(.Row, 15) To .ValueMatrix(.Row, 14) Step -1
                strSQL = "UPDATE 자재입출내역 SET " _
                              & "사용구분 = 9, " _
                              & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                              & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                        & "WHERE 사업장코드 = '" & vsfg2.TextMatrix(lngRR, 1) & "' " _
                          & "AND 분류코드 = '" & Mid(vsfg2.TextMatrix(lngRR, 6), 1, 2) & "' " _
                          & "AND 세부코드 = '" & Mid(vsfg2.TextMatrix(lngRR, 6), 3) & "' " _
                          & "AND 입출고구분 = 1 " _
                          & "AND 입출고일자 = '" & DTOS(vsfg2.TextMatrix(lngRR, 2)) & "' " _
                          & "AND 입출고시간 = '" & vsfg2.TextMatrix(lngRR, 29) & "' " _
                          & "AND 매출처코드 = '" & vsfg2.TextMatrix(lngRR, 10) & "' " _
                          & "AND 거래번호 = " & vsfg2.ValueMatrix(lngRR, 3) & " "
                On Error GoTo ERROR_TABLE_UPDATE
                PB_adoCnnSQL.Execute strSQL
                '최종입출고일자갱신(사업장코드, 분류코드, 세부코드, 입출고구분)
                strSQL = "sp자재최종입출고일자갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                       & "'" & Mid(vsfg2.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(vsfg2.TextMatrix(lngRR, 6), 3) & "', 1 "
                On Error GoTo ERROR_STORED_PROCEDURE
                PB_adoCnnSQL.Execute strSQL
                '자재최종단가갱신(사업장코드, 분류코드, 세부코드, 입출고구분, 업체코드, 단가, 거래일자)
                'If vsfg2.ValueMatrix(lngRR, 19) > 0 And PB_intIAutoPriceGbn = 1 Then
                '   strSQL = "sp자재최종단가갱신 '" & PB_regUserinfoU.UserBranchCode & "', " _
                '          & "'" & Mid(vsfg2.TextMatrix(lngRR, 6), 1, 2) & "', '" & Mid(vsfg2.TextMatrix(lngRR, 6), 3) & "', 1, " _
                '          & "'" & vsfg2.TextMatrix(lngRR, 13) & "', " _
                '          & "" & vsfg2.ValueMatrix(lngRR, 19) & ", '" & vsfg2.ValueMatrix(lngRR, 2) & "' "
                '   On Error GoTo ERROR_STORED_PROCEDURE
                '   PB_adoCnnSQL.Execute strSQL
                'End If
                vsfg2.RemoveItem lngRR
            Next lngRR
            lblTotMny.Caption = Format(Vals(lblTotMny.Caption) - .ValueMatrix(.Row, 6), "#,#.00") '전체금액에서 제외
            .RemoveItem .Row
            For lngRRR = 1 To .Rows - 1
                If lngDelCntS < .ValueMatrix(lngRRR, 14) Then
                   .TextMatrix(lngRRR, 14) = .ValueMatrix(lngRRR, 14) - (lngDelCntE - lngDelCntS + 1)
                   .TextMatrix(lngRRR, 15) = .ValueMatrix(lngRRR, 15) - (lngDelCntE - lngDelCntS + 1)
                End If
            Next lngRRR
            .Row = 0
       End With
       PB_adoCnnSQL.CommitTrans
       cmdFind.SetFocus
       Screen.MousePointer = vbDefault
    End If
    cmdDelete.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "저장 실패"
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "거래내역 삭제 실패"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "서버 연결 실패"
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
    Set frm매입수정 = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    Text1(0).Text = "": Text1(1).Text = ""
    dtpF_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
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
    With vsfg1              'Rows 1, Cols 16, RowHeightMax(Min) 300
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
         .Cols = 16
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1200   '사업장코드 'H
         .ColWidth(1) = 1200   '거래일자
         .ColWidth(2) = 1000   '거래번호
         .ColWidth(3) = 1730   '사업장코드+거래일자+거래번호(KEY) H
         .ColWidth(4) = 1500   '매입처코드
         .ColWidth(5) = 6800   '매입처명
         .ColWidth(6) = 2000   '금액(단가)
         .ColWidth(7) = 2000   '금액(부가) 'H
         .ColWidth(8) = 2000   '금액(합계) 'H
         .ColWidth(9) = 1000   '매입구분(현금구분)
         .ColWidth(10) = 1200  '매입일자(변경후)
         .ColWidth(11) = 1000  '           'H(Not Used)
         .ColWidth(12) = 2000  '           'H(Not Used)
         .ColWidth(13) = 1500  '           'H(Not Used)
         .ColWidth(14) = 1000  'ROW(vsfg2.Row) 'H(Not Used)
         .ColWidth(15) = 1000  'COL(vsfg2.Row) 'H(Not Used)
         
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "사업장코드" 'H
         .TextMatrix(0, 1) = "거래일자"
         .TextMatrix(0, 2) = "거래번호"
         .TextMatrix(0, 3) = "거래(KEY)"  'H (KEY)
         .TextMatrix(0, 4) = "매입처코드"
         .TextMatrix(0, 5) = "매입처명"
         .TextMatrix(0, 6) = "금액"
         .TextMatrix(0, 7) = "금액"       'H
         .TextMatrix(0, 8) = "금액"       'H
         .TextMatrix(0, 9) = "매입구분"
         .TextMatrix(0, 10) = "매입일자"
         .TextMatrix(0, 11) = "Not Used"  'H
         .TextMatrix(0, 12) = "Not Used"  'H
         .TextMatrix(0, 13) = "Not Used"  'H
         .TextMatrix(0, 14) = "Row"       'H
         .TextMatrix(0, 15) = "Col"       'H
         
         .ColHidden(0) = True: .ColHidden(3) = True:
         .ColHidden(7) = True: .ColHidden(8) = True
         .ColHidden(11) = True: .ColHidden(12) = True: .ColHidden(13) = True:
         .ColHidden(14) = True: .ColHidden(15) = True
         .ColFormat(6) = "#,#.00": .ColFormat(7) = "#,#.00": .ColFormat(8) = "#,#.00"
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 5
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 2, 3, 4, 9, 10
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictRows  'flexMergeFixedOnly
         .MergeRow(0) = True
         For lngC = 0 To 3
             .MergeCol(lngC) = True
         Next lngC
    End With
End Sub
Private Sub Subvsfg2_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg2              'Rows 1, Cols 30, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .ScrollBars = flexScrollBarBoth
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 6
         .Rows = 1             'Subvsfg2_Fill수행시에 설정
         .Cols = 30
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 400    'No
         .ColWidth(1) = 1000   '사업장코드
         .ColWidth(2) = 1200   '거래일자
         .ColWidth(3) = 1000   '거래번호
         .ColWidth(4) = 1900   '품목코드(변경후)
         .ColWidth(5) = 2600   '품명(변경후)
         .ColWidth(6) = 1900   '퓸목코드(변경전)   'H
         .ColWidth(7) = 2600   '품명(변경전)       'H
         .ColWidth(8) = 1000   '매출처코드(변경후) 'H
         .ColWidth(9) = 2000   '매출처명(변경후)   'H
         .ColWidth(10) = 1000  '매출처코드(변경전) 'H
         .ColWidth(11) = 2000  '매출처명(변경전) 'H
         .ColWidth(12) = 2000  '사업장코드+거래일자+거래번호+입출고시간+자재코드+매출처코드+(KEY) 'H
         .ColWidth(13) = 1000  '매입처코드 'H
         .ColWidth(14) = 2500  '매입처명   'H
         .ColWidth(15) = 2200  '자재규격
         .ColWidth(16) = 1000  '수량
         .ColWidth(17) = 800   '단위
         .ColWidth(18) = 800   '직송       'H
         .ColWidth(19) = 1600  '매입단가
         .ColWidth(20) = 1200  '매입부가   'H
         .ColWidth(21) = 1600  '매입가격(단가+부가) 'H
         .ColWidth(22) = 1700  '매입금액
         .ColWidth(23) = 1600  '매출단가   'H
         .ColWidth(24) = 1200  '매출부가   'H
         .ColWidth(25) = 1600  '매출가격(단가+부가) 'H
         .ColWidth(26) = 1700  '매출금액   'H
         .ColWidth(27) = 5000  '적요
         .ColWidth(28) = 800   'SQL구분
         .ColWidth(29) = 1000  '입출고시간 'H
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "No"
         .TextMatrix(0, 1) = "사업장코드"   'H
         .TextMatrix(0, 2) = "거래일자"     'H
         .TextMatrix(0, 3) = "거래번호"     'H
         .TextMatrix(0, 4) = "코드"         '변경후(Or 변경전)
         .TextMatrix(0, 5) = "품명"         '변경후(Or 변경전)
         .TextMatrix(0, 6) = "자재코드"     'H, 변경전
         .TextMatrix(0, 7) = "품명"         'H, 변경전
         .TextMatrix(0, 8) = "매출처코드"   'H, 변경후
         .TextMatrix(0, 9) = "매출처명"     'H, 변경후
         .TextMatrix(0, 10) = "매출처코드"  'H, 변경전
         .TextMatrix(0, 11) = "매출처명"    'H, 변경전
         .TextMatrix(0, 12) = "KEY"         'H
         .TextMatrix(0, 13) = "매입처코드"  'H
         .TextMatrix(0, 14) = "매입처명"    'H
         .TextMatrix(0, 15) = "규격"
         .TextMatrix(0, 16) = "수량"
         .TextMatrix(0, 17) = "단위"
         .TextMatrix(0, 18) = "직송"        'H
         .TextMatrix(0, 19) = "매입단가"
         .TextMatrix(0, 20) = "매입부가"    'H
         .TextMatrix(0, 21) = "매입가격"    '(단가 + 부가) 'H
         .TextMatrix(0, 22) = "매입금액"
         .TextMatrix(0, 23) = "매출단가"    'H
         .TextMatrix(0, 24) = "매출부가"    'H
         .TextMatrix(0, 25) = "매출가격"    '(단가 + 부가) 'H
         .TextMatrix(0, 26) = "매출금액"    'H
         .TextMatrix(0, 27) = "적요"
         .TextMatrix(0, 28) = "구분"        'H(SQL구분:I.Insert, U.Update, D.Delete)
         .TextMatrix(0, 29) = "입출고시간"  'H
         .ColHidden(1) = True: .ColHidden(2) = True: .ColHidden(3) = True
         .ColHidden(6) = True: .ColHidden(7) = True
         .ColHidden(8) = True: .ColHidden(9) = True: .ColHidden(10) = True: .ColHidden(11) = True: .ColHidden(12) = True
         .ColHidden(13) = True: .ColHidden(14) = True: .ColHidden(18) = True: .ColHidden(20) = True:: .ColHidden(21) = True
         .ColHidden(23) = True: .ColHidden(24) = True: .ColHidden(25) = True: .ColHidden(26) = True
         .ColHidden(28) = True: .ColHidden(29) = True
         .ColFormat(16) = "#,#"
         For lngC = 19 To 26
             .ColFormat(lngC) = "#,#.00"
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 27
                         .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 0, 1, 2, 3, 18, 28, 29
                         .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                         .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictRows 'flexMergeFixedOnly
         .MergeRow(0) = True
         For lngC = 0 To 5
             .MergeCol(lngC) = True
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
    vsfg1.Rows = 1
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.사업장코드 AS 사업장코드, T1.거래일자 AS 거래일자, T1.거래번호 AS 거래번호, " _
                  & "T1.매입처코드 AS 매입처코드, ISNULL(T2.매입처명, '') AS 매입처명, " _
                  & "T1.현금구분 AS 현금구분, SUM(T1.출고수량 * T1.출고단가) AS 단가금액, SUM(T1.출고수량 * T1.출고부가) AS 부가금액 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 매입처 T2 ON T2.사업장코드 = T1.사업장코드 AND T2.매입처코드 = T1.매입처코드 " _
            & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.사용구분 = 0 AND T1.입출고구분 = 1 " _
              & "AND T1.입출고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
              & "AND T1.매입처코드 LIKE '%" & Trim(Text1(0).Text) & "%' " _
            & "GROUP BY T1.사업장코드, T1.거래일자, T1.거래번호, " _
                     & "T1.매입처코드, ISNULL(T2.매입처명, ''), T1.현금구분 " _
            & "ORDER BY T1.사업장코드, T1.거래일자, T1.거래번호 "
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
               .TextMatrix(lngR, 1) = Format(P_adoRec("거래일자"), "0000-00-00")
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("거래번호")), 0, P_adoRec("거래번호"))
               .TextMatrix(lngR, 3) = P_adoRec("사업장코드") & "-" & Format(P_adoRec("거래일자"), "0000/00/00") _
                                    & "-" & CStr(P_adoRec("거래번호"))
               .Cell(flexcpData, lngR, 3, lngR, 3) = Trim(.TextMatrix(lngR, 3)) 'FindRow 사용을 위해
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("매입처코드")), "", P_adoRec("매입처코드"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("매입처명")), "", P_adoRec("매입처명"))
               '금액은 vsfg2에서 계산
               '.TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("단가금액")), 0, P_adoRec("단가금액"))
               '.TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("부가금액")), 0, P_adoRec("부가금액"))
               '.TextMatrix(lngR, 8) = .ValueMatrix(lngR, 6) + .ValueMatrix(lngR, 7)
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("현금구분")), 0, P_adoRec("현금구분"))
               If P_adoRec("현금구분") = 1 Then
                  .Cell(flexcpChecked, lngR, 9) = flexChecked    '1
               Else
                  .Cell(flexcpChecked, lngR, 9) = flexUnchecked  '2
               End If
               .Cell(flexcpText, lngR, 9) = "현금매입"
               '.Cell(flexcpAlignment, lngR, 9, lngR, 9) = flexAlignCenterCenter
               .TextMatrix(lngR, 10) = Format(P_adoRec("거래일자"), "0000-00-00")
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
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
            .TopRow = 1           '무조건 TOP ROW를 첫행으로 이동
            vsfg1_EnterCell       'vsfg1_EnterCell 자동실행(만약 한건 일때도 강제로 자동실행)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "거래 읽기 실패"
    Unload Me
    Exit Sub
End Sub

Private Sub Subvsfg2_FILL()
Dim strSQL      As String
Dim strWhere    As String
Dim strOrderBy  As String
Dim lngR        As Long
Dim lngC        As Long
Dim lngRR       As Long
Dim lngRRR      As Long
Dim strCell     As String
Dim strSubTotal As String
    vsfg2.Rows = 1
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.사업장코드, T1.거래일자, T1.거래번호, T1.입출고시간, " _
                 & "(T1.분류코드 + T1.세부코드) AS 자재코드, ISNULL(T4.자재명,'') AS 자재명, " _
                 & "ISNULL(T1.매출처코드, '') AS 매출처코드, ISNULL(T3.매출처명,'') AS 매출처명, " _
                 & "T1.매입처코드 AS 매입처코드, T2.매입처명 AS 매입처명, T4.규격 AS 자재규격, T1.입고수량, " _
                 & "T4.단위 AS 자재단위, T1.직송구분 AS 직송구분, T1.입고단가, T1.입고부가, " _
                 & "T1.출고단가 , T1.출고부가, T1.적요 AS 적요 " _
            & "FROM 자재입출내역 T1 " _
            & "LEFT JOIN 매입처 T2 ON T2.사업장코드 = T1.사업장코드 AND T2.매입처코드 = T1.매입처코드 " _
            & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
            & "LEFT JOIN 자재 T4 ON (T4.분류코드 = T1.분류코드 AND T4.세부코드 = T1.세부코드) " _
           & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.사용구분 = 0 AND T1.입출고구분 = 1 " _
             & "AND T1.입출고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
           & "ORDER BY T1.사업장코드, T1.거래일자, T1.거래번호, T1.입출고시간 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg2.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       cmdExit.SetFocus
       Exit Sub
    Else
       With vsfg2
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt2 Then
               '.ScrollBars = flexScrollBarHorizontal
            Else
               '.ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               '.TextMatrix(lngR, 0) = Format(P_adoRec("XX"), "0000-00-00")
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("사업장코드")), "", P_adoRec("사업장코드"))
               .TextMatrix(lngR, 2) = Format(P_adoRec("거래일자"), "0000-00-00")
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("거래번호")), 0, P_adoRec("거래번호"))
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("자재코드")), "", P_adoRec("자재코드"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("자재명")), "", P_adoRec("자재명"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("자재코드")), "", P_adoRec("자재코드"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("자재명")), "", P_adoRec("자재명"))
               .TextMatrix(lngR, 8) = IIf(IsNull(P_adoRec("매출처코드")), "", P_adoRec("매출처코드"))
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("매출처명")), "", P_adoRec("매출처명"))
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("매출처코드")), "", P_adoRec("매출처코드"))
               .TextMatrix(lngR, 11) = IIf(IsNull(P_adoRec("매출처명")), "", P_adoRec("매출처명"))
               .TextMatrix(lngR, 12) = P_adoRec("사업장코드") & "-" & Format(P_adoRec("거래일자"), "0000/00/00") _
                                     & "-" & CStr(P_adoRec("거래번호")) & "-" & P_adoRec("입출고시간") _
                                     & "-" & P_adoRec("자재코드") & "-" & P_adoRec("매출처코드")
               .TextMatrix(lngR, 13) = IIf(IsNull(P_adoRec("매입처코드")), "", P_adoRec("매입처코드"))
               .TextMatrix(lngR, 14) = IIf(IsNull(P_adoRec("매입처명")), "", P_adoRec("매입처명"))
               .TextMatrix(lngR, 15) = IIf(IsNull(P_adoRec("자재규격")), "", P_adoRec("자재규격"))
               .TextMatrix(lngR, 16) = Format(P_adoRec("입고수량"), "#,#")
               .TextMatrix(lngR, 17) = IIf(IsNull(P_adoRec("자재단위")), "", P_adoRec("자재단위"))
               If P_adoRec("직송구분") = 0 Then
                  .Cell(flexcpChecked, lngR, 18) = flexUnchecked
               Else
                  .Cell(flexcpChecked, lngR, 18) = flexChecked
               End If
               .Cell(flexcpText, lngR, 18) = "직 송"
               .Cell(flexcpAlignment, lngR, 18, lngR, 18) = flexAlignLeftCenter
               .TextMatrix(lngR, 19) = Format(P_adoRec("입고단가"), "#,#.00")
               .TextMatrix(lngR, 20) = Format(P_adoRec("입고부가"), "#,#.00")
               .TextMatrix(lngR, 21) = .ValueMatrix(lngR, 19) + .ValueMatrix(lngR, 20) '입고단가 + 입고부가
               .TextMatrix(lngR, 22) = .ValueMatrix(lngR, 16) * .ValueMatrix(lngR, 19) '입고금액(수량*단가)
               .TextMatrix(lngR, 23) = Format(P_adoRec("출고단가"), "#,#.00")
               .TextMatrix(lngR, 24) = Format(P_adoRec("출고부가"), "#,#.00")
               .TextMatrix(lngR, 25) = .ValueMatrix(lngR, 23) + .ValueMatrix(lngR, 24) '출고단가 + 출고부가
               .TextMatrix(lngR, 26) = .ValueMatrix(lngR, 16) * .ValueMatrix(lngR, 23) '출고금액(수량*단가)
               .TextMatrix(lngR, 27) = IIf(IsNull(P_adoRec("적요")), "", P_adoRec("적요"))
               .TextMatrix(lngR, 29) = IIf(IsNull(P_adoRec("입출고시간")), "", P_adoRec("입출고시간"))
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserBranchCode Then
               '   lngRR = lngR
               'End If
               .RowHidden(lngR) = True
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg2_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt2 Then
                  '.TopRow = .Rows - PC_intRowCnt2 + .FixedRows
                  .TopRow = 1
               End If
            Else
               .Row = lngRR       'vsfg2_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt2 Then
                  .TopRow = .Row
               End If
            End If
            '.MultiTotals = True '(default value : true)
            '.Subtotal flexSTClear
            '.SubtotalPosition = flexSTBelow
            '.Subtotal flexSTCount, 6, 8, "#", vbRed, vbWhite, , "%s", , False
            '.Subtotal flexSTSum, 6, 10, , vbRed, vbWhite, , "%s", , False
            For lngR = 1 To .Rows - 1
                strCell = .TextMatrix(lngR, 1) & "-" & Format(DTOS(.TextMatrix(lngR, 2)), "0000/00/00") & "-" & .TextMatrix(lngR, 3)
                For lngRRR = 1 To vsfg1.Rows - 1
                    If strCell = vsfg1.TextMatrix(lngRRR, 3) Then
                       If vsfg1.ValueMatrix(lngRRR, 14) = 0 Then
                          vsfg1.TextMatrix(lngRRR, 14) = lngR
                       End If
                       vsfg1.TextMatrix(lngRRR, 15) = lngR
                       '거래 합계금액 계산
                       vsfg1.TextMatrix(lngRRR, 6) = vsfg1.ValueMatrix(lngRRR, 6) + .ValueMatrix(lngR, 22)
                       lblTotMny.Caption = Format(Vals(lblTotMny.Caption) + .ValueMatrix(lngR, 22), "#,#.00")
                       Exit For
                    End If
                Next lngRRR
            Next lngR
            'vsfg2_EnterCell       'vsfg1_EnterCell 자동실행(만약 한건 일때도 강제로 자동실행) 'Not Used
            '.SetFocus                                                                         'Not Used
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "거래내역 읽기 실패"
    Unload Me
    Exit Sub
End Sub

