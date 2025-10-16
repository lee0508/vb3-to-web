VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm계정코드관리 
   BorderStyle     =   0  '없음
   Caption         =   "제경비코드"
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10100
   ScaleMode       =   0  '사용자
   ScaleWidth      =   15405
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   13
      Top             =   660
      Width           =   15195
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   2
         Left            =   7755
         MaxLength       =   1
         TabIndex        =   2
         Top             =   225
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   3
         Left            =   7755
         MaxLength       =   60
         TabIndex        =   3
         Top             =   593
         Width           =   6975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   1755
         MaxLength       =   4
         TabIndex        =   0
         Top             =   225
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   10  '한글 
         Index           =   1
         Left            =   1755
         MaxLength       =   30
         TabIndex        =   1
         Top             =   593
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   285
         IMEMode         =   8  '영문
         Index           =   4
         Left            =   7755
         MaxLength       =   1
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "합계시산표연결여부"
         Height          =   240
         Index           =   3
         Left            =   5640
         TabIndex        =   21
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   3000
         TabIndex        =   20
         Top             =   265
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "(0.정상, 9.사용불가)"
         Height          =   240
         Index           =   18
         Left            =   8640
         TabIndex        =   19
         Top             =   1005
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "적요"
         Height          =   240
         Index           =   2
         Left            =   6510
         TabIndex        =   18
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "계정코드"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   510
         TabIndex        =   16
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "계정명"
         Height          =   240
         Index           =   1
         Left            =   510
         TabIndex        =   15
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "사용구분"
         Height          =   240
         Index           =   7
         Left            =   6510
         TabIndex        =   14
         ToolTipText     =   "0.정상, 기타.사용불가"
         Top             =   1005
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton optPrtChk0 
         Caption         =   "코드순"
         Height          =   255
         Left            =   6840
         TabIndex        =   23
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton optPrtChk1 
         Caption         =   "이름순"
         Height          =   255
         Left            =   6840
         TabIndex        =   22
         Top             =   390
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   390
         Left            =   7980
         Picture         =   "계정코드관리.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   390
         Left            =   9120
         Picture         =   "계정코드관리.frx":0963
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "계정코드관리.frx":1308
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "계정코드관리.frx":1C56
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "계정코드관리.frx":25DA
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "계정코드관리.frx":2E61
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   195
         Width           =   1095
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "계 정 코 드 관 리"
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
         TabIndex        =   8
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   7886
      Left            =   60
      TabIndex        =   6
      Top             =   2100
      Width           =   15195
      _cx             =   26802
      _cy             =   13910
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777215
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
End
Attribute VB_Name = "frm계정코드관리"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 계정코드관리
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 계정과목
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 25  '그리드 한 페이지 당 행수(FixedRows 포함)

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
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계정코드관리(서버와의 연결 실패)"
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
    If (Index = 0 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '계정코드검색
       PB_strAccCode = UPPER(Trim(Text1(Index).Text))
       PB_strAccName = ""  'Trim(Text1(Index + 1).Text)
       frm계정코드검색.Show vbModal
       If (Len(PB_strAccCode) + Len(PB_strAccName)) = 0 Then '검색에서 취소(ESC)
       Else
          Text1(Index).Text = PB_strAccCode
          Text1(Index + 1).Text = PB_strAccName
       End If
       If PB_strAccCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strAccCode = "": PB_strAccName = ""
    ElseIf _
       KeyCode = vbKeyDelete Then
       If Len(Text1(0).Text) = 0 Then
          Text1(1).Text = ""
          Exit Sub
       End If
    Else
       If KeyCode = vbKeyReturn Then
          Select Case Index
                 Case Text1.UBound
                      If cmdSave.Enabled = True And vsfg1.Rows > 1 Then
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계정코드 읽기 실패"
    Unload Me
    Exit Sub
End Sub
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         .Text = Trim(.Text)
         Select Case Index
                Case 0  '계정코드
                     .Text = UPPER(.Text)
                     If Len(.Text) < 1 Then
                        Text1(1).Text = ""
                     End If
                Case 2  '합계시산표연결여부
                     .Text = UPPER(.Text)
                     If .Text <> "N" Then
                        .Text = "Y"
                     End If
                Case 4  '사용구분
                     If .Text <> "0" Then
                        .Text = "9"
                     End If
                Case Else
                     .Text = Trim(.Text)
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계정코드 읽기 실패"
    Unload Me
    Exit Sub
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
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         If .Row < .FixedRows Then
            Text1(Text1.LBound).Enabled = True
         Else
            Text1(Text1.LBound).Enabled = False
            For lngC = 0 To .Cols - 1
                Select Case lngC
                       Case Is <= 4
                            Text1(lngC).Text = .TextMatrix(.Row, lngC)
                End Select
            Next lngC
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
    '
End Sub

'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
    SubClearText
    vsfg1.Row = 0
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
         If Text1(Text1.LBound).Enabled = True Then '계정코드정보 추가면 검색
            strSQL = "SELECT * FROM 계정과목 " _
                    & "WHERE 계정코드 = '" & Trim(Text1(Text1.LBound).Text) & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount <> 0 Then
               P_adoRec.Close
               Text1(Text1.LBound).SetFocus
               cmdSave.Enabled = True
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            P_adoRec.Close
         End If
         PB_adoCnnSQL.BeginTrans
         If Text1(Text1.LBound).Enabled = True Then '계정코드 추가
            strSQL = "INSERT INTO 계정과목(계정코드, 계정명, 합계시산표연결여부, 적요, " _
                                        & "사용구분, 수정일자, 사용자코드 ) Values( " _
                    & "'" & Trim(Text1(0).Text) & "','" & Trim(Text1(1).Text) & "', " _
                    & "'" & Trim(Text1(2).Text) & "', '" & Trim(Text1(3).Text) & "', " _
                    & "" & Val(Trim(Text1(4).Text)) & ", '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' )"
            On Error GoTo ERROR_TABLE_INSERT
            .AddItem .Rows
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 4
                            .TextMatrix(.Rows - 1, lngC) = Val(Trim(Text1(lngC).Text))
                            Select Case Val(Text1(lngC).Text)
                                   Case 0
                                        .TextMatrix(.Rows - 1, lngC + 1) = "정상"
                                   Case 9
                                        .TextMatrix(.Rows - 1, lngC + 1) = "사용불가"
                                   Case Else
                                        .TextMatrix(.Rows - 1, lngC + 1) = "구분오류"
                            End Select
                       Case Else
                            .TextMatrix(.Rows - 1, lngC) = Text1(lngC).Text
                            If lngC = 0 Then .Cell(flexcpData, .Rows - 1, lngC, .Rows - 1, lngC) = Text1(lngC).Text
                End Select
            Next lngC
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(Text1.LBound).Enabled = False
            .Row = .Rows - 1          '자동으로 vsfg1_EnterCell Event 발생
         Else                                          '계정코드 변경
            strSQL = "UPDATE 계정과목 SET " _
                          & "계정명 = '" & Trim(Text1(1).Text) & "', " _
                          & "합계시산표연결여부 = '" & Trim(Text1(2).Text) & "', " _
                          & "적요 = '" & Trim(Text1(3).Text) & "', " _
                          & "사용구분 = " & Val(Text1(4).Text) & ", " _
                          & "수정일자 = '" & PB_regUserinfoU.UserServerDate & "', " _
                          & "사용자코드 = '" & PB_regUserinfoU.UserCode & "' " _
                    & "WHERE 계정코드 = '" & Trim(Text1(Text1.LBound).Text) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 4
                            .TextMatrix(.Row, lngC) = Vals(Trim(Text1(lngC).Text))
                            Select Case Val(Text1(lngC).Text)
                                   Case 0
                                        .TextMatrix(.Row, lngC + 1) = "정상"
                                   Case 9
                                        .TextMatrix(.Row, lngC + 1) = "사용불가"
                                   Case Else
                                        .TextMatrix(.Row, lngC + 1) = "구분오류"
                            End Select
                       Case Else
                            .TextMatrix(.Row, lngC) = Text1(lngC).Text
                End Select
            Next lngC
         End If
         PB_adoCnnSQL.Execute strSQL
         PB_adoCnnSQL.CommitTrans
         .SetFocus
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계정과목 읽기 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계정과목 추가 실패"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계정과목 변경 실패"
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
               'strSQL = "SELECT Count(*) AS 해당건수 FROM 제경비내역 " _
               '        & "WHERE 제경비코드 = " & .TextMatrix(.Row, 0) & " "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("해당건수")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "건)가 있으므로 삭제할 수 없습니다.", vbCritical, "계정코드 삭제 불가"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "DELETE FROM 계정과목 WHERE 계정코드 = " & .TextMatrix(.Row, 0) & " "
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
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계정코드 삭제 실패"
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
    Set frm계정코드관리 = Nothing
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
    Text1(Text1.LBound).Enabled = False                '계정코드 FLASE
    With vsfg1              'Rows 0, Cols 6, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarVertical
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 1
         .Rows = 1             'Subvsfg1_Fill수행시에 설정
         .Cols = 6
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1000   '계정코드
         .ColWidth(1) = 3500   '계정명
         .ColWidth(2) = 2000   '합계시산표연결여부
         .ColWidth(3) = 6000   '적요
         .ColWidth(4) = 1      '사용구분
         .ColWidth(5) = 1000   '사용구분
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "계정코드"
         .TextMatrix(0, 1) = "계정명"
         .TextMatrix(0, 2) = "합계시산표연결여부"
         .TextMatrix(0, 3) = "적요"
         .TextMatrix(0, 4) = "사용구분"
         .TextMatrix(0, 5) = "사용구분"
         .ColHidden(4) = True
         If PB_regUserinfoU.UserAuthority <> "99" Then
            '.ColHidden(4) = True: .ColHidden(5) = True
         End If
         .ColAlignment(0) = flexAlignCenterCenter
         .ColAlignment(1) = flexAlignLeftCenter
         .ColAlignment(2) = flexAlignCenterCenter
         .ColAlignment(3) = flexAlignLeftCenter
         For lngC = 4 To 5
             .ColAlignment(lngC) = flexAlignCenterCenter
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
          & "FROM 계정과목 T1 " _
         & "ORDER BY T1.계정코드 "
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
               .TextMatrix(lngR, 0) = P_adoRec("계정코드")
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow 사용을 위해
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("계정명")), "", P_adoRec("계정명"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("합계시산표연결여부")), "", P_adoRec("합계시산표연결여부"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("적요")), "", P_adoRec("적요"))
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("사용구분")), "", P_adoRec("사용구분"))
               Select Case .ValueMatrix(lngR, 4)
                      Case 0
                           .TextMatrix(lngR, 5) = "정상"
                      Case 9
                           .TextMatrix(lngR, 5) = "사용불가"
                      Case Else
                           .TextMatrix(lngR, 5) = "구분오류"
               End Select
               'If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserCode Then
               '   lngRR = lngR
               'End If
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Rows - PC_intRowCnt + .FixedRows
                  .TopRow = 1
               End If
               Text1(Text1.LBound).Enabled = True
               Text1(Text1.LBound).SetFocus
               Exit Sub
            Else
               .Row = lngRR       'vsfg1_EnterCell 자동실행(만약 한건 일때는 자동실행 안함)
               If .Rows > PC_intRowCnt Then
                  '.TopRow = .Row
               End If
            End If
            vsfg1_EnterCell       'vsfg1_EnterCell 자동실행(만약 한건 일때도 강제로 자동실행)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계정과목 읽기 실패"
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
               Case 0
                    If Len(Text1(lngC).Text) < 1 Then
                       Exit Function
                    End If
               Case 1  '계정명
                    If LenH(Trim(Text1(lngC).Text)) > 30 Then
                       Exit Function
                    End If
               Case 2  '합계시산표연결여부
                    If Not (Text1(lngC).Text = "Y" Or Text1(lngC).Text = "N") Then
                       Exit Function
                    End If
               Case 3 '적요
                    If LenH(Trim(Text1(lngC).Text)) > 60 Then
                       Exit Function
                    End If
               Case 4  '사용구분
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0")
                    If Not (Text1(lngC).Text = "0" Or Text1(lngC).Text = "9") Then
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
            strExeFile = App.Path & ".\계정코드관리보고서.rpt"
         Else
            strExeFile = App.Path & ".\계정코드관리보고서.rpt"
         End If
         varRetVal = Dir(strExeFile)
         If Len(varRetVal) = 0 Then
            intRetCHK = 0
         Else
            .ReportFileName = strExeFile
            On Error GoTo ERROR_CRYSTAL_REPORTS
            '--- Formula Fields ---
            .Formulas(0) = "ForBranchName= '" & PB_regUserinfoU.UserBranchName & "'"                       '사업장명
            .Formulas(1) = "ForPrtDateTime = '" & strForPrtDateTime & "'"                                  '출력일시
            .Formulas(2) = "ForSortGbn = 1"                                          '정렬순서(0.계정코드, 1.계정명)
            '--- Parameter Fields ---
            '프로그램실행일자
            .StoredProcParam(0) = PB_regUserinfoU.UserBranchCode '사업장코드
            .StoredProcParam(1) = 0                              '사용구분                                 'Mot Used
            .StoredProcParam(2) = IIf(optPrtChk0.Value = True, 0, 1)                 '정렬순서(0.계정코드, 1.계정명)
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
            .WindowTitle = PB_regUserinfoU.UserBranchName & " - " & "계정코드관리보고서"
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


