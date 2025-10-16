VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frm내역시세검색 
   Appearance      =   0  '평면
   BackColor       =   &H00008000&
   BorderStyle     =   1  '단일 고정
   Caption         =   "내역 단가 검색"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "내역시세검색.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3720
   StartUpPosition =   2  '화면 가운데
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   2385
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   3495
      _cx             =   6165
      _cy             =   4207
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorFixed  =   -2147483638
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483638
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
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
Attribute VB_Name = "frm내역시세검색"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 내역시세검색
' 사용된 Control :
' 참조된 Table   :
' 업  무  설  명 :
'
'+-------------------------------------------------------------------------------------------------------+
Private P_blnActived       As Boolean
Private P_intFindGbn       As Integer      '0.수동검색, 1.코드, 2.이름(명), 3.코드 + 이름(명)
Private P_intButton        As Integer
Private P_strFindString2   As String
Private P_strFirstCode     As String
Private P_strAppDate       As String       '적용일자
Private P_blnSelect        As Boolean
Private P_adoRec           As New ADODB.Recordset
Private Const PC_intRowCnt As Integer = 10  '그리드 한 페이지 당 행수(FixedRows 포함)

'+--------------------------------+
'/// LOAD FORM ( 한번만 실행 ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
    P_strFirstCode = PB_strMaterialsCode '최초 찾고자한 자재코드
    If (PB_strFMCCallFormName = "frm발주서작성") Then
       With frm발주서작성
            P_strAppDate = PB_regUserinfoU.UserClientDate
       End With
    End If
    If (PB_strFMCCallFormName = "frm발주서관리") Then
       With frm발주서관리
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매입작성1") Then
       With frm매입작성1
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매입작성2") Then
       With frm매입작성2
            P_strAppDate = DTOS(.dtpJ_Date.Value)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매입수정") Then
       With frm매입수정
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    
    If (PB_strFMCCallFormName = "frm견적서작성") Then
       With frm견적서작성
            P_strAppDate = PB_regUserinfoU.UserClientDate
       End With
    End If
    If (PB_strFMCCallFormName = "frm견적서관리") Then
       With frm견적서관리
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매출작성1") Then
       With frm매출작성1
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매출작성2") Then
       With frm매출작성2
            P_strAppDate = DTOS(.dtpJ_Date.Value)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매출수정") Then
       With frm매출수정
            P_strAppDate = .vsfg2.TextMatrix(.vsfg2.Row, 2)
       End With
    End If
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
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재시세 (서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+------------+
'/// vsfg1 ///
'+------------+
Private Sub vsfg1_DblClick()
    vsfg1_KeyDown vbKeyReturn, 0
End Sub
Private Sub vsfg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
       SubEscape
    End If
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
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         If .Row >= .FixedRows Then
         End If
    End With
End Sub
'+---------------------------------------+
'/// 시세 검색후 Return
'+---------------------------------------+
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL         As String
Dim intRetVal   As Integer
Dim lngR        As Long
Dim blnDupOK    As Boolean
Dim varFormName As Variant
    With vsfg1
         If KeyCode = vbKeyReturn Then
            If .Row <= 0 Then
               PB_strMaterialsCode = ""
               PB_strMaterialsName = ""
            Else
               PB_strMaterialsCode = .TextMatrix(.Row, 0)
               PB_strMaterialsName = .TextMatrix(.Row, 1)
            End If
            varFormName = PB_strFMCCallFormName
            If .Row >= .FixedRows Then
               If (PB_strFMCCallFormName = "frm발주서작성") Then
                  With frm발주서작성
                       .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '단가
                       .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '부가
                       .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 7)    '가격
                       '합계금액 계산(금액 다르면)
                       If .vsfg1.ValueMatrix(.vsfg1.Row, 11) <> _
                          (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                          .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 11) _
                                             + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                        End If
                        .vsfg1.TextMatrix(.vsfg1.Row, 11) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                          * vsfg1.ValueMatrix(vsfg1.Row, 5)   '입고금액
                        P_blnSelect = True
                        PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm발주서관리") Then
                  With frm발주서관리
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '단가 다르면
                           .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '가격
                          '합계금액 계산(금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '입고금액
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매입작성1") Then
                  With frm매입작성1
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '단가 다르면
                           .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '가격
                          '합계금액 계산(금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)     '금액
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매입작성2") Then
                  With frm매입작성2
                       If (.vsfg1.TextMatrix(.vsfg1.Row, 8) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then   '단가 다르면
                          .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 5)          '단가
                          .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 6)          '부가
                          .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 7)         '가격
                          '합계금액 계산(금액이 다르면)
                          If .vsfg1.ValueMatrix(.vsfg1.Row, 11) <> _
                            (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 11) _
                                                + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg1.TextMatrix(.vsfg1.Row, 11) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '입고금액
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매입수정") Then
                  With frm매입수정
                       If (.vsfg2.TextMatrix(.vsfg2.Row, 19) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '단가 다르면
                          .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 5)   '단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 6)   '부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 7)   '가격
                          '합계금액 계산(금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 6) = .vsfg1.ValueMatrix(.vsfg1.Row, 6) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '금액
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm견적서작성") Then
                  With frm견적서작성
                       .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '단가
                       .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '부가
                       .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '가격
                       '합계금액 계산(금액 다르면)
                       If .vsfg1.ValueMatrix(.vsfg1.Row, 15) <> _
                          (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                          .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 15) _
                                             + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                       End If
                       .vsfg1.TextMatrix(.vsfg1.Row, 15) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                         * vsfg1.ValueMatrix(vsfg1.Row, 5)
                       P_blnSelect = True
                       PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm견적서관리") Then
                  With frm견적서관리
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '단가 다르면
                           .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '가격
                          '합계금액 계산(금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)     '금액
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매출작성1") Then
                  With frm매출작성1관리
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '단가 다르면
                           .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 5)     '단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 6)     '부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 7)     '가격
                          '합계금액 계산(금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)     '금액
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매출작성2") Then
                  With frm매출작성2
                       If (.vsfg1.TextMatrix(.vsfg1.Row, 12) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then
                          .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 5)   '단가
                          .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 6)   '부가
                          .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 7)   '가격
                          '합계금액 계산(금액이 다르면)
                          If .vsfg1.ValueMatrix(.vsfg1.Row, 15) <> _
                            (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 15) _
                                               + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg1.TextMatrix(.vsfg1.Row, 15) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '금액
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매출수정") Then
                  With frm매출수정
                       If (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.TextMatrix(vsfg1.Row, 5)) Then '단가 다르면
                          .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                           .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 5)   '단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 6)   '부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 7)   '가격
                          '합계금액 계산(금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)) Then
                            .vsfg1.TextMatrix(.vsfg1.Row, 6) = .vsfg1.ValueMatrix(.vsfg1.Row, 6) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5))
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                               + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 5)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 5)   '금액
                          P_blnSelect = True
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               End If
            End If
            Unload Me
         End If
    End With
    Exit Sub
End Sub

'+-----------+
'/// 조회 ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    'P_strFindString2 = Trim(Text1(1).Text)  '조회할 경우 검색할 자재명 보관
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

'+-----------+
'/// 종료 ///
'+-----------+
Private Sub Form_Unload(Cancel As Integer)
    If P_blnSelect = False Then
       SubEscape
    End If
    Screen.MousePointer = vbDefault
    If P_adoRec.State <> adStateClosed Then
       P_adoRec.Close
    End If
    Set P_adoRec = Nothing
    Set frm내역시세검색 = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfgGrid) 초기화 ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    With vsfg1           'Rows 1, Cols 8, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         '.BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .Ellipsis = flexEllipsisEnd
         '.ExplorerBar = flexExSortShow
         .ExtendLastCol = True
         .FocusRect = flexFocusHeavy
         .FontSize = 9
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 0
         .Rows = 1             'SubvsfgUpGrid_Fill수행시에 설정
         .Cols = 8
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1900   '자재코드(분류코드+세부코드) 'H
         .ColWidth(1) = 2500   '자재명                      'H
         .ColWidth(2) = 2200   '규격                        'H
         .ColWidth(3) = 900    '단위                        'H
         .ColWidth(4) = 1200   '일자
         .ColWidth(5) = 1500   '단가
         .ColWidth(6) = 1200   '부가                        'H
         .ColWidth(7) = 1500   '가격                        'H
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "코드"
         .TextMatrix(0, 1) = "품명"
         .TextMatrix(0, 2) = "규격"
         .TextMatrix(0, 3) = "단위"
         
         If (PB_strFMCCallFormName = "frm발주서작성") Or (PB_strFMCCallFormName = "frm발주서관리") Then
            .TextMatrix(0, 4) = "발주일자"
            .TextMatrix(0, 5) = "발주단가"
            .TextMatrix(0, 6) = "발주부가"
            .TextMatrix(0, 7) = "발주가격"
         ElseIf _
            (PB_strFMCCallFormName = "frm매입작성1") Or (PB_strFMCCallFormName = "frm매입작성2") Or _
            (PB_strFMCCallFormName = "frm매입수정") Then
            .TextMatrix(0, 4) = "매입일자"
            .TextMatrix(0, 5) = "매입단가"
            .TextMatrix(0, 6) = "매입부가"
            .TextMatrix(0, 7) = "매입가격"
         ElseIf _
            (PB_strFMCCallFormName = "frm견적서작성") Or (PB_strFMCCallFormName = "frm견적서관리") Then
            .TextMatrix(0, 4) = "견적일자"
            .TextMatrix(0, 5) = "견적단가"
            .TextMatrix(0, 6) = "견적부가"
            .TextMatrix(0, 7) = "견적가격"
         ElseIf _
            (PB_strFMCCallFormName = "frm매출작성1") Or (PB_strFMCCallFormName = "frm매출작성2") Or _
            (PB_strFMCCallFormName = "frm매출수정") Then
            .TextMatrix(0, 4) = "매출일자"
            .TextMatrix(0, 5) = "매출단가"
            .TextMatrix(0, 6) = "매출부가"
            .TextMatrix(0, 7) = "매출가격"
         End If
         .ColHidden(0) = True: .ColHidden(1) = True: .ColHidden(2) = True:: .ColHidden(3) = True
         .ColHidden(6) = True: .ColHidden(7) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 5 To 7
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 1, 2, 3
                        .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 4
                        .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                        .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictAll
         'For lngC = 0 To 5
         '    .MergeCol(lngC) = True
         'Next lngC
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfgGrid) 채우기///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL     As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
Dim strAppDate As Long
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    If (PB_strFMCCallFormName = "frm발주서작성") Or (PB_strFMCCallFormName = "frm발주서관리") Then
       strSQL = "SELECT TOP 5 " _
                     & "T1.자재코드 AS 자재코드, J1.자재명 AS 품명, " _
                     & "J1.규격 AS 규격, J1.단위 AS 단위, T1.발주일자 AS 일자, " _
                     & "T1.입고단가 AS 단가, T1.입고부가 AS 부가 " _
                & "FROM 발주내역 T1 " _
                & "LEFT JOIN 자재 J1 ON (J1.분류코드 + J1.세부코드) = T1.자재코드 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.발주일자 <= '" & P_strAppDate & "' " _
                 & "AND T1.매입처코드 = '" & PB_strSupplierCode & "' " _
                 & "AND T1.자재코드 = '" & PB_strMaterialsCode & "' " _
                 & "AND T1.사용구분 = 0 AND T1.입고단가 > 0 " _
               & "GROUP BY T1.자재코드, J1.자재명, J1.규격, J1.단위, T1.발주일자, T1.입고단가, T1.입고부가 " _
               & "ORDER BY T1.발주일자 DESC "
    End If
    If (PB_strFMCCallFormName = "frm매입작성1") Or (PB_strFMCCallFormName = "frm매입작성2") Or _
       (PB_strFMCCallFormName = "frm매입수정") Then
       strSQL = "SELECT TOP 5 " _
                     & "(T1.분류코드 + T1.세부코드) AS 자재코드, J1.자재명 AS 품명, " _
                     & "J1.규격 AS 규격, J1.단위 AS 단위, T1.입출고일자 AS 일자, " _
                     & "T1.입고단가 AS 단가, T1.입고부가 AS 부가 " _
                & "FROM 자재입출내역 T1 " _
                & "LEFT JOIN 자재 J1 ON J1.분류코드 = T1.분류코드 AND J1.세부코드 = T1.세부코드 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.입출고일자 <= '" & P_strAppDate & "' " _
                 & "AND T1.매입처코드 = '" & PB_strSupplierCode & "' " _
                 & "AND T1.분류코드 = '" & Mid(PB_strMaterialsCode, 1, 2) & "' " _
                 & "AND T1.세부코드 = '" & Mid(PB_strMaterialsCode, 3) & "' " _
                 & "AND T1.사용구분 = 0 AND T1.입출고구분 = 1 AND T1.입고단가 > 0 " _
               & "GROUP BY (T1.분류코드 + T1.세부코드), J1.자재명, J1.규격, J1.단위, T1.입출고일자, T1.입고단가, T1.입고부가 " _
               & "ORDER BY T1.입출고일자 DESC "
    End If
    
    If (PB_strFMCCallFormName = "frm견적서작성") Or (PB_strFMCCallFormName = "frm견적서관리") Then
       strSQL = "SELECT TOP 5 " _
                     & "T1.자재코드 AS 자재코드, J1.자재명 AS 품명, " _
                     & "J1.규격 AS 규격, J1.단위 AS 단위, T1.견적일자 AS 일자, " _
                     & "T1.출고단가 AS 단가, T1.출고부가 AS 부가 " _
                & "FROM 견적내역 T1 " _
                & "LEFT JOIN 자재 J1 ON (J1.분류코드 + J1.세부코드) = T1.자재코드 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.견적일자 <= '" & P_strAppDate & "' " _
                 & "AND T1.매출처코드 = '" & PB_strSupplierCode & "' " _
                 & "AND T1.자재코드 = '" & PB_strMaterialsCode & "' " _
                 & "AND T1.사용구분 = 0 AND T1.출고단가 > 0 " _
               & "GROUP BY T1.자재코드, J1.자재명, J1.규격, J1.단위, T1.견적일자, T1.출고단가, T1.출고부가 " _
               & "ORDER BY T1.견적일자 DESC "
    End If
    If (PB_strFMCCallFormName = "frm매출작성1") Or (PB_strFMCCallFormName = "frm매출작성2") Or _
       (PB_strFMCCallFormName = "frm매출수정") Then
       strSQL = "SELECT TOP 5 " _
                     & "(T1.분류코드 + T1.세부코드) AS 자재코드, J1.자재명 AS 품명, " _
                     & "J1.규격 AS 규격, J1.단위 AS 단위, T1.입출고일자 AS 일자, " _
                     & "T1.출고단가 AS 단가, T1.출고부가 AS 부가 " _
                & "FROM 자재입출내역 T1 " _
                & "LEFT JOIN 자재 J1 ON J1.분류코드 = T1.분류코드 AND J1.세부코드 = T1.세부코드 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND T1.입출고일자 <= '" & P_strAppDate & "' " _
                 & "AND T1.매출처코드 = '" & PB_strSupplierCode & "' " _
                 & "AND T1.분류코드 = '" & Mid(PB_strMaterialsCode, 1, 2) & "' " _
                 & "AND T1.세부코드 = '" & Mid(PB_strMaterialsCode, 3) & "' " _
                 & "AND T1.사용구분 = 0 AND T1.입출고구분 = 2 AND T1.출고단가 > 0 " _
               & "GROUP BY (T1.분류코드 + T1.세부코드), J1.자재명, J1.규격, J1.단위, T1.입출고일자, T1.출고단가, T1.출고부가 " _
               & "ORDER BY T1.입출고일자 DESC "
    End If
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + vsfg1.FixedRows
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = vbBlack
            .Rows = P_adoRec.RecordCount + 1
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("자재코드")
               .TextMatrix(lngR, 1) = P_adoRec("품명")
               .TextMatrix(lngR, 2) = P_adoRec("규격")
               .TextMatrix(lngR, 3) = P_adoRec("단위")
               .TextMatrix(lngR, 4) = Format(P_adoRec("일자"), "0000-00-00")
               .TextMatrix(lngR, 5) = P_adoRec("단가")
               .Cell(flexcpData, lngR, 4, lngR, 4) = .TextMatrix(lngR, 4) + "-" + .TextMatrix(lngR, 5)
               .TextMatrix(lngR, 6) = P_adoRec("부가")
               .TextMatrix(lngR, 7) = .TextMatrix(lngR, 5) + .TextMatrix(lngR, 6)
               'If PB_strMaterialsCode = Trim(.TextMatrix(lngR, 0)) Then
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
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "자재시세 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+----------+
'/// ESC ///
'+----------+
Private Sub SubEscape()
    PB_strFMCCallFormName = ""
    PB_strMaterialsCode = ""
    PB_strMaterialsName = ""
    PB_strSupplierCode = ""
    Unload Me
End Sub

