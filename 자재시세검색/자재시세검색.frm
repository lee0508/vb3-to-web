VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frm자재시세검색 
   Appearance      =   0  '평면
   BackColor       =   &H00008000&
   BorderStyle     =   1  '단일 고정
   Caption         =   "자재 시세 검색"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "자재시세검색.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   13320
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   13095
      Begin VB.TextBox txtUnit 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   1  '입력 상태 설정
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "추가(&A)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11160
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtSize 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         MaxLength       =   30
         TabIndex        =   2
         Top             =   240
         Width           =   2355
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "조회(&F)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11160
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtBarCode 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   4680
         MaxLength       =   8
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         MaxLength       =   30
         TabIndex        =   1
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   1320
         MaxLength       =   18
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "규격"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   12
         Left            =   7440
         TabIndex        =   14
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "단위"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "바코드"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   8
         Left            =   3360
         TabIndex        =   12
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "품명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   7
         Left            =   3600
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   3240
         TabIndex        =   10
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "코드"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   3345
      Left            =   120
      TabIndex        =   6
      Top             =   1155
      Width           =   13095
      _cx             =   23098
      _cy             =   5900
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
Attribute VB_Name = "frm자재시세검색"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 자재시세검색
' 사용된 Control :
' 참조된 Table   : 자재, 자재원장
' 업  무  설  명 :
'
'+-------------------------------------------------------------------------------------------------------+
Private P_blnActived       As Boolean
Private P_intFindGbn       As Integer      '0.수동검색, 1.코드, 2.이름(명), 3.코드 + 이름(명)
Private P_intButton        As Integer
Private P_strFindString2   As String
Private P_strFirstCode     As String
Private P_intIOGbn         As Integer      '입출고구분 : 1.입고, 2.출고
Private P_blnSelect        As Boolean
Private P_adoRec           As New ADODB.Recordset
Private Const PC_intRowCnt As Integer = 10  '그리드 한 페이지 당 행수(FixedRows 포함)

'+--------------------------------+
'/// LOAD FORM ( 한번만 실행 ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
    P_strFirstCode = PB_strMaterialsCode '최초 찾고자한 자재코드
    txtCode.Text = PB_strMaterialsCode: txtName.Text = PB_strMaterialsName
    If (PB_strFMCCallFormName = "frm발주서작성") Then
       With frm발주서작성
            'txtSize.Text = .vsfg1.TextMatrix(.vsfg1.Row, 2)
            'txtUnit.Text = .vsfg1.TextMatrix(.vsfg1.Row, 4)
       End With
    End If
    If (PB_strFMCCallFormName = "frm발주서관리") Then
       With frm발주서관리
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매입작성1") Then
       With frm매입작성1
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매입작성2") Then
       With frm매입작성2
            'txtSize.Text = .vsfg1.TextMatrix(.vsfg1.Row, 2)
            'txtUnit.Text = .vsfg1.TextMatrix(.vsfg1.Row, 4)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매입수정") Then
       With frm매입수정
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    
    If (PB_strFMCCallFormName = "frm견적서작성") Then
       With frm견적서작성
            'txtSize.Text = .vsfg1.TextMatrix(.vsfg1.Row, 2)
            'txtUnit.Text = .vsfg1.TextMatrix(.vsfg1.Row, 4)
       End With
    End If
    If (PB_strFMCCallFormName = "frm견적서관리") Then
       With frm견적서관리
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매출작성1") Then
       With frm매출작성1
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매출작성2") Then
       With frm매출작성2
            'txtSize.Text = .vsfg1.TextMatrix(.vsfg1.Row, 2)
            'txtUnit.Text = .vsfg1.TextMatrix(.vsfg1.Row, 4)
       End With
    End If
    If (PB_strFMCCallFormName = "frm매출수정") Then
       With frm매출수정
            'txtSize.Text = .vsfg2.TextMatrix(.vsfg2.Row, 15)
            'txtUnit.Text = .vsfg2.TextMatrix(.vsfg2.Row, 17)
       End With
    End If
    If Len(PB_strMaterialsCode) = 0 And Len(PB_strMaterialsName) = 0 Then
       P_intFindGbn = 0    '수동검색
    ElseIf _
       Len(PB_strMaterialsCode) <> 0 And Len(PB_strMaterialsName) = 0 Then
       P_intFindGbn = 1    '코드로만 자동검색
    ElseIf _
       Len(PB_strMaterialsCode) = 0 And Len(PB_strMaterialsName) <> 0 Then
       P_intFindGbn = 2    '이름(명)으로만 자동검색
    ElseIf _
       Len(PB_strMaterialsCode) <> 0 And Len(PB_strMaterialsName) <> 0 Then
       P_intFindGbn = 3    '코드와 이름(명)을 동시에 자동검색
    Else
       P_intFindGbn = 0    '수동검색
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
       Select Case Val(PB_regUserinfoU.UserAuthority)
              'Case Is <= 10 '조회
              '     txtCode.Enabled = False: txtSupplierCode.Enabled = False: cmdFind.Enabled = True
              'Case Is <= 20 '인쇄, 조회
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 40 '추가, 저장, 인쇄, 조회
              '     cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
              '     cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              'Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
              '     cmdPrint.Enabled = True: cmdClear.Enabled = True: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              'Case Else
              '     cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = False
              '     cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       SubOther_FILL
       If P_intFindGbn <> 0 Then
          Subvsfg1_FILL
       End If
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

'+--------------+
'/// txtCode ///
'+--------------+
Private Sub txtCode_GotFocus()
    With txtCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
    If KeyCode = vbKeyEscape Then
       SubEscape
    End If
    If KeyCode = vbKeyF1 Then
       'PB_strCallFormName = "frm자재시세검색"
       'PB_strMaterialsCode = Trim(txtCode.Text)
       'PB_strMaterialsName = Trim(txtName.Text)
       'frm자재검색.Show vbModal
       'If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '검색에서 취소(ESC)
       'Else
       '   txtCode.Text = PB_strMaterialsCode
       '   txtName.Text = PB_strMaterialsName
       'End If
       'If PB_strMaterialsCode <> "" Then
       '   SendKeys "{tab}"
       'End If
       'PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    End If
    txtCode.Text = Trim(txtCode.Text)
    If KeyCode = vbKeyReturn Then
       If IsNumeric(Mid(txtCode.Text, 1, 2)) And Mid(UPPER(txtCode.Text), 3, 4) = "CODE" And Len(Mid(txtCode, 7)) = 0 Then
          SendKeys "{tab}"
       ElseIf _
          Len(Trim(txtCode.Text)) > 1 Then
          cmdFind_Click
          vsfg1.SetFocus
       Else
          SendKeys "{tab}"
       End If
    End If
End Sub

'+--------------+
'/// txtName ///
'+--------------+
Private Sub txtName_GotFocus()
    With txtName
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
    If KeyCode = vbKeyEscape Then
       SubEscape
    End If
    'If KeyCode = vbKeyF1 Then
    '   PB_strCallFormName = "frm자재시세검색"
    '   PB_strMaterialsCode = Trim(txtCode.Text)
    '   PB_strMaterialsName = Trim(txtName.Text)
    '   frm자재검색.Show vbModal
    '   If (Len(PB_strMaterialsCode) + Len(PB_strMaterialsName)) = 0 Then '검색에서 취소(ESC)
    '   Else
    '      txtCode.Text = PB_strMaterialsCode
    '      txtName.Text = PB_strMaterialsName
    '   End If
    '   If PB_strMaterialsCode <> "" Then
    '      SendKeys "{tab}"
    '   End If
    '   PB_strCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
    'End If
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub

'+--------------+
'/// txtSize ///
'+--------------+
Private Sub txtSize_GotFocus()
    With txtSize
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtSize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtUnit.SetFocus
    ElseIf _
       KeyCode = vbKeyEscape Then
       SubEscape
    End If
End Sub

'+--------------+
'/// txtUnit ///
'+--------------+
Private Sub txtUnit_GotFocus()
    With txtUnit
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    ElseIf _
       KeyCode = vbKeyEscape Then
       SubEscape
    End If
End Sub

'+----------------+
'/// txtBarCode ///
'+----------------+
Private Sub txtBarCode_GotFocus()
    With txtBarCode
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
    If KeyCode = vbKeyEscape Then
       SubEscape
    End If
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub txtBarCode_LostFocus()
    With txtBarCode
         .Text = Trim(txtBarCode.Text)
    End With
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
Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            .Cell(flexcpForeColor, 0, 0, .FixedRows - 1, .Cols - 1) = vbBlack '.ForeColorSel
            .Cell(flexcpForeColor, .MouseRow, .MouseCol, .MouseRow, .MouseCol) = vbRed
            strData = .TextMatrix(.Row, 6)
            Select Case .MouseCol
                   'Case 3           '(1.적용일자, 2.매입처명)
                   '     .ColSel = 5
                   '     .ColSort(0) = flexSortNone
                   '     .ColSort(1) = flexSortNone
                   '     .ColSort(2) = flexSortNone
                   '     .ColSort(3) = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
                   '     .ColSort(4) = flexSortNone
                   '     .ColSort(5) = flexSortGenericAscending
                   '     .Sort = flexSortUseColSort
                   Case Else
                        .Sort = IIf(P_intButton = 1, flexSortGenericAscending, flexSortGenericDescending)
            End Select
            If .FindRow(strData, , 6) > 0 Then
               .Row = .FindRow(strData, , 6)
            End If
            If PC_intRowCnt < .Rows Then
               .TopRow = .Row
            End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         If .Row < .FixedRows Then
         '   Text1(2).Enabled = True
         Else
         '   Text1(2).Enabled = False
         End If
         If .Row >= .FixedRows Then
            'For lngC = 0 To .Cols - 1
            '    Select Case lngC
            '           Case 0 '자재코드
            '                txtCode.Text = .TextMatrix(.Row, lngC)
            '           Case 1 '자재명
            '                txtName.Text = .TextMatrix(.Row, lngC)
            '           Case 7 '규격
            '                txtSize.Text = .TextMatrix(.Row, lngC)
            '           'Case 4 '주매입처코드
            '           '     txtSupplierCode.Text = .TextMatrix(.Row, lngC)
            '           Case 8 '단위
            '                txtUnit.Text = .TextMatrix(.Row, lngC)
            '           Case Else
            '    End Select
            'Next lngC
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
                       'If Trim(.Text1(0).Text) = vsfg1.TextMatrix(vsfg1.Row, 4) Then                '매입처코드 같으면
                          'For lngR = 1 To .vsfg1.Rows - 1
                          '    If .vsfg1.Row <> lngR Then
                          '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg1.TextMatrix(lngR, 0) And _
                          '          .vsfg1.TextMatrix(.vsfg1.Row, 5) = .vsfg1.TextMatrix(lngR, 5) Then '자재코드 + 매출처코드
                          '          blnDupOK = True
                          '          Exit For
                          '       End If
                          '    End If
                          'Next lngR
                          If (blnDupOK = False) Then 'And _
                             '(.vsfg1.TextMatrix(.vsfg1.Row, 0) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then '자재코드 다르면
                             .vsfg1.TextMatrix(.vsfg1.Row, 0) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                             .vsfg1.TextMatrix(.vsfg1.Row, 1) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                             .vsfg1.TextMatrix(.vsfg1.Row, 2) = vsfg1.TextMatrix(vsfg1.Row, 7)      '자재규격
                             .vsfg1.TextMatrix(.vsfg1.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 8)      '자재단위
                             .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 10)    '입고단가
                             .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 11)    '입고부가
                             .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                             '합계금액 계산(입고금액이 다르면)
                             If .vsfg1.ValueMatrix(.vsfg1.Row, 11) <> _
                               (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                                .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 11) _
                                                   + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                             End If
                             .vsfg1.TextMatrix(.vsfg1.Row, 11) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고금액
                          
                             .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                             .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                             .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                             P_blnSelect = True
                          Else '중복이면
                             PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                          End If
                       'Else '매입처코드가 다르면
                       '   PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       'End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm발주서관리") Then
                  With frm발주서관리
                       'If .vsfg1.TextMatrix(.vsfg1.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 4) Then                  '매입처코드 같으면
                          'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 15) To .vsfg1.ValueMatrix(.vsfg1.Row, 16)
                          '    If .vsfg2.Row <> lngR Then
                          '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                          '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '자재코드 + 매출처코드
                          '          blnDupOK = True
                          '          Exit For
                          '       End If
                          '    End If
                          'Next lngR
                          If (blnDupOK = False) And _
                             (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Or _
                             (.vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.ValueMatrix(vsfg1.Row, 10)) Then '자재코드, 단가 다르면
                             If .vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0) Then
                                .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                                .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                             End If
                             If .vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.ValueMatrix(vsfg1.Row, 10) Then
                                .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                                .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                             End If
                             If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                                .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                             End If
                             .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                             .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                             'Grid Key(사업장코드+발/주/일/자+발주번호+자재코드+매출처코드
                             .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 1) _
                                       & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                             .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '자재규격
                             .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '자재단위
                             
                             .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고단가
                             .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '입고부가
                             .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                             '합계금액 계산(입고금액이 다르면)
                             If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                               (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                                .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10))
                                .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                   + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                             End If
                             .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고금액
                                                               
                             .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                             .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                             .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                             .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고금액
                             P_blnSelect = True
                          Else '중복이면
                             PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                          End If
                       'Else '매입처코드가 다르면
                       '   PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       'End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매입작성1") Then
                  With frm매입작성1
                       'If .vsfg1.TextMatrix(.vsfg1.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 4) Then                '매입처코드
                          'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 15) To .vsfg1.ValueMatrix(.vsfg1.Row, 16)
                          '    If .vsfg2.Row <> lngR Then
                          '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                          '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '자재코드 + 매출처코드
                          '          blnDupOK = True
                          '          Exit For
                          '       End If
                          '    End If
                          'Next lngR
                          'if 중복 = false And (0.발주자재코드 <> 자재코드)           '적용
                          If (blnDupOK = False) And _
                             (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Or _
                             (.vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.ValueMatrix(vsfg1.Row, 10)) Then '자재코드, 단가 다르면
                             If .vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0) Then
                                .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                                .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                             End If
                             If .vsfg2.ValueMatrix(.vsfg2.Row, 19) <> vsfg1.ValueMatrix(vsfg1.Row, 10) Then
                                .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbRed
                                .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 19, .vsfg2.Row, 19) = vbWhite
                             End If
                             If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                                .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                             End If
                             .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                             .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                             'Grid Key(사업장코드+발/주/일/자+발주번호+자재코드+매출처코드
                             .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 1) _
                                       & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                             .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '자재규격
                             .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '자재단위
                             
                             .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고단가
                             .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '입고부가
                             .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                             '합계금액 계산(입고금액이 다르면)
                             If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                               (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                                .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10))
                                .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                   + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                             End If
                             .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고금액
                                                               
                             .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                             .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                             .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                             .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고금액
                             P_blnSelect = True
                          Else '중복이면
                             PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                          End If
                       'Else '매입처코드가 다르면
                       '   PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       'End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매입작성2") Then
                  With frm매입작성2
                       'If Trim(.Text1(0).Text) = vsfg1.TextMatrix(vsfg1.Row, 4) Then                '매입처코드
                          'For lngR = 1 To .vsfg1.Rows - 1
                          '    If .vsfg1.Row <> lngR Then
                          '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg1.TextMatrix(lngR, 0) And _
                          '          .vsfg1.TextMatrix(.vsfg1.Row, 5) = .vsfg1.TextMatrix(lngR, 5) Then '자재코드 + 매출처코드
                          '          blnDupOK = True
                          '          Exit For
                          '       End If
                          '    End If
                          'Next lngR
                          If (blnDupOK = False) And _
                             (.vsfg1.TextMatrix(.vsfg1.Row, 0) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then '자재코드 다르면
                             .vsfg1.TextMatrix(.vsfg1.Row, 0) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                             .vsfg1.TextMatrix(.vsfg1.Row, 1) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                             .vsfg1.TextMatrix(.vsfg1.Row, 2) = vsfg1.TextMatrix(vsfg1.Row, 7)      '자재규격
                             .vsfg1.TextMatrix(.vsfg1.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 8)      '자재단위
                             .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 10)    '입고단가
                             .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 11)    '입고부가
                             .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                             '합계금액 계산(입고금액이 다르면)
                             If .vsfg1.ValueMatrix(.vsfg1.Row, 11) <> _
                               (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                                .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 11) _
                                                   + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                             End If
                             .vsfg1.TextMatrix(.vsfg1.Row, 11) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                               * vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고금액
                             .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                             .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                             .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                             P_blnSelect = True
                          Else '중복이면
                             PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                          End If
                       'Else '매입처코드가 다르면
                       '   PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       'End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매입수정") Then
                  With frm매입수정
                       'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 14) To .vsfg1.ValueMatrix(.vsfg1.Row, 15)
                       '    If .vsfg2.Row <> lngR Then
                       '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                       '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '자재코드 + 매출처코드
                       '          blnDupOK = True
                       '          Exit For
                       '       End If
                       '    End If
                       'Next lngR
                       If (blnDupOK = False) And _
                          (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then
                          .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                          .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                          .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                          'Grid Key(사업장코드+거/래/일/자+거래번호+자재코드+매출처코드
                          .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 3) _
                                    & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                          .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '자재규격
                          .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '자재단위
                             
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '입고부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                          '합계금액 계산(입고금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 22) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)) Then
                             .vsfg1.TextMatrix(.vsfg1.Row, 6) = .vsfg1.ValueMatrix(.vsfg1.Row, 6) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10))
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 22) _
                                                + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 10)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고금액
                                                               
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고금액
                          P_blnSelect = True
                       Else '중복이면
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm견적서작성") Then
                  With frm견적서작성
                       If (blnDupOK = False) Then 'And _
                          (.vsfg1.TextMatrix(.vsfg1.Row, 0) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then '자재코드 다르면
                          .vsfg1.TextMatrix(.vsfg1.Row, 0) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                          .vsfg1.TextMatrix(.vsfg1.Row, 1) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                          .vsfg1.TextMatrix(.vsfg1.Row, 2) = vsfg1.TextMatrix(vsfg1.Row, 7)      '자재규격
                          .vsfg1.TextMatrix(.vsfg1.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 8)      '자재단위
                          .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 10)    '입고단가
                          .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 11)    '입고부가
                          .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                          .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                          .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                          .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                          '합계금액 계산(출고금액이 다르면)
                          If .vsfg1.ValueMatrix(.vsfg1.Row, 15) <> _
                            (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                             .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 15) _
                                                + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg1.TextMatrix(.vsfg1.Row, 15) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고금액
                          P_blnSelect = True
                       Else '중복이면
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm견적서관리") Then
                  With frm견적서관리
                       If (blnDupOK = False) And _
                          (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Or _
                          (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.ValueMatrix(vsfg1.Row, 13)) Then '자재코드, 단가 다르면
                          If .vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0) Then
                             .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                             .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                          End If
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.ValueMatrix(vsfg1.Row, 13) Then
                             .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                             .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          End If
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                          .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                          'Grid Key(사업장코드+견/적/일/자+견적번호+자재코드+매출처코드
                          .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 1) _
                                    & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                          .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '자재규격
                          .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '자재단위
                             
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '입고부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고금액
                                                               
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                          '합계금액 계산(출고금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                            .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13))
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                               + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고금액
                          P_blnSelect = True
                       Else '중복이면
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매출작성1") Then
                  With frm매출작성1
                       'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 15) To .vsfg1.ValueMatrix(.vsfg1.Row, 16)
                       '    If .vsfg2.Row <> lngR Then
                       '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                       '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '자재코드 + 매출처코드
                       '          blnDupOK = True
                       '          Exit For
                       '       End If
                       '    End If
                       'Next lngR
                       If (blnDupOK = False) And _
                          (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Or _
                          (.vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.ValueMatrix(vsfg1.Row, 13)) Then '자재코드, 단가 다르면
                          If .vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0) Then
                             .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                             .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                          End If
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 23) <> vsfg1.ValueMatrix(vsfg1.Row, 13) Then
                             .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbRed
                             .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 23, .vsfg2.Row, 23) = vbWhite
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                          .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                          'Grid Key(사업장코드+견/적/일/자+견적번호+자재코드+매출처코드
                          .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 3) _
                                    & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                          .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '자재규격
                          .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '자재단위
                             
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '입고부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고금액
                                                               
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                          '합계금액 계산(출고금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                            .vsfg1.TextMatrix(.vsfg1.Row, 7) = .vsfg1.ValueMatrix(.vsfg1.Row, 7) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13))
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                               + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고금액
                          P_blnSelect = True
                       Else '중복이면
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매출작성2") Then
                  With frm매출작성2
                       'For lngR = 1 To .vsfg1.Rows - 1
                       '    If .vsfg1.Row <> lngR Then
                       '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg1.TextMatrix(lngR, 0) Then '자재코드
                       '          blnDupOK = True
                       '          Exit For
                       '       End If
                       '    End If
                       'Next lngR
                       If (blnDupOK = False) And _
                          (.vsfg1.TextMatrix(.vsfg1.Row, 0) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then
                          .vsfg1.TextMatrix(.vsfg1.Row, 0) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                          .vsfg1.TextMatrix(.vsfg1.Row, 1) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                          .vsfg1.TextMatrix(.vsfg1.Row, 2) = vsfg1.TextMatrix(vsfg1.Row, 7)      '자재규격
                          .vsfg1.TextMatrix(.vsfg1.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 8)      '자재단위
                          .vsfg1.TextMatrix(.vsfg1.Row, 8) = vsfg1.ValueMatrix(vsfg1.Row, 10)    '입고단가
                          .vsfg1.TextMatrix(.vsfg1.Row, 9) = vsfg1.ValueMatrix(vsfg1.Row, 11)    '입고부가
                          .vsfg1.TextMatrix(.vsfg1.Row, 10) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                          .vsfg1.TextMatrix(.vsfg1.Row, 12) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                          .vsfg1.TextMatrix(.vsfg1.Row, 13) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                          .vsfg1.TextMatrix(.vsfg1.Row, 14) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                          '합계금액 계산(출고금액이 다르면)
                          If .vsfg1.ValueMatrix(.vsfg1.Row, 15) <> _
                            (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg1.ValueMatrix(.vsfg1.Row, 15) _
                                               + (.vsfg1.ValueMatrix(.vsfg1.Row, 3) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg1.TextMatrix(.vsfg1.Row, 15) = .vsfg1.ValueMatrix(.vsfg1.Row, 3) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고금액
                          P_blnSelect = True
                       Else '중복이면
                          PB_strFMCCallFormName = "": PB_strMaterialsCode = "": PB_strMaterialsName = ""
                       End If
                  End With
               ElseIf _
                  (PB_strFMCCallFormName = "frm매출수정") Then
                  With frm매출수정
                       'For lngR = .vsfg1.ValueMatrix(.vsfg1.Row, 15) To .vsfg1.ValueMatrix(.vsfg1.Row, 16)
                       '    If .vsfg2.Row <> lngR Then
                       '       If vsfg1.TextMatrix(vsfg1.Row, 0) = .vsfg2.TextMatrix(lngR, 4) And _
                       '          .vsfg2.TextMatrix(.vsfg2.Row, 8) = .vsfg2.TextMatrix(lngR, 8) Then '자재코드 + 매출처코드
                       '          blnDupOK = True
                       '          Exit For
                       '       End If
                       '    End If
                       'Next lngR
                       'if 중복 = false And (0.발주자재코드 <> 자재코드)           '적용
                       If (blnDupOK = False) And _
                          (.vsfg2.TextMatrix(.vsfg2.Row, 4) <> vsfg1.TextMatrix(vsfg1.Row, 0)) Then
                          .vsfg2.Cell(flexcpBackColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbRed
                          .vsfg2.Cell(flexcpForeColor, .vsfg2.Row, 5, .vsfg2.Row, 5) = vbWhite
                          If .vsfg2.TextMatrix(.vsfg2.Row, 28) = "" Then
                             .vsfg2.TextMatrix(.vsfg2.Row, 28) = "U"
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 4) = vsfg1.TextMatrix(vsfg1.Row, 0)      '자재코드
                          .vsfg2.TextMatrix(.vsfg2.Row, 5) = vsfg1.TextMatrix(vsfg1.Row, 1)      '자재명
                          'Grid Key(사업장코드+거/래/일/자+거래번호+자재코드+매출처코드
                          .vsfg2.TextMatrix(.vsfg2.Row, 12) = .vsfg1.TextMatrix(.vsfg1.Row, 3) _
                                    & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 4) & "-" & .vsfg2.TextMatrix(.vsfg2.Row, 8)
                          .vsfg2.TextMatrix(.vsfg2.Row, 15) = vsfg1.TextMatrix(vsfg1.Row, 7)     '자재규격
                          .vsfg2.TextMatrix(.vsfg2.Row, 17) = vsfg1.TextMatrix(vsfg1.Row, 8)     '자재단위
                             
                          .vsfg2.TextMatrix(.vsfg2.Row, 19) = vsfg1.ValueMatrix(vsfg1.Row, 10)   '입고단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 20) = vsfg1.ValueMatrix(vsfg1.Row, 11)   '입고부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 21) = vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고가격
                          .vsfg2.TextMatrix(.vsfg2.Row, 22) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 12)   '입고금액
                                                               
                          .vsfg2.TextMatrix(.vsfg2.Row, 23) = vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고단가
                          .vsfg2.TextMatrix(.vsfg2.Row, 24) = vsfg1.ValueMatrix(vsfg1.Row, 14)   '출고부가
                          .vsfg2.TextMatrix(.vsfg2.Row, 25) = vsfg1.ValueMatrix(vsfg1.Row, 15)   '출고가격
                          '합계금액 계산(출고금액이 다르면)
                          If .vsfg2.ValueMatrix(.vsfg2.Row, 26) <> _
                            (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)) Then
                            .vsfg1.TextMatrix(.vsfg1.Row, 6) = .vsfg1.ValueMatrix(.vsfg1.Row, 6) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                                             + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13))
                            .lblTotMny.Caption = Format(Vals(.lblTotMny.Caption) - .vsfg2.ValueMatrix(.vsfg2.Row, 26) _
                                               + (.vsfg2.ValueMatrix(.vsfg2.Row, 16) * vsfg1.ValueMatrix(vsfg1.Row, 13)), "#,#.00")
                          End If
                          .vsfg2.TextMatrix(.vsfg2.Row, 26) = .vsfg2.ValueMatrix(.vsfg2.Row, 16) _
                                                            * vsfg1.ValueMatrix(vsfg1.Row, 13)   '출고금액
                          P_blnSelect = True
                       Else '중복이면
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
'/// 추가 ///
'+-----------+
Private Sub cmdAdd_Click()
Dim strSQL     As String
Dim strMtCode  As String
Dim lngCodeSeq As Long
Dim lngR       As Long
Dim blnExist   As Boolean
    If LenH(txtCode.Text) > 18 Then
       MsgBox "코드길이가 최대길이(18자)를 초과합니다. 다시 확인후 입력하여 주세요.", vbCritical, "코드"
       txtCode.SetFocus
       Exit Sub
    End If
    If LenH(txtName.Text) > 30 Then
       MsgBox "품명길이가 최대길이(30자)를 초과합니다. 다시 확인후 입력하여 주세요.", vbCritical, "품명"
       txtName.SetFocus
       Exit Sub
    End If
    If LenH(txtSize.Text) > 30 Then
       MsgBox "규격길이가 최대길이(30자)를 초과합니다. 다시 확인후 입력하여 주세요.", vbCritical, "규격"
       txtSize.SetFocus
       Exit Sub
    End If
    If LenH(txtUnit.Text) > 20 Then
       MsgBox "단위길이가 최대길이(20자)를 초과합니다. 다시 확인후 입력하여 주세요.", vbCritical, "단위"
       txtUnit.SetFocus
       Exit Sub
    End If
    If LenH(txtBarCode.Text) > 13 Then
       MsgBox "바코드길이가 최대길이(13자)를 초과합니다. 다시 확인후 입력하여 주세요.", vbCritical, "바코드"
       txtUnit.SetFocus
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    cmdAdd.Enabled = False
    If (Len(Trim(txtCode.Text)) > 5) And (Len(Trim(txtName.Text)) > 0) And (Len(Trim(txtUnit.Text)) > 0) Then
       PB_adoCnnSQL.BeginTrans
       strSQL = "SELECT 분류코드 FROM 자재분류 " _
               & "WHERE 분류코드 = '" & Mid(Trim(txtCode.Text), 1, 2) & "' "
       On Error GoTo ERROR_TABLE_SELECT
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       If P_adoRec.RecordCount = 1 Then
          P_adoRec.Close
          strMtCode = Mid(Trim(txtCode.Text), 3)
          strSQL = "SELECT 분류코드, 세부코드 FROM 자재 " _
                  & "WHERE 분류코드 = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND 자재명 = '" & Trim(txtName.Text) & "' " _
                    & "AND 규격 = '" & Trim(txtSize.Text) & "' AND 단위 = '" & Trim(txtUnit.Text) & "' "
          On Error GoTo ERROR_TABLE_SELECT
          P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          If P_adoRec.RecordCount = 0 Then
             P_adoRec.Close
             strSQL = "SELECT 분류코드, 세부코드 FROM 자재 " _
                     & "WHERE 분류코드 = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND 세부코드 = '" & strMtCode & "' "
             On Error GoTo ERROR_TABLE_SELECT
             P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
             If P_adoRec.RecordCount = 1 Then
                txtCode.Text = P_adoRec("분류코드") + "CODE"
                strMtCode = "CODE"
             End If
             P_adoRec.Close
             If UPPER(strMtCode) = "CODE" Then
                strMtCode = UPPER(strMtCode)
                strSQL = "SELECT ISNULL(MAX(세부코드),0) AS 세부코드 FROM 자재 " _
                        & "WHERE 분류코드 = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND 세부코드 LIKE 'CODE_____' " _
                          & "AND DATALENGTH(세부코드) = 9 " _
                          & "AND ISNUMERIC(SUBSTRING(세부코드, 5, 5)) = 1 "
                        '& "WHERE 분류코드 = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND 세부코드 LIKE'CODE%' "
                On Error GoTo ERROR_TABLE_SELECT
                P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                If P_adoRec.RecordCount = 0 Then
                  lngCodeSeq = 0
                Else
                  lngCodeSeq = Val(Mid(P_adoRec("세부코드"), 5))
                End If
                P_adoRec.Close
                lngCodeSeq = lngCodeSeq + 1
                strMtCode = strMtCode + Format(lngCodeSeq, "00000")
                txtCode.Text = Mid(Trim(txtCode.Text), 1, 2) + strMtCode
             End If
             strSQL = "INSERT INTO 자재 VALUES(" _
                                & "'" & Mid(Trim(txtCode.Text), 1, 2) & "', '" & strMtCode & "', " _
                                & "'" & Trim(txtName.Text) & "', '', '" & Trim(txtSize.Text) & "', '" & Trim(txtUnit.Text) & "', " _
                                & "0, 1, '', 0, '', '" & PB_regUserinfoU.UserCode & "') "
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
             '자재원장 추가
             strSQL = "INSERT INTO 자재원장 VALUES(" _
                                & "'" & PB_regUserinfoU.UserBranchCode & "', " _
                                & "'" & Mid(Trim(txtCode.Text), 1, 2) & "', '" & strMtCode & "', 0, 0, '', '', " _
                                & "0, '" & PB_regUserinfoU.UserClientDate & "', '" & PB_regUserinfoU.UserCode & "', " _
                                & "'', '', 0, 0, 0, 0, 0, 0 )"
             On Error GoTo ERROR_TABLE_INSERT
             PB_adoCnnSQL.Execute strSQL
          ElseIf _
             P_adoRec.RecordCount = 1 Then
             txtCode.Text = P_adoRec("분류코드") + P_adoRec("세부코드")
             strMtCode = P_adoRec("세부코드")
             P_adoRec.Close
             For lngR = 1 To vsfg1.Rows - 1
                 If vsfg1.TextMatrix(lngR, 0) = Trim(txtCode.Text) Then
                    vsfg1.Row = lngR
                    blnExist = True
                    Exit For
                 End If
             Next lngR
          Else
             P_adoRec.Close
          End If
          With vsfg1
               If blnExist = False Then
                  .AddItem ""
                  lngR = .Rows - 1
                  .TextMatrix(lngR, 0) = Trim(txtCode.Text): .TextMatrix(lngR, 1) = Trim(txtName.Text)
                  .TextMatrix(lngR, 2) = 0: .TextMatrix(lngR, 3) = 0
                  .TextMatrix(lngR, 5) = ""
                  .TextMatrix(lngR, 6) = .TextMatrix(lngR, 0)
                  .Cell(flexcpData, lngR, 6, lngR, 6) = .TextMatrix(lngR, 6)
                  .TextMatrix(lngR, 7) = Trim(txtSize.Text): .TextMatrix(lngR, 8) = Trim(txtUnit.Text)
                  .TextMatrix(lngR, 9) = 0: .TextMatrix(lngR, 10) = 0
                  .TextMatrix(lngR, 11) = 0: .TextMatrix(lngR, 12) = 0
                  .TextMatrix(lngR, 13) = 0: .TextMatrix(lngR, 14) = 0
                  .TextMatrix(lngR, 15) = 0: .TextMatrix(lngR, 16) = 0
                  .TextMatrix(lngR, 17) = 1: .TextMatrix(lngR, 18) = "과  세"
                  .Row = lngR
                  .SetFocus
               End If
          End With
       Else
          P_adoRec.Close
       End If
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
                      & "WHERE T1.분류코드 = '" & Mid(Trim(txtCode.Text), 1, 2) & "' AND T1.세부코드 = '" & strMtCode & "' "
              On Error GoTo ERROR_TABLE_INSERT
              PB_adoCnnSQL.Execute strSQL
              P_adoRec.MoveNext
           Loop
           P_adoRec.Close
       End If
       PB_adoCnnSQL.CommitTrans
    End If
    cmdAdd.Enabled = True
    vsfg1.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "읽기 실패"
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "추가 실패"
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "갱신 실패"
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
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
    Set frm자재시세검색 = Nothing
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
    With vsfg1           'Rows 1, Cols 19, RowHeightMax(Min) 300
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
         .FixedCols = 6
         .Rows = 1             'SubvsfgUpGrid_Fill수행시에 설정
         .Cols = 19
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 1700   '자재코드(분류코드+세부코드) '1900
         .ColWidth(1) = 2300   '자재명                      '2500
         .ColWidth(2) = 850    '가용재고
         .ColWidth(3) = 1000   '현재재고
         .ColWidth(4) = 1000   '주매입처코드                'Hidden
         .ColWidth(5) = 1800   '주매입처명
         .ColWidth(6) = 1000   '자재코드                    'Hidden
         .ColWidth(7) = 2200   '규격
         .ColWidth(8) = 550    '단위
         .ColWidth(9) = 1000   '폐기율
         .ColWidth(10) = 1200  '입고단가
         .ColWidth(11) = 1200  '입고부가
         .ColWidth(12) = 1500  '입고가격
         .ColWidth(13) = 1200  '출고단가
         .ColWidth(14) = 1200  '출고부가
         .ColWidth(15) = 1500  '출고가격
         .ColWidth(16) = 800   '마진율
         .ColWidth(17) = 1     '과세구분
         .ColWidth(18) = 500   '과세구분
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "코드"
         .TextMatrix(0, 1) = "품명"
         .TextMatrix(0, 2) = "가용재고"
         .TextMatrix(0, 3) = "현재재고"
         .TextMatrix(0, 4) = "주매입처코드"    'H
         .TextMatrix(0, 5) = "주매입처명"
         .TextMatrix(0, 6) = "KEY"             'H
         .TextMatrix(0, 7) = "규격"
         .TextMatrix(0, 8) = "단위"
         .TextMatrix(0, 9) = "폐기율"          'H
         .TextMatrix(0, 10) = "매입단가"
         .TextMatrix(0, 11) = "매입부가"       'H
         .TextMatrix(0, 12) = "매입가격"       'H
         .TextMatrix(0, 13) = "매출단가"
         .TextMatrix(0, 14) = "매출부가"       'H
         .TextMatrix(0, 15) = "매출가격"       'H
         .TextMatrix(0, 16) = "마진율"         'H
         .TextMatrix(0, 17) = "과세구분"       'H
         .TextMatrix(0, 18) = "과세"           'H
         .ColHidden(4) = True: .ColHidden(6) = True: .ColHidden(9) = True: .ColHidden(12) = True
         .ColHidden(11) = True: .ColHidden(14) = True: .ColHidden(15) = True: .ColHidden(16) = True
         .ColHidden(17) = True: .ColHidden(18) = True
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 2
                         .ColFormat(lngC) = "#,#"
                    Case 9 To 16
                         .ColFormat(lngC) = "#,#.00"
             End Select
         Next lngC
         For lngC = 0 To .Cols - 1
             Select Case lngC
                    Case 0, 1, 5, 7, 8
                        .ColAlignment(lngC) = flexAlignLeftCenter
                    Case 4, 18
                        .ColAlignment(lngC) = flexAlignCenterCenter
                    Case Else
                        .ColAlignment(lngC) = flexAlignRightCenter
             End Select
         Next lngC
         .MergeCells = flexMergeRestrictAll
         For lngC = 0 To 5
             .MergeCol(lngC) = True
         Next lngC
         If (PB_strFMCCallFormName = "frm발주서작성") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm발주서관리") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm매입작성1") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm매입작성2") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "반품관리") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm매입수정") Then
            .ColHidden(13) = True
         End If
         If (PB_strFMCCallFormName = "frm견적서작성") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm견적서관리") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm매출작성1") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm매출작성2") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm반입관리") Then
            '.ColHidden(10) = True
         End If
         If (PB_strFMCCallFormName = "frm매출수정") Then
            '.ColHidden(10) = True
         End If
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfgGrid) 채우기///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim strSQL     As String
Dim strSelect  As String
Dim strJoin    As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
Dim lngRRR     As Long
Dim strAppDate As Long
    Screen.MousePointer = vbHourglass
    txtCode.Text = Trim(txtCode.Text): txtName.Text = Trim(txtName.Text)
    If Len(txtCode.Text) = 0 And Len(txtName.Text) = 0 Then
       P_intFindGbn = 0    '수동검색
    ElseIf _
       Len(txtCode.Text) <> 0 And Len(txtName.Text) = 0 Then
       P_intFindGbn = 1    '코드로만 자동검색
    ElseIf _
       Len(txtCode.Text) = 0 And Len(txtName.Text) <> 0 Then
       P_intFindGbn = 2    '이름(명)으로만 자동검색
    ElseIf _
       Len(txtCode.Text) <> 0 And Len(txtName.Text) <> 0 Then
       P_intFindGbn = 3    '코드와 이름(명)을 동시에 자동검색
    Else
       P_intFindGbn = 0    '수동검색
    End If
    If P_intFindGbn = 1 Then '코드로 검색
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "(T1.분류코드 + T1.세부코드) LIKE '%" & Trim(txtCode.Text) & "%' " _
                & "AND T2.단위 LIKE '%" & Trim(txtUnit.Text) & "%' " _
                & "AND T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       strOrderBy = "ORDER BY T1.분류코드, T1.세부코드 "
    ElseIf _
       P_intFindGbn = 2 Then '이름으로 검색
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T2.자재명 LIKE '%" & Trim(txtName.Text) & "%' " _
                & "AND T2.규격 LIKE '%" & Trim(txtSize.Text) & "%' " _
                & "AND T2.단위 LIKE '%" & Trim(txtUnit.Text) & "%' " _
                & "AND T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       strOrderBy = "ORDER BY T2.자재명 "
    ElseIf _
       P_intFindGbn = 3 Then '코드와 이름으로 검색
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "(T1.분류코드 + T1.세부코드) LIKE '%" & Trim(txtCode.Text) & "%' " _
                & "AND T2.자재명 LIKE '%" & Trim(txtName.Text) & "%' " _
                & "AND T2.규격 LIKE '%" & Trim(txtSize.Text) & "%' " _
                & "AND T2.단위 LIKE '%" & Trim(txtUnit.Text) & "%' " _
                & "AND T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       strOrderBy = "ORDER BY T1.분류코드, T1.세부코드 "
    Else
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T2.자재명 LIKE '%" & Trim(txtName.Text) & "%' " _
                & "AND T2.규격 LIKE '%" & Trim(txtSize.Text) & "%' " _
                & "AND T2.단위 LIKE '%" & Trim(txtUnit.Text) & "%' " _
                & "AND T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       strOrderBy = "ORDER BY T1.분류코드, T1.세부코드 "
    End If
    If PB_strFMCCallFormName = "frm발주서작성" Or PB_strFMCCallFormName = "frm발주서관리" Or _
       PB_strFMCCallFormName = "frm매입작성1" Or PB_strFMCCallFormName = "frm매입작성2" Or PB_strFMCCallFormName = "frm매입수정" Or _
       PB_strFMCCallFormName = "frm반품관리" Then
       P_intIOGbn = 1
       strSelect = "입고단가 = CASE WHEN T4.단가구분 = 1 THEN T1.입고단가1 " _
                                 & "WHEN T4.단가구분 = 2 THEN T1.입고단가2 " _
                                 & "WHEN T4.단가구분 = 3 THEN T1.입고단가3 ELSE 0 END, " _
                 & "입고부가 = CASE WHEN T4.단가구분 = 1 THEN ROUND(T1.입고단가1 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.단가구분 = 2 THEN ROUND(T1.입고단가2 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.단가구분 = 3 THEN ROUND(T1.입고단가3 * (" & PB_curVatRate & "), 0, 1) ELSE 0 END, " _
                 & "입고가격 = CASE WHEN T4.단가구분 = 1 THEN T1.입고단가1 + ROUND(T1.입고단가1 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.단가구분 = 2 THEN T1.입고단가2 + ROUND(T1.입고단가2 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.단가구분 = 3 THEN T1.입고단가3 + ROUND(T1.입고단가3 * (" & PB_curVatRate & "), 0, 1) ELSE 0 END, " _
                 & "0 AS 출고단가, 0 AS 출고부가, 0 AS 출고가격, "
       strJoin = "INNER JOIN 매입처 T4 ON T4.사업장코드 = T1.사업장코드 AND T4.매입처코드 = '" & Trim(PB_strSupplierCode) & "' "
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T4.매입처코드 = '" & Trim(PB_strSupplierCode) & "' "
    Else
       P_intIOGbn = 2
       strSelect = "T1.입고단가1 AS 입고단가, (T1.입고단가1 * 0.1) AS 입고부가, (T1.입고단가1 + (T1.입고단가1 * 0.1)) AS 입고가격, " _
                 & "출고단가 = CASE WHEN T4.단가구분 = 1 THEN T1.출고단가1 " _
                                 & "WHEN T4.단가구분 = 2 THEN T1.출고단가2 " _
                                 & "WHEN T4.단가구분 = 3 THEN T1.출고단가3 ELSE 0 END, " _
                 & "출고부가 = CASE WHEN T4.단가구분 = 1 THEN ROUND(T1.출고단가1 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.단가구분 = 2 THEN ROUND(T1.출고단가2 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.단가구분 = 3 THEN ROUND(T1.출고단가3 * (" & PB_curVatRate & "), 0, 1) ELSE 0 END, " _
                 & "출고가격 = CASE WHEN T4.단가구분 = 1 THEN T1.출고단가1 + ROUND(T1.출고단가1 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.단가구분 = 2 THEN T1.출고단가2 + ROUND(T1.출고단가2 * (" & PB_curVatRate & "), 0, 1) " _
                                 & "WHEN T4.단가구분 = 3 THEN T1.출고단가3 + ROUND(T1.출고단가3 * (" & PB_curVatRate & "), 0, 1) ELSE 0 END, "
       strJoin = "INNER JOIN 매출처 T4 ON T4.사업장코드 = T1.사업장코드 AND T4.매출처코드 = '" & Trim(PB_strSupplierCode) & "' "
       strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") _
                & "T4.매출처코드 = '" & Trim(PB_strSupplierCode) & "' "
    End If
    strWhere = strWhere + IIf(Trim(strWhere) = "", "WHERE ", "AND ") & "T2.사용구분 = 0 "
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.분류코드, T1.세부코드, T2.자재명 AS 자재명, " _
                  & "T1.주매입처코드 AS 주매입처코드, ISNULL(T3.매입처명, '') AS 주매입처명, T2.규격 AS 규격, T2.단위 AS 단위," _
                  & "T2.폐기율 AS 폐기율, T2.과세구분 AS 과세구분, T2.사용구분 AS 사용구분, ISNULL(T1.적정재고, 0) AS 적정재고, "
    strSQL = strSQL + strSelect
    strSQL = strSQL _
                  & "(SELECT ISNULL(SUM(입고누계수량 - 출고누계수량), 0) FROM 자재원장마감 " _
                    & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                      & "AND 사업장코드 = T1.사업장코드 " _
                      & "AND 마감년월 >= '" & Left(PB_regUserinfoU.UserClientDate, 4) + "00" & "' " _
                      & "AND 마감년월 <  '" & Left(PB_regUserinfoU.UserClientDate, 6) & "') AS 이월재고,"
    strSQL = strSQL _
                 & "(SELECT ISNULL(SUM(입고수량 - 출고수량), 0) FROM 자재입출내역 " _
                   & "WHERE 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드 " _
                     & "AND 사업장코드 = T1.사업장코드 " _
                     & "AND 입출고일자 BETWEEN '" & Left(PB_regUserinfoU.UserClientDate, 6) + "01" & "' " _
                                            & "AND '" & PB_regUserinfoU.UserClientDate & "') AS 금월재고, "
    strSQL = strSQL _
                 & "(SELECT ISNULL(SUM(발주량), 0) FROM 발주내역 " _
                   & "WHERE 자재코드 = (T1.분류코드 + T1.세부코드) " _
                     & "AND 사업장코드 = T1.사업장코드 AND 상태코드 = 1 AND 사용구분 = 0 " _
                     & "AND 발주일자 >= '" & PB_regUserinfoU.UserClientDate & "') AS 입고예정, "
    strSQL = strSQL _
                 & "(SELECT ISNULL(SUM(수량), 0) FROM 견적내역 " _
                   & "WHERE 자재코드 = (T1.분류코드 + T1.세부코드) " _
                     & "AND 사업장코드 = T1.사업장코드 AND 상태코드 = 1 AND 사용구분 = 0 " _
                     & "AND 견적일자 >= '" & PB_regUserinfoU.UserClientDate & "') AS 출고예정 "
    strSQL = strSQL _
             & "FROM 자재원장 T1 " _
            & "INNER JOIN 자재 T2 " _
                    & "ON T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드 " _
             & "LEFT JOIN 매입처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매입처코드 = T1.주매입처코드 " _
            & "" & strJoin & " " _
            & "" & strWhere & " " _
            & "" & strOrderBy & " "
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
               .TextMatrix(lngR, 0) = P_adoRec("분류코드") + P_adoRec("세부코드")
               .TextMatrix(lngR, 1) = P_adoRec("자재명")
               .TextMatrix(lngR, 2) = P_adoRec("이월재고") + P_adoRec("금월재고") + P_adoRec("입고예정") + P_adoRec("출고예정")
               If P_adoRec("적정재고") <> 0 And _
                  .ValueMatrix(lngR, 2) < P_adoRec("적정재고") Then
                  .Cell(flexcpForeColor, lngR, 2, lngR, 2) = vbRed
                  .Cell(flexcpFontBold, lngR, 2, lngR, 2) = True
               End If
               .TextMatrix(lngR, 3) = P_adoRec("이월재고") + P_adoRec("금월재고")
               .TextMatrix(lngR, 4) = P_adoRec("주매입처코드")
               .TextMatrix(lngR, 5) = P_adoRec("주매입처명")
               .TextMatrix(lngR, 6) = .TextMatrix(lngR, 0)
               .Cell(flexcpData, lngR, 6, lngR, 6) = .TextMatrix(lngR, 6)
               .TextMatrix(lngR, 7) = P_adoRec("규격")
               .TextMatrix(lngR, 8) = P_adoRec("단위")
               
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("폐기율")), 0, P_adoRec("폐기율"))
               .TextMatrix(lngR, 10) = P_adoRec("입고단가")
               .TextMatrix(lngR, 11) = P_adoRec("입고부가")
               .TextMatrix(lngR, 12) = P_adoRec("입고가격")
               .TextMatrix(lngR, 13) = P_adoRec("출고단가")
               .TextMatrix(lngR, 14) = P_adoRec("출고부가")
               .TextMatrix(lngR, 15) = P_adoRec("출고가격")
               '.TextMatrix(lngR, 16) = P_adoRec("마진율")
               .TextMatrix(lngR, 17) = IIf(IsNull(P_adoRec("과세구분")), 0, P_adoRec("과세구분"))
               If .ValueMatrix(lngR, 17) = 0 Then
                  .TextMatrix(lngR, 18) = "비과세"
               Else
                  .TextMatrix(lngR, 18) = "과  세"
               End If
               lngRR = 1
               'If P_intFindGbn = 1 Then
               '   If txtCode.Text = Trim(.TextMatrix(lngR, 0)) And _
               '      txtSupplierCode.Text = Trim(.TextMatrix(lngR, 4)) Then
               '      lngRR = lngR
               '   End If
               'ElseIf _
               '   P_intFindGbn = 2 Then
               '   If txtName.Text = Trim(.TextMatrix(lngR, 1)) And _
               '      txtSupplierCode.Text = Trim(.TextMatrix(lngR, 4)) And _
               '      txtSize.Text = Trim(.TextMatrix(lngR, 7)) And _
               '      txtUnit.Text = Trim(.TextMatrix(lngR, 8)) Then
               '      lngRR = lngR
               '   End If
               'ElseIf _
               '   P_intFindGbn = 3 Then
               '   If txtCode.Text = Trim(.TextMatrix(lngR, 0)) And _
               '      txtName.Text = Trim(.TextMatrix(lngR, 1)) And _
               '      txtSupplierCode.Text = Trim(.TextMatrix(lngR, 4)) And _
               '      txtSize.Text = Trim(.TextMatrix(lngR, 7)) And _
               '      txtUnit.Text = Trim(.TextMatrix(lngR, 8)) Then
               '      lngRR = lngR
               '   End If
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

Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
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

