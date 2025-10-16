VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15315
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11100
   ScaleMode       =   0  '사용자
   ScaleWidth      =   15405
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  '아래 맞춤
      Height          =   348
      Left            =   0
      TabIndex        =   0
      Top             =   10095
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3683
            MinWidth        =   3683
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3507
            MinWidth        =   3507
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "도움말"
            TextSave        =   "도움말"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17887
            MinWidth        =   17887
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   0
      X2              =   14997.62
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   15405
      Y1              =   31.897
      Y2              =   31.897
   End
   Begin VB.Menu 기초자료정보관리 
      Caption         =   "기초자료정보관리(&1)"
      Begin VB.Menu 사업장정보 
         Caption         =   "사 업 장  정  보   등   록"
      End
      Begin VB.Menu 사용자정보 
         Caption         =   "사 용 자  정  보   등   록"
      End
      Begin VB.Menu Filler1_1 
         Caption         =   "-"
      End
      Begin VB.Menu 제조처정보 
         Caption         =   "제 조 처  정  보   등   록"
         Enabled         =   0   'False
      End
      Begin VB.Menu Filler1_2 
         Caption         =   "-"
      End
      Begin VB.Menu 은행정보 
         Caption         =   "은   행   코   드   등   록"
      End
      Begin VB.Menu Filler1_3 
         Caption         =   "-"
      End
      Begin VB.Menu 자재분류 
         Caption         =   "자   재   분   류   등   록"
      End
      Begin VB.Menu 자재정보 
         Caption         =   "자   재   코   드   등   록"
      End
      Begin VB.Menu Filler1_4 
         Caption         =   "-"
      End
      Begin VB.Menu 종료 
         Caption         =   "종                           료"
      End
   End
   Begin VB.Menu 거래처별장부관리 
      Caption         =   "거래처별장부관리(&2)"
      Begin VB.Menu 매입세금계산서장부입력 
         Caption         =   "매입세금계산서장부  입           력"
      End
      Begin VB.Menu 매입세금계산서조회및수정 
         Caption         =   "매입세금계산서장부  조회 및 수정"
      End
      Begin VB.Menu 미지급금원장 
         Caption         =   "매    입    처     출    금    처    리"
      End
      Begin VB.Menu 매입처보조부 
         Caption         =   "매  입 처  별     매    입    현    황"
      End
      Begin VB.Menu Filler2_1 
         Caption         =   "-"
      End
      Begin VB.Menu 매출세금계산서장부입력 
         Caption         =   "매 출 처   미 수 금    장  부 입  력"
      End
      Begin VB.Menu 매출세금계산서조회및수정 
         Caption         =   "매출처미수금 장부    조회 및 수정"
      End
      Begin VB.Menu 미수금원장 
         Caption         =   "매    출    처     수    금    처    리"
      End
      Begin VB.Menu 계산서건별 
         Caption         =   "세 금 계산서    (건    별)   처    리"
      End
      Begin VB.Menu 계산서일괄 
         Caption         =   "매  출 장  부     일  괄  처  리(NO)"
      End
      Begin VB.Menu 세금계산서 
         Caption         =   "세 금 계산서     조  회  및   수  정"
      End
      Begin VB.Menu 매출처보조부 
         Caption         =   "매  출 처  별     매    출    현    황"
      End
   End
   Begin VB.Menu 매입관리 
      Caption         =   "매입관리(&3)"
      Begin VB.Menu 매입작성2 
         Caption         =   "매   입   전   표   입   력"
      End
      Begin VB.Menu 매입처정보 
         Caption         =   "매      입      처   등   록"
      End
      Begin VB.Menu 매입수정 
         Caption         =   "매 입 전 표   조회및수정"
      End
      Begin VB.Menu Filler3_1 
         Caption         =   "-"
      End
      Begin VB.Menu 발주서작성 
         Caption         =   "발      주      서   작   성"
      End
      Begin VB.Menu 발주서관리 
         Caption         =   "발   주   서   조회및수정"
      End
      Begin VB.Menu 매입작성1 
         Caption         =   "발주서 매 입 전 표 처 리"
      End
      Begin VB.Menu Filler3_2 
         Caption         =   "-"
      End
      Begin VB.Menu 반품관리 
         Caption         =   "매 입 전 표   반 품 처 리"
      End
      Begin VB.Menu Filler3_3 
         Caption         =   "-"
      End
      Begin VB.Menu 매입처별단가조회 
         Caption         =   "매 입 처 별   단 가 조 회"
      End
      Begin VB.Menu 품목별매입수량조회 
         Caption         =   "품목별 매 입 수 량 조 회"
      End
   End
   Begin VB.Menu 매출관리 
      Caption         =   "매출관리(&4)"
      Begin VB.Menu 매출작성2 
         Caption         =   "거래 명세서  작         성"
      End
      Begin VB.Menu 매출처정보 
         Caption         =   "매      출      처   등   록"
      End
      Begin VB.Menu 매출수정 
         Caption         =   "거래 명세서  조회및수정"
      End
      Begin VB.Menu Filler4_1 
         Caption         =   "-"
      End
      Begin VB.Menu 견적서작성 
         Caption         =   "견      적      서   작   성"
      End
      Begin VB.Menu 견적서관리 
         Caption         =   "견   적   서   조회및수정"
      End
      Begin VB.Menu 매출작성1 
         Caption         =   "견적서 거래명세서 처 리"
      End
      Begin VB.Menu Filler4_2 
         Caption         =   "-"
      End
      Begin VB.Menu 반입관리 
         Caption         =   "거래명세서   반 품 처 리"
      End
      Begin VB.Menu Filler4_3 
         Caption         =   "-"
      End
      Begin VB.Menu 매출처별단가조회 
         Caption         =   "매 출 처 별   단 가 조 회"
      End
      Begin VB.Menu 품목별매출수량조회 
         Caption         =   "품목별 매 출 수 량 조 회"
      End
   End
   Begin VB.Menu 회계관리 
      Caption         =   "회계관리(&5)"
      Begin VB.Menu 금전출납등록관리 
         Caption         =   "금전출납 등 록 관 리"
      End
      Begin VB.Menu 계정코드등록관리 
         Caption         =   "계정코드 등 록 관 리"
      End
   End
   Begin VB.Menu 재고관리 
      Caption         =   "재고관리(&6)"
      Begin VB.Menu 자재원장 
         Caption         =   "자  재   원  장"
      End
      Begin VB.Menu Filler6_1 
         Caption         =   "-"
      End
      Begin VB.Menu 미달상품 
         Caption         =   "미  달   상  품"
      End
      Begin VB.Menu 수불부 
         Caption         =   "수     불     부"
      End
      Begin VB.Menu Filler6_2 
         Caption         =   "-"
      End
      Begin VB.Menu 재고이동 
         Caption         =   "재  고   이  동"
         Enabled         =   0   'False
      End
      Begin VB.Menu 재고조정 
         Caption         =   "재  고   조  정"
      End
   End
   Begin VB.Menu 마감관리 
      Caption         =   "마감관리(&8)"
      Begin VB.Menu 거래내역취소 
         Caption         =   "거래내역 취소"
      End
      Begin VB.Menu 지급수금취소 
         Caption         =   "지급/수금취소"
      End
      Begin VB.Menu Filler8_1 
         Caption         =   "-"
      End
      Begin VB.Menu 출력물관리 
         Caption         =   "출 력 물 관 리"
      End
      Begin VB.Menu Filler8_2 
         Caption         =   "-"
      End
      Begin VB.Menu 월마감작업 
         Caption         =   "월 마 감 작 업"
      End
      Begin VB.Menu Filler8_3 
         Caption         =   "-"
      End
      Begin VB.Menu 자료백업 
         Caption         =   "자  료   백  업"
      End
   End
   Begin VB.Menu 운영관리 
      Caption         =   "운영관리(&9)"
   End
   Begin VB.Menu 도움말 
      Caption         =   "도움말"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : Main
' 사용된 Control : StatusBar
' 참조된 Table   : 사용자
' 업  무  설  명 :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived     As Boolean
Private P_adoRec         As New ADODB.Recordset

Private Sub Form_Activate()
    P_blnActived = False
    Me.Caption = PB_strSystemName & " Ver5.1.26a " & " - " & PB_regUserinfoU.UserBranchName
End Sub

Private Sub Form_Load()
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       Me.Caption = Me.Caption & " - " & PB_regUserinfoU.UserBranchName
       With frmMain
            .Left = 0: .Top = 0
            .Height = 11100: .ScaleHeight = 11100
            .Width = 15405: .ScaleWidth = 15405
       End With
       SBar.Panels(1).Text = "작업일자 : " & Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       SBar.Panels(2).Text = "사용자 : " & PB_regUserinfoU.UserName
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case Is <= 10 '
              Case Is <= 20 '
              Case Is <= 40 '
              Case Is <= 50 '
              Case Is <= 99 '
              Case Else
       End Select
       P_blnActived = True
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
   Dim x, Y
   If Me.WindowState = vbMinimized Then
      'Me.Icon = LoadPicture()
      'Do While Me.WindowState = vbMinimized
      '   Me.DrawWidth = 10
      '   Me.ForeColor = QBColor(Int(Rnd * 15))
      '   X = Me.Width * Rnd
      '   Y = Me.Height * Rnd
      '   PSet (X, Y)
      '   DoEvents
      'Loop
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '로그인 또는 메인이 아닌 폼이 떠있는 경우 종료할수 없도록 한다.
    'If Forms.Count > 2 Then
    '   If UnloadMode > 0 Then
    '   Else
    '   End If
    '   Cancel = True
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strSQL  As String
Dim inti    As Integer
Dim StrDate As String
Dim strTime As String
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT CONVERT(VARCHAR(8),GETDATE(), 112) AS 서버일자, " _
                  & "RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    StrDate = P_adoRec("서버일자")
    strTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) _
            + Mid(P_adoRec("서버시간"), 7, 2) + Mid(P_adoRec("서버시간"), 10)
    P_adoRec.Close
    PB_adoCnnSQL.BeginTrans
    strSQL = "UPDATE 사용자 SET " _
                  & "로그인여부 = 'N', " _
                  & "종료일시 = '" & StrDate & strTime & "' " _
            & "WHERE 사용자코드 = '" & PB_regUserinfoU.UserCode & "' "
    On Error GoTo ERROR_TABLE_UPDATE
    PB_adoCnnSQL.Execute strSQL
    PB_adoCnnSQL.CommitTrans
    Screen.MousePointer = vbDefault
    Unload frmLogin
    If PB_adoCnnSQL.State <> 0 Then
       PB_adoCnnSQL.Close
    End If
    Set PB_adoCnnSQL = Nothing
    Set frmLogin = Nothing
    Set frmMain = Nothing
    End
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사용자정보 변경 실패"
    Screen.MousePointer = vbDefault
    Unload frmLogin
    If PB_adoCnnSQL.State <> adStateClosed Then
       PB_adoCnnSQL.Close
    End If
    Set PB_adoCnnSQL = Nothing
    Set frmLogin = Nothing
    Set frmMain = Nothing
    End
    Exit Sub
End Sub

'+---------------------------------------------+
'| 기초자료정보관리(메뉴명:기초자료정보관리(1))
'+---------------------------------------------+
Private Sub 사업장정보_Click()                   '메뉴명:사업장정보등록
Dim iRet
    'API 이용
    'iRet = GetSystemMenu(frmMain.hwnd, 0)
    'DeleteMenu iRet, SC_MAXIMIZE, MF_BYCOMMAND
    'DeleteMenu iRet, SC_MINIMIZE, MF_BYCOMMAND
    'DeleteMenu iRet, SC_CLOSE, MF_BYCOMMAND
    With frm사업장정보
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         '.Show vbModeless
         .Show vbModal
    End With
End Sub
Private Sub 사용자정보_Click()                   '메뉴명:사용자정보등록
    With frm사용자정보
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

Private Sub 제조처정보_Click()                   '메뉴명:제조처정보등록
    With frm제조처정보
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 은행정보_Click()                     '메뉴명:은행코드등록
    With frm은행정보
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 자재분류_Click()                     '메뉴명:자재분류등록
    With frm자재분류
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 자재정보_Click()                     '메뉴명:자재코드등록
    With frm자재정보
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

Private Sub 종료_Click()
    Unload Me
End Sub

'+---------------------------------------------+
'| 거래처별장부관리(메뉴명:거래처별장부관리(2))
'|---------------------------------------------+
'| 매입세금계산서장부 입력
'| 매입세금계산서장부 조회및수정
'| 매입처별 출금처리
'| 매입처별 매입현황
'| ---------------------
'| 매출처미수금장부 입력
'| 매출처미수금장부 조회및수정
'| 매출처별 수금처리
'| 세금계산서(건별)처리
'| 매출장부일괄처리(NO)
'| 세금계산서조회및수정
'| 매출처별 매출현황
'+-----------------------------+
Private Sub 매입세금계산서장부입력_Click()       '메뉴명:매입세금계산서장부 입력
    With frm매입세금계산서장부입력
         '.Left = 0: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매입세금계산서조회및수정_Click()     '메뉴명:매입세금계산서장부 조회및수정
    With frm매입세금계산서
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 미지급금원장_Click()                 '메뉴명:매입처별 출금처리
    With frm미지급금원장
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매입처보조부_Click()                 '메뉴명:매입처별 매입현황
    With frm매입처보조부
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

Private Sub 매출세금계산서장부입력_Click()       '메뉴명:매출처미수금장부 입력
    With frm매출세금계산서장부입력
         '.Left = 0: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매출세금계산서조회및수정_Click()     '메뉴명:매출처미수금장부 조회및수정
    With frm매출세금계산서
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 미수금원장_Click()                   '메뉴명:매출처별 수금처리
    With frm미수금원장
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 계산서건별_Click()                   '메뉴명:세금계산서(건별)처리
    With frm계산서건별
         '.Left = 0: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 계산서일괄_Click()                   '메뉴명:매출장부일괄처리(NO)
    With frm계산서일괄
         '.Left = 0: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 세금계산서_Click()                   '메뉴명:세금계산서조회및수정
    With frm세금계산서
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매출처보조부_Click()                 '메뉴명:매출처별 매출현황
    With frm매출처보조부
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| 매입관리(메뉴명:매입관리(3))
'|-----------------------------+
Private Sub 매입작성2_Click()                    '메뉴명:매입전표입력
    With frm매입작성2
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매입처정보_Click()                   '메뉴명:매입처등록
    With frm매입처정보
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매입수정_Click()                     '메뉴명:매입전표 조회및수정
    With frm매입수정
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 발주서작성_Click()                   '메뉴명:발주서작성
    With frm발주서작성
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 발주서관리_Click()                   '메뉴명:발주서 조회및수정
    With frm발주서관리
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매입작성1_Click()                    '메뉴명:발주서 매입전표 처리
    With frm매입작성1
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 반품관리_Click()                     '메뉴명:매입전표 반품처리
    With frm반품관리
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매입처별단가조회_Click()            '메뉴명:매입처별단가조회
    With frm매입처별단가조회
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 품목별매입수량조회_Click()           '메뉴명:품목별매입수량조회
    With frm품목별매입수량조회
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| 매출관리(메뉴명:매출관리(4))
'|-----------------------------+
Private Sub 매출작성2_Click()                    '메뉴명:거래명세서작성
    With frm매출작성2
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매출처정보_Click()                   '메뉴명:매출처등록
    With frm매출처정보
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매출수정_Click()                     '메뉴명:거래명세서조회및수정
    With frm매출수정
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 견적서작성_Click()                   '메뉴명:견적서거래명세서처리
    With frm견적서작성
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 견적서관리_Click()                   '메뉴명:견적서조회및수정
    With frm견적서관리
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매출작성1_Click()                    '메뉴명:견적서거래명세서처리
    With frm매출작성1
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 반입관리_Click()                     '메뉴명:거래명세서반품처리
    With frm반입관리
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 매출처별단가조회_Click()            '메뉴명:매출처별단가조회
    With frm매출처별단가조회
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 품목별매출수량조회_Click()          '메뉴명:품목별매출수량조회
    With frm품목별매출수량조회
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| 회계관리(메뉴명:회계관리(5))
'+-----------------------------+
Private Sub 금전출납등록관리_Click()             '메뉴명:금전출납등록관리
    With frm금전출납관리
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 계정코드등록관리_Click()             '메뉴명:계정코드등록관리
    With frm계정코드관리
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| 재고관리(메뉴명:재고관리(6))
'+-----------------------------+
Private Sub 자재원장_Click()                     '메뉴명:자재원장
    With frm자재원장
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 미달상품_Click()                     '메뉴명:미달상품
    With frm미달상품
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 수불부_Click()                       '메뉴명:수불부
    With frm수불부
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 재고조정_Click()                     '메뉴명:재고조정
    With frm재고조정
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| 마감관리(메뉴명:마감관리(8))
'+-----------------------------+
Private Sub 거래내역취소_Click()                 '메뉴명:거래내역취소
    With frm거래내역취소
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 지급수금취소_Click()                 '메뉴명:지급수금취소
    With frm지급수금취소
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub 출력물관리_Click()                   '메뉴명:출력물관리
    With frm출력물관리
         .Left = (15405 - .Width) / 2: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

Private Sub 월마감작업_Click()                   '메뉴명:월마감작업
    With frm월마감작업
         .Show vbModal
    End With
End Sub

Private Sub 자료백업_Click()                     '메뉴명:자료백업
    With frm자료백업
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| 운영관리(메뉴명:운영관리(9))
'+-----------------------------+

