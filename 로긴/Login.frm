VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3750
   StartUpPosition =   2  '화면 가운데
   Begin MSComCtl2.DTPicker dtpClientDate 
      Height          =   285
      Left            =   1995
      TabIndex        =   0
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   56623105
      CurrentDate     =   37761
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "종  료(&X)"
      Height          =   375
      Left            =   1920
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdExec 
      BackColor       =   &H00FFFFFF&
      Caption         =   "연  결(&C)"
      Height          =   375
      Left            =   120
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtPasswd 
      Alignment       =   2  '가운데 맞춤
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   1995
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1635
      Width           =   1095
   End
   Begin VB.TextBox txtUserId 
      Alignment       =   2  '가운데 맞춤
      Height          =   270
      Left            =   1995
      TabIndex        =   1
      Top             =   1245
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "휴먼편지체"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1785
      TabIndex        =   8
      Top             =   260
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "사용자 작업일자 :"
      Height          =   255
      Index           =   2
      Left            =   315
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "사용자 비밀번호 :"
      Height          =   255
      Index           =   1
      Left            =   315
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "사용자 코드번호 :"
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   5
      Top             =   1275
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   45
      Picture         =   "Login.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3670
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-----------------------------------------------------------------------------+
'| 1. 참조(Project 표준 exe : 1),2),3),4) 기본
'|    1) Visual Basic For Applications -> msvbvm60.dll
'|    2) Visual Basic runtime objects and procedures -> msvbvm60.dll\3
'|    3) Visual Basic objects and procedures -> VB6.OLE
'|    4) OLE automation -> stdole2.tlb
'|
'|    5) Microsoft ActiveX Data Objects 2.1 Library -> msado21.tlb
'|    6) VideoSoft VSFlexGrid7.0 (OLEDB) -> Vsflex7.oca
'|    7) Microsoft Data Formatting Objects Library 6.0 -> MMSTDFMT.DLL
'|
'| 2. 구성요소
'|    1) VideoSoft VSFlexGrid 7.0(OLEDB) -> Vsflex7.ocx
'|    2) Crystal Report Control -> Crystl32.OCX
'|    3) Microsoft Windows Common Controls 6.0 -> Mscomctl.ocx
'|    4) Microsoft Windows Common Controls-2 6.0(SP4) -> MSCPMCT2.OCX
'|       Animation, UpDown, MonthView, DTPicker, FlatscrollBar
'+-----------------------------------------------------------------------------+
Option Explicit
Private P_blnActived     As Boolean
Private P_adoRec         As New ADODB.Recordset
Private P_intConnWay     As Integer '0.비밀번호, 1.결재비밀번호

Private Sub Form_Initialize()
    '
End Sub
'+---------------------------+
'| LOAD FORM ( 한번만 실행 )
'+---------------------------+
Private Sub Form_Load()
    P_blnActived = False
    If App.PrevInstance = True Then
       MsgBox "이미 실행중인 프로그램입니다.", vbCritical, "판매 관리 시스템"
       End
    End If
End Sub

'+--------------------------------------------+
'| Server 연결 방법
'+--------------------------------------------+
'    With adoCnn
'            P_RegU = UserinfoUr_Read()
'           'Set adoCnn = New ADODB.Connection (Private adocnn As ADODB.Connection)
'           'adoCnn.ConnectionString = "driver={SQL Server};server=Server;uid=sa;pwd="
'            'adoCnn.Open
'            'adoCnn.DefaultDatabase = "YmhDB"
'            P_OpenText = "Server=" & P_RegU.ServerName & ";DSN=YmhDB;uid=ymhuser;pwd=userymh;DataBase=YmhDB"
'            On Error GoTo CONNECTION_ERROR_SERVER
'            '.ConnectionString = "Driver={SQL Server}=" & P_RegU.ServerName & ";uid=sa;pwd=;DataBase=YmhDB"
'            .Open P_OpenText
'    End With
'
'+--------------------------------------------+
'| ACTIVATE FORM 활성화 ( 한번만 실행해야 함 )
'+--------------------------------------------+
Private Sub Form_Activate()
Dim strSQL As String
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       PB_regUserinfoU = UserinfoU_Read()
       'Set Connection String
       'DoConnection
       '서버와을 최초연결
       PB_Fnc_AdoCnnSQL
       If PB_varErrCode <> 0 Or PB_blnStatusOfConn <> True Then
          GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       End If
       PB_strSystemName = "판매 관리 시스템": Me.Caption = PB_strSystemName
       lblTitle.Caption = PB_regUserinfoU.UserBranchName
       '서버시간을 가져옴
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT CONVERT(VARCHAR(8),GETDATE(), 112) AS 서버일자, " _
                     & "RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS 서버시간 "
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       PB_regUserinfoU.UserServerDate = P_adoRec("서버일자")
       PB_regUserinfoU.UserServerTime = Mid(P_adoRec("서버시간"), 1, 2) + Mid(P_adoRec("서버시간"), 4, 2) _
                                      + Mid(P_adoRec("서버시간"), 7, 2) + Mid(P_adoRec("서버시간"), 10)
       dtpClientDate.Format = dtpShortDate
       dtpClientDate.Value = Format(P_adoRec("서버일자"), "0000-00-00")
       P_adoRec.Close
       txtUserId.SetFocus
    End If
    P_blnActived = True
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox PB_varErrCode & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "판매 관리 시스템 (서버와의 연결 실패)"
    Unload frmLogin
    Exit Sub
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    If P_adoRec.State <> 0 Then
       P_adoRec.Close
    End If
    If PB_adoCnnSQL.State <> adStateClosed Then
       PB_adoCnnSQL.Close
    End If
    Set frmLogin = Nothing
    Set P_adoRec = Nothing
    Set PB_adoCnnSQL = Nothing
    End
End Sub
Private Sub Form_Terminate()
    '
End Sub

'+-----------------+
'| 사용자 작업일자
'+-----------------+
Private Sub dtpClientDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}"
       Exit Sub
    End If
End Sub
'+-----------------+
'| 사용자 코드번호
'+-----------------+
Private Sub txtUserId_GotFocus()
    With txtUserId
         .SelStart = 0
         .SelLength = Len(.Text)
    End With
    
End Sub
Private Sub txtUserId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}"
       Exit Sub
    End If
End Sub

'+--------------------+
'| 사용자 로긴비밀번호
'+--------------------+
Private Sub txtPasswd_GotFocus()
    With txtPasswd
         .SelStart = 0
         .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtPasswd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       cmdExec_Click
       Exit Sub
    End If
End Sub

'+--------------------+
'| 실행(연결)
'+--------------------+
Private Sub cmdExec_Click()
Dim strSQL      As String
Dim strLpBuffer As String * 256
Dim lngCnt      As Long
Dim intRetVal   As Integer
Dim strWhere    As String
    If Len(Trim(txtUserId.Text)) = 0 Then
       txtUserId.SetFocus
       Exit Sub
    End If
    If Len(Trim(txtPasswd.Text)) = 0 Then
       txtPasswd.SetFocus
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    If P_intConnWay = 0 Then
       strWhere = "AND T1.로그인비밀번호 = '" & Trim(txtPasswd.Text) & "' "
    Else
       strWhere = "AND T1.결재비밀번호 = '" & Trim(txtPasswd.Text) & "' "
    End If
    strSQL = "SELECT T1.사용자코드 AS 사용자코드, T1.사용자명 AS 사용자명, " _
                  & "T1.사업장코드 AS 사업장코드, T1.결재비밀번호 AS 결재비밀번호, " _
                  & "T1.사용자권한 AS 사용자권한, T1.로그인여부 AS 로그인여부," _
                  & "T2.부가세율 AS 부가세율, T2.사업장명 AS 사업장명, T2.사업자번호 AS 사업자번호, T2.대표자명 AS 대표자명, " _
                  & "(T2.주소 + SPACE(1) + T2.번지) AS 주소번지, T2.업태 AS 업태, T2.업종 AS 업종, " _
                  & "T2.미지급금발생구분 AS 미지급금발생구분, T2.미수금발생구분 AS 미수금발생구분, " _
                  & "T2.출력타입구분 AS 출력타입구분, T2.거래명세서상단마진 AS 거래명세서상단마진, " _
                  & "T2.거래명세서왼쪽마진 AS 거래명세서왼쪽마진, T2.세금계산서상단마진 AS 세금계산서상단마진, " _
                  & "T2.세금계산서왼쪽마진 AS 세금계산서왼쪽마진, " _
                  & "T2.최종입고단가자동갱신구분 AS 최종입고단가자동갱신구분, T2.최종출고단가자동갱신구분 AS 최종출고단가자동갱신구분 " _
             & "FROM 사용자 T1 " _
            & "INNER JOIN 사업장 T2 " _
                    & "ON T1.사업장코드 = T2.사업장코드 " _
            & "WHERE T1.사용자코드 = '" & Trim(txtUserId.Text) & "' " _
              & "" & strWhere & " " _
              & "AND T1.사용구분 = 0 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       MsgBox "사용자코드와 로그인비밀번호를 다시 입력하세요.", vbCritical, "사용자 입력 오류"
       P_adoRec.Close
       txtUserId.SetFocus
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       If P_intConnWay = 0 Then
          If UPPER(P_adoRec("로그인여부")) = "Y" Then
             MsgBox "이미 로그인된 사용자입니다. 비밀번호대신에 결재비밀번호를 입력하세요).", vbCritical, "사용자 연결 오류"
             P_adoRec.Close
             P_intConnWay = 1
             txtPasswd.Text = ""
             txtPasswd.SetFocus
             Screen.MousePointer = vbDefault
             Exit Sub
          End If
       End If
    End If
    '+--------------+
    '전역변수에 등록
    '+--------------+
    PB_curVatRate = (P_adoRec("부가세율") / 100) '부가세율
    
    PB_intIAutoPriceGbn = P_adoRec("최종입고단가자동갱신구분")
    PB_intOAutoPriceGbn = P_adoRec("최종출고단가자동갱신구분")
    
    PB_strEnterNo = P_adoRec("사업자번호")
    PB_strEnterName = P_adoRec("사업장명")
    PB_strRepName = P_adoRec("대표자명")
    PB_strEnterAddress = P_adoRec("주소번지")
    PB_strUptae = P_adoRec("업태")
    PB_strUpjong = P_adoRec("업종")
    
    PB_intPrtTypeGbn = P_adoRec("출력타입구분")
    PB_intDTopMargin = P_adoRec("거래명세서상단마진")
    PB_intDLeftMargin = P_adoRec("거래명세서왼쪽마진")
    PB_intTTopMargin = P_adoRec("세금계산서상단마진")
    PB_intTLeftMargin = P_adoRec("세금계산서왼쪽마진")
    
    'Registery 등록
    'UserComputerName   As String   '1. WorkStation Name
    'UserClientName     As String   '2. Client Wondows Login Name
    'UserServerDate     As String   '3. 연결시의 서버일자
    'UserServerTime     As String   '4. 연결시의 서버시간
    'UserClientDate     As sting    '5. 프로그램 실행일자
    'UserClientTime     As sting    '6. 프로그램 실행시간
    '+------------------------+
    'UserBranchCode     As String   '7. 사업장코드
    'UserBranchName     As String   '8. 사업장명
    'UserCode           As String   '9. 사용자코드
    'UserName           As String   '10.사용자성명
    'UserLoginPasswd    As String   '11.사용자비밀번호
    'UserSanctionPasswd As String   '12.사용자결재비밀번호
    'UserAuthority      As String   '13.사용자권한
    strLpBuffer = "-"
    lngCnt = GetComputerName(strLpBuffer, 256)
    PB_regUserinfoU.UserComputerName = Trim(strLpBuffer)
    lngCnt = GetUserName(strLpBuffer, 256)
    PB_regUserinfoU.UserClientName = Trim(strLpBuffer)
    'PB_regUserinfoU.UserServerDate = Form_Activate 실행시에 이미 결정
    'PB_regUserinfoU.UserServerTime = Form_Activate 실행시에 이미 결정
    PB_regUserinfoU.UserClientDate = Format(dtpClientDate.Value, "yyyymmdd") 'Format(Date, "yyyymmdd")
    PB_regUserinfoU.UserClientTime = PB_regUserinfoU.UserServerTime           'Format(Time, "hhmmss")
    '+------------------------+
    PB_regUserinfoU.UserBranchCode = P_adoRec("사업장코드")
    PB_regUserinfoU.UserBranchName = P_adoRec("사업장명")
    PB_regUserinfoU.UserCode = P_adoRec("사용자코드")
    PB_regUserinfoU.UserName = P_adoRec("사용자명")
    PB_regUserinfoU.UserLoginPasswd = Trim(txtPasswd.Text)
    PB_regUserinfoU.UserName = P_adoRec("사용자명")
    PB_regUserinfoU.UserSanctionPasswd = P_adoRec("결재비밀번호")
    PB_regUserinfoU.UserAuthority = P_adoRec("사용자권한")
    '+------------------------+
    PB_regUserinfoU.UserMJGbn = P_adoRec("미지급금발생구분")
    PB_regUserinfoU.UserMSGbn = P_adoRec("미수금발생구분")
    P_adoRec.Close
    UserinfoU_Save PB_regUserinfoU
    Screen.MousePointer = vbHourglass
    PB_adoCnnSQL.BeginTrans
    strSQL = "UPDATE 사용자 SET " _
                  & "로그인여부 = 'Y', " _
                  & "시작일시 = '" & PB_regUserinfoU.UserServerDate & PB_regUserinfoU.UserServerTime & "' " _
            & "WHERE 사용자코드 = '" & Trim(txtUserId.Text) & "' "
    On Error GoTo ERROR_TABLE_UPDATE
    PB_adoCnnSQL.Execute strSQL
    PB_adoCnnSQL.CommitTrans
    Screen.MousePointer = vbDefault
    frmLogin.Hide    '로긴화면 숨기고 메인화면으로 이동
    frmMain.Show     'vbModal
    Exit Sub
ERROR_TABLE_SELECT:
    If P_adoRec.State <> 0 Then
       P_adoRec.Close
    End If
    MsgBox Err.Description & "/" & strSQL
    MsgBox "서버 접속중 오류가 발생하였습니다. 담당자에게 문의하여주세요.", vbCritical, "서버 접속 오류1"
    Screen.MousePointer = vbDefault
    cmdExit_Click
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Description & "/" & strSQL
    MsgBox "서버 접속중 오류가 발생하였습니다. 담당자에게 문의하여주세요.", vbCritical, "서버 접속 오류2"
    Screen.MousePointer = vbDefault
    cmdExit_Click
    Exit Sub
End Sub

'+-------------------+
'| 프로그램 완전 종료
'+-------------------+
Private Sub cmdExit_Click()
    Unload frmLogin
End Sub

