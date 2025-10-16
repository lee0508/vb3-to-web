VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm계산서일괄 
   BorderStyle     =   1  '단일 고정
   Caption         =   "매 출 장 부 일 괄 처 리"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "계산서일괄.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3106.538
   ScaleMode       =   0  '사용자
   ScaleWidth      =   12255
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   12075
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   8640
         Style           =   2  '드롭다운 목록
         TabIndex        =   20
         Top             =   240
         Width           =   3315
      End
      Begin VB.ListBox lstPort 
         Height          =   240
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   7920
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
         Caption         =   "매 출 장 부 일 괄 처 리"
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
         Left            =   3765
         TabIndex        =   10
         Top             =   180
         Width           =   4650
      End
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
      Height          =   2355
      Left            =   60
      TabIndex        =   8
      Top             =   630
      Width           =   12075
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   6780
         Picture         =   "계산서일괄.frx":030A
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   5580
         Picture         =   "계산서일괄.frx":0C58
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   4380
         Picture         =   "계산서일괄.frx":14DF
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtTaxBillMny 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '영문
         Left            =   1680
         TabIndex        =   4
         Top             =   1260
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   2160
         TabIndex        =   1
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57081857
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   4320
         TabIndex        =   2
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57081857
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpP_Date 
         Height          =   270
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   57081857
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   11760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblJanMny 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "0"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1680
         TabIndex        =   3
         Top             =   1000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "계산서금액"
         Height          =   240
         Index           =   19
         Left            =   360
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "현장부잔액"
         Height          =   240
         Index           =   18
         Left            =   360
         TabIndex        =   18
         Top             =   1000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "발행일자"
         Height          =   240
         Index           =   17
         Left            =   360
         TabIndex        =   17
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "까지"
         Height          =   240
         Index           =   6
         Left            =   5760
         TabIndex        =   15
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "부터"
         Height          =   240
         Index           =   5
         Left            =   3600
         TabIndex        =   14
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   4
         Left            =   6675
         TabIndex        =   13
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   3
         Left            =   1680
         TabIndex        =   12
         Top             =   645
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "발행기간"
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   640
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm계산서일괄"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 계산서일괄
' 사용된 Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' 참조된 Table   : 사업장, 매출처, 자재입출내역, 세금계산서
' 업  무  설  명 :
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

Dim P                  As Printer
Dim strDefaultPrinter  As String
Dim aryPrinter()       As String
Dim strBuffer          As String

    frmMain.SBar.Panels(4).Text = "현금매출과 매출처가 계산서 미발행인 경우는 집계에서 제외되며, 매출세금계산서장부에도 저장됩니다. "
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       
       'Printer Setting and Seaching(API)
       strBuffer = Space(1024)
       inti = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
       aryPrinter = Split(strBuffer, ",")
       strDefaultPrinter = Trim(Trim(aryPrinter(0)))
       For Each P In Printers
           cboPrinter.AddItem Trim(P.DeviceName)
           lstPort.AddItem P.Port
       Next
       For inti = 0 To cboPrinter.ListCount - 1
           cboPrinter.ListIndex = inti
           If UCase(Trim(cboPrinter.Text)) = UCase(Trim(strDefaultPrinter)) Then
              Exit For
           End If
       Next inti
       '---
       Select Case Val(PB_regUserinfoU.UserAuthority)
              'Case Is <= 10 '조회
              '     cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 20 '인쇄, 조회
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 40 '추가, 저장, 인쇄, 조회
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 50 '삭제, 추가, 저장, 조회, 인쇄
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              'Case Is <= 99 '삭제, 추가, 저장, 조회, 인쇄
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              'Case Else
              '     cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = False
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
       End Select
       SubOther_FILL
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "계산서건별(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+--------------------+
'--- Select Printer ---
'+--------------------+
Private Sub cboPrinter_Click()
    lstPort.ListIndex = cboPrinter.ListIndex
End Sub

'+---------------+
'/// 발행일자 ///
'+---------------+
Private Sub dtpP_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
'+---------------+
'/// 발행기간 ///
'+---------------+
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
Dim intRetVal      As Integer
    If KeyCode = vbKeyReturn Then
       intRetVal = MsgBox("현미수금잔액과 세금세산서금액을 계산하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton1, "현미수금잔액/계산서금액 계산")
       If intRetVal = vbYes Then
          SubCompute_JanMny PB_regUserinfoU.UserBranchCode, DTOS(dtpT_Date.Value) '미수금잔액계산
          SubCompute_TaxBillMny PB_regUserinfoU.UserBranchCode, DTOS(dtpF_Date.Value), DTOS(dtpT_Date.Value) '세금계산서금액계산
       End If
       SendKeys "{tab}"
    End If
End Sub

'+-----------+
'/// 추가 ///
'+-----------+
Private Sub cmdClear_Click()
    SubClearText
End Sub

'+-----------+
'/// 저장 ///
'+-----------+
Private Sub cmdSave_Click()
Dim blnSaveOK      As Boolean
Dim strSQL         As String
Dim lngR           As Long
Dim lngC           As Long
Dim blnOK          As Boolean
Dim intRetVal      As Integer
Dim lngLogCnt      As Long
Dim strMakeYear    As String
Dim lngLogCnt1     As Long
Dim lngLogCnt2     As Long
Dim curInputMoney  As Long
Dim curOutputMoney As Long
Dim strServerTime  As String
Dim strTime        As String
    If dtpF_Date.Value > dtpT_Date.Value Then
       dtpF_Date.SetFocus
       Exit Sub
    End If
    If Vals(txtTaxBillMny.Text) < 1 Then
       cmdExit.SetFocus
       Exit Sub
    End If
    intRetVal = MsgBox("세금계산서를 일괄 저장하시겠습니까 ?", vbQuestion + vbYesNo + vbDefaultButton2, "자료 저장")
    If intRetVal = vbNo Then
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
    strSQL = "SELECT T1.사업장코드, T1.매출처코드, " _
                 & "(SELECT TOP 1 (S2.자재명 + SPACE(1) + S2.규격)  FROM 자재입출내역 S1 " _
                    & "LEFT JOIN 자재 S2 ON S1.분류코드 = S2.분류코드 AND S1.세부코드 = S2.세부코드 " _
                   & "WHERE S1.사업장코드 = T1.사업장코드 " _
                     & "AND S1.입출고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
                     & "AND S1.입출고구분 = 2 AND S1.사용구분 = 0 AND S1.현금구분 = 0 AND S1.계산서발행여부 = 0 " _
                     & "AND S1.매출처코드 = T1.매출처코드 " _
                   & "ORDER BY S1.입출고일자, S1.입출고시간) AS 품목및규격, " _
                 & "(SELECT (COUNT(S1.세부코드) - 1) FROM 자재입출내역 S1 " _
                   & "WHERE S1.사업장코드 = T1.사업장코드 " _
                     & "AND S1.입출고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
                     & "AND S1.입출고구분 = 2 AND S1.사용구분 = 0 AND S1.현금구분 = 0 AND S1.계산서발행여부 = 0 " _
                     & "AND S1.매출처코드 = T1.매출처코드) AS 수량, " _
                  & "ISNULL(SUM(ISNULL(T1.출고단가, 0) * ISNULL(T1.출고수량, 0)), 0) AS 공급가액, " _
                     & "ROUND((ISNULL(SUM(ISNULL(T1.출고단가, 0) * ISNULL(T1.출고수량, 0)), 0) * " & PB_curVatRate & "), 0, 1) AS 세액 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 자재 T2 ON T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드 " _
             & "LEFT JOIN 매출처 T3 ON T3.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
              & "AND T1.입출고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
              & "AND T1.입출고구분 = 2 AND T1.사용구분 = 0 AND T1.현금구분 = 0 AND T1.계산서발행여부 = 0 " _
              & "AND T3.계산서발행여부 = 1 " _
            & "GROUP BY T1.사업장코드, T1.매출처코드 " _
            & "ORDER BY T1.사업장코드, T1.매출처코드 "
           '& "HAVING SUM(ISNULL(T1.출고수량, 0) * ISNULL(T1.출고단가, 0)) > 0 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          '세금계산서번호 구하기
          strMakeYear = Mid(DTOS(dtpP_Date.Value), 1, 4)
          strSQL = "spLogCounter '세금계산서', '" & PB_regUserinfoU.UserBranchCode + strMakeYear & "', " _
                              & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
          On Error GoTo ERROR_STORED_PROCEDURE
          P_adoRecW.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          lngLogCnt1 = P_adoRecW(0)
          lngLogCnt2 = P_adoRecW(1)
          P_adoRecW.Close
          '세금계산서 추가
          strSQL = "INSERT INTO 세금계산서(사업장코드, 작성년도, 책번호, 일련번호, " _
                                        & "매출처코드, 작성일자, 품목및규격, 수량, " _
                                        & "공급가액, 세액, 금액구분, 영청구분, " _
                                        & "발행여부, 작성구분, 미수구분, 적요, 사용구분, " _
                                        & "수정일자, 사용자코드) VALUES(" _
          & "'" & PB_regUserinfoU.UserBranchCode & "', '" & strMakeYear & "', " & lngLogCnt1 & ", " & lngLogCnt2 & "," _
          & "'" & P_adoRec("매출처코드") & "', '" & DTOS(dtpP_Date.Value) & "','" & P_adoRec("품목및규격") & "'," & P_adoRec("수량") & ", " _
          & "" & P_adoRec("공급가액") & ", " & P_adoRec("세액") & ", 3, 1, " _
          & "0, 2, 1, '', 0, " _
          & "'" & PB_regUserinfoU.UserServerDate & "','" & PB_regUserinfoU.UserCode & "')"
          On Error GoTo ERROR_TABLE_INSERT
          PB_adoCnnSQL.Execute strSQL
          '2.매출세금계산서장부(작성시간:strTime)
          strSQL = "INSERT INTO 매출세금계산서장부 " _
                 & "SELECT T1.사업장코드, T1.작성일자, '" & strTime & "', T1.매출처코드, " _
                        & "T1.품목및규격, T1.수량, T1.공급가액, T1.세액, " _
                        & "T1.금액구분, T1.영청구분, T1.발행여부, T1.작성구분, " _
                        & "T1.미수구분, T1.적요, T1.사용구분, T1.수정일자, " _
                        & "T1.사용자코드, T1.작성년도, T1.책번호, T1.일련번호 " _
                   & "FROM 세금계산서 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND T1.작성년도  = '" & strMakeYear & "' AND 책번호 = " & lngLogCnt1 & " AND T1.일련번호 = " & lngLogCnt2 & " "
          On Error GoTo ERROR_TABLE_INSERT
          PB_adoCnnSQL.Execute strSQL
    
          strSQL = "UPDATE 자재입출내역 SET " _
                        & "작성년도 = '" & strMakeYear & "', 책번호 = " & lngLogCnt1 & ", 일련번호 = " & lngLogCnt2 & ", " _
                        & "계산서발행여부 = 1 " _
                  & "WHERE 사업장코드 = '" & P_adoRec("사업장코드") & "' AND 매출처코드 = '" & P_adoRec("매출처코드") & "' " _
                    & "AND 입출고일자 BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
                    & "AND 입출고구분 = 2 AND 사용구분 = 0 AND 현금구분 = 0 AND 계산서발행여부 = 0 "
          On Error GoTo ERROR_TABLE_UPDATE
          PB_adoCnnSQL.Execute strSQL
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    PB_adoCnnSQL.CommitTrans
    Screen.MousePointer = vbDefault
    cmdSave.Enabled = True
    cmdClear_Click
    cmdExit.SetFocus
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
    Set frm계산서일괄 = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    lblJanMny.Caption = "0.00": txtTaxBillMny.Text = ""
    dtpP_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpF_Date.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
    dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "사업장 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+-------------------------+
'/// Clear text1(index) ///
'+-------------------------+
Private Sub SubClearText()
    lblJanMny.Caption = "0.00"
    txtTaxBillMny.Text = ""
End Sub

'+-----------------+
'/// 미수금잔액 ///
'+-----------------+
Private Sub SubCompute_JanMny(strBranchCode As String, strT_Date As String)
Dim strSQL As String
    lblJanMny.Caption = "0.00"
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.매출처코드 AS 매출처코드, " _
                  & "SUM(T1.미수금누계금액) AS 미수금금액, SUM(T1.미수금입금누계금액) AS 미수금입금금액 " _
             & "FROM 미수금원장마감 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' " _
              & "AND T1.마감년월 >= (SUBSTRING('" & strT_Date & "', 1, 4) + '00') " _
              & "AND T1.마감년월 < SUBSTRING('" & strT_Date & "', 1, 6) " _
            & "GROUP BY T1.매출처코드 " _
            & "UNION ALL "
    If PB_regUserinfoU.UserMSGbn = "1" Then  '미수금발생구분 1.전표이면
       strSQL = strSQL + "SELECT T1.매출처코드 AS 매출처코드, " _
                  & "(SUM(T1.출고수량 * T1.출고단가) * " & (PB_curVatRate + 1) & ") AS 미수금금액, 0 AS 미수금입금금액 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.사용구분 = 0 " _
              & "AND T1.입출고일자 BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
              & "AND (T1.입출고구분 = 2 OR T1.입출고구분 = 8) " _
            & "GROUP BY T1.매출처코드 "
            '& "AND (T1.현금구분 = 0 OR (T1.현금구분 = 1 AND T1.계산서발행여부 = 1)) "
    Else                                     '미수금발생구분 2.계산서이면
       strSQL = strSQL _
           & "SELECT T1.매출처코드 AS 매출처코드, " _
                 & "(SUM(T1.공급가액 + T1.세액)) AS 미수금금액, 0 AS 미수금입금금액 " _
             & "FROM 매출세금계산서장부 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' AND T1.사용구분 = 0 " _
              & "AND T1.작성일자 BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
              & "AND T1.미수구분 = 1 " _
            & "GROUP BY T1.매출처코드 "
    End If
    strSQL = strSQL + "UNION ALL " _
           & "SELECT T1.매출처코드 AS 매출처코드, " _
                  & "0 AS 미수금금액, " _
                  & "ISNULL(SUM(T1.미수금입금금액), 0) As 미수금입금금액 " _
             & "FROM 미수금내역 T1 " _
             & "LEFT JOIN 사업장 T2 ON T2.사업장코드 = T1.사업장코드 " _
             & "LEFT JOIN 매출처 T3 ON T3.사업장코드 = T1.사업장코드 AND T3.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' " _
              & "AND T1.미수금입금일자 BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
            & "GROUP BY T1.매출처코드 " _
            & "ORDER BY T1.매출처코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          lblJanMny.Caption = Format(Vals(lblJanMny.Caption) + P_adoRec("미수금금액") - P_adoRec("미수금입금금액"), "#,0.00")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "TABLE 읽기 실패"
    Unload Me
    Exit Sub
End Sub

'+---------------------+
'/// 세금계산서금액 ///
'+---------------------+
Private Sub SubCompute_TaxBillMny(strBranchCode As String, strF_Date As String, strT_Date As String)
Dim strSQL As String
    txtTaxBillMny.Text = "0.00"
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    'strSQL = "SELECT ISNULL(SUM(ISNULL(T1.출고수량, 0) * ISNULL(T1.출고단가, 0)), 0) AS 세금계산서금액 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 매출처 T2 " _
               & "ON T2.사업장코드 = T1.사업장코드 AND T2.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' " _
              & "AND T1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
              & "AND T1.입출고구분 = 2 AND T1.사용구분 = 0 AND T1.현금구분 = 0 AND T1.계산서발행여부 = 0 " _
              & "AND T2.계산서발행여부 = 1 "
    strSQL = "SELECT T1.매출처코드 AS 매출처코드, ISNULL(SUM(ISNULL(T1.출고수량, 0) * ISNULL(T1.출고단가, 0)), 0) AS 세금계산서금액 " _
             & "FROM 자재입출내역 T1 " _
             & "LEFT JOIN 매출처 T2 " _
               & "ON T2.사업장코드 = T1.사업장코드 AND T2.매출처코드 = T1.매출처코드 " _
            & "WHERE T1.사업장코드 = '" & strBranchCode & "' " _
              & "AND T1.입출고일자 BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
              & "AND T1.입출고구분 = 2 AND T1.사용구분 = 0 AND T1.현금구분 = 0 AND T1.계산서발행여부 = 0 " _
              & "AND T2.계산서발행여부 = 1 " _
            & "GROUP BY T1.매출처코드 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          txtTaxBillMny.Text = Format(Vals(txtTaxBillMny) + P_adoRec("세금계산서금액"), "#,0.00")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "TABLE 읽기 실패"
    Unload Me
    Exit Sub
End Sub

