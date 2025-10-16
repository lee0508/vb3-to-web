VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm월마감작업 
   BorderStyle     =   1  '단일 고정
   Caption         =   "월마감작업"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "월마감작업.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4080
   StartUpPosition =   1  '소유자 가운데
   Begin VB.OptionButton optGbn4 
      Caption         =   "미수금"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "미수금내역[장부관련]"
      Top             =   915
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "마감취소(&C)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optGbn1 
      Caption         =   "자재"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "매입(출)내역[재고관련]"
      Top             =   550
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optGbn2 
      Caption         =   "회계"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      ToolTipText     =   "회계전표내역[회계관련]"
      Top             =   550
      Width           =   735
   End
   Begin VB.OptionButton optGbn3 
      Caption         =   "미지급금"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "미지급금내역[장부관련]"
      Top             =   915
      Width           =   1095
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "마감실행(&E)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "종료(&X)"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpT_Date 
      Height          =   270
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   476
      _Version        =   393216
      Format          =   56950785
      CurrentDate     =   37763
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "월마감구분"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "월마감일자"
      Height          =   240
      Index           =   10
      Left            =   120
      TabIndex        =   6
      Top             =   165
      Width           =   1095
   End
End
Attribute VB_Name = "frm월마감작업"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' 프로그램 제 목 : 월마감작업
' 사용된 Control :
' 참조된 Table   :
' 업  무  설  명 :
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
    frmMain.SBar.Panels(4).Text = "금월이전의 자료를 수정한 경우 반드시 월마감 작업을 해당 월마감구분으로 작업하셔야만 합니다."
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       cmdExec.Enabled = False: cmdExit.Enabled = True
       'If dtpT_Date.Value = _
       '   DateAdd("d", -1, DateAdd("m", 1, Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00"))) Then
       '   If Val(PB_regUserinfoU.UserClientDate) >= 50 Then
       '      cmdExec.Enabled = True
       '      cmdExec.SetFocus
       '   End If
       'End If
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "(서버와의 연결 실패)"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// 일자 ///
'+-----------+
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub dtpT_Date_LostFocus()
    If dtpT_Date.Value = _
       DateAdd("d", -1, DateAdd("m", 1, Format(Mid(DTOS(dtpT_Date), 1, 6) & "01", "0000-00-00"))) Then
       If Mid(DTOS(dtpT_Date.Value), 1, 6) >= "190001" Then
          cmdExec.Enabled = True: cmdCancel.Enabled = True
          Exit Sub
       Else
          cmdExec.Enabled = False: cmdCancel.Enabled = False
       End If
    Else
       cmdExec.Enabled = False: cmdCancel.Enabled = False
    End If
End Sub

Private Sub optGbn1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub optGbn2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub optGbn3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub optGbn4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub

'+-----------+
'/// 실행 ///
'+-----------+
Private Sub cmdExec_Click()
Dim strSQL      As String
Dim lngR        As Long
Dim lngC        As Long
Dim blnOK       As Boolean
Dim intRetVal   As Integer
Dim strNextYM   As String
Dim strMagamGbn As String
    If Year(dtpT_Date.Value) = Year(Format(PB_regUserinfoU.UserClientDate, "0000-00-00")) Then
       blnOK = True
    Else
       If PB_regUserinfoU.UserAuthority >= "50" Then
          blnOK = True
       End If
    End If
    '기초마감년월
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.자재기초마감년월 AS 자재기초마감년월, T1.회계기초마감년월 AS 회계기초마감년월, " _
                  & "T1.미지급금기초마감년월 AS 미지급금기초마감년월, T1.미수금기초마감년월 AS 미수금기초마감년월 " _
              & "FROM 사업장 T1 WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
    On Error GoTo ERROR_MONTH_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 1 Then
       If optGbn1.Value = True Then
          strMagamGbn = optGbn1.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("자재기초마감년월") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn2.Value = True Then
          strMagamGbn = optGbn2.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("회계기초마감년월") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn3.Value = True Then
          strMagamGbn = optGbn3.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("미지급금기초마감년월") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn4.Value = True Then
          strMagamGbn = optGbn4.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("미수금기초마감년월") Then
             blnOK = True
          Else
             blnOK = False
          End If
       End If
       P_adoRec.Close
    Else
      P_adoRec.Close
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
    If blnOK = True Then
       intRetVal = MsgBox("월마감 작업(" + strMagamGbn + ")을 월마감구분으로 작업하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton2, "월마감 작업")
    Else
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If intRetVal = vbNo Then
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    cmdExec.Enabled = False
    strNextYM = CStr(Year(DateAdd("y", 1, dtpT_Date.Value))) + "00"
    Screen.MousePointer = vbHourglass
    '월마감
    PB_adoCnnSQL.BeginTrans
    If optGbn1.Value = True Then '1.자재마감
       '1. 자재원장마감
       strSQL = "DELETE 자재원장마감 WHERE 마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                    & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
       strSQL = "INSERT 자재원장마감 " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.분류코드, " _
                      & "T1.세부코드, '" & Left(DTOS(dtpT_Date.Value), 6) & "', " _
                      & "SUM(T1.입고수량), SUM(T1.출고수량), " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM 자재입출내역 T1 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.사용구분 = 0 " _
                 & "AND SUBSTRING(T1.입출고일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
               & "GROUP BY T1.사업장코드, T1.분류코드, T1.세부코드 "
               '& "AND (T1.입출고구분 BETWEEN 1 AND 6) "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
    ElseIf _
       optGbn2.Value = True Then '2.회계전표내역마감
       '2. 회계전표내역마감
       strSQL = "DELETE 회계전표내역마감 WHERE 마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                        & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
       strSQL = "INSERT 회계전표내역마감 " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.계정코드, '" & Left(DTOS(dtpT_Date.Value), 6) & "', " _
                      & "SUM(T1.입금금액), SUM(T1.출금금액), " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM 회계전표내역 T1 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND SUBSTRING(T1.작성일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                 & "AND T1.사용구분 = 0 " _
               & "GROUP BY T1.사업장코드, T1.계정코드 "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
    ElseIf _
       optGbn3.Value = True Then '3.미지급금마감
       '3. 미지급금원장마감
       strSQL = "DELETE 미지급금원장마감 WHERE 마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                        & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
       If PB_regUserinfoU.UserMJGbn = "1" Then '미지급금발생구분 1.전표, 2.(세금)계산서
          strSQL = "INSERT 미지급금원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매입처코드, " _
                         & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', SUM(T1.입고수량*(T1.입고단가+T1.입고부가)), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 자재입출내역 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.사용구분 = 0 AND T1.매입처코드 <> '' " _
                    & "AND (T1.입출고구분 = 1) AND T1.현금구분 = 0 " _
                    & "AND SUBSTRING(T1.입출고일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                  & "GROUP BY T1.사업장코드, T1.매입처코드 "
       Else
          strSQL = "INSERT 미지급금원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매입처코드, " _
                         & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', SUM(T1.공급가액 + T1.세액), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 매입세금계산서장부 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.사용구분 = 0 AND T1.매입처코드 <> '' " _
                    & "AND T1.미지급구분 = 1 " _
                    & "AND SUBSTRING(T1.작성일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                  & "GROUP BY T1.사업장코드, T1.매입처코드 "
       End If
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
       strSQL = "UPDATE 미지급금원장마감 SET " _
                     & "미지급금지급누계금액 = 미지급금지급누계금액 + Z.F5 " _
                & "FROM " _
                    & "(SELECT '" & PB_regUserinfoU.UserBranchCode & "' AS F0, T1.매입처코드 AS F2, " _
                             & "'" & Left(DTOS(dtpT_Date.Value), 6) & "' AS F3, 0 AS F4, " _
                             & "SUM(T1.미지급금지급금액) AS F5, " _
                             & "'" & PB_regUserinfoU.UserServerDate & "' AS F6, '" & PB_regUserinfoU.UserCode & "' AS F7 " _
                       & "FROM 미지급금내역 T1 " _
                      & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                        & "AND SUBSTRING(T1.미지급금지급일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                        & "AND EXISTS " _
                           & "(SELECT T2.마감년월 " _
                              & "FROM 미지급금원장마감 T2 " _
                             & "WHERE T2.사업장코드 = T1.사업장코드 AND T2.매입처코드 = T1.매입처코드 " _
                               & "AND T2.마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' ) " _
                      & "GROUP BY T1.매입처코드) AS Z " _
               & "WHERE 사업장코드 = Z.F0 AND 매입처코드 = Z.F2 AND 마감년월 = Z.F3 "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
       strSQL = "INSERT 미지급금원장마감 " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매입처코드, " _
                      & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', 0, " _
                      & "SUM(T1.미지급금지급금액), " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM 미지급금내역 T1 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND SUBSTRING(T1.미지급금지급일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                 & "AND NOT EXISTS " _
                        & "(SELECT T2.마감년월 " _
                           & "FROM 미지급금원장마감 T2 " _
                          & "WHERE T2.사업장코드 = T1.사업장코드 AND T2.매입처코드 = T1.매입처코드 " _
                            & "AND T2.마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' ) " _
               & "GROUP BY T1.사업장코드, T1.매입처코드 "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
    Else                         '4.미수금마감
       '4. 미수금원장마감
       strSQL = "DELETE 미수금원장마감 WHERE 마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                      & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
       If PB_regUserinfoU.UserMSGbn = "1" Then '미수금발생구분 1.전표, 2.(세금)계산서
          strSQL = "INSERT 미수금원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매출처코드, " _
                         & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', (SUM(T1.출고수량*T1.출고단가) * " & (PB_curVatRate + 1) & ") , " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 자재입출내역 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.사용구분 = 0 AND T1.매출처코드 <> '' " _
                    & "AND (T1.입출고구분 = 2 OR T1.입출고구분 = 8) " _
                    & "AND SUBSTRING(T1.입출고일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                  & "GROUP BY T1.사업장코드, T1.매출처코드 "
                  '& "AND (T1.현금구분 = 0 OR (T1.현금구분 = 1 AND T1.계산서발행여부 = 1)) "
       Else
          strSQL = "INSERT 미수금원장마감 " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매출처코드, " _
                      & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', SUM(T1.공급가액 + T1.세액), " _
                      & "0, " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM 매출세금계산서장부 T1 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.사용구분 = 0 AND T1.매출처코드 <> '' " _
                 & "AND T1.미수구분 = 1 " _
                 & "AND SUBSTRING(T1.작성일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
               & "GROUP BY T1.사업장코드, T1.매출처코드 "
          'strSQL = "INSERT 미수금원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매출처코드, " _
                         & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', SUM(T1.공급가액 + T1.세액) , " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 세금계산서 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.사용구분 = 0 AND T1.매출처코드 <> '' " _
                    & "AND (T1.미수구분 = 1) " _
                    & "AND SUBSTRING(T1.작성일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                  & "GROUP BY T1.사업장코드, T1.매출처코드 "
       End If
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
       strSQL = "UPDATE 미수금원장마감 SET " _
                     & "미수금입금누계금액 = 미수금입금누계금액 + Z.F5 " _
                & "FROM " _
                    & "(SELECT '" & PB_regUserinfoU.UserBranchCode & "' AS F0, T1.매출처코드 AS F2, " _
                             & "'" & Left(DTOS(dtpT_Date.Value), 6) & "' AS F3, 0 AS F4, " _
                             & "SUM(T1.미수금입금금액) AS F5, " _
                             & "'" & PB_regUserinfoU.UserServerDate & "' AS F6, '" & PB_regUserinfoU.UserCode & "' AS F7 " _
                       & "FROM 미수금내역 T1 " _
                      & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                        & "AND SUBSTRING(T1.미수금입금일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                        & "AND EXISTS " _
                           & "(SELECT T2.마감년월 " _
                              & "FROM 미수금원장마감 T2 " _
                             & "WHERE T2.사업장코드 = T1.사업장코드 AND T2.매출처코드 = T1.매출처코드 " _
                               & "AND T2.마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' ) " _
                      & "GROUP BY T1.매출처코드) AS Z " _
               & "WHERE 사업장코드 = Z.F0 AND 매출처코드 = Z.F2 AND 마감년월 = Z.F3 "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
       strSQL = "INSERT 미수금원장마감 " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매출처코드, " _
                      & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', 0, " _
                      & "SUM(T1.미수금입금금액), " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM 미수금내역 T1 " _
               & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND SUBSTRING(T1.미수금입금일자, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                 & "AND NOT EXISTS " _
                        & "(SELECT T2.마감년월 " _
                           & "FROM 미수금원장마감 T2 " _
                          & "WHERE T2.사업장코드 = T1.사업장코드 AND T2.매출처코드 = T1.매출처코드 " _
                            & "AND T2.마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' ) " _
               & "GROUP BY T1.사업장코드, T1.매출처코드 "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
    End If
    '년마감
    If Mid(DTOS(dtpT_Date.Value), 5, 2) = "12" Then
       If optGbn1.Value = True Then '1.자재마감
          '1. 자재원장마감
          strSQL = "DELETE 자재원장마감 WHERE 마감년월 = '" & strNextYM & "' " _
                                       & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT 자재원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.분류코드, " _
                         & "T1.세부코드, '" & strNextYM & "', " _
                         & "SUM(T1.입고누계수량), SUM(T1.출고누계수량), " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 자재원장마감 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND SUBSTRING(T1.마감년월, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.사업장코드, T1.분류코드, T1.세부코드 "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       ElseIf _
          optGbn2.Value = True Then '2.회계전표내역마감
          '2. 회계전표내역마감
          strSQL = "DELETE 회계전표내역마감 WHERE 마감년월 = '" & strNextYM & "' " _
                                           & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT 회계전표내역마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.계정코드, '" & strNextYM & "', " _
                         & "SUM(T1.입금누계금액), SUM(T1.출금누계금액), " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 회계전표내역마감 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND SUBSTRING(T1.마감년월, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.사업장코드, T1.계정코드 "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
          strSQL = "UPDATE 회계전표내역마감 SET " _
                        & "입금누계금액 = CASE WHEN (입금누계금액 > 출금누계금액) THEN (입금누계금액 - 출금누계금액) ELSE 0 END, " _
                        & "출금누계금액 = CASE WHEN (입금누계금액 < 출금누계금액) THEN (출금누계금액 - 입금누계금액) ELSE 0 END " _
                  & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND 마감년월 = '" & strNextYM & "' "
          On Error GoTo ERROR_MONTH_UPDATE
          PB_adoCnnSQL.Execute strSQL
       ElseIf _
          optGbn3.Value = True Then '3.미지급금마감
          '3. 미지급금원장마감
          strSQL = "DELETE 미지급금원장마감 WHERE 마감년월 = '" & strNextYM & "' " _
                                           & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT 미지급금원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매입처코드, " _
                         & "'" & strNextYM & "', SUM(T1.미지급금누계금액 - T1.미지급금지급누계금액), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 미지급금원장마감 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.매입처코드 <> '' " _
                    & "AND SUBSTRING(T1.마감년월, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.사업장코드, T1.매입처코드 "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       Else
          '4. 미수금원장마감
          strSQL = "DELETE 미수금원장마감 WHERE 마감년월 = '" & strNextYM & "' " _
                                         & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT 미수금원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매출처코드, " _
                         & "'" & strNextYM & "', SUM(T1.미수금누계금액 - T1.미수금입금누계금액), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 미수금원장마감 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.매출처코드 <> '' " _
                    & "AND SUBSTRING(T1.마감년월, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.사업장코드, T1.매출처코드 "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       End If
    End If
    PB_adoCnnSQL.CommitTrans
    cmdExec.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "마감읽기 (서버와의 연결 실패)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "마감삭제 (서버와의 연결 실패)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "마감추가 (서버와의 연결 실패)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "마감조정 (서버와의 연결 실패)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

'+-----------+
'/// 취소 ///
'+-----------+
Private Sub cmdCancel_Click()
Dim strSQL    As String
Dim lngR      As Long
Dim lngC      As Long
Dim blnOK     As Boolean
Dim intRetVal As Integer
Dim strNextYM As String
Dim strMagamGbn As String
    If Year(dtpT_Date.Value) = Year(Format(PB_regUserinfoU.UserClientDate, "0000-00-00")) Then
       blnOK = True
    Else
       If PB_regUserinfoU.UserAuthority >= "50" Then
          blnOK = True
       End If
    End If
    '기초마감년월
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.자재기초마감년월 AS 자재기초마감년월, T1.회계기초마감년월 AS 회계기초마감년월, " _
                  & "T1.미지급금기초마감년월 AS 미지급금기초마감년월, T1.미수금기초마감년월 AS 미수금기초마감년월 " _
              & "FROM 사업장 T1 WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
    On Error GoTo ERROR_MONTH_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 1 Then
       If optGbn1.Value = True Then
          strMagamGbn = optGbn1.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("자재기초마감년월") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn2.Value = True Then
          strMagamGbn = optGbn2.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("회계기초마감년월") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn3.Value = True Then
          strMagamGbn = optGbn3.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("미지급금기초마감년월") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn4.Value = True Then
          strMagamGbn = optGbn4.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("미수금기초마감년월") Then
             blnOK = True
          Else
             blnOK = False
          End If
       End If
       P_adoRec.Close
    Else
      P_adoRec.Close
      Screen.MousePointer = vbDefault
      Exit Sub
    End If
    If blnOK = True Then
       intRetVal = MsgBox("월마감 작업(" + strMagamGbn + ")을 월마감구분으로 작업을 취소하시겠습니까 ?", vbCritical + vbYesNo + vbDefaultButton2, "월마감 작업 취소")
    Else
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If intRetVal = vbNo Then
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    cmdExec.Enabled = False
    strNextYM = CStr(Year(DateAdd("y", 1, dtpT_Date.Value))) + "00"
    Screen.MousePointer = vbHourglass
    '월마감
    PB_adoCnnSQL.BeginTrans
    If optGbn1.Value = True Then '1.자재마감
       '1. 자재원장마감
       strSQL = "DELETE 자재원장마감 WHERE 마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                    & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
    ElseIf _
       optGbn2.Value = True Then '2.회계전표내역마감
       '2. 회계전표내역마감
       strSQL = "DELETE 회계전표내역마감 WHERE 마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                        & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
    ElseIf _
       optGbn3.Value = True Then '3.미지급금마감
       '3. 미지급금원장마감
       strSQL = "DELETE 미지급금원장마감 WHERE 마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                        & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
    Else                         '4.미수금마감
       '4. 미수금원장마감
       strSQL = "DELETE 미수금원장마감 WHERE 마감년월 = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                      & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
    End If
    '년마감
    If Mid(DTOS(dtpT_Date.Value), 5, 2) = "12" Then
       If optGbn1.Value = True Then '1.자재마감
          '1. 자재원장마감
          strSQL = "DELETE 자재원장마감 WHERE 마감년월 = '" & strNextYM & "' " _
                                       & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT 자재원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.분류코드, " _
                         & "T1.세부코드, '" & strNextYM & "', " _
                         & "SUM(T1.입고누계수량), SUM(T1.출고누계수량), " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 자재원장마감 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND SUBSTRING(T1.마감년월, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.사업장코드, T1.분류코드, T1.세부코드 "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       ElseIf _
          optGbn2.Value = True Then '2.회계전표내역마감
          '2. 회계전표내역마감
          strSQL = "DELETE 회계전표내역마감 WHERE 마감년월 = '" & strNextYM & "' " _
                                           & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT 회계전표내역마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.계정코드, '" & strNextYM & "', " _
                         & "SUM(T1.입금누계금액), SUM(T1.출금누계금액), " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 회계전표내역마감 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND SUBSTRING(T1.마감년월, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.사업장코드, T1.계정코드 "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
          strSQL = "UPDATE 회계전표내역마감 SET " _
                        & "입금누계금액 = CASE WHEN (입금누계금액 > 출금누계금액) THEN (입금누계금액 - 출금누계금액) ELSE 0 END, " _
                        & "출금누계금액 = CASE WHEN (입금누계금액 < 출금누계금액) THEN (출금누계금액 - 입금누계금액) ELSE 0 END " _
                  & "WHERE 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND 마감년월 = '" & strNextYM & "' "
          On Error GoTo ERROR_MONTH_UPDATE
          PB_adoCnnSQL.Execute strSQL
       ElseIf _
          optGbn3.Value = True Then '3.미지급금마감
          '3. 미지급금원장마감
          strSQL = "DELETE 미지급금원장마감 WHERE 마감년월 = '" & strNextYM & "' " _
                                           & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT 미지급금원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매입처코드, " _
                         & "'" & strNextYM & "', SUM(T1.미지급금누계금액 - T1.미지급금지급누계금액), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 미지급금원장마감 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.매입처코드 <> '' " _
                    & "AND SUBSTRING(T1.마감년월, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.사업장코드, T1.매입처코드 "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       Else
          '4. 미수금원장마감
          strSQL = "DELETE 미수금원장마감 WHERE 마감년월 = '" & strNextYM & "' " _
                                         & "AND 사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT 미수금원장마감 " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.매출처코드, " _
                         & "'" & strNextYM & "', SUM(T1.미수금누계금액 - T1.미수금입금누계금액), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM 미수금원장마감 T1 " _
                  & "WHERE T1.사업장코드 = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.매출처코드 <> '' " _
                    & "AND SUBSTRING(T1.마감년월, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.사업장코드, T1.매출처코드 "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       End If
    End If
    PB_adoCnnSQL.CommitTrans
    cmdCancel.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "마감읽기 (서버와의 연결 실패)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "마감삭제 (서버와의 연결 실패)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "마감추가 (서버와의 연결 실패)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "서버와의 연결에 실패했습니다. 프로그램을 종료합니다.", vbCritical, "마감조정 (서버와의 연결 실패)"
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
    Set frm월마감작업 = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+

