VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�������۾� 
   BorderStyle     =   1  '���� ����
   Caption         =   "�������۾�"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "�������۾�.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4080
   StartUpPosition =   1  '������ ���
   Begin VB.OptionButton optGbn4 
      Caption         =   "�̼���"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "�̼��ݳ���[��ΰ���]"
      Top             =   915
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�������(&C)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optGbn1 
      Caption         =   "����"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "����(��)����[������]"
      Top             =   550
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optGbn2 
      Caption         =   "ȸ��"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      ToolTipText     =   "ȸ����ǥ����[ȸ�����]"
      Top             =   550
      Width           =   735
   End
   Begin VB.OptionButton optGbn3 
      Caption         =   "�����ޱ�"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "�����ޱݳ���[��ΰ���]"
      Top             =   915
      Width           =   1095
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "��������(&E)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
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
      Alignment       =   1  '������ ����
      Caption         =   "����������"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Caption         =   "����������"
      Height          =   240
      Index           =   10
      Left            =   120
      TabIndex        =   6
      Top             =   165
      Width           =   1095
   End
End
Attribute VB_Name = "frm�������۾�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : �������۾�
' ���� Control :
' ������ Table   :
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 22  '�׸��� �� ������ �� ���(FixedRows ����)

'+--------------------------------+
'/// LOAD FORM ( �ѹ��� ���� ) ///
'+--------------------------------+
Private Sub Form_Load()
    P_blnActived = False
End Sub

'+-------------------------------------------+
'/// ACTIVATE FORM Ȱ��ȭ ( �ѹ��� ���� ) ///
'+-------------------------------------------+
Private Sub Form_Activate()
Dim SQL     As String
    frmMain.SBar.Panels(4).Text = "�ݿ������� �ڷḦ ������ ��� �ݵ�� ������ �۾��� �ش� �������������� �۾��ϼž߸� �մϴ�."
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// ���� ///
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
'/// ���� ///
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
    '���ʸ������
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ʸ������ AS ������ʸ������, T1.ȸ����ʸ������ AS ȸ����ʸ������, " _
                  & "T1.�����ޱݱ��ʸ������ AS �����ޱݱ��ʸ������, T1.�̼��ݱ��ʸ������ AS �̼��ݱ��ʸ������ " _
              & "FROM ����� T1 WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
    On Error GoTo ERROR_MONTH_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 1 Then
       If optGbn1.Value = True Then
          strMagamGbn = optGbn1.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("������ʸ������") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn2.Value = True Then
          strMagamGbn = optGbn2.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("ȸ����ʸ������") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn3.Value = True Then
          strMagamGbn = optGbn3.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("�����ޱݱ��ʸ������") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn4.Value = True Then
          strMagamGbn = optGbn4.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("�̼��ݱ��ʸ������") Then
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
       intRetVal = MsgBox("������ �۾�(" + strMagamGbn + ")�� �������������� �۾��Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton2, "������ �۾�")
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
    '������
    PB_adoCnnSQL.BeginTrans
    If optGbn1.Value = True Then '1.���縶��
       '1. ������帶��
       strSQL = "DELETE ������帶�� WHERE ������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                    & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
       strSQL = "INSERT ������帶�� " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.�з��ڵ�, " _
                      & "T1.�����ڵ�, '" & Left(DTOS(dtpT_Date.Value), 6) & "', " _
                      & "SUM(T1.�԰����), SUM(T1.������), " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM �������⳻�� T1 " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.��뱸�� = 0 " _
                 & "AND SUBSTRING(T1.���������, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
               & "GROUP BY T1.������ڵ�, T1.�з��ڵ�, T1.�����ڵ� "
               '& "AND (T1.������� BETWEEN 1 AND 6) "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
    ElseIf _
       optGbn2.Value = True Then '2.ȸ����ǥ��������
       '2. ȸ����ǥ��������
       strSQL = "DELETE ȸ����ǥ�������� WHERE ������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                        & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
       strSQL = "INSERT ȸ����ǥ�������� " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.�����ڵ�, '" & Left(DTOS(dtpT_Date.Value), 6) & "', " _
                      & "SUM(T1.�Աݱݾ�), SUM(T1.��ݱݾ�), " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM ȸ����ǥ���� T1 " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND SUBSTRING(T1.�ۼ�����, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                 & "AND T1.��뱸�� = 0 " _
               & "GROUP BY T1.������ڵ�, T1.�����ڵ� "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
    ElseIf _
       optGbn3.Value = True Then '3.�����ޱݸ���
       '3. �����ޱݿ��帶��
       strSQL = "DELETE �����ޱݿ��帶�� WHERE ������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                        & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
       If PB_regUserinfoU.UserMJGbn = "1" Then '�����ޱݹ߻����� 1.��ǥ, 2.(����)��꼭
          strSQL = "INSERT �����ޱݿ��帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                         & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', SUM(T1.�԰����*(T1.�԰�ܰ�+T1.�԰�ΰ�)), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM �������⳻�� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.��뱸�� = 0 AND T1.����ó�ڵ� <> '' " _
                    & "AND (T1.������� = 1) AND T1.���ݱ��� = 0 " _
                    & "AND SUBSTRING(T1.���������, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
       Else
          strSQL = "INSERT �����ޱݿ��帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                         & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', SUM(T1.���ް��� + T1.����), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM ���Լ��ݰ�꼭��� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.��뱸�� = 0 AND T1.����ó�ڵ� <> '' " _
                    & "AND T1.�����ޱ��� = 1 " _
                    & "AND SUBSTRING(T1.�ۼ�����, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
       End If
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
       strSQL = "UPDATE �����ޱݿ��帶�� SET " _
                     & "�����ޱ����޴���ݾ� = �����ޱ����޴���ݾ� + Z.F5 " _
                & "FROM " _
                    & "(SELECT '" & PB_regUserinfoU.UserBranchCode & "' AS F0, T1.����ó�ڵ� AS F2, " _
                             & "'" & Left(DTOS(dtpT_Date.Value), 6) & "' AS F3, 0 AS F4, " _
                             & "SUM(T1.�����ޱ����ޱݾ�) AS F5, " _
                             & "'" & PB_regUserinfoU.UserServerDate & "' AS F6, '" & PB_regUserinfoU.UserCode & "' AS F7 " _
                       & "FROM �����ޱݳ��� T1 " _
                      & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                        & "AND SUBSTRING(T1.�����ޱ���������, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                        & "AND EXISTS " _
                           & "(SELECT T2.������� " _
                              & "FROM �����ޱݿ��帶�� T2 " _
                             & "WHERE T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� " _
                               & "AND T2.������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' ) " _
                      & "GROUP BY T1.����ó�ڵ�) AS Z " _
               & "WHERE ������ڵ� = Z.F0 AND ����ó�ڵ� = Z.F2 AND ������� = Z.F3 "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
       strSQL = "INSERT �����ޱݿ��帶�� " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                      & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', 0, " _
                      & "SUM(T1.�����ޱ����ޱݾ�), " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM �����ޱݳ��� T1 " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND SUBSTRING(T1.�����ޱ���������, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                 & "AND NOT EXISTS " _
                        & "(SELECT T2.������� " _
                           & "FROM �����ޱݿ��帶�� T2 " _
                          & "WHERE T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� " _
                            & "AND T2.������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' ) " _
               & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
    Else                         '4.�̼��ݸ���
       '4. �̼��ݿ��帶��
       strSQL = "DELETE �̼��ݿ��帶�� WHERE ������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                      & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
       If PB_regUserinfoU.UserMSGbn = "1" Then '�̼��ݹ߻����� 1.��ǥ, 2.(����)��꼭
          strSQL = "INSERT �̼��ݿ��帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                         & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', (SUM(T1.������*T1.���ܰ�) * " & (PB_curVatRate + 1) & ") , " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM �������⳻�� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.��뱸�� = 0 AND T1.����ó�ڵ� <> '' " _
                    & "AND (T1.������� = 2 OR T1.������� = 8) " _
                    & "AND SUBSTRING(T1.���������, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
                  '& "AND (T1.���ݱ��� = 0 OR (T1.���ݱ��� = 1 AND T1.��꼭���࿩�� = 1)) "
       Else
          strSQL = "INSERT �̼��ݿ��帶�� " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                      & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', SUM(T1.���ް��� + T1.����), " _
                      & "0, " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM ���⼼�ݰ�꼭��� T1 " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.��뱸�� = 0 AND T1.����ó�ڵ� <> '' " _
                 & "AND T1.�̼����� = 1 " _
                 & "AND SUBSTRING(T1.�ۼ�����, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
               & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
          'strSQL = "INSERT �̼��ݿ��帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                         & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', SUM(T1.���ް��� + T1.����) , " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM ���ݰ�꼭 T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.��뱸�� = 0 AND T1.����ó�ڵ� <> '' " _
                    & "AND (T1.�̼����� = 1) " _
                    & "AND SUBSTRING(T1.�ۼ�����, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
       End If
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
       strSQL = "UPDATE �̼��ݿ��帶�� SET " _
                     & "�̼����Աݴ���ݾ� = �̼����Աݴ���ݾ� + Z.F5 " _
                & "FROM " _
                    & "(SELECT '" & PB_regUserinfoU.UserBranchCode & "' AS F0, T1.����ó�ڵ� AS F2, " _
                             & "'" & Left(DTOS(dtpT_Date.Value), 6) & "' AS F3, 0 AS F4, " _
                             & "SUM(T1.�̼����Աݱݾ�) AS F5, " _
                             & "'" & PB_regUserinfoU.UserServerDate & "' AS F6, '" & PB_regUserinfoU.UserCode & "' AS F7 " _
                       & "FROM �̼��ݳ��� T1 " _
                      & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                        & "AND SUBSTRING(T1.�̼����Ա�����, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                        & "AND EXISTS " _
                           & "(SELECT T2.������� " _
                              & "FROM �̼��ݿ��帶�� T2 " _
                             & "WHERE T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� " _
                               & "AND T2.������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' ) " _
                      & "GROUP BY T1.����ó�ڵ�) AS Z " _
               & "WHERE ������ڵ� = Z.F0 AND ����ó�ڵ� = Z.F2 AND ������� = Z.F3 "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
       strSQL = "INSERT �̼��ݿ��帶�� " _
              & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                      & "'" & Left(DTOS(dtpT_Date.Value), 6) & "', 0, " _
                      & "SUM(T1.�̼����Աݱݾ�), " _
                      & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                & "FROM �̼��ݳ��� T1 " _
               & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND SUBSTRING(T1.�̼����Ա�����, 1, 6) = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                 & "AND NOT EXISTS " _
                        & "(SELECT T2.������� " _
                           & "FROM �̼��ݿ��帶�� T2 " _
                          & "WHERE T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� " _
                            & "AND T2.������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' ) " _
               & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
       On Error GoTo ERROR_MONTH_INSERT
       PB_adoCnnSQL.Execute strSQL
    End If
    '�⸶��
    If Mid(DTOS(dtpT_Date.Value), 5, 2) = "12" Then
       If optGbn1.Value = True Then '1.���縶��
          '1. ������帶��
          strSQL = "DELETE ������帶�� WHERE ������� = '" & strNextYM & "' " _
                                       & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT ������帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.�з��ڵ�, " _
                         & "T1.�����ڵ�, '" & strNextYM & "', " _
                         & "SUM(T1.�԰������), SUM(T1.��������), " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM ������帶�� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND SUBSTRING(T1.�������, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.�з��ڵ�, T1.�����ڵ� "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       ElseIf _
          optGbn2.Value = True Then '2.ȸ����ǥ��������
          '2. ȸ����ǥ��������
          strSQL = "DELETE ȸ����ǥ�������� WHERE ������� = '" & strNextYM & "' " _
                                           & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT ȸ����ǥ�������� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.�����ڵ�, '" & strNextYM & "', " _
                         & "SUM(T1.�Աݴ���ݾ�), SUM(T1.��ݴ���ݾ�), " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM ȸ����ǥ�������� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND SUBSTRING(T1.�������, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.�����ڵ� "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
          strSQL = "UPDATE ȸ����ǥ�������� SET " _
                        & "�Աݴ���ݾ� = CASE WHEN (�Աݴ���ݾ� > ��ݴ���ݾ�) THEN (�Աݴ���ݾ� - ��ݴ���ݾ�) ELSE 0 END, " _
                        & "��ݴ���ݾ� = CASE WHEN (�Աݴ���ݾ� < ��ݴ���ݾ�) THEN (��ݴ���ݾ� - �Աݴ���ݾ�) ELSE 0 END " _
                  & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND ������� = '" & strNextYM & "' "
          On Error GoTo ERROR_MONTH_UPDATE
          PB_adoCnnSQL.Execute strSQL
       ElseIf _
          optGbn3.Value = True Then '3.�����ޱݸ���
          '3. �����ޱݿ��帶��
          strSQL = "DELETE �����ޱݿ��帶�� WHERE ������� = '" & strNextYM & "' " _
                                           & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT �����ޱݿ��帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                         & "'" & strNextYM & "', SUM(T1.�����ޱݴ���ݾ� - T1.�����ޱ����޴���ݾ�), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM �����ޱݿ��帶�� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.����ó�ڵ� <> '' " _
                    & "AND SUBSTRING(T1.�������, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       Else
          '4. �̼��ݿ��帶��
          strSQL = "DELETE �̼��ݿ��帶�� WHERE ������� = '" & strNextYM & "' " _
                                         & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT �̼��ݿ��帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                         & "'" & strNextYM & "', SUM(T1.�̼��ݴ���ݾ� - T1.�̼����Աݴ���ݾ�), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM �̼��ݿ��帶�� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.����ó�ڵ� <> '' " _
                    & "AND SUBSTRING(T1.�������, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����б� (�������� ���� ����)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�������� (�������� ���� ����)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����߰� (�������� ���� ����)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�������� (�������� ���� ����)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

'+-----------+
'/// ��� ///
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
    '���ʸ������
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ʸ������ AS ������ʸ������, T1.ȸ����ʸ������ AS ȸ����ʸ������, " _
                  & "T1.�����ޱݱ��ʸ������ AS �����ޱݱ��ʸ������, T1.�̼��ݱ��ʸ������ AS �̼��ݱ��ʸ������ " _
              & "FROM ����� T1 WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
    On Error GoTo ERROR_MONTH_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 1 Then
       If optGbn1.Value = True Then
          strMagamGbn = optGbn1.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("������ʸ������") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn2.Value = True Then
          strMagamGbn = optGbn2.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("ȸ����ʸ������") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn3.Value = True Then
          strMagamGbn = optGbn3.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("�����ޱݱ��ʸ������") Then
             blnOK = True
          Else
             blnOK = False
          End If
       ElseIf _
          optGbn4.Value = True Then
          strMagamGbn = optGbn4.Caption
          If Mid(DTOS(dtpT_Date.Value), 1, 6) > P_adoRec("�̼��ݱ��ʸ������") Then
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
       intRetVal = MsgBox("������ �۾�(" + strMagamGbn + ")�� �������������� �۾��� ����Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton2, "������ �۾� ���")
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
    '������
    PB_adoCnnSQL.BeginTrans
    If optGbn1.Value = True Then '1.���縶��
       '1. ������帶��
       strSQL = "DELETE ������帶�� WHERE ������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                    & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
    ElseIf _
       optGbn2.Value = True Then '2.ȸ����ǥ��������
       '2. ȸ����ǥ��������
       strSQL = "DELETE ȸ����ǥ�������� WHERE ������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                        & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
    ElseIf _
       optGbn3.Value = True Then '3.�����ޱݸ���
       '3. �����ޱݿ��帶��
       strSQL = "DELETE �����ޱݿ��帶�� WHERE ������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                        & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
    Else                         '4.�̼��ݸ���
       '4. �̼��ݿ��帶��
       strSQL = "DELETE �̼��ݿ��帶�� WHERE ������� = '" & Left(DTOS(dtpT_Date.Value), 6) & "' " _
                                      & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
       On Error GoTo ERROR_MONTH_DELETE
       PB_adoCnnSQL.Execute strSQL
    End If
    '�⸶��
    If Mid(DTOS(dtpT_Date.Value), 5, 2) = "12" Then
       If optGbn1.Value = True Then '1.���縶��
          '1. ������帶��
          strSQL = "DELETE ������帶�� WHERE ������� = '" & strNextYM & "' " _
                                       & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT ������帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.�з��ڵ�, " _
                         & "T1.�����ڵ�, '" & strNextYM & "', " _
                         & "SUM(T1.�԰������), SUM(T1.��������), " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM ������帶�� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND SUBSTRING(T1.�������, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.�з��ڵ�, T1.�����ڵ� "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       ElseIf _
          optGbn2.Value = True Then '2.ȸ����ǥ��������
          '2. ȸ����ǥ��������
          strSQL = "DELETE ȸ����ǥ�������� WHERE ������� = '" & strNextYM & "' " _
                                           & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT ȸ����ǥ�������� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.�����ڵ�, '" & strNextYM & "', " _
                         & "SUM(T1.�Աݴ���ݾ�), SUM(T1.��ݴ���ݾ�), " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM ȸ����ǥ�������� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND SUBSTRING(T1.�������, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.�����ڵ� "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
          strSQL = "UPDATE ȸ����ǥ�������� SET " _
                        & "�Աݴ���ݾ� = CASE WHEN (�Աݴ���ݾ� > ��ݴ���ݾ�) THEN (�Աݴ���ݾ� - ��ݴ���ݾ�) ELSE 0 END, " _
                        & "��ݴ���ݾ� = CASE WHEN (�Աݴ���ݾ� < ��ݴ���ݾ�) THEN (��ݴ���ݾ� - �Աݴ���ݾ�) ELSE 0 END " _
                  & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND ������� = '" & strNextYM & "' "
          On Error GoTo ERROR_MONTH_UPDATE
          PB_adoCnnSQL.Execute strSQL
       ElseIf _
          optGbn3.Value = True Then '3.�����ޱݸ���
          '3. �����ޱݿ��帶��
          strSQL = "DELETE �����ޱݿ��帶�� WHERE ������� = '" & strNextYM & "' " _
                                           & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT �����ޱݿ��帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                         & "'" & strNextYM & "', SUM(T1.�����ޱݴ���ݾ� - T1.�����ޱ����޴���ݾ�), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM �����ޱݿ��帶�� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.����ó�ڵ� <> '' " _
                    & "AND SUBSTRING(T1.�������, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
          On Error GoTo ERROR_MONTH_INSERT
          PB_adoCnnSQL.Execute strSQL
       Else
          '4. �̼��ݿ��帶��
          strSQL = "DELETE �̼��ݿ��帶�� WHERE ������� = '" & strNextYM & "' " _
                                         & "AND ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
          On Error GoTo ERROR_MONTH_DELETE
          PB_adoCnnSQL.Execute strSQL
          strSQL = "INSERT �̼��ݿ��帶�� " _
                 & "SELECT '" & PB_regUserinfoU.UserBranchCode & "', T1.����ó�ڵ�, " _
                         & "'" & strNextYM & "', SUM(T1.�̼��ݴ���ݾ� - T1.�̼����Աݴ���ݾ�), " _
                         & "0, " _
                         & "'" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' " _
                   & "FROM �̼��ݿ��帶�� T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND T1.����ó�ڵ� <> '' " _
                    & "AND SUBSTRING(T1.�������, 1, 4) = '" & Left(DTOS(dtpT_Date.Value), 4) & "' " _
                  & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����б� (�������� ���� ����)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�������� (�������� ���� ����)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�����߰� (�������� ���� ����)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_MONTH_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�������� (�������� ���� ����)"
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

'+-----------+
'/// ���� ///
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
    Set frm�������۾� = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+

