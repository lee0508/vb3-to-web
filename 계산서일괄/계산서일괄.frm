VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm��꼭�ϰ� 
   BorderStyle     =   1  '���� ����
   Caption         =   "�� �� �� �� �� �� ó ��"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "��꼭�ϰ�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3106.538
   ScaleMode       =   0  '�����
   ScaleWidth      =   12255
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame1 
      Appearance      =   0  '���
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   12075
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   8640
         Style           =   2  '��Ӵٿ� ���
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
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� �� �� �� �� ó ��"
         BeginProperty Font 
            Name            =   "����ü"
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
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
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
         Picture         =   "��꼭�ϰ�.frx":030A
         Style           =   1  '�׷���
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   5580
         Picture         =   "��꼭�ϰ�.frx":0C58
         Style           =   1  '�׷���
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   4380
         Picture         =   "��꼭�ϰ�.frx":14DF
         Style           =   1  '�׷���
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtTaxBillMny 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
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
         Alignment       =   1  '������ ����
         Caption         =   "0"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1680
         TabIndex        =   3
         Top             =   1000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��꼭�ݾ�"
         Height          =   240
         Index           =   19
         Left            =   360
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "������ܾ�"
         Height          =   240
         Index           =   18
         Left            =   360
         TabIndex        =   18
         Top             =   1000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   17
         Left            =   360
         TabIndex        =   17
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   6
         Left            =   5760
         TabIndex        =   15
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
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
         Alignment       =   1  '������ ����
         Caption         =   "����Ⱓ"
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   640
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm��꼭�ϰ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ��꼭�ϰ�
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : �����, ����ó, �������⳻��, ���ݰ�꼭
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived         As Boolean
Private P_adoRec             As New ADODB.Recordset
Private P_adoRecW            As New ADODB.Recordset
Private P_intButton          As Integer
Private P_strFindString2     As String
Private Const PC_intRowCnt1  As Integer = 25   '�׸���1�� �� ������ �� ���(FixedRows ����)

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
Dim strSQL             As String
Dim inti               As Integer

Dim P                  As Printer
Dim strDefaultPrinter  As String
Dim aryPrinter()       As String
Dim strBuffer          As String

    frmMain.SBar.Panels(4).Text = "���ݸ���� ����ó�� ��꼭 �̹����� ���� ���迡�� ���ܵǸ�, ���⼼�ݰ�꼭��ο��� ����˴ϴ�. "
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
              'Case Is <= 10 '��ȸ
              '     cmdPrint.Enabled = False: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 20 '�μ�, ��ȸ
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = False: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 40 '�߰�, ����, �μ�, ��ȸ
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = False: cmdExit.Enabled = True
              'Case Is <= 50 '����, �߰�, ����, ��ȸ, �μ�
              '     cmdPrint.Enabled = True: cmdClear.Enabled = False: cmdFind.Enabled = True
              '     cmdSave.Enabled = True: cmdDelete.Enabled = True: cmdExit.Enabled = True
              'Case Is <= 99 '����, �߰�, ����, ��ȸ, �μ�
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "��꼭�Ǻ�(�������� ���� ����)"
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
'/// �������� ///
'+---------------+
Private Sub dtpP_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
'+---------------+
'/// ����Ⱓ ///
'+---------------+
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
Dim intRetVal      As Integer
    If KeyCode = vbKeyReturn Then
       intRetVal = MsgBox("���̼����ܾװ� ���ݼ��꼭�ݾ��� ����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "���̼����ܾ�/��꼭�ݾ� ���")
       If intRetVal = vbYes Then
          SubCompute_JanMny PB_regUserinfoU.UserBranchCode, DTOS(dtpT_Date.Value) '�̼����ܾװ��
          SubCompute_TaxBillMny PB_regUserinfoU.UserBranchCode, DTOS(dtpF_Date.Value), DTOS(dtpT_Date.Value) '���ݰ�꼭�ݾװ��
       End If
       SendKeys "{tab}"
    End If
End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
    SubClearText
End Sub

'+-----------+
'/// ���� ///
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
    intRetVal = MsgBox("���ݰ�꼭�� �ϰ� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "�ڷ� ����")
    If intRetVal = vbNo Then
       Exit Sub
    End If
    '�����ð� ���ϱ�
    cmdSave.Enabled = False
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) + Mid(P_adoRec("�����ð�"), 7, 2) _
                  + Mid(P_adoRec("�����ð�"), 10)
    P_adoRec.Close
    strTime = strServerTime
    
    PB_adoCnnSQL.BeginTrans
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.������ڵ�, T1.����ó�ڵ�, " _
                 & "(SELECT TOP 1 (S2.����� + SPACE(1) + S2.�԰�)  FROM �������⳻�� S1 " _
                    & "LEFT JOIN ���� S2 ON S1.�з��ڵ� = S2.�з��ڵ� AND S1.�����ڵ� = S2.�����ڵ� " _
                   & "WHERE S1.������ڵ� = T1.������ڵ� " _
                     & "AND S1.��������� BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
                     & "AND S1.������� = 2 AND S1.��뱸�� = 0 AND S1.���ݱ��� = 0 AND S1.��꼭���࿩�� = 0 " _
                     & "AND S1.����ó�ڵ� = T1.����ó�ڵ� " _
                   & "ORDER BY S1.���������, S1.�����ð�) AS ǰ��ױ԰�, " _
                 & "(SELECT (COUNT(S1.�����ڵ�) - 1) FROM �������⳻�� S1 " _
                   & "WHERE S1.������ڵ� = T1.������ڵ� " _
                     & "AND S1.��������� BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
                     & "AND S1.������� = 2 AND S1.��뱸�� = 0 AND S1.���ݱ��� = 0 AND S1.��꼭���࿩�� = 0 " _
                     & "AND S1.����ó�ڵ� = T1.����ó�ڵ�) AS ����, " _
                  & "ISNULL(SUM(ISNULL(T1.���ܰ�, 0) * ISNULL(T1.������, 0)), 0) AS ���ް���, " _
                     & "ROUND((ISNULL(SUM(ISNULL(T1.���ܰ�, 0) * ISNULL(T1.������, 0)), 0) * " & PB_curVatRate & "), 0, 1) AS ���� " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ���� T2 ON T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
              & "AND T1.��������� BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
              & "AND T1.������� = 2 AND T1.��뱸�� = 0 AND T1.���ݱ��� = 0 AND T1.��꼭���࿩�� = 0 " _
              & "AND T3.��꼭���࿩�� = 1 " _
            & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� " _
            & "ORDER BY T1.������ڵ�, T1.����ó�ڵ� "
           '& "HAVING SUM(ISNULL(T1.������, 0) * ISNULL(T1.���ܰ�, 0)) > 0 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          '���ݰ�꼭��ȣ ���ϱ�
          strMakeYear = Mid(DTOS(dtpP_Date.Value), 1, 4)
          strSQL = "spLogCounter '���ݰ�꼭', '" & PB_regUserinfoU.UserBranchCode + strMakeYear & "', " _
                              & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
          On Error GoTo ERROR_STORED_PROCEDURE
          P_adoRecW.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
          lngLogCnt1 = P_adoRecW(0)
          lngLogCnt2 = P_adoRecW(1)
          P_adoRecW.Close
          '���ݰ�꼭 �߰�
          strSQL = "INSERT INTO ���ݰ�꼭(������ڵ�, �ۼ��⵵, å��ȣ, �Ϸù�ȣ, " _
                                        & "����ó�ڵ�, �ۼ�����, ǰ��ױ԰�, ����, " _
                                        & "���ް���, ����, �ݾױ���, ��û����, " _
                                        & "���࿩��, �ۼ�����, �̼�����, ����, ��뱸��, " _
                                        & "��������, ������ڵ�) VALUES(" _
          & "'" & PB_regUserinfoU.UserBranchCode & "', '" & strMakeYear & "', " & lngLogCnt1 & ", " & lngLogCnt2 & "," _
          & "'" & P_adoRec("����ó�ڵ�") & "', '" & DTOS(dtpP_Date.Value) & "','" & P_adoRec("ǰ��ױ԰�") & "'," & P_adoRec("����") & ", " _
          & "" & P_adoRec("���ް���") & ", " & P_adoRec("����") & ", 3, 1, " _
          & "0, 2, 1, '', 0, " _
          & "'" & PB_regUserinfoU.UserServerDate & "','" & PB_regUserinfoU.UserCode & "')"
          On Error GoTo ERROR_TABLE_INSERT
          PB_adoCnnSQL.Execute strSQL
          '2.���⼼�ݰ�꼭���(�ۼ��ð�:strTime)
          strSQL = "INSERT INTO ���⼼�ݰ�꼭��� " _
                 & "SELECT T1.������ڵ�, T1.�ۼ�����, '" & strTime & "', T1.����ó�ڵ�, " _
                        & "T1.ǰ��ױ԰�, T1.����, T1.���ް���, T1.����, " _
                        & "T1.�ݾױ���, T1.��û����, T1.���࿩��, T1.�ۼ�����, " _
                        & "T1.�̼�����, T1.����, T1.��뱸��, T1.��������, " _
                        & "T1.������ڵ�, T1.�ۼ��⵵, T1.å��ȣ, T1.�Ϸù�ȣ " _
                   & "FROM ���ݰ�꼭 T1 " _
                  & "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                    & "AND T1.�ۼ��⵵  = '" & strMakeYear & "' AND å��ȣ = " & lngLogCnt1 & " AND T1.�Ϸù�ȣ = " & lngLogCnt2 & " "
          On Error GoTo ERROR_TABLE_INSERT
          PB_adoCnnSQL.Execute strSQL
    
          strSQL = "UPDATE �������⳻�� SET " _
                        & "�ۼ��⵵ = '" & strMakeYear & "', å��ȣ = " & lngLogCnt1 & ", �Ϸù�ȣ = " & lngLogCnt2 & ", " _
                        & "��꼭���࿩�� = 1 " _
                  & "WHERE ������ڵ� = '" & P_adoRec("������ڵ�") & "' AND ����ó�ڵ� = '" & P_adoRec("����ó�ڵ�") & "' " _
                    & "AND ��������� BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
                    & "AND ������� = 2 AND ��뱸�� = 0 AND ���ݱ��� = 0 AND ��꼭���࿩�� = 0 "
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_TABLE_SELECT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�˻� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "���� ����"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�α� ���� ����"
    Unload Me
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
    Set frm��꼭�ϰ� = Nothing
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �б� ����"
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
'/// �̼����ܾ� ///
'+-----------------+
Private Sub SubCompute_JanMny(strBranchCode As String, strT_Date As String)
Dim strSQL As String
    lblJanMny.Caption = "0.00"
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.����ó�ڵ� AS ����ó�ڵ�, " _
                  & "SUM(T1.�̼��ݴ���ݾ�) AS �̼��ݱݾ�, SUM(T1.�̼����Աݴ���ݾ�) AS �̼����Աݱݾ� " _
             & "FROM �̼��ݿ��帶�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' " _
              & "AND T1.������� >= (SUBSTRING('" & strT_Date & "', 1, 4) + '00') " _
              & "AND T1.������� < SUBSTRING('" & strT_Date & "', 1, 6) " _
            & "GROUP BY T1.����ó�ڵ� " _
            & "UNION ALL "
    If PB_regUserinfoU.UserMSGbn = "1" Then  '�̼��ݹ߻����� 1.��ǥ�̸�
       strSQL = strSQL + "SELECT T1.����ó�ڵ� AS ����ó�ڵ�, " _
                  & "(SUM(T1.������ * T1.���ܰ�) * " & (PB_curVatRate + 1) & ") AS �̼��ݱݾ�, 0 AS �̼����Աݱݾ� " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' AND T1.��뱸�� = 0 " _
              & "AND T1.��������� BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
              & "AND (T1.������� = 2 OR T1.������� = 8) " _
            & "GROUP BY T1.����ó�ڵ� "
            '& "AND (T1.���ݱ��� = 0 OR (T1.���ݱ��� = 1 AND T1.��꼭���࿩�� = 1)) "
    Else                                     '�̼��ݹ߻����� 2.��꼭�̸�
       strSQL = strSQL _
           & "SELECT T1.����ó�ڵ� AS ����ó�ڵ�, " _
                 & "(SUM(T1.���ް��� + T1.����)) AS �̼��ݱݾ�, 0 AS �̼����Աݱݾ� " _
             & "FROM ���⼼�ݰ�꼭��� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' AND T1.��뱸�� = 0 " _
              & "AND T1.�ۼ����� BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
              & "AND T1.�̼����� = 1 " _
            & "GROUP BY T1.����ó�ڵ� "
    End If
    strSQL = strSQL + "UNION ALL " _
           & "SELECT T1.����ó�ڵ� AS ����ó�ڵ�, " _
                  & "0 AS �̼��ݱݾ�, " _
                  & "ISNULL(SUM(T1.�̼����Աݱݾ�), 0) As �̼����Աݱݾ� " _
             & "FROM �̼��ݳ��� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' " _
              & "AND T1.�̼����Ա����� BETWEEN (SUBSTRING('" & strT_Date & "', 1, 6) + '01') AND '" & strT_Date & "' " _
            & "GROUP BY T1.����ó�ڵ� " _
            & "ORDER BY T1.����ó�ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          lblJanMny.Caption = Format(Vals(lblJanMny.Caption) + P_adoRec("�̼��ݱݾ�") - P_adoRec("�̼����Աݱݾ�"), "#,0.00")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "TABLE �б� ����"
    Unload Me
    Exit Sub
End Sub

'+---------------------+
'/// ���ݰ�꼭�ݾ� ///
'+---------------------+
Private Sub SubCompute_TaxBillMny(strBranchCode As String, strF_Date As String, strT_Date As String)
Dim strSQL As String
    txtTaxBillMny.Text = "0.00"
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    'strSQL = "SELECT ISNULL(SUM(ISNULL(T1.������, 0) * ISNULL(T1.���ܰ�, 0)), 0) AS ���ݰ�꼭�ݾ� " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ����ó T2 " _
               & "ON T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' " _
              & "AND T1.��������� BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
              & "AND T1.������� = 2 AND T1.��뱸�� = 0 AND T1.���ݱ��� = 0 AND T1.��꼭���࿩�� = 0 " _
              & "AND T2.��꼭���࿩�� = 1 "
    strSQL = "SELECT T1.����ó�ڵ� AS ����ó�ڵ�, ISNULL(SUM(ISNULL(T1.������, 0) * ISNULL(T1.���ܰ�, 0)), 0) AS ���ݰ�꼭�ݾ� " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ����ó T2 " _
               & "ON T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' " _
              & "AND T1.��������� BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
              & "AND T1.������� = 2 AND T1.��뱸�� = 0 AND T1.���ݱ��� = 0 AND T1.��꼭���࿩�� = 0 " _
              & "AND T2.��꼭���࿩�� = 1 " _
            & "GROUP BY T1.����ó�ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          txtTaxBillMny.Text = Format(Vals(txtTaxBillMny) + P_adoRec("���ݰ�꼭�ݾ�"), "#,0.00")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "TABLE �б� ����"
    Unload Me
    Exit Sub
End Sub

