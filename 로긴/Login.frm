VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "����ü"
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
   StartUpPosition =   2  'ȭ�� ���
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
         Name            =   "����ü"
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
      Caption         =   "��  ��(&X)"
      Height          =   375
      Left            =   1920
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdExec 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��  ��(&C)"
      Height          =   375
      Left            =   120
      Style           =   1  '�׷���
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtPasswd 
      Alignment       =   2  '��� ����
      Height          =   270
      IMEMode         =   3  '��� ����
      Left            =   1995
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1635
      Width           =   1095
   End
   Begin VB.TextBox txtUserId 
      Alignment       =   2  '��� ����
      Height          =   270
      Left            =   1995
      TabIndex        =   1
      Top             =   1245
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�޸�����ü"
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
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����� �۾����� :"
      Height          =   255
      Index           =   2
      Left            =   315
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "����� ��й�ȣ :"
      Height          =   255
      Index           =   1
      Left            =   315
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����� �ڵ��ȣ :"
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
'| 1. ����(Project ǥ�� exe : 1),2),3),4) �⺻
'|    1) Visual Basic For Applications -> msvbvm60.dll
'|    2) Visual Basic runtime objects and procedures -> msvbvm60.dll\3
'|    3) Visual Basic objects and procedures -> VB6.OLE
'|    4) OLE automation -> stdole2.tlb
'|
'|    5) Microsoft ActiveX Data Objects 2.1 Library -> msado21.tlb
'|    6) VideoSoft VSFlexGrid7.0 (OLEDB) -> Vsflex7.oca
'|    7) Microsoft Data Formatting Objects Library 6.0 -> MMSTDFMT.DLL
'|
'| 2. �������
'|    1) VideoSoft VSFlexGrid 7.0(OLEDB) -> Vsflex7.ocx
'|    2) Crystal Report Control -> Crystl32.OCX
'|    3) Microsoft Windows Common Controls 6.0 -> Mscomctl.ocx
'|    4) Microsoft Windows Common Controls-2 6.0(SP4) -> MSCPMCT2.OCX
'|       Animation, UpDown, MonthView, DTPicker, FlatscrollBar
'+-----------------------------------------------------------------------------+
Option Explicit
Private P_blnActived     As Boolean
Private P_adoRec         As New ADODB.Recordset
Private P_intConnWay     As Integer '0.��й�ȣ, 1.�����й�ȣ

Private Sub Form_Initialize()
    '
End Sub
'+---------------------------+
'| LOAD FORM ( �ѹ��� ���� )
'+---------------------------+
Private Sub Form_Load()
    P_blnActived = False
    If App.PrevInstance = True Then
       MsgBox "�̹� �������� ���α׷��Դϴ�.", vbCritical, "�Ǹ� ���� �ý���"
       End
    End If
End Sub

'+--------------------------------------------+
'| Server ���� ���
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
'| ACTIVATE FORM Ȱ��ȭ ( �ѹ��� �����ؾ� �� )
'+--------------------------------------------+
Private Sub Form_Activate()
Dim strSQL As String
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       PB_regUserinfoU = UserinfoU_Read()
       'Set Connection String
       'DoConnection
       '�������� ���ʿ���
       PB_Fnc_AdoCnnSQL
       If PB_varErrCode <> 0 Or PB_blnStatusOfConn <> True Then
          GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       End If
       PB_strSystemName = "�Ǹ� ���� �ý���": Me.Caption = PB_strSystemName
       lblTitle.Caption = PB_regUserinfoU.UserBranchName
       '�����ð��� ������
       P_adoRec.CursorLocation = adUseClient
       strSQL = "SELECT CONVERT(VARCHAR(8),GETDATE(), 112) AS ��������, " _
                     & "RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       PB_regUserinfoU.UserServerDate = P_adoRec("��������")
       PB_regUserinfoU.UserServerTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) _
                                      + Mid(P_adoRec("�����ð�"), 7, 2) + Mid(P_adoRec("�����ð�"), 10)
       dtpClientDate.Format = dtpShortDate
       dtpClientDate.Value = Format(P_adoRec("��������"), "0000-00-00")
       P_adoRec.Close
       txtUserId.SetFocus
    End If
    P_blnActived = True
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox PB_varErrCode & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�Ǹ� ���� �ý��� (�������� ���� ����)"
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
'| ����� �۾�����
'+-----------------+
Private Sub dtpClientDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}"
       Exit Sub
    End If
End Sub
'+-----------------+
'| ����� �ڵ��ȣ
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
'| ����� �α��й�ȣ
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
'| ����(����)
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
       strWhere = "AND T1.�α��κ�й�ȣ = '" & Trim(txtPasswd.Text) & "' "
    Else
       strWhere = "AND T1.�����й�ȣ = '" & Trim(txtPasswd.Text) & "' "
    End If
    strSQL = "SELECT T1.������ڵ� AS ������ڵ�, T1.����ڸ� AS ����ڸ�, " _
                  & "T1.������ڵ� AS ������ڵ�, T1.�����й�ȣ AS �����й�ȣ, " _
                  & "T1.����ڱ��� AS ����ڱ���, T1.�α��ο��� AS �α��ο���," _
                  & "T2.�ΰ����� AS �ΰ�����, T2.������ AS ������, T2.����ڹ�ȣ AS ����ڹ�ȣ, T2.��ǥ�ڸ� AS ��ǥ�ڸ�, " _
                  & "(T2.�ּ� + SPACE(1) + T2.����) AS �ּҹ���, T2.���� AS ����, T2.���� AS ����, " _
                  & "T2.�����ޱݹ߻����� AS �����ޱݹ߻�����, T2.�̼��ݹ߻����� AS �̼��ݹ߻�����, " _
                  & "T2.���Ÿ�Ա��� AS ���Ÿ�Ա���, T2.�ŷ�������ܸ��� AS �ŷ�������ܸ���, " _
                  & "T2.�ŷ��������ʸ��� AS �ŷ��������ʸ���, T2.���ݰ�꼭��ܸ��� AS ���ݰ�꼭��ܸ���, " _
                  & "T2.���ݰ�꼭���ʸ��� AS ���ݰ�꼭���ʸ���, " _
                  & "T2.�����԰�ܰ��ڵ����ű��� AS �����԰�ܰ��ڵ����ű���, T2.�������ܰ��ڵ����ű��� AS �������ܰ��ڵ����ű��� " _
             & "FROM ����� T1 " _
            & "INNER JOIN ����� T2 " _
                    & "ON T1.������ڵ� = T2.������ڵ� " _
            & "WHERE T1.������ڵ� = '" & Trim(txtUserId.Text) & "' " _
              & "" & strWhere & " " _
              & "AND T1.��뱸�� = 0 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       MsgBox "������ڵ�� �α��κ�й�ȣ�� �ٽ� �Է��ϼ���.", vbCritical, "����� �Է� ����"
       P_adoRec.Close
       txtUserId.SetFocus
       Screen.MousePointer = vbDefault
       Exit Sub
    Else
       If P_intConnWay = 0 Then
          If UPPER(P_adoRec("�α��ο���")) = "Y" Then
             MsgBox "�̹� �α��ε� ������Դϴ�. ��й�ȣ��ſ� �����й�ȣ�� �Է��ϼ���).", vbCritical, "����� ���� ����"
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
    '���������� ���
    '+--------------+
    PB_curVatRate = (P_adoRec("�ΰ�����") / 100) '�ΰ�����
    
    PB_intIAutoPriceGbn = P_adoRec("�����԰�ܰ��ڵ����ű���")
    PB_intOAutoPriceGbn = P_adoRec("�������ܰ��ڵ����ű���")
    
    PB_strEnterNo = P_adoRec("����ڹ�ȣ")
    PB_strEnterName = P_adoRec("������")
    PB_strRepName = P_adoRec("��ǥ�ڸ�")
    PB_strEnterAddress = P_adoRec("�ּҹ���")
    PB_strUptae = P_adoRec("����")
    PB_strUpjong = P_adoRec("����")
    
    PB_intPrtTypeGbn = P_adoRec("���Ÿ�Ա���")
    PB_intDTopMargin = P_adoRec("�ŷ�������ܸ���")
    PB_intDLeftMargin = P_adoRec("�ŷ��������ʸ���")
    PB_intTTopMargin = P_adoRec("���ݰ�꼭��ܸ���")
    PB_intTLeftMargin = P_adoRec("���ݰ�꼭���ʸ���")
    
    'Registery ���
    'UserComputerName   As String   '1. WorkStation Name
    'UserClientName     As String   '2. Client Wondows Login Name
    'UserServerDate     As String   '3. ������� ��������
    'UserServerTime     As String   '4. ������� �����ð�
    'UserClientDate     As sting    '5. ���α׷� ��������
    'UserClientTime     As sting    '6. ���α׷� ����ð�
    '+------------------------+
    'UserBranchCode     As String   '7. ������ڵ�
    'UserBranchName     As String   '8. ������
    'UserCode           As String   '9. ������ڵ�
    'UserName           As String   '10.����ڼ���
    'UserLoginPasswd    As String   '11.����ں�й�ȣ
    'UserSanctionPasswd As String   '12.����ڰ����й�ȣ
    'UserAuthority      As String   '13.����ڱ���
    strLpBuffer = "-"
    lngCnt = GetComputerName(strLpBuffer, 256)
    PB_regUserinfoU.UserComputerName = Trim(strLpBuffer)
    lngCnt = GetUserName(strLpBuffer, 256)
    PB_regUserinfoU.UserClientName = Trim(strLpBuffer)
    'PB_regUserinfoU.UserServerDate = Form_Activate ����ÿ� �̹� ����
    'PB_regUserinfoU.UserServerTime = Form_Activate ����ÿ� �̹� ����
    PB_regUserinfoU.UserClientDate = Format(dtpClientDate.Value, "yyyymmdd") 'Format(Date, "yyyymmdd")
    PB_regUserinfoU.UserClientTime = PB_regUserinfoU.UserServerTime           'Format(Time, "hhmmss")
    '+------------------------+
    PB_regUserinfoU.UserBranchCode = P_adoRec("������ڵ�")
    PB_regUserinfoU.UserBranchName = P_adoRec("������")
    PB_regUserinfoU.UserCode = P_adoRec("������ڵ�")
    PB_regUserinfoU.UserName = P_adoRec("����ڸ�")
    PB_regUserinfoU.UserLoginPasswd = Trim(txtPasswd.Text)
    PB_regUserinfoU.UserName = P_adoRec("����ڸ�")
    PB_regUserinfoU.UserSanctionPasswd = P_adoRec("�����й�ȣ")
    PB_regUserinfoU.UserAuthority = P_adoRec("����ڱ���")
    '+------------------------+
    PB_regUserinfoU.UserMJGbn = P_adoRec("�����ޱݹ߻�����")
    PB_regUserinfoU.UserMSGbn = P_adoRec("�̼��ݹ߻�����")
    P_adoRec.Close
    UserinfoU_Save PB_regUserinfoU
    Screen.MousePointer = vbHourglass
    PB_adoCnnSQL.BeginTrans
    strSQL = "UPDATE ����� SET " _
                  & "�α��ο��� = 'Y', " _
                  & "�����Ͻ� = '" & PB_regUserinfoU.UserServerDate & PB_regUserinfoU.UserServerTime & "' " _
            & "WHERE ������ڵ� = '" & Trim(txtUserId.Text) & "' "
    On Error GoTo ERROR_TABLE_UPDATE
    PB_adoCnnSQL.Execute strSQL
    PB_adoCnnSQL.CommitTrans
    Screen.MousePointer = vbDefault
    frmLogin.Hide    '�α�ȭ�� ����� ����ȭ������ �̵�
    frmMain.Show     'vbModal
    Exit Sub
ERROR_TABLE_SELECT:
    If P_adoRec.State <> 0 Then
       P_adoRec.Close
    End If
    MsgBox Err.Description & "/" & strSQL
    MsgBox "���� ������ ������ �߻��Ͽ����ϴ�. ����ڿ��� �����Ͽ��ּ���.", vbCritical, "���� ���� ����1"
    Screen.MousePointer = vbDefault
    cmdExit_Click
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Description & "/" & strSQL
    MsgBox "���� ������ ������ �߻��Ͽ����ϴ�. ����ڿ��� �����Ͽ��ּ���.", vbCritical, "���� ���� ����2"
    Screen.MousePointer = vbDefault
    cmdExit_Click
    Exit Sub
End Sub

'+-------------------+
'| ���α׷� ���� ����
'+-------------------+
Private Sub cmdExit_Click()
    Unload frmLogin
End Sub

