VERSION 5.00
Begin VB.Form frm�ڷ��� 
   BorderStyle     =   1  '���� ����
   Caption         =   "�ڷ� ���"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "�ڷ���.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3255
   StartUpPosition =   1  '������ ���
   Begin VB.Frame Frame1 
      Caption         =   "[ ����� ���� ]"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "����� ������ �̸� ���¸� ����"
      Top             =   100
      Width           =   3015
      Begin VB.OptionButton optName 
         Caption         =   "yyyy/mm/dd hh:mm:ss (��)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton optName 
         Caption         =   "yyyy/mm/dd hh       (��)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton optName 
         Caption         =   "yyyy/mm/dd          (��)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "����(&E)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frm�ڷ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : �ڷ���
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
Dim strSQL      As String
Dim strDateTime As String '�����Ͻ�
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
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
Private Sub cmdExec_Click()
Dim strSQL            As String
Dim lngR              As Long
Dim lngC              As Long
Dim blnOK             As Boolean
Dim intRetVal         As Integer
Dim strNextYM         As String
Dim strServerDateTime As String
Dim strExecDateTime   As String
Dim strToBackUpPath   As String
Dim strToBackUpFile   As String
Dim intParR_Status    As Integer
Dim strParR_MSG       As String
    P_adoRec.CursorLocation = adUseClient
    'PATH
    strSQL = "SELECT ISNULL(T1.�������, '') AS ������� FROM ����� T1 WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 1 Then
       strToBackUpPath = P_adoRec("�������")
    End If
    P_adoRec.Close
    If Len(strToBackUpPath) = 0 Then
       MsgBox "����������� ��������� ����Ȯ���ϼ���!", vbCritical + vbOKOnly, "������� ����"
       Exit Sub
    End If
    '�����Ͻ�
    strSQL = "SELECT CONVERT(VARCHAR(19),GETDATE(), 120) AS �����Ͻ� "
    On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    strServerDateTime = P_adoRec("�����Ͻ�")
    P_adoRec.Close
    For lngR = optName.LBound To optName.UBound
        If optName(lngR).Value = True Then
           If lngR = 0 Then '��
              strServerDateTime = Format(DTOS(Mid(strServerDateTime, 1, 10)), "0000/00/00")
              strExecDateTime = DTOS(Mid(strServerDateTime, 1, 10))
           ElseIf _
              lngR = 1 Then '��
              strServerDateTime = Format(DTOS(Mid(strServerDateTime, 1, 10)), "0000/00/00") + " " + Mid(strServerDateTime, 12, 2)
              strExecDateTime = DTOS(Mid(strServerDateTime, 1, 10)) + Mid(strServerDateTime, 12, 2)
           Else             '��
              strServerDateTime = Format(DTOS(Mid(strServerDateTime, 1, 10)), "0000/00/00") + " " + Right(strServerDateTime, 8)
              strExecDateTime = DTOS(Mid(strServerDateTime, 1, 10)) + Format(Right(strServerDateTime, 8), "hhmmss")
           End If
        End If
    Next lngR
    intRetVal = MsgBox("[" + strServerDateTime + "] �ڷḦ ����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton2, "�ڷ���")
    If intRetVal = vbNo Then
       Exit Sub
    End If
    cmdExec.Enabled = False
    Screen.MousePointer = vbHourglass
    PB_adoCnnSQL.BeginTrans
    '�ڷ���
    strToBackUpFile = strExecDateTime + ".bak"
    strSQL = "spBackUpYmhDB '" & strToBackUpPath & "', '" & strToBackUpFile & "', '', 0, '' "
    On Error GoTo ERROR_STORED_PROCEDURE
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
    Else
       intParR_Status = P_adoRec(0)
       strParR_MSG = P_adoRec(1)
    End If
    P_adoRec.Close
    '��� ����
    MsgBox strParR_MSG, IIf(intParR_Status = 0, vbCritical, vbInformation), "�ڷ� ��� ���"
    PB_adoCnnSQL.CommitTrans
    cmdExec.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "(�������� ���� ����)"
    Unload Me
    Exit Sub
ERROR_STORED_PROCEDURE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "�ڷ� ��� ����"
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
    Set frm�ڷ��� = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
