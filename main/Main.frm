VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  '���� ����
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15315
   BeginProperty Font 
      Name            =   "����ü"
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
   ScaleMode       =   0  '�����
   ScaleWidth      =   15405
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  '�Ʒ� ����
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
            Text            =   "����"
            TextSave        =   "����"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17887
            MinWidth        =   17887
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
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
   Begin VB.Menu �����ڷ��������� 
      Caption         =   "�����ڷ���������(&1)"
      Begin VB.Menu ��������� 
         Caption         =   "�� �� ��  ��  ��   ��   ��"
      End
      Begin VB.Menu ��������� 
         Caption         =   "�� �� ��  ��  ��   ��   ��"
      End
      Begin VB.Menu Filler1_1 
         Caption         =   "-"
      End
      Begin VB.Menu ����ó���� 
         Caption         =   "�� �� ó  ��  ��   ��   ��"
         Enabled         =   0   'False
      End
      Begin VB.Menu Filler1_2 
         Caption         =   "-"
      End
      Begin VB.Menu �������� 
         Caption         =   "��   ��   ��   ��   ��   ��"
      End
      Begin VB.Menu Filler1_3 
         Caption         =   "-"
      End
      Begin VB.Menu ����з� 
         Caption         =   "��   ��   ��   ��   ��   ��"
      End
      Begin VB.Menu �������� 
         Caption         =   "��   ��   ��   ��   ��   ��"
      End
      Begin VB.Menu Filler1_4 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "��                           ��"
      End
   End
   Begin VB.Menu �ŷ�ó����ΰ��� 
      Caption         =   "�ŷ�ó����ΰ���(&2)"
      Begin VB.Menu ���Լ��ݰ�꼭����Է� 
         Caption         =   "���Լ��ݰ�꼭���  ��           ��"
      End
      Begin VB.Menu ���Լ��ݰ�꼭��ȸ�׼��� 
         Caption         =   "���Լ��ݰ�꼭���  ��ȸ �� ����"
      End
      Begin VB.Menu �����ޱݿ��� 
         Caption         =   "��    ��    ó     ��    ��    ó    ��"
      End
      Begin VB.Menu ����ó������ 
         Caption         =   "��  �� ó  ��     ��    ��    ��    Ȳ"
      End
      Begin VB.Menu Filler2_1 
         Caption         =   "-"
      End
      Begin VB.Menu ���⼼�ݰ�꼭����Է� 
         Caption         =   "�� �� ó   �� �� ��    ��  �� ��  ��"
      End
      Begin VB.Menu ���⼼�ݰ�꼭��ȸ�׼��� 
         Caption         =   "����ó�̼��� ���    ��ȸ �� ����"
      End
      Begin VB.Menu �̼��ݿ��� 
         Caption         =   "��    ��    ó     ��    ��    ó    ��"
      End
      Begin VB.Menu ��꼭�Ǻ� 
         Caption         =   "�� �� ��꼭    (��    ��)   ó    ��"
      End
      Begin VB.Menu ��꼭�ϰ� 
         Caption         =   "��  �� ��  ��     ��  ��  ó  ��(NO)"
      End
      Begin VB.Menu ���ݰ�꼭 
         Caption         =   "�� �� ��꼭     ��  ȸ  ��   ��  ��"
      End
      Begin VB.Menu ����ó������ 
         Caption         =   "��  �� ó  ��     ��    ��    ��    Ȳ"
      End
   End
   Begin VB.Menu ���԰��� 
      Caption         =   "���԰���(&3)"
      Begin VB.Menu �����ۼ�2 
         Caption         =   "��   ��   ��   ǥ   ��   ��"
      End
      Begin VB.Menu ����ó���� 
         Caption         =   "��      ��      ó   ��   ��"
      End
      Begin VB.Menu ���Լ��� 
         Caption         =   "�� �� �� ǥ   ��ȸ�׼���"
      End
      Begin VB.Menu Filler3_1 
         Caption         =   "-"
      End
      Begin VB.Menu ���ּ��ۼ� 
         Caption         =   "��      ��      ��   ��   ��"
      End
      Begin VB.Menu ���ּ����� 
         Caption         =   "��   ��   ��   ��ȸ�׼���"
      End
      Begin VB.Menu �����ۼ�1 
         Caption         =   "���ּ� �� �� �� ǥ ó ��"
      End
      Begin VB.Menu Filler3_2 
         Caption         =   "-"
      End
      Begin VB.Menu ��ǰ���� 
         Caption         =   "�� �� �� ǥ   �� ǰ ó ��"
      End
      Begin VB.Menu Filler3_3 
         Caption         =   "-"
      End
      Begin VB.Menu ����ó���ܰ���ȸ 
         Caption         =   "�� �� ó ��   �� �� �� ȸ"
      End
      Begin VB.Menu ǰ�񺰸��Լ�����ȸ 
         Caption         =   "ǰ�� �� �� �� �� �� ȸ"
      End
   End
   Begin VB.Menu ������� 
      Caption         =   "�������(&4)"
      Begin VB.Menu �����ۼ�2 
         Caption         =   "�ŷ� ����  ��         ��"
      End
      Begin VB.Menu ����ó���� 
         Caption         =   "��      ��      ó   ��   ��"
      End
      Begin VB.Menu ������� 
         Caption         =   "�ŷ� ����  ��ȸ�׼���"
      End
      Begin VB.Menu Filler4_1 
         Caption         =   "-"
      End
      Begin VB.Menu �������ۼ� 
         Caption         =   "��      ��      ��   ��   ��"
      End
      Begin VB.Menu ���������� 
         Caption         =   "��   ��   ��   ��ȸ�׼���"
      End
      Begin VB.Menu �����ۼ�1 
         Caption         =   "������ �ŷ����� ó ��"
      End
      Begin VB.Menu Filler4_2 
         Caption         =   "-"
      End
      Begin VB.Menu ���԰��� 
         Caption         =   "�ŷ�����   �� ǰ ó ��"
      End
      Begin VB.Menu Filler4_3 
         Caption         =   "-"
      End
      Begin VB.Menu ����ó���ܰ���ȸ 
         Caption         =   "�� �� ó ��   �� �� �� ȸ"
      End
      Begin VB.Menu ǰ�񺰸��������ȸ 
         Caption         =   "ǰ�� �� �� �� �� �� ȸ"
      End
   End
   Begin VB.Menu ȸ����� 
      Caption         =   "ȸ�����(&5)"
      Begin VB.Menu �����ⳳ��ϰ��� 
         Caption         =   "�����ⳳ �� �� �� ��"
      End
      Begin VB.Menu �����ڵ��ϰ��� 
         Caption         =   "�����ڵ� �� �� �� ��"
      End
   End
   Begin VB.Menu ������ 
      Caption         =   "������(&6)"
      Begin VB.Menu ������� 
         Caption         =   "��  ��   ��  ��"
      End
      Begin VB.Menu Filler6_1 
         Caption         =   "-"
      End
      Begin VB.Menu �̴޻�ǰ 
         Caption         =   "��  ��   ��  ǰ"
      End
      Begin VB.Menu ���Һ� 
         Caption         =   "��     ��     ��"
      End
      Begin VB.Menu Filler6_2 
         Caption         =   "-"
      End
      Begin VB.Menu ����̵� 
         Caption         =   "��  ��   ��  ��"
         Enabled         =   0   'False
      End
      Begin VB.Menu ������� 
         Caption         =   "��  ��   ��  ��"
      End
   End
   Begin VB.Menu �������� 
      Caption         =   "��������(&8)"
      Begin VB.Menu �ŷ�������� 
         Caption         =   "�ŷ����� ���"
      End
      Begin VB.Menu ���޼������ 
         Caption         =   "����/�������"
      End
      Begin VB.Menu Filler8_1 
         Caption         =   "-"
      End
      Begin VB.Menu ��¹����� 
         Caption         =   "�� �� �� �� ��"
      End
      Begin VB.Menu Filler8_2 
         Caption         =   "-"
      End
      Begin VB.Menu �������۾� 
         Caption         =   "�� �� �� �� ��"
      End
      Begin VB.Menu Filler8_3 
         Caption         =   "-"
      End
      Begin VB.Menu �ڷ��� 
         Caption         =   "��  ��   ��  ��"
      End
   End
   Begin VB.Menu ����� 
      Caption         =   "�����(&9)"
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : Main
' ���� Control : StatusBar
' ������ Table   : �����
' ��  ��  ��  �� :
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
       SBar.Panels(1).Text = "�۾����� : " & Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
       SBar.Panels(2).Text = "����� : " & PB_regUserinfoU.UserName
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
    '�α��� �Ǵ� ������ �ƴ� ���� ���ִ� ��� �����Ҽ� ������ �Ѵ�.
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
    strSQL = "SELECT CONVERT(VARCHAR(8),GETDATE(), 112) AS ��������, " _
                  & "RIGHT(CONVERT(VARCHAR(23),GETDATE(), 121),12) AS �����ð� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    StrDate = P_adoRec("��������")
    strTime = Mid(P_adoRec("�����ð�"), 1, 2) + Mid(P_adoRec("�����ð�"), 4, 2) _
            + Mid(P_adoRec("�����ð�"), 7, 2) + Mid(P_adoRec("�����ð�"), 10)
    P_adoRec.Close
    PB_adoCnnSQL.BeginTrans
    strSQL = "UPDATE ����� SET " _
                  & "�α��ο��� = 'N', " _
                  & "�����Ͻ� = '" & StrDate & strTime & "' " _
            & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserCode & "' "
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
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "��������� ���� ����"
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
'| �����ڷ���������(�޴���:�����ڷ���������(1))
'+---------------------------------------------+
Private Sub ���������_Click()                   '�޴���:������������
Dim iRet
    'API �̿�
    'iRet = GetSystemMenu(frmMain.hwnd, 0)
    'DeleteMenu iRet, SC_MAXIMIZE, MF_BYCOMMAND
    'DeleteMenu iRet, SC_MINIMIZE, MF_BYCOMMAND
    'DeleteMenu iRet, SC_CLOSE, MF_BYCOMMAND
    With frm���������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         '.Show vbModeless
         .Show vbModal
    End With
End Sub
Private Sub ���������_Click()                   '�޴���:������������
    With frm���������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

Private Sub ����ó����_Click()                   '�޴���:����ó�������
    With frm����ó����
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ��������_Click()                     '�޴���:�����ڵ���
    With frm��������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ����з�_Click()                     '�޴���:����з����
    With frm����з�
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ��������_Click()                     '�޴���:�����ڵ���
    With frm��������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

Private Sub ����_Click()
    Unload Me
End Sub

'+---------------------------------------------+
'| �ŷ�ó����ΰ���(�޴���:�ŷ�ó����ΰ���(2))
'|---------------------------------------------+
'| ���Լ��ݰ�꼭��� �Է�
'| ���Լ��ݰ�꼭��� ��ȸ�׼���
'| ����ó�� ���ó��
'| ����ó�� ������Ȳ
'| ---------------------
'| ����ó�̼������ �Է�
'| ����ó�̼������ ��ȸ�׼���
'| ����ó�� ����ó��
'| ���ݰ�꼭(�Ǻ�)ó��
'| ��������ϰ�ó��(NO)
'| ���ݰ�꼭��ȸ�׼���
'| ����ó�� ������Ȳ
'+-----------------------------+
Private Sub ���Լ��ݰ�꼭����Է�_Click()       '�޴���:���Լ��ݰ�꼭��� �Է�
    With frm���Լ��ݰ�꼭����Է�
         '.Left = 0: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ���Լ��ݰ�꼭��ȸ�׼���_Click()     '�޴���:���Լ��ݰ�꼭��� ��ȸ�׼���
    With frm���Լ��ݰ�꼭
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub �����ޱݿ���_Click()                 '�޴���:����ó�� ���ó��
    With frm�����ޱݿ���
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ����ó������_Click()                 '�޴���:����ó�� ������Ȳ
    With frm����ó������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

Private Sub ���⼼�ݰ�꼭����Է�_Click()       '�޴���:����ó�̼������ �Է�
    With frm���⼼�ݰ�꼭����Է�
         '.Left = 0: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ���⼼�ݰ�꼭��ȸ�׼���_Click()     '�޴���:����ó�̼������ ��ȸ�׼���
    With frm���⼼�ݰ�꼭
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub �̼��ݿ���_Click()                   '�޴���:����ó�� ����ó��
    With frm�̼��ݿ���
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ��꼭�Ǻ�_Click()                   '�޴���:���ݰ�꼭(�Ǻ�)ó��
    With frm��꼭�Ǻ�
         '.Left = 0: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ��꼭�ϰ�_Click()                   '�޴���:��������ϰ�ó��(NO)
    With frm��꼭�ϰ�
         '.Left = 0: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ���ݰ�꼭_Click()                   '�޴���:���ݰ�꼭��ȸ�׼���
    With frm���ݰ�꼭
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ����ó������_Click()                 '�޴���:����ó�� ������Ȳ
    With frm����ó������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| ���԰���(�޴���:���԰���(3))
'|-----------------------------+
Private Sub �����ۼ�2_Click()                    '�޴���:������ǥ�Է�
    With frm�����ۼ�2
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ����ó����_Click()                   '�޴���:����ó���
    With frm����ó����
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ���Լ���_Click()                     '�޴���:������ǥ ��ȸ�׼���
    With frm���Լ���
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ���ּ��ۼ�_Click()                   '�޴���:���ּ��ۼ�
    With frm���ּ��ۼ�
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ���ּ�����_Click()                   '�޴���:���ּ� ��ȸ�׼���
    With frm���ּ�����
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub �����ۼ�1_Click()                    '�޴���:���ּ� ������ǥ ó��
    With frm�����ۼ�1
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ��ǰ����_Click()                     '�޴���:������ǥ ��ǰó��
    With frm��ǰ����
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ����ó���ܰ���ȸ_Click()            '�޴���:����ó���ܰ���ȸ
    With frm����ó���ܰ���ȸ
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ǰ�񺰸��Լ�����ȸ_Click()           '�޴���:ǰ�񺰸��Լ�����ȸ
    With frmǰ�񺰸��Լ�����ȸ
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| �������(�޴���:�������(4))
'|-----------------------------+
Private Sub �����ۼ�2_Click()                    '�޴���:�ŷ������ۼ�
    With frm�����ۼ�2
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ����ó����_Click()                   '�޴���:����ó���
    With frm����ó����
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub �������_Click()                     '�޴���:�ŷ�������ȸ�׼���
    With frm�������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub �������ۼ�_Click()                   '�޴���:�������ŷ�����ó��
    With frm�������ۼ�
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ����������_Click()                   '�޴���:��������ȸ�׼���
    With frm����������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub �����ۼ�1_Click()                    '�޴���:�������ŷ�����ó��
    With frm�����ۼ�1
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ���԰���_Click()                     '�޴���:�ŷ�������ǰó��
    With frm���԰���
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ����ó���ܰ���ȸ_Click()            '�޴���:����ó���ܰ���ȸ
    With frm����ó���ܰ���ȸ
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ǰ�񺰸��������ȸ_Click()          '�޴���:ǰ�񺰸��������ȸ
    With frmǰ�񺰸��������ȸ
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| ȸ�����(�޴���:ȸ�����(5))
'+-----------------------------+
Private Sub �����ⳳ��ϰ���_Click()             '�޴���:�����ⳳ��ϰ���
    With frm�����ⳳ����
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub �����ڵ��ϰ���_Click()             '�޴���:�����ڵ��ϰ���
    With frm�����ڵ����
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| ������(�޴���:������(6))
'+-----------------------------+
Private Sub �������_Click()                     '�޴���:�������
    With frm�������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub �̴޻�ǰ_Click()                     '�޴���:�̴޻�ǰ
    With frm�̴޻�ǰ
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ���Һ�_Click()                       '�޴���:���Һ�
    With frm���Һ�
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub �������_Click()                     '�޴���:�������
    With frm�������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| ��������(�޴���:��������(8))
'+-----------------------------+
Private Sub �ŷ��������_Click()                 '�޴���:�ŷ��������
    With frm�ŷ��������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ���޼������_Click()                 '�޴���:���޼������
    With frm���޼������
         .Left = 0: .Top = 650
         .Height = 10100: .ScaleHeight = 10100
         .Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub
Private Sub ��¹�����_Click()                   '�޴���:��¹�����
    With frm��¹�����
         .Left = (15405 - .Width) / 2: .Top = 650
         '.Height = 10100: .ScaleHeight = 10100
         '.Width = 15405: .ScaleWidth = 15405
         .Show vbModal
    End With
End Sub

Private Sub �������۾�_Click()                   '�޴���:�������۾�
    With frm�������۾�
         .Show vbModal
    End With
End Sub

Private Sub �ڷ���_Click()                     '�޴���:�ڷ���
    With frm�ڷ���
         .Show vbModal
    End With
End Sub

'+-----------------------------+
'| �����(�޴���:�����(9))
'+-----------------------------+

