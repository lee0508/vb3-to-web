VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm��꼭�Ǻ� 
   BorderStyle     =   1  '���� ����
   Caption         =   "���ݰ�꼭(�Ǻ�)ó��"
   ClientHeight    =   5790
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
   Icon            =   "��꼭�Ǻ�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5792.867
   ScaleMode       =   0  '�����
   ScaleWidth      =   12255
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame1 
      Appearance      =   0  '���
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   20
      Top             =   0
      Width           =   12075
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   8640
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   44
         Top             =   240
         Width           =   3315
      End
      Begin VB.ListBox lstPort 
         Height          =   240
         Left            =   3960
         TabIndex        =   30
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
         Caption         =   "���ݰ�꼭(�Ǻ�)ó��"
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
         TabIndex        =   21
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
      Height          =   5115
      Left            =   60
      TabIndex        =   19
      Top             =   630
      Width           =   12075
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         Index           =   10
         Left            =   1680
         TabIndex        =   16
         Top             =   4000
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   9
         Left            =   6840
         TabIndex        =   15
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   4
         Left            =   4920
         TabIndex        =   8
         Top             =   2100
         Width           =   1815
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   4080
         Picture         =   "��꼭�Ǻ�.frx":030A
         Style           =   1  '�׷���
         TabIndex        =   43
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   2880
         Picture         =   "��꼭�Ǻ�.frx":0C58
         Style           =   1  '�׷���
         TabIndex        =   17
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   1680
         Picture         =   "��꼭�Ǻ�.frx":14DF
         Style           =   1  '�׷���
         TabIndex        =   18
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   6
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   10
         Top             =   2450
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   7
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   11
         Top             =   2800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   8
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   12
         Top             =   3200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   9
         Top             =   2450
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   3
         Left            =   1680
         TabIndex        =   7
         Top             =   2100
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   0
         Top             =   300
         Width           =   735
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "������ �μ�"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   4575
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   2
         Left            =   3720
         TabIndex        =   5
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   1
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpF_Date 
         Height          =   270
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpT_Date 
         Height          =   270
         Left            =   4320
         TabIndex        =   3
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpP_Date 
         Height          =   270
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpS_Date 
         Height          =   270
         Left            =   1680
         TabIndex        =   13
         Top             =   3600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin MSComCtl2.DTPicker dtpM_Date 
         Height          =   270
         Left            =   4320
         TabIndex        =   14
         Top             =   3600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   56623105
         UpDown          =   -1  'True
         CurrentDate     =   37763
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   14
         Left            =   360
         TabIndex        =   49
         Top             =   4050
         Width           =   1095
      End
      Begin VB.Label lblMDate 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Left            =   3000
         TabIndex        =   48
         Top             =   3645
         Width           =   1095
      End
      Begin VB.Label lblBillNo 
         Alignment       =   1  '������ ����
         Caption         =   "������ȣ"
         Height          =   240
         Left            =   5640
         TabIndex        =   47
         Top             =   3645
         Width           =   1095
      End
      Begin VB.Label lblSDate 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Left            =   360
         TabIndex        =   46
         Top             =   3650
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�ΰ�����"
         Height          =   240
         Index           =   7
         Left            =   3600
         TabIndex        =   45
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   11760
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label1 
         Caption         =   "(0.����, 1.��ǥ, 2.����, 3.�ܻ�̼���)"
         Height          =   240
         Index           =   13
         Left            =   2640
         TabIndex        =   42
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "��"
         Height          =   240
         Index           =   12
         Left            =   6480
         TabIndex        =   41
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         Caption         =   "��"
         Height          =   240
         Index           =   8
         Left            =   5280
         TabIndex        =   40
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblJanMny 
         Alignment       =   1  '������ ����
         Caption         =   "0"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1680
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "(0.������, 1.û����)"
         Height          =   240
         Index           =   21
         Left            =   2640
         TabIndex        =   39
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���ݾ���"
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   38
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   9
         Left            =   360
         TabIndex        =   37
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "ǰ��ױ԰�"
         Height          =   240
         Index           =   20
         Left            =   360
         TabIndex        =   36
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��꼭�ݾ�"
         Height          =   240
         Index           =   19
         Left            =   360
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "������ܾ�"
         Height          =   240
         Index           =   18
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ó�ڵ�"
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��������"
         Height          =   240
         Index           =   17
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "(0.�ŷ�, 1.����)"
         Height          =   360
         Index           =   11
         Left            =   2280
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "���౸��"
         Height          =   240
         Index           =   10
         Left            =   360
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   6
         Left            =   5760
         TabIndex        =   27
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����"
         Height          =   240
         Index           =   5
         Left            =   3600
         TabIndex        =   26
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "]"
         Height          =   240
         Index           =   4
         Left            =   6675
         TabIndex        =   25
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "["
         Height          =   240
         Index           =   3
         Left            =   1680
         TabIndex        =   24
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����Ⱓ"
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "[Home]"
         Height          =   240
         Index           =   15
         Left            =   1680
         TabIndex        =   22
         Top             =   1485
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm��꼭�Ǻ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ��꼭�Ǻ�
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : �����, ����ó, �������⳻��, ���ݰ�꼭
' ��  ��  ��  �� : �������(8.�̼��ݹ߻��ݾ׿��� ����)
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

Dim p                  As Printer
Dim strDefaultPrinter  As String
Dim aryPrinter()       As String
Dim strBuffer          As String

    frmMain.SBar.Panels(4).Text = "������ ���迡�� ����, ���⼼�ݰ�꼭��ο��� ����, 3.�ܻ�̼����� �ƴѰ�� �̼��� ���ݳ����� ����˴ϴ�. "
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       
       'Printer Setting and Seaching(API)
       strBuffer = Space(1024)
       inti = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
       aryPrinter = Split(strBuffer, ",")
       strDefaultPrinter = Trim(Trim(aryPrinter(0)))
       For Each p In Printers
           cboPrinter.AddItem Trim(p.DeviceName)
           lstPort.AddItem p.Port
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

'+-------------------+
'/// Text1(index) ///
'+-------------------+
Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
         .SelStart = 0
         .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim strSQL     As String
Dim strExeFile As String
Dim varRetVal  As Variant
    If (Index = 1 And (KeyCode = vbKeyHome Or KeyCode = vbKeyReturn)) Then  '����ó�˻�
       PB_strSupplierCode = UPPER(Trim(Text1(Index).Text))
       PB_strSupplierName = "" 'Trim(Text1(Index + 1).Text)
       frm����ó�˻�.Show vbModal
       If (Len(PB_strSupplierCode) + Len(PB_strSupplierName)) = 0 Then '�˻����� ���(ESC)
       Else
          Text1(Index).Text = PB_strSupplierCode
          Text1(Index + 1).Text = PB_strSupplierName
       End If
       If PB_strSupplierCode <> "" Then
          SendKeys "{TAB}"
       End If
       PB_strSupplierCode = "": PB_strSupplierName = ""
    Else
       If KeyCode = vbKeyReturn Then
          Select Case Index
          End Select
          SendKeys "{tab}"
       End If
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����ó �б� ����"
    Unload Me
    Exit Sub
End Sub

'+-------------+
'/// ����ó ///
'+-------------+
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL   As String
Dim lngR     As Long
Dim intIndex As Integer
    intIndex = Index
    With Text1(Index)
         .Text = Trim(.Text)
         Select Case Index
                Case 0 '���౸��
                     If Trim(.Text) = "0" Then '�ŷ�
                        If Vals(Text1(3).Text) = 0 Then
                           Text1(3).Enabled = False: Text1(4).Enabled = False
                        Else
                           Text1(3).Enabled = True: Text1(4).Enabled = True
                        End If
                     ElseIf _
                        Trim(.Text) = "1" Then '����
                        Text1(3).Enabled = True: Text1(4).Enabled = True
                     Else
                        Text1(3).Enabled = False: Text1(4).Enabled = False
                     End If
                Case 1 '����ó�ڵ�
                     .Text = UPPER(Trim(.Text))
                     If Len(Trim(.Text)) < 1 Then
                        .Text = ""
                        Text1(2).Text = ""
                        Exit Sub
                     End If
                     'P_adoRec.CursorLocation = adUseClient
                     'strSQL = "SELECT * FROM ����ó " _
                             & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                               & "AND ����ó�ڵ� = '" & Trim(.Text) & "' "
                     'On Error GoTo ERROR_TABLE_SELECT
                     'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                     'If P_adoRec.RecordCount = 0 Then
                     '   P_adoRec.Close
                     '   .Text = ""
                     '   .SetFocus
                     '   Exit Sub
                     'End If
                     'Text1(2).Text = P_adoRec("����ó��")
                     'P_adoRec.Close
                     SubCompute_JanMny PB_regUserinfoU.UserBranchCode, PB_regUserinfoU.UserClientDate, Text1(1).Text '�̼����ܾ�
                     If Trim(Text1(0).Text) = "0" Then '�ŷ�(���ݰ�꼭�ݾ�)
                        SubCompute_TaxBillMny PB_regUserinfoU.UserBranchCode, DTOS(dtpF_Date.Value), _
                                              DTOS(dtpT_Date.Value), Trim(Text1(1).Text)
                        If Vals(Text1(3).Text) = 0 Then
                           Text1(3).Enabled = False: Text1(4).Enabled = False
                        Else
                           Text1(3).Enabled = True: Text1(4).Enabled = True
                        End If
                     End If
                Case 3 '���ݰ�꼭�ݾ�(���ް���)
                     .Text = Format(Vals(.Text), "#,0.00")
                     Text1(4).Text = Format(Fix(Vals(.Text) * (PB_curVatRate)), "#,0.00")
                Case 4 '���ݰ�꼭�ݾ�(�ΰ���)
                     .Text = Format(Vals(.Text), "#,0.00")
                Case 7 '����(7), ��������(8)
                     dtpS_Date.Enabled = False: dtpM_Date.Enabled = False: Text1(9).Enabled = False
                     If Text1(7).Text = "0" Or Text1(7).Text = "1" Then   '���� �Ǵ� ��ǥ
                        Text1(8).Text = "0"
                        dtpS_Date.Enabled = True: dtpM_Date.Enabled = False: Text1(9).Enabled = False
                     ElseIf _
                        Text1(7).Text = "2" Then                           '����
                        Text1(8).Text = "0"
                        dtpS_Date.Enabled = True: dtpM_Date.Enabled = True: Text1(9).Enabled = True
                     Else
                        Text1(7).Text = "3"
                        Text1(8).Text = "1"
                     End If
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "TABLE �б� ����"
    Unload Me
    Exit Sub
End Sub

'+---------------+
'/// �������� ///
'+---------------+
Private Sub dtpP_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpP_Date_LostFocus()
    dtpS_Date.Value = dtpP_Date.Value
End Sub
'+---------------+
'/// ����Ⱓ ///
'+---------------+
Private Sub dtpF_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub dtpT_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
'+---------------+
'/// �������� ///
'+---------------+
Private Sub dtpS_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
'+---------------+
'/// �������� ///
'+---------------+
Private Sub dtpM_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
    SubClearText
    Text1(Text1.LBound).SetFocus
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
Dim p              As Printer
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
Dim strJukyo       As String  '����(�������⳻��. �̼��ݳ���)
Dim intSactionWay  As Integer '�������(�̼��ݳ���)
Dim curInputMoney  As Long
Dim curOutputMoney As Long
Dim strServerTime  As String
Dim strTime        As String
Dim intKijang      As Integer '1.����
    
    For Each p In Printers
        If Trim(p.DeviceName) = Trim(cboPrinter.Text) And lstPort.List(cboPrinter.ListIndex) = p.Port Then
           Set Printer = p
           Exit For
        End If
    Next
    
    If dtpF_Date.Value > dtpT_Date.Value Then
       dtpF_Date.SetFocus
       Exit Sub
    End If
    '�Է³��� �˻�
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       If Text1(lngC).Enabled = False Then
          Text1(lngC).Enabled = True
       End If
       Text1(lngC).SetFocus
       Exit Sub
    End If
    intRetVal = MsgBox("���ݰ�꼭�� �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
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
    'å��ȣ, �Ϸù�ȣ ���ϱ�
    PB_adoCnnSQL.BeginTrans
    strMakeYear = Mid(DTOS(dtpP_Date.Value), 1, 4)
    strSQL = "spLogCounter '���ݰ�꼭', '" & PB_regUserinfoU.UserBranchCode + strMakeYear & "', " _
                         & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
    On Error GoTo ERROR_STORED_PROCEDURE
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    lngLogCnt1 = P_adoRec(0)
    lngLogCnt2 = P_adoRec(1)
    P_adoRec.Close
    '���ݰ�꼭 �߰�
    strSQL = "INSERT INTO ���ݰ�꼭(������ڵ�, �ۼ��⵵, å��ȣ, �Ϸù�ȣ, " _
                                  & "����ó�ڵ�, �ۼ�����, ǰ��ױ԰�, ����, " _
                                  & "���ް���, ����, �ݾױ���, ��û����, " _
                                  & "���࿩��, �ۼ�����, �̼�����, ����, ��뱸��, " _
                                  & "��������, ������ڵ�) VALUES(" _
    & "'" & PB_regUserinfoU.UserBranchCode & "', '" & strMakeYear & "', " & lngLogCnt1 & ", " & lngLogCnt2 & "," _
    & "'" & Trim(Text1(1).Text) & "', '" & DTOS(dtpP_Date.Value) & "','" & Trim(Text1(5).Text) & "'," & Vals(Text1(6).Text) & ", " _
    & "" & Vals(Text1(3).Text) & ", " & Vals(Text1(4).Text) & ", " & Vals(Text1(7).Text) & "," & Vals(Text1(8).Text) & ", " _
    & "0, " & Vals(Text1(0).Text) & ", 1, '" & Text1(10).Text & "', 0, " _
    & "'" & PB_regUserinfoU.UserServerDate & "','" & PB_regUserinfoU.UserCode & "') "
    On Error GoTo ERROR_TABLE_INSERT
    PB_adoCnnSQL.Execute strSQL
    If Trim(Text1(0).Text = "0") Then '�ŷ��� ���ݰ�꼭�̸� (�������⳻�� ����)
       strSQL = "UPDATE �������⳻�� SET " _
                     & "��꼭���࿩�� = 1, " _
                     & "�ۼ��⵵ = '" & strMakeYear & "', " _
                     & "å��ȣ = " & lngLogCnt1 & ", " _
                     & "�Ϸù�ȣ = " & lngLogCnt2 & " " _
               & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' AND ����ó�ڵ� = '" & Trim(Text1(1).Text) & "' " _
                 & "AND ��������� BETWEEN '" & DTOS(dtpF_Date.Value) & "' AND '" & DTOS(dtpT_Date.Value) & "' " _
                 & "AND ������� = 2 AND ��뱸�� = 0 AND ���ݱ��� = 0 AND ��꼭���࿩�� = 0 "
       On Error GoTo ERROR_TABLE_UPDATE
       PB_adoCnnSQL.Execute strSQL
    Else                                 '���Ǻ� ���ݰ�꼭
       intRetVal = MsgBox("�̼��ݿ� �����ϰڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton1, "�̼��� ����")
       If intRetVal = vbYes Then
          intKijang = 1
       End If
       If Vals(Trim(Text1(6).Text)) = 0 Then '����
          strJukyo = Trim(Text1(5).Text)
       Else
          strJukyo = Trim(Text1(5).Text) & " �� " + CStr(Vals(Trim(Text1(6).Text))) + "��"
       End If
       '�ŷ���ȣ ���ϱ�
       'P_adoRec.CursorLocation = adUseClient
       'strSQL = "spLogCounter '�������⳻��', '" & PB_regUserinfoU.UserBranchCode + DTOS(dtpP_Date.Value) + "2" & "', " _
       '                        & "0, 0, '" & PB_regUserinfoU.UserServerDate & "', '" & PB_regUserinfoU.UserCode & "' "
       'On Error GoTo ERROR_STORED_PROCEDURE
       'P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
       'lngLogCnt = P_adoRec(0)
       'P_adoRec.Close
       If intKijang = 1 Then
          'strSQL = "INSERT INTO �������⳻��(������ڵ�, �з��ڵ�, �����ڵ�, �������," _
                                          & "���������, �����ð�, �԰����, �԰�ܰ�," _
                                          & "�԰�ΰ�, ������, ���ܰ�, ���ΰ�," _
                                          & "����ó�ڵ�, ����ó�ڵ�, �������������, ���۱���," _
                                          & "�߰�����, �߰߹�ȣ, �ŷ�����, �ŷ���ȣ," _
                                          & "��꼭���࿩��, ���ݱ���, ��������, ����, �ۼ��⵵, å��ȣ," _
                                          & "�Ϸù�ȣ, ��뱸��, ��������, ������ڵ�, ����̵�������ڵ�) VALUES(" _
                 & "'" & PB_regUserinfoU.UserBranchCode & "', '', '', 8," _
                 & "'" & DTOS(dtpP_Date.Value) & "','" & strServerTime & "', 0, 0," _
                 & "0, " & IIf(Vals(Text1(3).Text) < 0, -1, 1) & ", " & Abs(Vals(Text1(3).Text)) & ", " & Abs(Vals(Text1(4).Text)) & "," _
                 & "'', '" & Trim(Text1(1).Text) & "', '" & DTOS(dtpP_Date.Value) & "', 0," _
                 & "'', 0, '" & DTOS(dtpP_Date.Value) & "', " & lngLogCnt & "," _
                 & "1, " & IIf(Vals(Trim(Text1(7).Text)) = 0 Or Vals(Trim(Text1(7).Text)) = 1, 1, 0) & ", 0, '" & strJukyo & "', " _
                 & "'" & strMakeYear & "', " & lngLogCnt1 & ", " _
                 & "" & lngLogCnt2 & ", 0,'" & PB_regUserinfoU.UserServerDate & "','" & PB_regUserinfoU.UserCode & "', '')"
          'On Error GoTo ERROR_TABLE_INSERT
          'PB_adoCnnSQL.Execute strSQL
       End If
       strSQL = "UPDATE ���ݰ�꼭 SET " _
                     & "�̼����� = " & intKijang & " " _
               & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND �ۼ��⵵ = '" & strMakeYear & "' " _
                 & "AND å��ȣ = " & lngLogCnt1 & " " _
                 & "AND �Ϸù�ȣ = " & lngLogCnt2 & " "
       On Error GoTo ERROR_TABLE_UPDATE
       PB_adoCnnSQL.Execute strSQL
    End If
    'if (�ŷ� AND ������) OR (���� AND ������ AND ������) then  '�̼����Աݳ��� �߰�
    If (Trim(Text1(0).Text = "0") And Trim(Text1(8).Text) = "0") Or _
       (Trim(Text1(0).Text = "1") And Trim(Text1(8).Text) = "0" And intKijang = 1) Then
       If Vals(Trim(Text1(7).Text)) = 0 Or Vals(Trim(Text1(7).Text)) = 1 Then '���� �Ǵ� ��ǥ
          intSactionWay = 0
       ElseIf _
          Vals(Trim(Text1(7).Text)) = 2 Then '����
          intSactionWay = 1
       End If
       strJukyo = PB_regUserinfoU.UserBranchCode + "-" + strMakeYear + "-" + CStr(lngLogCnt1) + "-" + CStr(lngLogCnt2)
       strSQL = "INSERT INTO �̼��ݳ���(������ڵ�, ����ó�ڵ�, " _
                                     & "�̼����Ա�����, �̼����Աݽð�," _
                                     & "�̼����Աݱݾ�, �������, " _
                                     & "��������, ������ȣ, " _
                                     & "����, ��������, " _
                                     & "������ڵ�) VALUES(" _
                        & "'" & PB_regUserinfoU.UserBranchCode & "', '" & Trim(Text1(1).Text) & "', " _
                        & "'" & DTOS(dtpS_Date.Value) & "', '" & strServerTime & "', " _
                        & "" & (Vals(Text1(3).Text) + Vals(Text1(4).Text)) & ", " & intSactionWay & ", " _
                        & "'" & IIf(intSactionWay = 0, "", DTOS(dtpM_Date.Value)) & "', " _
                        & "'" & IIf(intSactionWay = 0, "", Text1(9).Text) & "', " _
                        & "'" & strJukyo & "', '" & PB_regUserinfoU.UserServerDate & "', " _
                        & "'" & PB_regUserinfoU.UserCode & "' )"
       On Error GoTo ERROR_TABLE_INSERT
       PB_adoCnnSQL.Execute strSQL
    End If
    '��꼭 ���࿩��
    If chkPrint.Value = 1 Then
       strSQL = "UPDATE ���ݰ�꼭 SET " _
                     & "���࿩�� = 1 " _
               & "WHERE ������ڵ� = '" & PB_regUserinfoU.UserBranchCode & "' " _
                 & "AND �ۼ��⵵ = '" & strMakeYear & "' " _
                 & "AND å��ȣ = " & lngLogCnt1 & " " _
                 & "AND �Ϸù�ȣ = " & lngLogCnt2 & " "
       On Error GoTo ERROR_TABLE_UPDATE
       PB_adoCnnSQL.Execute strSQL
    End If
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
    
    PB_adoCnnSQL.CommitTrans
    If (chkPrint.Value = 1) And (cboPrinter.ListIndex >= 0) Then                         '���ݰ�꼭 ���
       SubPubPrint_TaxBill p, PB_intPrtTypeGbn, PB_regUserinfoU.UserBranchCode, Mid(DTOS(dtpP_Date.Value), 1, 4), lngLogCnt1, lngLogCnt2, _
                           0, "", "", ""
    End If
    Screen.MousePointer = vbDefault
    cmdSave.Enabled = True
    cmdClear_Click
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
    Set frm��꼭�Ǻ� = Nothing
    frmMain.SBar.Panels(4).Text = ""
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
Private Sub SubOther_FILL()
Dim strSQL        As String
Dim intIndex      As Integer
    For intIndex = 0 To Text1.Count - 1: Text1(intIndex).Text = "": Next intIndex
    Text1(0).Text = "": Text1(1).Text = ""
    dtpP_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpF_Date.Value = Format(Mid(PB_regUserinfoU.UserClientDate, 1, 6) & "01", "0000-00-00")
    dtpT_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpS_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
    dtpM_Date.Value = Format(PB_regUserinfoU.UserClientDate, "0000-00-00")
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
Dim lngC As Long
    For lngC = Text1.LBound To Text1.UBound
        Text1(lngC).Text = ""
    Next lngC
    lblJanMny.Caption = "0"
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Text1(lngC).Text = Trim(Text1(lngC).Text)
        Select Case lngC
               Case 0  '���౸��(0.�ŷ�, 1.����)
                    If Not (Text1(lngC).Text = "0" Or Text1(lngC).Text = "1") Then
                       Exit Function
                    End If
               Case 1  '����ó�ڵ�
                    If Len(Trim(Text1(lngC).Text)) < 1 Then
                       Exit Function
                    End If
               Case 3  '��꼭�ݾ�
                    Text1(lngC).Text = Format(Vals(Text1(lngC).Text), "#,0.00")
                    If (Vals(Trim(Text1(lngC).Text)) = 0) And (Trim(Text1(0).Text) = "0") Then '0.�ŷ�
                       lngC = 1
                       Exit Function
                    End If
                    If (Vals(Trim(Text1(lngC).Text)) = 0) And (Trim(Text1(0).Text) = "1") Then '1.����
                       Exit Function
                    End If
               Case 5  'ǰ��ױ԰�
                    If (Len(Trim(Text1(lngC).Text)) < 1) Or (Len(Trim(Text1(lngC).Text)) > 40) Then
                       Exit Function
                    End If
               Case 6  '����
                    If Vals(Trim(Text1(lngC).Text)) < 0 Then
                       Exit Function
                    End If
               Case 7  '����(0.����, 1.��ǥ, 2.����, 3.�ܻ�̼���)
                    If Not (Text1(lngC).Text = "0" Or Text1(lngC).Text = "1" Or Text1(lngC).Text = "2" Or Text1(lngC).Text = "3") Then
                       Exit Function
                    End If
                    If Text1(lngC).Text = "3" And Trim(Text1(8).Text) = "0" Then '�ܻ�̼��� and ������
                       Exit Function
                    End If
               Case 8  '��û����
                    If Not (Text1(lngC).Text = "0" Or Text1(lngC).Text = "1") Then
                       Exit Function
                    End If
                    If Not (Text1(7).Text = "3") And Trim(Text1(lngC).Text) = "1" Then '�ܻ�̼��ݾƴϰ� and û����
                       Exit Function
                    End If
               Case 9  '������ȣ
                    If Len(Trim(Text1(lngC).Text)) > 20 Then
                       Exit Function
                    End If
               Case 10 '����
                    If Len(Trim(Text1(lngC).Text)) > 50 Then
                       Exit Function
                    End If
               Case Else
        End Select
    Next lngC
    blnOK = True
End Function

'+-----------------+
'/// �̼����ܾ� ///
'+-----------------+
Private Sub SubCompute_JanMny(strBranchCode As String, strT_Date As String, strSupplierCode As String)
Dim strSQL As String
    lblJanMny.Caption = "0"
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT T1.����ó�ڵ� AS ����ó�ڵ�, " _
                  & "SUM(T1.�̼��ݴ���ݾ�) AS �̼��ݱݾ�, SUM(T1.�̼����Աݴ���ݾ�) AS �̼����Աݱݾ� " _
             & "FROM �̼��ݿ��帶�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' " _
              & "AND T1.����ó�ڵ� = '" & strSupplierCode & "' " _
              & "AND T1.������� >= (SUBSTRING('" & strT_Date & "', 1, 4) + '00') " _
              & "AND T1.������� < SUBSTRING('" & strT_Date & "', 1, 6) " _
            & "GROUP BY T1.����ó�ڵ� " _
            & "UNION ALL "
    If PB_regUserinfoU.UserMSGbn = "1" Then  '�̼��ݹ߻����� 1.��ǥ�̸�
       strSQL = strSQL _
           & "SELECT T1.����ó�ڵ� AS ����ó�ڵ�, " _
                  & "(SUM(T1.������ * T1.���ܰ�) * " & (PB_curVatRate + 1) & ") AS �̼��ݱݾ�, 0 AS �̼����Աݱݾ� " _
             & "FROM �������⳻�� T1 " _
             & "LEFT JOIN ����� T2 ON T2.������ڵ� = T1.������ڵ� " _
             & "LEFT JOIN ����ó T3 ON T3.������ڵ� = T1.������ڵ� AND T3.����ó�ڵ� = T1.����ó�ڵ� " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' AND T1.��뱸�� = 0 " _
              & "AND T1.����ó�ڵ� = '" & strSupplierCode & "' " _
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
              & "AND T1.����ó�ڵ� = '" & strSupplierCode & "' " _
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
              & "AND T1.����ó�ڵ� = '" & strSupplierCode & "' " _
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
Private Sub SubCompute_TaxBillMny(strBranchCode As String, strF_Date As String, strT_Date As String, strSupplierCode As String)
Dim strSQL As String
    Text1(3).Text = "0.00": Text1(4).Text = "0.00": Text1(5).Text = "": Text1(6).Text = ""
    Screen.MousePointer = vbHourglass
    P_adoRec.CursorLocation = adUseClient
    strSQL = "SELECT ISNULL(SUM(ISNULL(������, 0) * ISNULL(���ܰ�, 0)), 0) AS ���ް��� " _
             & "FROM �������⳻�� T1 " _
            & "WHERE T1.������ڵ� = '" & strBranchCode & "' AND T1.����ó�ڵ� = '" & strSupplierCode & "' " _
              & "AND T1.��������� BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
              & "AND T1.������� = 2 AND T1.��뱸�� = 0 AND T1.���ݱ��� = 0 AND T1.��꼭���࿩�� = 0 "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          Text1(3).Text = Format(Vals(Text1(3).Text) + P_adoRec("���ް���"), "#,0.00")
          P_adoRec.MoveNext
       Loop
       P_adoRec.Close
    End If
    If Vals(Text1(3).Text) > 1 Then
       Text1(4).Text = Format(Fix(Vals(Text1(3).Text) * (PB_curVatRate)), "#,0.00")
    End If
    'ǰ��ױ԰� �� ����(��)
    'strSQL = "SELECT ISNULL(SUM(ISNULL(���ܰ�, 0) * ISNULL(������, 0)), 0) AS ���ݰ�꼭�ݾ�, "
    strSQL = "SELECT (SELECT TOP 1 (ISNULL(S2.�����,'') + SPACE(1) + ISNULL(S2.�԰�, '')) FROM �������⳻�� S1 " _
                     & "LEFT JOIN ���� S2 ON S1.�з��ڵ� = S2.�з��ڵ� AND S1.�����ڵ� = S2.�����ڵ� " _
                    & "WHERE S1.������ڵ� = T1.������ڵ� " _
                      & "AND S1.��������� BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
                      & "AND S1.������� = 2 AND S1.��뱸�� = 0 AND S1.���ݱ��� = 0 AND S1.��꼭���࿩�� = 0 " _
                      & "AND S1.����ó�ڵ� = T1.����ó�ڵ� " _
                    & "ORDER BY S1.���������, S1.�����ð�) AS ǰ��ױ԰�, " _
                  & "(SELECT (COUNT(S1.�����ڵ�) - 1) FROM �������⳻�� S1 " _
                    & "WHERE S1.������ڵ� = T1.������ڵ� " _
                      & "AND S1.��������� BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
                      & "AND S1.������� = 2 AND S1.��뱸�� = 0 AND S1.���ݱ��� = 0 AND S1.��꼭���࿩�� = 0 " _
                      & "AND S1.����ó�ڵ� = T1.����ó�ڵ�) AS ���� " _
              & "FROM �������⳻�� T1 " _
              & "LEFT JOIN ���� T2 ON T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ� " _
             & "WHERE T1.������ڵ� = '" & strBranchCode & "' AND T1.����ó�ڵ� = '" & strSupplierCode & "' " _
               & "AND T1.��������� BETWEEN '" & strF_Date & "' AND '" & strT_Date & "' " _
               & "AND T1.������� = 2 AND T1.��뱸�� = 0 AND T1.��꼭���࿩�� = 0 AND T1.���ݱ��� = 0 " _
             & "GROUP BY T1.������ڵ�, T1.����ó�ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
    Else
       Do Until P_adoRec.EOF
          Text1(5).Text = P_adoRec("ǰ��ױ԰�")
          Text1(6).Text = Format(P_adoRec("����"), "#,0")
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

