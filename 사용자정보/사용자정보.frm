VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frm��������� 
   BorderStyle     =   0  '����
   Caption         =   "���������"
   ClientHeight    =   10095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15405
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10100
   ScaleMode       =   0  '�����
   ScaleWidth      =   15405
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  '���
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdClear 
         Height          =   390
         Left            =   9120
         Picture         =   "���������.frx":0000
         Style           =   1  '�׷���
         TabIndex        =   24
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   390
         Left            =   13635
         Picture         =   "���������.frx":09A5
         Style           =   1  '�׷���
         TabIndex        =   23
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   390
         Left            =   12510
         Picture         =   "���������.frx":12F3
         Style           =   1  '�׷���
         TabIndex        =   22
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Height          =   390
         Left            =   11385
         Picture         =   "���������.frx":1C77
         Style           =   1  '�׷���
         TabIndex        =   21
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Height          =   390
         Left            =   10260
         Picture         =   "���������.frx":24FE
         Style           =   1  '�׷���
         TabIndex        =   20
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00008000&
         BorderStyle     =   1  '���� ����
         Caption         =   "�� �� �� �� �� �� ��"
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
         Left            =   165
         TabIndex        =   19
         Top             =   180
         Width           =   4650
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vsfg1 
      Height          =   7886
      Left            =   60
      TabIndex        =   17
      Top             =   2100
      Width           =   15195
      _cx             =   26802
      _cy             =   13910
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
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
      Height          =   1395
      Left            =   60
      TabIndex        =   8
      Top             =   630
      Width           =   15195
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   3  '��� ����
         Index           =   4
         Left            =   5550
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   7
         Left            =   8790
         MaxLength       =   1
         TabIndex        =   7
         Top             =   585
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   6
         Left            =   8790
         MaxLength       =   2
         TabIndex        =   6
         Top             =   225
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   3  '��� ����
         Index           =   5
         Left            =   5550
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   945
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   8  '����
         Index           =   3
         Left            =   5550
         MaxLength       =   1
         TabIndex        =   3
         Top             =   225
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   2
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   2
         Top             =   945
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '���
         Height          =   285
         IMEMode         =   10  '�ѱ� 
         Index           =   1
         Left            =   1755
         MaxLength       =   20
         TabIndex        =   1
         Top             =   593
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   8  '����
         Index           =   0
         Left            =   1755
         MaxLength       =   4
         TabIndex        =   0
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "(Y.�α���)"
         Height          =   240
         Index           =   8
         Left            =   6120
         TabIndex        =   26
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "(0.����)"
         Height          =   240
         Index           =   18
         Left            =   9480
         TabIndex        =   25
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "��뱸��"
         Height          =   240
         Index           =   7
         Left            =   7350
         TabIndex        =   16
         ToolTipText     =   "0.����, ��Ÿ.���Ұ�"
         Top             =   645
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "������ڵ�"
         Height          =   240
         Index           =   6
         Left            =   7350
         TabIndex        =   15
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�����й�ȣ"
         Height          =   240
         Index           =   5
         Left            =   4110
         TabIndex        =   14
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�α��κ�й�ȣ"
         Height          =   240
         Index           =   4
         Left            =   4110
         TabIndex        =   13
         Top             =   645
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "�α��ο���"
         Height          =   240
         Index           =   3
         Left            =   4110
         TabIndex        =   12
         ToolTipText     =   "���� �α����(Y.�α���, ��Ÿ.�̷α���)"
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ڱ���"
         Height          =   240
         Index           =   2
         Left            =   555
         TabIndex        =   11
         Top             =   983
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "����ڸ�"
         Height          =   240
         Index           =   1
         Left            =   555
         TabIndex        =   10
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Caption         =   "������ڵ�"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   555
         TabIndex        =   9
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm���������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------------------------------------------------------+
' ���α׷� �� �� : ���������
' ���� Control : VideoSoft VSFlexGrid 7.0(OLEDB) = vsflex7.ocx
' ������ Table   : �����
' ��  ��  ��  �� :
'+-------------------------------------------------------------------------------------------------------+
Option Explicit
Private P_blnActived       As Boolean
Private P_adoRec           As New ADODB.Recordset
Private P_intButton        As Integer
Private Const PC_intRowCnt As Integer = 25  '�׸��� �� ������ �� ���(FixedRows ����)

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
    If P_blnActived = False Then
       Screen.MousePointer = vbHourglass
       On Error GoTo ERROR_FORM_ACTIVATE_CONNECTION_SERVER
       Subvsfg1_INIT
       Subvsfg1_FILL
       Select Case Val(PB_regUserinfoU.UserAuthority)
              Case 99    '�߰�, ��ȸ, ����, ���� ����
                   cmdClear.Enabled = True: cmdFind.Enabled = True: cmdSave.Enabled = True
                   cmdDelete.Enabled = True: cmdExit.Enabled = True
              Case Else  '���� ����
                   cmdClear.Enabled = False: cmdFind.Enabled = False: cmdSave.Enabled = True
                   cmdDelete.Enabled = False: cmdExit.Enabled = True
                   Text1(2).Enabled = False: Text1(3).Enabled = False
                   Text1(6).Enabled = False: Text1(7).Enabled = False
       End Select
       P_blnActived = True
       Screen.MousePointer = vbDefault
    End If
    Exit Sub
ERROR_FORM_ACTIVATE_CONNECTION_SERVER:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� ����(�������� ���� ����)"
    Unload Me
    Exit Sub
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
Dim strExeFile As String
Dim varRetVal  As Variant
    If KeyCode = vbKeyReturn Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
Dim strSQL As String
    With Text1(Index)
         Select Case Index
                Case 0  '������ڵ�
                     .Text = Format(Val(Trim(.Text)), "0000")
                     If Not (.Text >= "0001" And .Text <= "9999") Then
                        .Text = ""
                     End If
                     If .Enabled = True Then
                        P_adoRec.CursorLocation = adUseClient
                        strSQL = "SELECT * FROM ����� " _
                                & "WHERE ������ڵ� = '" & Trim(.Text) & "' "
                        On Error GoTo ERROR_TABLE_SELECT
                        P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
                        If P_adoRec.RecordCount <> 0 Then
                           P_adoRec.Close
                           .Text = ""
                           .SetFocus
                           Exit Sub
                        End If
                        P_adoRec.Close
                     End If
                Case 6  '������ڵ�
                    .Text = Format(Val(Trim(.Text)), "00")
                     If Not (.Text >= "01" And .Text <= "99") Then
                        .Text = ""
                     End If
                Case Else
         End Select
    End With
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "��������� �б� ����"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// Grid ///
'+-----------+
Private Sub vsfg1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With vsfg1
         .ToolTipText = ""
         If .MouseRow < .FixedRows Or .MouseCol < 0 Then
            Exit Sub
         End If
         .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
    End With
End Sub
Private Sub vsfg1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    P_intButton = Button
End Sub

Private Sub vsfg1_Click()
Dim strData As String
    With vsfg1
         If (.MouseRow >= 0 And .MouseRow < .FixedRows) Then
            .Col = .MouseCol
            '.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack '.ForeColorSel
            '.Cell(flexcpForeColor, 0, .MouseCol, 0, .MouseCol) = vbRed
            'strData = Trim(.Cell(flexcpData, .Row, 0))
            'If P_intButton = 1 Then
            '   .Sort = flexSortGenericAscending
            'Else
            '   .Sort = flexSortGenericDescending
            'End If
            'If .FindRow(strData, , 0) > 0 Then
            '   .Row = .FindRow(strData, , 0)
            'End If
            'If PC_intRowCnt < .Rows Then
            '   .TopRow = .Row
            'End If
         End If
    End With
End Sub
Private Sub vsfg1_EnterCell()
Dim lngR As Long
Dim lngC As Long
    With vsfg1
         If .Row < .FixedRows Then
            Text1(Text1.LBound).Enabled = True
         Else
            Text1(Text1.LBound).Enabled = False
            For lngC = 0 To 8
                Select Case lngC
                       Case Is <= 7
                            Text1(lngC).Text = .TextMatrix(.Row, lngC)
                End Select
            Next lngC
         End If
    End With
End Sub
Private Sub vsfg1_KeyDown(KeyCode As Integer, Shift As Integer)
    '
End Sub

'+-----------+
'/// �߰� ///
'+-----------+
Private Sub cmdClear_Click()
    SubClearText
    vsfg1.Row = 0
    Text1(Text1.LBound).SetFocus
End Sub
'+-----------+
'/// ��ȸ ///
'+-----------+
Private Sub cmdFind_Click()
    Screen.MousePointer = vbHourglass
    cmdFind.Enabled = False
    SubClearText
    Subvsfg1_FILL
    cmdFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub
'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdSave_Click()
Dim strSQL    As String
Dim lngR      As Long
Dim lngC      As Long
Dim blnOK     As Boolean
Dim intRetVal As Integer
    '�Է³��� �˻�
    FncCheckTextBox lngC, blnOK
    If blnOK = False Then
       Text1(lngC).SetFocus
       Exit Sub
    End If
    If Text1(Text1.LBound).Enabled = True Then
       intRetVal = MsgBox("�Էµ� �ڷḦ �߰��Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� �߰�")
    Else
       intRetVal = MsgBox("������ �ڷḦ �����Ͻðڽ��ϱ� ?", vbQuestion + vbYesNo + vbDefaultButton1, "�ڷ� ����")
    End If
    If intRetVal = vbNo Then
       vsfg1.SetFocus
       Exit Sub
    End If
    cmdSave.Enabled = False
    With vsfg1
         Screen.MousePointer = vbHourglass
         P_adoRec.CursorLocation = adUseClient
         If Text1(Text1.LBound).Enabled = True Then '����� �߰��� �˻�
            strSQL = "SELECT * FROM ����� " _
                    & "WHERE ������ڵ� = '" & Trim(Text1(Text1.LBound).Text) & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount <> 0 Then
               P_adoRec.Close
               Text1(Text1.LBound).SetFocus
               Screen.MousePointer = vbDefault
               cmdSave.Enabled = True
               Exit Sub
            End If
            P_adoRec.Close
         End If
         PB_adoCnnSQL.BeginTrans
         If Text1(Text1.LBound).Enabled = True Then '����� �߰�
            strSQL = "INSERT INTO �����(������ڵ�, ����ڸ�, ����ڱ���," _
                                          & "�α��ο���, �α��κ�й�ȣ, �����й�ȣ," _
                                          & "������ڵ�, ��뱸��, ��������," _
                                          & "�����Ͻ�, �����Ͻ�) Values( " _
                    & "'" & Trim(Text1(0).Text) & "','" & Trim(Text1(1).Text) & "','" & Trim(Text1(2).Text) & "', " _
                    & "'" & Trim(Text1(3).Text) & "','" & Trim(Text1(4).Text) & "','" & Trim(Text1(5).Text) & "', " _
                    & "'" & Trim(Text1(6).Text) & "'," & Val(Text1(7).Text) & ", '" & PB_regUserinfoU.UserServerDate & "', " _
                    & "'','' )"
            On Error GoTo ERROR_TABLE_INSERT
            .AddItem .Rows
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 7
                            .TextMatrix(.Rows - 1, lngC) = Val(Trim(Text1(lngC).Text))
                            Select Case Val(Text1(lngC).Text)
                                   Case 0
                                        .TextMatrix(.Rows - 1, lngC + 1) = "����"
                                   Case 9
                                        .TextMatrix(.Rows - 1, lngC + 1) = "���Ұ�"
                                   Case Else
                                        .TextMatrix(.Rows - 1, lngC + 1) = "���п���"
                            End Select
                       Case Else
                            .TextMatrix(.Rows - 1, lngC) = Text1(lngC).Text
                            If lngC = 0 Then .Cell(flexcpData, .Rows - 1, lngC, .Rows - 1, lngC) = Text1(lngC).Text
                End Select
            Next lngC
            If .Rows > PC_intRowCnt Then
               .ScrollBars = flexScrollBarBoth
               .TopRow = .Rows - 1
            End If
            Text1(Text1.LBound).Enabled = False
            .Row = .Rows - 1          '�ڵ����� vsfg1_EnterCell Event �߻�
         Else                                          '����� ����
            strSQL = "UPDATE ����� SET " _
                          & "����ڸ� = '" & Trim(Text1(1).Text) & "', " _
                          & "����ڱ��� = '" & Trim(Text1(2).Text) & "', " _
                          & "�α��ο��� = '" & Trim(Text1(3).Text) & "', " _
                          & "�α��κ�й�ȣ = '" & Trim(Text1(4).Text) & "', " _
                          & "�����й�ȣ = '" & Trim(Text1(5).Text) & "', " _
                          & "������ڵ� = '" & Trim(Text1(6).Text) & "', " _
                          & "��뱸�� = " & Val(Text1(7).Text) & ", " _
                          & "�������� = '" & PB_regUserinfoU.UserServerDate & "' " _
                    & "WHERE ������ڵ� = '" & Trim(Text1(Text1.LBound).Text) & "' "
            On Error GoTo ERROR_TABLE_UPDATE
            For lngC = Text1.LBound To Text1.UBound
                Select Case lngC
                       Case 7
                            .TextMatrix(.Row, lngC) = Vals(Trim(Text1(lngC).Text))
                            Select Case Val(Text1(lngC).Text)
                                   Case 0
                                        .TextMatrix(.Row, lngC + 1) = "����"
                                   Case 9
                                        .TextMatrix(.Row, lngC + 1) = "���Ұ�"
                                   Case Else
                                        .TextMatrix(.Row, lngC + 1) = "���п���"
                            End Select
                       Case Else
                            .TextMatrix(.Row, lngC) = Text1(lngC).Text
                End Select
            Next lngC
         End If
         PB_adoCnnSQL.Execute strSQL
         PB_adoCnnSQL.CommitTrans
         '����������� �ٲ� ��� ������Ʈ�� ����
         If PB_regUserinfoU.UserCode = Trim(Text1(0).Text) And PB_regUserinfoU.UserAuthority = "99" And _
            PB_regUserinfoU.UserBranchCode <> Trim(Text1(6).Text) Then
            P_adoRec.CursorLocation = adUseClient
            strSQL = "SELECT ISNULL(������,'') AS ������ FROM ����� " _
                    & "WHERE ������ڵ� = '" & Trim(Trim(Text1(6).Text)) & "' "
            On Error GoTo ERROR_TABLE_SELECT
            P_adoRec.Open strSQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
            If P_adoRec.RecordCount = 1 Then
               PB_regUserinfoU.UserBranchName = P_adoRec("������")
            End If
            P_adoRec.Close
         End If
         If PB_regUserinfoU.UserCode = Trim(Text1(0).Text) Then
            If PB_regUserinfoU.UserAuthority = "99" Then
               If (PB_regUserinfoU.UserName <> Trim(Text1(1).Text) Or _
                   PB_regUserinfoU.UserAuthority <> Trim(Text1(2).Text) Or _
                   PB_regUserinfoU.UserLoginPasswd <> Trim(Text1(4).Text) Or _
                   PB_regUserinfoU.UserSanctionPasswd <> Trim(Text1(5).Text) Or _
                   PB_regUserinfoU.UserBranchCode <> Trim(Text1(6).Text)) Then
                   PB_regUserinfoU.UserName = Trim(Text1(1).Text)
                   PB_regUserinfoU.UserAuthority = Trim(Text1(2).Text)
                   PB_regUserinfoU.UserLoginPasswd = Trim(Text1(4).Text)
                   PB_regUserinfoU.UserSanctionPasswd = Trim(Text1(5).Text)
                   PB_regUserinfoU.UserBranchCode = Trim(Text1(6).Text)
                   UserinfoU_Save PB_regUserinfoU
                   frmMain.Caption = PB_strSystemName & " - " & PB_regUserinfoU.UserBranchName
               End If
            Else
               If (PB_regUserinfoU.UserName <> Trim(Text1(1).Text) Or _
                   PB_regUserinfoU.UserLoginPasswd <> Trim(Text1(4).Text) Or _
                   PB_regUserinfoU.UserSanctionPasswd <> Trim(Text1(5).Text)) Then
                   PB_regUserinfoU.UserName = Trim(Text1(1).Text)
                   PB_regUserinfoU.UserLoginPasswd = Trim(Text1(4).Text)
                   PB_regUserinfoU.UserSanctionPasswd = Trim(Text1(5).Text)
                   UserinfoU_Save PB_regUserinfoU
               End If
            End If
            frmMain.SBar.Panels(2).Text = "�� �� �� : " & PB_regUserinfoU.UserName
         End If
         .SetFocus
         Screen.MousePointer = vbDefault
    End With
    cmdSave.Enabled = True
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �б� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_INSERT:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� �߰� ����"
    Unload Me
    Exit Sub
ERROR_TABLE_UPDATE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� ���� ����"
    Unload Me
    Exit Sub
End Sub

'+-----------+
'/// ���� ///
'+-----------+
Private Sub cmdDelete_Click()
Dim strSQL       As String
Dim intRetVal    As Integer
Dim lngCnt       As Long
    With vsfg1
         If .Row >= .FixedRows Then
            intRetVal = MsgBox("��ϵ� �ڷḦ �����Ͻðڽ��ϱ� ?", vbCritical + vbYesNo + vbDefaultButton2, "�ڷ� ����")
            If intRetVal = vbYes Then
               Screen.MousePointer = vbHourglass
               cmdDelete.Enabled = False
               '������ �������̺� �˻�
               'P_adoRec.CursorLocation = adUseClient
               'strSQL = "SELECT Count(*) AS �ش�Ǽ� FROM TableName " _
               '        & "WHERE ����ڱ��� = " & .TextMatrix(.Row, 0) & " "
               'On Error GoTo ERROR_TABLE_SELECT
               'P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
               'lngCnt = P_adoRec("�ش�Ǽ�")
               'P_adoRec.Close
               If lngCnt <> 0 Then
                  MsgBox "XXX(" & Format(lngCnt, "#,#") & "��)�� �����Ƿ� ������ �� �����ϴ�.", vbCritical, "����� ���� �Ұ�"
                  cmdDelete.Enabled = True
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               PB_adoCnnSQL.BeginTrans
               P_adoRec.CursorLocation = adUseClient
               strSQL = "DELETE FROM ����� WHERE ������ڵ� = " & .TextMatrix(.Row, 0) & " "
               On Error GoTo ERROR_TABLE_DELETE
               PB_adoCnnSQL.Execute strSQL
               PB_adoCnnSQL.CommitTrans
               .RemoveItem .Row
               Text1(Text1.LBound).Enabled = False
               If .Rows <= PC_intRowCnt Then
                  .ScrollBars = flexScrollBarHorizontal
               End If
               Screen.MousePointer = vbDefault
               If .Rows = 1 Then
                  SubClearText
                  .Row = 0
                  Text1(Text1.LBound).Enabled = True
                  Text1(Text1.LBound).SetFocus
                  Exit Sub
               End If
               vsfg1_EnterCell
               cmdDelete.Enabled = True
            End If
            .SetFocus
         End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_TABLE_DELETE:
    PB_adoCnnSQL.RollbackTrans
    MsgBox Err.Number & Err.Description & _
           vbCr & strSQL & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "����� ���� ����"
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
    Set frm��������� = Nothing
End Sub

'+--------------------+
'/// Sub Procedure ///
'+--------------------+
'+----------------------------------+
'/// VsFlexGrid(vsfg1) �ʱ�ȭ ///
'+----------------------------------+
Private Sub Subvsfg1_INIT()
Dim lngR    As Long
Dim lngC    As Long
    Text1(Text1.LBound).Enabled = False                '������ڵ� FLASE
    With vsfg1              'Rows 1, Cols 11, RowHeightMax(Min) 300
         .AllowBigSelection = False
         .AllowSelection = False
         .AllowUserResizing = flexResizeColumns
         .BackColorBkg = &H8000000F
         .BackColorSel = &H8000&
         .ExtendLastCol = True
         .ScrollBars = flexScrollBarHorizontal
         .ScrollTrack = True
         .SelectionMode = flexSelectionByRow
         .FixedRows = 1
         .FixedCols = 2
         .Rows = 1             'Subvsfg1_Fill����ÿ� ����
         .Cols = 11
         .RowHeightMax = 300
         .RowHeightMin = 300
         .ColWidth(0) = 850    '������ڵ�
         .ColWidth(1) = 1500   '����ڸ�
         .ColWidth(2) = 1500   '����ڱ���
         .ColWidth(3) = 1500   '�α��ο���
         .ColWidth(4) = 1500   '�α��κ�й�ȣ
         .ColWidth(5) = 1500   '�����й�ȣ
         .ColWidth(6) = 1500   '������ڵ�
         .ColWidth(7) = 1      '��뱸��
         .ColWidth(8) = 1000   '��뱸��
         .ColWidth(9) = 2000   '�����Ͻ�
         .ColWidth(10) = 2000  '�����Ͻ�
         .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .TextMatrix(0, 0) = "�ڵ�"
         .TextMatrix(0, 1) = "����ڸ�"
         .TextMatrix(0, 2) = "����ڱ���"
         .TextMatrix(0, 3) = "�α��ο���"
         .TextMatrix(0, 4) = "�α��κ�й�ȣ"
         .TextMatrix(0, 5) = "�����й�ȣ"
         .TextMatrix(0, 6) = "������ڵ�"
         .TextMatrix(0, 7) = "��뱸��"
         .TextMatrix(0, 8) = "��뱸��"
         .TextMatrix(0, 9) = "�����Ͻ�"
         .TextMatrix(0, 10) = "�����Ͻ�"
         .ColHidden(7) = True
         If PB_regUserinfoU.UserAuthority <> "99" Then
            .ColHidden(4) = True: .ColHidden(5) = True
         End If
         .ColAlignment(0) = flexAlignCenterCenter
         .ColAlignment(1) = flexAlignLeftCenter
         For lngC = 2 To 10
             .ColAlignment(lngC) = flexAlignCenterCenter
         Next lngC
    End With
End Sub

'+---------------------------------+
'/// VsFlexGrid(vsfg1) ä���///
'+---------------------------------+
Private Sub Subvsfg1_FILL()
Dim SQL        As String
Dim strWhere   As String
Dim strOrderBy As String
Dim lngR       As Long
Dim lngC       As Long
Dim lngRR      As Long
    P_adoRec.CursorLocation = adUseClient
    If PB_regUserinfoU.UserAuthority < "99" Then
       strWhere = "WHERE T1.������ڵ� = '" & PB_regUserinfoU.UserCode & "' "
    End If
    SQL = "SELECT * " _
          & "FROM ����� T1 " _
         & "" & strWhere & " " _
         & "ORDER BY T1.������ڵ� "
    On Error GoTo ERROR_TABLE_SELECT
    P_adoRec.Open SQL, PB_adoCnnSQL, adOpenStatic, adLockReadOnly
    vsfg1.Rows = P_adoRec.RecordCount + 1
    If P_adoRec.RecordCount = 0 Then
       P_adoRec.Close
       Text1(Text1.LBound).Enabled = True
       Text1(Text1.LBound).SetFocus
       Exit Sub
    Else
       With vsfg1
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlack
            If .Rows <= PC_intRowCnt Then
               .ScrollBars = flexScrollBarHorizontal
            Else
               .ScrollBars = flexScrollBarBoth
            End If
            Do Until P_adoRec.EOF
               lngR = lngR + 1
               .TextMatrix(lngR, 0) = P_adoRec("������ڵ�")
               .Cell(flexcpData, lngR, 0, lngR, 0) = Trim(.TextMatrix(lngR, 0)) 'FindRow ����� ����
               .TextMatrix(lngR, 1) = IIf(IsNull(P_adoRec("����ڸ�")), "", P_adoRec("����ڸ�"))
               .TextMatrix(lngR, 2) = IIf(IsNull(P_adoRec("����ڱ���")), "", P_adoRec("����ڱ���"))
               .TextMatrix(lngR, 3) = IIf(IsNull(P_adoRec("�α��ο���")), "", P_adoRec("�α��ο���"))
               .TextMatrix(lngR, 4) = IIf(IsNull(P_adoRec("�α��κ�й�ȣ")), "", P_adoRec("�α��κ�й�ȣ"))
               .TextMatrix(lngR, 5) = IIf(IsNull(P_adoRec("�����й�ȣ")), "", P_adoRec("�����й�ȣ"))
               .TextMatrix(lngR, 6) = IIf(IsNull(P_adoRec("������ڵ�")), "", P_adoRec("������ڵ�"))
               .TextMatrix(lngR, 7) = IIf(IsNull(P_adoRec("��뱸��")), "", P_adoRec("��뱸��"))
               Select Case .ValueMatrix(lngR, 7)
                      Case 0
                           .TextMatrix(lngR, 8) = "����"
                      Case 9
                           .TextMatrix(lngR, 8) = "���Ұ�"
                      Case Else
                           .TextMatrix(lngR, 8) = "���п���"
               End Select
               .TextMatrix(lngR, 9) = IIf(IsNull(P_adoRec("�����Ͻ�")), "", P_adoRec("�����Ͻ�"))
               If Len(Trim(.TextMatrix(lngR, 9))) <> 0 Then
                  .TextMatrix(lngR, 9) = Format(Left(.TextMatrix(lngR, 9), 14), "0000-00-00 00:00:00")
               End If
               .TextMatrix(lngR, 10) = IIf(IsNull(P_adoRec("�����Ͻ�")), "", P_adoRec("�����Ͻ�"))
               If Len(Trim(.TextMatrix(lngR, 10))) <> 0 Then
                  .TextMatrix(lngR, 10) = Format(Left(.TextMatrix(lngR, 10), 14), "0000-00-00 00:00:00")
               End If
               If .TextMatrix(lngR, 0) = PB_regUserinfoU.UserCode Then
                  lngRR = lngR
               End If
               P_adoRec.MoveNext
            Loop
            P_adoRec.Close
            If lngRR = 0 Then
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Rows - PC_intRowCnt + .FixedRows
               End If
               Text1(Text1.LBound).Enabled = True
               Text1(Text1.LBound).SetFocus
               Exit Sub
            Else
               .Row = lngRR       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� �ڵ����� ����)
               If .Rows > PC_intRowCnt Then
                  .TopRow = .Row
               End If
            End If
            vsfg1_EnterCell       'vsfg1_EnterCell �ڵ�����(���� �Ѱ� �϶��� ������ �ڵ�����)
            .SetFocus
       End With
    End If
    Exit Sub
ERROR_TABLE_SELECT:
    MsgBox Err.Number & Err.Description & _
           vbCr & "�������� ���ῡ �����߽��ϴ�. ���α׷��� �����մϴ�.", vbCritical, "��������� �б� ����"
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
End Sub

'+-------------------------+
'/// Check text1(index) ///
'+-------------------------+
Private Function FncCheckTextBox(lngC As Long, blnOK As Boolean)
    For lngC = Text1.LBound To Text1.UBound
        Select Case lngC
               Case 0
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0000")
                    If Not (Text1(lngC).Text >= "0001" And Text1(lngC).Text <= "9999") Then
                       Exit Function
                    End If
               Case 2  '����ڱ���
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "00")
                    If Not (Text1(lngC).Text >= "01" And Text1(lngC).Text <= "99") Then
                       Exit Function
                    End If
               Case 3  '�α��ο���
                    Text1(lngC).Text = UPPER(Trim(Text1(lngC).Text))
                    If Text1(lngC).Text <> "Y" Then
                       Text1(lngC).Text = "N"
                    End If
                    If Not (Text1(lngC).Text = "Y" Or Text1(lngC).Text = "N") Then
                       Exit Function
                    End If
               Case 4  '�α��κ�й�ȣ
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0000")
                    If Not (Text1(lngC).Text >= "0001" And Text1(lngC).Text <= "9999") Then
                       Exit Function
                    End If
               Case 5  '�����й�ȣ
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0000")
                    If Not (Text1(lngC).Text >= "0001" And Text1(lngC).Text <= "9999") Then
                       Exit Function
                    End If
               Case 6  '������ڵ�
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "00")
                    If Not (Text1(lngC).Text >= "01" And Text1(lngC).Text <= "99") Then
                       Exit Function
                    End If
               Case 7  '��뱸��
                    Text1(lngC).Text = Format(Val(Trim(Text1(lngC).Text)), "0")
                    If Not (Text1(lngC).Text >= "0" And Text1(lngC).Text <= "9") Then
                       Exit Function
                    End If
               Case Else
        End Select
    Next lngC
    blnOK = True
End Function
