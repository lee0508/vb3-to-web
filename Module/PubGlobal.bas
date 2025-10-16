Attribute VB_Name = "PubGlobal"
Option Explicit

'+--------------+
'| API Function
'+--------------+
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
                 
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
                 
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
                
'Printer
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
                
'+----------------------------------+
'| API �Լ����� ���
'+----------------------------------+
Public Const SC_CLOSE = &HF060
Public Const SC_MAXIMIZE = &HF030
Public Const SC_MINIMIZE = &HF020
Public Const MF_BYCOMMAND = &H0&

Public Const CB_SHOWDROPDOWN = &H14F

'+--------------+
'| ���뺯��
'+--------------+
Public PB_strSystemName        As String               '��) �Ǹ� ���� �ý���
Public PB_Test                 As Integer              'Test Mode = 1, Real Mode = 0
Public PB_strConnServerName    As String
Public PB_strConnDSN           As String
Public PB_strConnUserId        As String
Public PB_strConnDataBaseName  As String
Public PB_adoCnnSQL            As New ADODB.Connection 'MS-SQL SERVER ���� ���� (object 2.1 liblary)
Public PB_blnStatusOfConn      As Boolean              'MS SQL SERVER ���� ������� (True:����, False:����)
Public PB_varErrCode           As Variant              'Error Check Code
Public PB_blnNew               As Boolean
Public PB_curVatRate           As Currency             '�ΰ�����(��.10% -> (10.00/100))

'�ܰ�����
Public PB_intIAutoPriceGbn     As Integer              '�����԰�ܰ��ڵ����ű���(1.�ڵ�����)
Public PB_intOAutoPriceGbn     As Integer              '�������ܰ��ڵ����ű���(1.�ڵ�����)
'����
Public PB_adoCnnMDB            As New ADODB.Connection 'MDB ����
Public PB_strFileAccessMDB     As String               '*.mdb(Access) ������ ��ġ
'������Ʈ��
Public PB_regUserinfoU         As UserinfoU
'�����ȣ�˻�
Public PB_strPostCode          As String               '�����ȣ
Public PB_strPostName          As String               '��    ��
'����(��)ó�˻�
Public PB_strSupplierCode      As String               '����ó�ڵ�
Public PB_strSupplierName      As String               '����ó��
'����˻�
Public PB_strMaterialsCode     As String               '�����ڵ�(�з��ڵ�+�����ڵ�)
Public PB_strMaterialsName     As String               '�����
'�����ڵ�˻�
Public PB_strAccCode           As String               '�����ڵ�
Public PB_strAccName           As String               '������

Public PB_strFMCCallFormName   As String               '����ü��˻��� �ε��� ���� �̸�
Public PB_strCallFormName      As String               '����˻��� �ε��� ���� �̸�
Public PB_strFMWCallFormName   As String               '�������˻��� �ε��� ���� �̸�

'������ȯ�漳��
Public PB_intPrtTypeGbn        As Integer              '���Ÿ�Ա���(�ŷ�����/���ݰ�꼭)
Public PB_intDLeftMargin       As Integer              '���ʸ���
Public PB_intDTopMargin        As Integer              '��ܸ���
Public PB_intTLeftMargin       As Integer              '���ʸ���
Public PB_intTTopMargin        As Integer              '��ܸ���

Public PB_strEnterNo           As String               '����ڹ�ȣ
Public PB_strEnterName         As String               '��ȣ
Public PB_strRepName           As String               '��ǥ�ڸ�
Public PB_strEnterAddress      As String               '�ּ�
Public PB_strUptae             As String               '����
Public PB_strUpjong            As String               '����
'+--------------+
'| OBDC üũ
'+--------------+
Public Sub InitConnection()
    PB_strConnDSN = ""
End Sub
Public Sub DoConnection()
Dim varArr    As Variant
    varArr = GetCommandLine()
    If UBound(varArr) > 0 Then
       '��) ����� �μ� = "localhost YmhDB ymhuser" �϶�(1.�����̸�, 2.ODBC DSN, 3.userid, 4.DataBaseName
       'varArr(1) = localhost, varArr(2) = "YmhDB", varArr(3) = "ymhuser", varArr(4) = "YmhDB"
       If Len(varArr(1)) = 0 Then
          PB_strConnServerName = "localhost"
       Else
          PB_strConnServerName = varArr(1)
       End If
       If Len(varArr(2)) = 0 Then
          PB_strConnDSN = "YmhDB"
       Else
          PB_strConnDSN = varArr(2)
       End If
       If Len(varArr(3)) = 0 Then
          PB_strConnUserId = "ymhuser"
       Else
          PB_strConnUserId = varArr(3)
       End If
       If Len(varArr(4)) = 0 Then
          PB_strConnDataBaseName = PB_strConnDSN
       Else
          PB_strConnDataBaseName = varArr(4)
       End If
    Else
       PB_strConnDSN = ""
       MsgBox "����� ODBC(DSN) �̸��� �Ű������� ������ �ּ���.", vbInformation, "ODBC(DSN) ����"
       End
    End If
End Sub

'+------------------------------------------------------------------------------------------------------+
'| Get Command Line   : ������Ʈ(P) - XX ���� �Ӽ�(E) - ����� - ������μ�
'| IsMissing(argname) : �������� Variant �μ��� ���ν����� ���޵Ǿ����� ��Ÿ���� Boolean ���� ��ȯ�մϴ�.
'+------------------------------------------------------------------------------------------------------+
Public Function GetCommandLine(Optional MaxArgs) As Variant
Dim blnArg                      As Boolean
Dim strCmdLine, strString       As String
Dim intArgs, intCmdLnLen, inti  As Integer
Dim intMinArgs                  As Integer
    intMinArgs = 4
    If IsMissing(MaxArgs) = True Then
       MaxArgs = 10
    End If
    ReDim ArgArray(0 To MaxArgs)
    intArgs = 0: blnArg = False
    strCmdLine = Trim(Command())   'strCmdLine = "localhost YmhDB ymhuser"
    intCmdLnLen = Len(strCmdLine)  'intCmdLnLen = 13
    For inti = 1 To intCmdLnLen
        strString = Mid(strCmdLine, inti, 1)
        If (strString <> " " And strString <> vbTab) Then
            If Not blnArg Then
               If intArgs = MaxArgs Then Exit For
               intArgs = intArgs + 1
               blnArg = True
            End If
            ArgArray(intArgs) = ArgArray(intArgs) & strString
        Else
            blnArg = False
        End If
    Next inti
    ReDim Preserve ArgArray(0 To IIf(intArgs < intMinArgs, intMinArgs, intArgs))
    GetCommandLine = ArgArray()
End Function

'+-----------------------+
'| MS ACCESS SERVER ���� |
'+-----------------------+
Public Function PB_Fnc_AdoCnnMDB()
    If PB_Test = 1 Then
       'PB_strFileSaroMDB = "\\XX\c\XXDB.mdb"
    Else
       'PB_strFileSaroMDB = "\\YY\SXXDB.mdb"
    End If
    On Error GoTo DRIVER_ERROR_HANDLER
    PB_varErrCode = Dir(PB_strFileAccessMDB)
    On Error GoTo DRIVER_ERROR_HANDLER
    PB_adoCnnMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & PB_strFileAccessMDB & ";User Id=admin;Password=;"
    On Error GoTo ERROR_MDB_CONNECTION
    Exit Function
ERROR_MDB_CONNECTION:
    PB_varErrCode = Err.Number
    Screen.MousePointer = vbDefault
    Exit Function
DRIVER_ERROR_HANDLER:
    PB_varErrCode = Err.Number
    Screen.MousePointer = vbDefault
    Exit Function
End Function

'+-------------------------+
'/// MS-SQL SERVER ���� ///
'+-------------------------+
Public Function PB_Fnc_AdoCnnSQL()
    InitConnection
    DoConnection
    PB_Test = 0
    'PB_strConnServerName = "DANSEPO"
    PB_adoCnnSQL.CursorLocation = adUseClient
    If PB_Test = 0 Then
       'MS-SQL SERVER CONNECTION(ODBC User DSN��)
       On Error GoTo ERROR_SQL_CONNECTION
       'PB_adoCnnSQL.Open "Server=218.38.206.2;DSN=YmhDB;uid=sa;pwd=;DataBase=YmhDB"
       'PB_adoCnnSQL.Open "Provider=MSDASQL;Driver={SQLSERVER};DSN=YmhDB;uid=ymhuser;pwd=userymh;DataBase=YmhDB"
       'PB_adoCnnSQL.Open "Server=localhost;DSN=YmhDB;uid=ymhuser;pwd=userymh;DataBase=YmhDB"
       PB_adoCnnSQL.Open "Server=" & PB_strConnServerName & ";" _
                       & "DSN=" & PB_strConnDSN & ";uid=" & PB_strConnUserId & ";pwd=userymh;DataBase=" & PB_strConnDataBaseName & ""
       PB_adoCnnSQL.DefaultDatabase = PB_strConnDataBaseName
    Else
       'PB_adoCnnSQL.ConnectionString = "Driver={SQL Server};Server=218.38.206.2;Uid=sa;Pwd=1730"
       On Error GoTo ERROR_SQL_CONNECTION
       'PB_adoCnnSQL.Open "Server=218.38.206.2;DSN=TestYmhDB;uid=sa;pwd=;DataBase=TestYmhDB"
       PB_adoCnnSQL.Open "Server=Dansepo;DSN=TestYmhDB;uid=ymhuser;pwd=userymh;DataBase=TestYmhDB"
       PB_adoCnnSQL.DefaultDatabase = "TestYmhDB"
    End If
    PB_blnStatusOfConn = True
    Exit Function
ERROR_SQL_CONNECTION:
    PB_blnStatusOfConn = False
    PB_varErrCode = Err.Number & vbCr & Err.Description
    Screen.MousePointer = vbDefault
    Exit Function
DRIVER_ERROR_HANDLER:
    PB_blnStatusOfConn = False
    PB_varErrCode = Err.Number
    Screen.MousePointer = vbDefault
    Exit Function
End Function

