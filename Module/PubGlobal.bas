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
'| API 함수에서 사용
'+----------------------------------+
Public Const SC_CLOSE = &HF060
Public Const SC_MAXIMIZE = &HF030
Public Const SC_MINIMIZE = &HF020
Public Const MF_BYCOMMAND = &H0&

Public Const CB_SHOWDROPDOWN = &H14F

'+--------------+
'| 공용변수
'+--------------+
Public PB_strSystemName        As String               '예) 판매 관리 시스템
Public PB_Test                 As Integer              'Test Mode = 1, Real Mode = 0
Public PB_strConnServerName    As String
Public PB_strConnDSN           As String
Public PB_strConnUserId        As String
Public PB_strConnDataBaseName  As String
Public PB_adoCnnSQL            As New ADODB.Connection 'MS-SQL SERVER 와의 연결 (object 2.1 liblary)
Public PB_blnStatusOfConn      As Boolean              'MS SQL SERVER 와의 연결상태 (True:정상, False:실패)
Public PB_varErrCode           As Variant              'Error Check Code
Public PB_blnNew               As Boolean
Public PB_curVatRate           As Currency             '부가세율(예.10% -> (10.00/100))

'단가변경
Public PB_intIAutoPriceGbn     As Integer              '최종입고단가자동갱신구분(1.자동변경)
Public PB_intOAutoPriceGbn     As Integer              '최종출고단가자동갱신구분(1.자동변경)
'연결
Public PB_adoCnnMDB            As New ADODB.Connection 'MDB 접속
Public PB_strFileAccessMDB     As String               '*.mdb(Access) 파일의 위치
'레지스트리
Public PB_regUserinfoU         As UserinfoU
'우편번호검색
Public PB_strPostCode          As String               '우편번호
Public PB_strPostName          As String               '주    소
'매입(출)처검색
Public PB_strSupplierCode      As String               '매입처코드
Public PB_strSupplierName      As String               '매입처명
'자재검색
Public PB_strMaterialsCode     As String               '자재코드(분류코드+세부코드)
Public PB_strMaterialsName     As String               '자재명
'계정코드검색
Public PB_strAccCode           As String               '계정코드
Public PB_strAccName           As String               '계정명

Public PB_strFMCCallFormName   As String               '자재시세검색을 로드한 폼의 이름
Public PB_strCallFormName      As String               '자재검색을 로드한 폼의 이름
Public PB_strFMWCallFormName   As String               '자재원장검색을 로드한 폼의 이름

'프린터환경설정
Public PB_intPrtTypeGbn        As Integer              '출력타입구분(거래명세서/세금계산서)
Public PB_intDLeftMargin       As Integer              '왼쪽마진
Public PB_intDTopMargin        As Integer              '상단마진
Public PB_intTLeftMargin       As Integer              '왼쪽마진
Public PB_intTTopMargin        As Integer              '상단마진

Public PB_strEnterNo           As String               '사업자번호
Public PB_strEnterName         As String               '상호
Public PB_strRepName           As String               '대표자명
Public PB_strEnterAddress      As String               '주소
Public PB_strUptae             As String               '업태
Public PB_strUpjong            As String               '업종
'+--------------+
'| OBDC 체크
'+--------------+
Public Sub InitConnection()
    PB_strConnDSN = ""
End Sub
Public Sub DoConnection()
Dim varArr    As Variant
    varArr = GetCommandLine()
    If UBound(varArr) > 0 Then
       '예) 명령줄 인수 = "localhost YmhDB ymhuser" 일때(1.서버이름, 2.ODBC DSN, 3.userid, 4.DataBaseName
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
       MsgBox "사용할 ODBC(DSN) 이름을 매개변수로 전달해 주세요.", vbInformation, "ODBC(DSN) 연결"
       End
    End If
End Sub

'+------------------------------------------------------------------------------------------------------+
'| Get Command Line   : 프로젝트(P) - XX 정보 속성(E) - 만들기 - 명령줄인수
'| IsMissing(argname) : 선택적인 Variant 인수가 프로시저에 전달되었는지 나타내는 Boolean 값을 반환합니다.
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
'| MS ACCESS SERVER 연결 |
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
'/// MS-SQL SERVER 연결 ///
'+-------------------------+
Public Function PB_Fnc_AdoCnnSQL()
    InitConnection
    DoConnection
    PB_Test = 0
    'PB_strConnServerName = "DANSEPO"
    PB_adoCnnSQL.CursorLocation = adUseClient
    If PB_Test = 0 Then
       'MS-SQL SERVER CONNECTION(ODBC User DSN만)
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

