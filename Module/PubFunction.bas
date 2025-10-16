Attribute VB_Name = "PubFunction"
'********************************************************************************
'* 상수                                                                         *
'********************************************************************************
 Option Explicit
 Public Const CURRENCY_MIN As Currency = 0
 Public Const CURRENCY_MAX As Currency = 922337203685477#
 Public Const VBCheckUnselected = 0
 Public Const VBCheckSelected = 1
 Public Const VBCheckGrayed = 2
 Public Const intResDriveNotReady As Integer = -1
 Public Const intResDriveReady    As Integer = 0
 Public Const intResDiskNotReady  As Integer = 1

 Public Const misTRUE  As String = "TRUE"
 Public Const misFALSE As String = "FALSE"


'********************************************************************************
'* iLsu()                                                                       *
'* 일수를 구한다.                                                               *
'********************************************************************************
 Public Function ilsu(ByVal sDate As String) As Long

    '선언
     Dim nDaysInYear      As Single      '년간 총일수 평균
     Dim nDaysInMonth     As Single      '월간 총일수 평균
     Dim nYear            As Integer     '당해 년도
     Dim nMonth           As Integer     '당해 월
     Dim nDay             As Integer     '당해 일
     Dim nThisDaysInYear  As Single      '당해 년도까지의 총일수
     Dim nThisDaysInMonth As Single      '당해 월까지의 총일수
     Dim nThisDaysInTotal As Long        '총일수

    '정의
     nDaysInYear = 365.25
     nDaysInMonth = 30.6
     nYear = Val(Mid(sDate, 1, 4))
     nMonth = Val(Mid(sDate, 5, 2))
     nDay = Val(Mid(sDate, 7, 2))

    '년,월,일 재조정
     If nMonth < 3 Then
        nYear = nYear - 1
        nMonth = nMonth + 12
     End If
     nMonth = nMonth + 1

    '일수계산 (0000년 1월 1일 기준)
     nThisDaysInYear = nYear * nDaysInYear    '당해 년도까지의 총일수를 구한다.
     nThisDaysInYear = Fix(nThisDaysInYear)
     nThisDaysInMonth = nMonth * nDaysInMonth '당해 월까지의 총일수를 구한다.
     nThisDaysInMonth = Fix(nThisDaysInMonth)
     nThisDaysInTotal = nThisDaysInYear + nThisDaysInMonth + nDay '총일수
     nThisDaysInTotal = Fix(nThisDaysInTotal)
            
    '돌림
     ilsu = nThisDaysInTotal

 End Function

'********************************************************************************
'* mkDate()                                                                     *
'* 일자를 만든다.                                                               *
'********************************************************************************
 Public Function Mkdate(ByVal nThisDaysInTotal As Single) As String

    '선언
     Dim nDaysInYear      As Single      '년간 총일수 평균
     Dim nDaysInMonth     As Single      '월간 총일수 평균
     Dim aDaysInMonth(12) As Integer     '월별 일수
     Dim nYear            As Single      '당해 년도
     Dim nMonth           As Single      '당해 월
     Dim nDay             As Single      '당해 일
     Dim nTempYear        As Single      '임시 년도

    '정의
     nDaysInYear = 365.25
     nDaysInMonth = 30.6
     aDaysInMonth(1) = 31
     aDaysInMonth(2) = 28
     aDaysInMonth(3) = 31
     aDaysInMonth(4) = 30
     aDaysInMonth(5) = 31
     aDaysInMonth(6) = 30
     aDaysInMonth(7) = 31
     aDaysInMonth(8) = 31
     aDaysInMonth(9) = 30
     aDaysInMonth(10) = 31
     aDaysInMonth(11) = 30
     aDaysInMonth(12) = 31

    '날짜를 구한다.
     nYear = nThisDaysInTotal - 122.1
     nYear = Fix(nYear / nDaysInYear)
     nTempYear = Fix(nYear * nDaysInYear)
     nTempYear = nThisDaysInTotal - nTempYear

     nMonth = Fix(nTempYear / nDaysInMonth)
     nDay = Fix(nDaysInMonth * nMonth)
     nDay = nTempYear - nDay
     
    '날짜조정
     If nDay = 0 Then
        nMonth = nMonth - 1
        nDay = 31
     End If
     If nMonth >= 14 Then
        nMonth = nMonth - 12
     End If
     nMonth = nMonth - 1
     If nMonth < 3 Then
        nYear = nYear + 1
     End If

    '돌림
     Mkdate = Format(nYear, "0000") + Format(nMonth, "00") + Format(nDay, "00")

 End Function
'********************************************************************************
'* mkMonth()                                                                    *
'* 일자를 만든다.                                                               *
'********************************************************************************
Public Function MkMonth(ByVal H_Date As String, ByVal Ijm As Single, _
                              D_Date As String, ByVal Mon_Ck As Integer) As String
Dim Cb        As Integer
Dim M_1       As Integer
Dim M_2       As Integer
Dim M_3       As Integer
Dim Mok       As Integer
Dim Mal_Ck    As Integer
Dim Im_YY     As Single
Dim Im_MM     As Single
Dim c(12)     As Integer
    c(1) = 31: c(2) = 28: c(3) = 31: c(4) = 30: c(5) = 31: c(6) = 30
    c(7) = 31: c(8) = 31: c(9) = 30: c(10) = 31: c(11) = 30: c(12) = 31
    Mal_Ck = 0: D_Date = "": M_1 = 0: M_2 = 0: M_3 = 0: Im_YY = 0: Im_MM = 0
    If Right(H_Date, 2) = c(Val(Mid(H_Date, 5, 2))) Then
       Mal_Ck = 1
    Else
       Mal_Ck = 0
    End If
    Im_MM = Val(Mid(H_Date, 5, 2)) + Ijm
Job_1:
    If Im_MM > 12 Then
       H_Date = STUFF(H_Date, 1, 4, Format(Vals(Left(H_Date, 4)) + 1, "0000"))
       Im_MM = Im_MM - 12
       GoTo Job_1
    End If
Job_2:
    If Im_MM < 1 Then
       H_Date = STUFF(H_Date, 1, 4, Format(Vals(Left(H_Date, 4)) - 1, "0000"))
       Im_MM = Im_MM + 12
       GoTo Job_2
    End If
Job_2_1:
    If Mon_Ck <> 1 Then
       GoTo Job_3
    End If
    H_Date = STUFF(H_Date, 7, 2, Format(Vals(Right(H_Date, 2)) - 1, "00"))
    If Val(Right(H_Date, 2)) < 1 Then
       Im_MM = Im_MM - 1
       Mon_Ck = 0
       GoTo Job_2
    End If
Job_3:
    Mok = Val(Right(H_Date, 4)) / 4: M_1 = Val(Right(H_Date, 4)) Mod 4
    Mok = Val(Right(H_Date, 4)) / 100: M_2 = Val(Right(H_Date, 4)) Mod 100
    Mok = Val(Right(H_Date, 4)) / 400: M_3 = Val(Right(H_Date, 4)) Mod 400
    If Val(Mid(H_Date, 5, 2)) = 2 And M_1 = 0 Then
       Cb = 29
    Else
       Cb = c(Im_MM)
    End If
    If Val(Mid(H_Date, 5, 2)) = 2 And M_2 = 0 Then
       Cb = 28
    Else
       Cb = c(Im_MM)
    End If
    If Val(Mid(H_Date, 5, 2)) = 2 And M_3 = 0 Then
       Cb = 29
    Else
       Cb = c(Im_MM)
    End If
    If Val(Right(H_Date, 2)) = 0 Then
       Cb = c(Im_MM)
    End If
Job_4:
    If Val(Right(H_Date, 2)) > Cb Then
       H_Date = STUFF(H_Date, 7, 2, Format(Cb, "00"))
    End If
    If Val(Right(H_Date, 2)) < 1 Then
       H_Date = STUFF(H_Date, 7, 2, Format(Cb, "00"))
    End If
Job_5:
    If (Mal_Ck = 1) And (Mon_Ck = 0) Then
       H_Date = STUFF(H_Date, 7, 2, Format(c(Im_MM), "00"))
    End If
    D_Date = H_Date
    D_Date = STUFF(D_Date, 5, 2, Format(Im_MM, "00"))
End Function

'+-------------------+
'/// 한글금액변환 ///
'+-------------------+
Function hValH(ByVal v As Double) As String

     Dim sUnit(0 To 2) As String
     Dim sStep(0 To 3) As String
     Dim sCall(0 To 9) As String
     Dim s As String
     Dim i As Integer
     Dim j As Integer
     Dim l As Integer
     Dim m As Integer
     Dim n As Integer
     Dim c As String
     Dim k As String

     sUnit(0) = ""
     sUnit(1) = "만"
     sUnit(2) = "억"
     sStep(0) = ""
     sStep(1) = "십"
     sStep(2) = "백"
     sStep(3) = "천"
     sCall(0) = ""
     sCall(1) = "일"
     sCall(2) = "이"
     sCall(3) = "삼"
     sCall(4) = "사"
     sCall(5) = "오"
     sCall(6) = "육"
     sCall(7) = "칠"
     sCall(8) = "팔"
     sCall(9) = "구"

     hValH = ""

     s = Trim(Str(v))
     n = Len(s)
     If n > 0 Then
        j = -1
        l = -1
        m = 1
        k = ""
        For i = n To m Step -1
            l = IIf(l = 3, 0, l + 1)
            c = sCall(Val(Mid(s, i, 1)))
            c = c + IIf(c = "", "", sStep(l))
            If l = 0 Then
               j = j + 1
               c = c + sUnit(j)
            End If
             k = c + k
        Next i

        hValH = k
     End If

 End Function

'+-------------------+
'/// 한문금액변환 ///
'+-------------------+
Function hMValH(ByVal v As Double) As String

     Dim sUnit(0 To 2) As String
     Dim sStep(0 To 3) As String
     Dim sCall(0 To 9) As String
     Dim s As String
     Dim i As Integer
     Dim j As Integer
     Dim l As Integer
     Dim m As Integer
     Dim n As Integer
     Dim c As String
     Dim k As String

     sUnit(0) = ""
     sUnit(1) = "萬"
     sUnit(2) = "億"
     sStep(0) = ""
     sStep(1) = "拾"
     sStep(2) = "百" '"佰"
     sStep(3) = "阡"
     sCall(0) = ""
     sCall(1) = "壹"
     sCall(2) = "貳"
     sCall(3) = "參"
     sCall(4) = "四"
     sCall(5) = "五"
     sCall(6) = "六"
     sCall(7) = "七"
     sCall(8) = "八"
     sCall(9) = "九"

     hMValH = ""

     s = Trim(Str(v))
     n = Len(s)
     If n > 0 Then
        j = -1
        l = -1
        m = 1
        k = ""
        For i = n To m Step -1
            l = IIf(l = 3, 0, l + 1)
            c = sCall(Val(Mid(s, i, 1)))
            c = c + IIf(c = "", "", sStep(l))
            If l = 0 Then
               j = j + 1
               c = c + sUnit(j)
            End If
             k = c + k
        Next i

        hMValH = k
     End If

 End Function

'********************************************************************************
'* MaxH()                                                                       *
'* 값 중에서 큰 값을 돌려준다.                                                  *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     MaxH(<값1>, <값2>) --> 큰 값                                             *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <값1>은 수치.                                                            *
'*                                                                              *
'*     <값2>는 수치.                                                            *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     MaxH()는 <값1>과 <값2>중에서 큰 값을 돌려준다.                           *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     MaxH()는 <값1>과 <값2>를 비교하여, 큰 값을 돌려준다.                     *
'*     <값1>과 <값2>는 Double형 데이터형식을 사용하며, 결과값도 Double형으로 돌 *
'*     려준다.                                                                  *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     $ 다음 예제들은, 여러가지 매개변수에 따른 결과를 보여준다.               *
'*                                                                              *
'*       MaxH(-1, -2)                  '-1                                      *
'*       MaxH(0, -1)                   '0                                       *
'*       MaxH(1, 2)                    '2                                       *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     SQL문에서 사용하는 Max()함수와는 다르다.                                 *
'*     ----------------------------------------                                 *
'*                                                                              *
'********************************************************************************
 Public Function MaxH(ByVal n1 As Double, ByVal n2 As Double) As Double

     MaxH = IIf(n1 >= n2, n1, n2)

 End Function


'********************************************************************************
'* MinH()                                                                       *
'* 값 중에서 작은 값을 돌려준다.                                                *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     MinH(<값1>, <값2>) --> 작은 값                                           *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <값1>은 수치.                                                            *
'*                                                                              *
'*     <값2>는 수치.                                                            *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     MinH()은 <값1>과 <값2>중에서 작은 값을 돌려준다.                         *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     MinH()는 <값1>과 <값2>를 비교하여, 작은 값을 돌려준다.                   *
'*     <값1>과 <값2>는 Double형 데이터형식을 사용하며, 결과값도 Double형으로 돌 *
'*     려준다.                                                                  *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     $ 다음 예제들은, 여러가지 매개변수에 따른 결과를 보여준다.               *
'*                                                                              *
'*       MinH(-1, -2)                  '-2                                      *
'*       MinH(0, -1)                   '-1                                      *
'*       MinH(1, 2)                    '1                                       *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     SQL문에서 사용하는 Min()함수와는 다르다.                                 *
'*     ----------------------------------------                                 *
'*                                                                              *
'********************************************************************************
 Public Function MinH(ByVal n1 As Double, ByVal n2 As Double) As Double

     MinH = IIf(n1 >= n2, n2, n1)

 End Function


'********************************************************************************
'* LenH()                                                                       *
'* 문자열의 길이를 돌려준다.                                                    *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     LenH(<문자열>) --> 길이                                                  *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <문자열>은 길이를 계산할 문자열.                                         *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     LenH()는 문자열의 길이를 돌려준다.                                       *
'*     <문자열>이 널문자 ("")이면, LenH()는 ZERO를 돌려준다.                    *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     LenH()는 문자열의 길이를 돌려주는 문자함수이다. 널바이트 (CHR(0))를 포함 *
'*     해서, 각 문자를 1로 계산한다. 반대로, 널문자는 ("") ZERO로 측정한다.     *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     $ 다음 예제들은, 여러가지 매개변수에 따른 결과를 보여준다.               *
'*                                                                              *
'*       LenH("")                      '0                                       *
'*       LenH("1234567890")            '10                                      *
'*       LenH(" 1234 ")                '6                                       *
'*       LenH(" 한글AlphaNumeric ")    '18                                      *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     LenH()는 유니코드를 2바이트로 계산한다.                                  *
'*     ---------------------------------------                                  *
'*                                                                              *
'********************************************************************************
 Public Function LenH(ByVal sStr As String) As Long

     LenH = MaxH(LenB(StrConv(sStr, vbFromUnicode)), 0)

 End Function


'********************************************************************************
'* LeftH()                                                                      *
'* 문자열에서 첫번째문자부터 부분문자열을 추출한다.                             *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     LeftH(<문자열>, <개수>) --> 부분문자열                                   *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <문자열>은 부분문자열을 추출할 전체문자열.                               *
'*                                                                              *
'*     <개수>는 추출할 문자의 개수. 범위는 0 에서 65536 까지이다.               *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     LeftH()는 <문자열>의 가장 왼쪽부터 <개수>만큼의 부분문자열을 추출한다.   *
'*     부분문자열을 추출하는 조건은 아래와 같다.                                *
'*                                                                              *
'*     $ 조건 1. 비주얼베이직에서 제공하는 조건으로 처리했을 때:                *
'*                                                                              *
'*       +-------------------+------------+-------------------+                 *
'*       | <문자열>          | <개수>     | 결과              |                 *
'*       +-------------------+------------+-------------------+                 *
'*       | 널문자|유효문자열 | < 0        | 오류              |                 *
'*       | 널문자|유효문자열 | = 0        | 널문자            |                 *
'*       | 널문자|유효문자열 | > 0        | 널문자|부분문자열 |                 *
'*       | 널문자|유효문자열 | > 전체길이 | 널문자|전체문자열 |                 *
'*       +-------+-----------+------------+-------------------+                 *
'*                                                                              *
'*     $ 조건 2. 수정된 조건으로 처리했을 때:                                   *
'*                                                                              *
'*       +-------------------+------------+-------------------+                 *
'*       | <문자열>          | <개수>     | 결과              |                 *
'*       +-------------------+------------+-------------------+                 *
'*       | 널문자|유효문자열 | < 0        | 널문자            |                 *
'*       | 널문자|유효문자열 | = 0        | 널문자            |                 *
'*       | 널문자|유효문자열 | > 0        | 널문자|부분문자열 |                 *
'*       | 널문자|유효문자열 | > 전체길이 | 널문자|전체문자열 |                 *
'*       +-------------------+------------+-------------------+                 *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     LeftH()는 명시된 문자열의 부분문자열을 돌려주는 문자함수이다. LeftH()는, *
'*     MidH(<문자열>, 1, <개수>)와 같은 역할을 한다. LeftH()는 또한, <문자열>의 *
'*     가장 오른쪽에서부터 부분문자열을 돌려주는 RightH()함수와 비슷하다.       *
'*                                                                              *
'*     LeftH(), RightH() 그리고 MidH()는 종종 AT()와 RAT()함수와 같이 사용된다. *
'*                                                                              *
'*     LeftH()는 2바이트문자가 잘리는 경우, 해당 문자를 포함시키지 않는다. 즉,  *
'*     <개수>만큼의 문자열을 추출할 때, 마지막부분이 2바이트문자이면서 절반으로 *
'*     나뉘어지게 되면, 마지막부분의 2바이트문자를 제외시킨 문자열을 돌려준다.  *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     $ 다음 예제들은, 여러가지 매개변수에 따른 결과를 보여준다:               *
'*                                                                              *
'*       sName = "James, William"                                               *
'*       LeftH("",    -1) '(조건1) error           , (조건2) ""                 *
'*       LeftH("",     0) '(조건1) ""              , (조건2) ""                 *
'*       LeftH("",     1) '(조건1) ""              , (조건2) ""                 *
'*       LeftH("",    14) '(조건1) ""              , (조건2) ""                 *
'*       LeftH("",    15) '(조건1) ""              , (조건2) ""                 *
'*       LeftH(sName, -1) '(조건1) error           , (조건2) ""                 *
'*       LeftH(sName,  0) '(조건1) ""              , (조건2) ""                 *
'*       LeftH(sName,  1) '(조건1) "J"             , (조건2) "J"                *
'*       LeftH(sName, 14) '(조건1) "James, William", (조건2) "James, William"   *
'*       LeftH(sName, 15) '(조건1) "James, William", (조건2) "James, William"   *
'*     > LeftH("가나", 3) '(조건1) "가"            , (조건2) "가"               *
'*                                                                              *
'*     $ 이 예제는, AT()함수를 사용한 경우이다:                                 *
'*                                                                              *
'*       sName = "James, William"                                               *
'*       LeftH(sName, At(",", sName) - 1)            '"James"                   *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     LeftH()는 유니코드를 2바이트로 계산한다.                                 *
'*     ----------------------------------------                                 *
'*                                                                              *
'********************************************************************************
 Public Function LeftH(ByVal sStr As String, ByVal nCount As Long) As String

    '매개변수
     nCount = MinH(MaxH(nCount, 0), 65536) '문자열의 길이를 정의

    '추출
     If sStr = "" Then
        LeftH = ""
     Else
        LeftH = StrConv(LeftB(StrConv(sStr, vbFromUnicode), nCount), vbUnicode)
     End If

 End Function


'********************************************************************************
'* RightH()                                                                     *
'* 문자열에서 마지막문자부터 부분문자열을 추출한다.                             *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     RightH(<문자열>, <개수>) --> 부분문자열                                  *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <문자열>은 부분문자열을 추출할 전체문자열.                               *
'*                                                                              *
'*     <개수>는 추출할 문자의 개수. 범위는 0 에서 65536 까지이다.               *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     RightH()는 <문자열>의 가장 오른쪽부터 <개수>만큼의 부분문자열을 왼쪽으로 *
'*     이동하면서 추출한다. 부분문자열을 추출하는 조건은 아래와 같다.           *
'*                                                                              *
'*     $ 조건 1. 비주얼베이직에서 제공하는 조건으로 처리했을 때:                *
'*                                                                              *
'*       +-------------------+------------+-------------------+                 *
'*       | <문자열>          | <개수>     | 결과              |                 *
'*       +-------------------+------------+-------------------+                 *
'*       | 널문자|유효문자열 | < 0        | 오류              *                 *
'*       | 널문자|유효문자열 | = 0        | 널문자            *                 *
'*       | 널문자|유효문자열 | > 0        | 널문자|부분문자열 *                 *
'*       | 널문자|유효문자열 | > 전체길이 | 널문자|전체문자열 *                 *
'*       +-------------------+------------+-------------------+                 *
'*                                                                              *
'*     $ 조건 2. 수정된 조건으로 처리했을 때:                                   *
'*                                                                              *
'*       +-------------------+------------+-------------------+                 *
'*       | <문자열>          | <개수>     | 결과              |                 *
'*       +-------------------+------------+-------------------+                 *
'*       | 널문자|유효문자열 | < 0        | 널문자            |                 *
'*       | 널문자|유효문자열 | = 0        | 널문자            |                 *
'*       | 널문자|유효문자열 | > 0        | 널문자|부분문자열 |                 *
'*       | 널문자|유효문자열 | > 전체길이 | 널문자|전체문자열 |                 *
'*       +-------------------+------------+-------------------+                 *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     RightH()는 명시된 문자열의 부분문자열을 돌려주는 문자함수이다. Right()는 *
'*     MidH(<문자열>, -<개수>)>와 같은 역할을 한다. RightH()는 또한, <문자열>의 *
'*     가장 왼쪽에서부터 부분문자열을 돌려주는 LeftH()함수와 비슷하다.          *
'*                                                                              *
'*     RightH(), LeftH() 그리고 MidH()는 종종 AT()와 RAT()함수와 같이 사용된다. *
'*                                                                              *
'*     RightH()는 2바이트문자가 잘리는 경우, 해당 문자를 포함시키지 않는다. 즉, *
'*     <개수>만큼의 문자열을 추출할 때, 마지막부분이 2바이트문자이면서 절반으로 *
'*     나뉘어지게 되면, 마지막부분의 2바이트문자를 제외시킨 문자열을 돌려준다.  *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     $ 다음 예제들은, 여러가지 매개변수에 따른 결과를 보여준다:               *
'*                                                                              *
'*       sName = "James, William"                                               *
'*       RightH("",    -1) '(조건1) error           , (조건2) ""                *
'*       RightH("",     0) '(조건1) ""              , (조건2) ""                *
'*       RightH("",     1) '(조건1) ""              , (조건2) ""                *
'*       RightH("",    14) '(조건1) ""              , (조건2) ""                *
'*       RightH("",    15) '(조건1) ""              , (조건2) ""                *
'*       RightH(sName, -1) '(조건1) error           , (조건2) ""                *
'*       RightH(sName,  0) '(조건1) ""              , (조건2) ""                *
'*       RightH(sName,  1) '(조건1) "m"             , (조건2) "m"               *
'*       RightH(sName, 14) '(조건1) "James, William", (조건2) "James, William"  *
'*       RightH(sName, 15) '(조건1) "James, William", (조건2) "James, William"  *
'*     > RightH("가나", 3) '(조건1) "나"            , (조건2) "나"              *
'*                                                                              *
'*     $ 이 예제는, RAT()함수를 사용한 경우이다:                                *
'*                                                                              *
'*       sName = "James, William"                                               *
'*       RightH(sName, LenH(sName) - Rat(",", sName) - 1) '", William"          *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     RightH()는 유니코드를 2바이트로 계산한다.                                *
'*     -----------------------------------------                                *
'*                                                                              *
'********************************************************************************
 Public Function RightH(ByVal sStr As String, ByVal nCount As Long) As String

    '매개변수
     nCount = MinH(MaxH(nCount, 0), 65536) '문자열의 길이를 정의

    '추출
     If sStr = "" Then
        RightH = ""
     Else
        RightH = StrConv(RightB(StrConv(sStr, vbFromUnicode), nCount), vbUnicode)
     End If

 End Function


'********************************************************************************
'* MidH()                                                                       *
'* 문자열에서 지정된 위치의 문자부터 부분문자열을 추출한다.                     *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     MidH(<문자열>, <시작위치>, [<개수>]) --> 부분문자열                      *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <문자열>은 부분문자열을 추출할 전체문자열.                               *
'*                                                                              *
'*     <시작위치>는 부분문자열을 추출하기 시작할 전체문자열내의 위치. 움수인 경 *
'*      우에는, 오른쪽부터 시작위치를 계산한다.                                 *
'*      범위는 -65536 에서 65536 까지이다.                                      *
'*                                                                              *
'*     <개수>는 추출할 문자의 개수. 범위는 0 에서 65536 까지이다.               *
'*      생략되면, <시작위치>부터 <문자열>의 마지막까지를 추출한다.              *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     MidH()는 <문자열>내의 지정된 위치부터, <개수>만큼의 부분 문자열을 오른쪽 *
'*     또는 왼쪽으로 이동하면서 추출한다.                                       *
'*     부분문자열을 추출하는 조건은 아래와 같다.                                *
'*                                                                              *
'*     [참고] L = 전체문자열길이 - 시작위치 + 1 ; L = 최대추출가능 문자수       *
'*            M = ABS(시작위치)                 ; M = 최대추출가능 문자수       *
'*                                                                              *
'*     $ 조건 1. 비주얼베이직에서 제공하는 조건으로 처리했을 때:                *
'*                                                                              *
'*       +-------------------+------------+------------+---------------------+  *
'*       | <문자열>          | <시작위치> | <개수>     | 결과                |  *
'*       +-------------------+------------+------------+---------------------+  *
'*       | 널문자|유효문자열 | < 0        | < 0        | 오류                |  *
'*       | 널문자|유효문자열 | < 0        | = 0        | 오류                |  *
'*       | 널문자|유효문자열 | < 0      ! | <  M       | 오류                |  *
'*       | 널문자|유효문자열 | < 0      ! | >= M       | 오류                |  *
'*       +-------+-----------+------------+------------+---------------------+  *
'*       | 널문자|유효문자열 | = 0        | < 0        | 오류                |  *
'*       | 널문자|유효문자열 | = 0        | = 0        | 오류                |  *
'*       | 널문자|유효문자열 | = 0        | > 0        | 오류                |  *
'*       | 널문자|유효문자열 | = 0        | > 전체길이 | 오류                |  *
'*       +-------------------+------------+------------+---------------------+  *
'*       | 널문자|유효문자열 | > 0        | < 0        | 오류                |  *
'*       | 널문자|유효문자열 | > 0        | = 0        | 널문자              |  *
'*       | 널문자|유효문자열 | > 0      ! | <  L       | 널문자|부분문자열   |  *
'*       | 널문자|유효문자열 | > 0      ! | >= L       | 널문자|시작위치이후 |  *
'*       +-------------------+------------+------------+---------------------+  *
'*       | 널문자|유효문자열 | > 전체길이 | < 0        | 오류                |  *
'*       | 널문자|유효문자열 | > 전체길이 | = 0        | 널문자              |  *
'*       | 널문자|유효문자열 | > 전체길이 | > 0        | 널문자              |  *
'*       | 널문자|유효문자열 | > 전체길이 | > 전체길이 | 널문자              |  *
'*       +-------------------+------------+------------+---------------------+  *
'*                                                                              *
'*     $ 조건 2. 수정된 조건으로 처리했을 때:                                   *
'*                                                                              *
'*       +-------------------+------------+------------+---------------------+  *
'*       | <문자열>          | <시작위치> | <개수>     | 결과                |  *
'*       +-------------------+------------+------------+---------------------+  *
'*       | 널문자|유효문자열 | < 0        | < 0        | 널문자              |  *
'*       | 널문자|유효문자열 | < 0        | = 0        | 널문자              |  *
'*       | 널문자|유효문자열 | < 0      ! | <  M       | 널문자|부분문자열   |  *
'*       | 널문자|유효문자열 | < 0      ! | >= M       | 널문자|시작위치이후 |  *
'*       +-------------------+------------+------------+---------------------+  *
'*       | 널문자|유효문자열 | = 0        | < 0        | 널문자              |  *
'*       | 널문자|유효문자열 | = 0        | = 0        | 널문자              |  *
'*       | 널문자|유효문자열 | = 0        | > 0        | 널문자              |  *
'*       | 널문자|유효문자열 | = 0        | > 전체길이 | 널문자              |  *
'*       +-------------------+------------+------------+---------------------+  *
'*       | 널문자|유효문자열 | > 0        | < 0        | 널문자              |  *
'*       | 널문자|유효문자열 | > 0        | = 0        | 널문자              |  *
'*       | 널문자|유효문자열 | > 0      ! | <  L       | 널문자|부분문자열   |  *
'*       | 널문자|유효문자열 | > 0      ! | >= L       | 널문자|시작위치이후 |  *
'*       +-------------------+------------+------------+---------------------+  *
'*       | 널문자|유효문자열 | > 전체길이 | < 0        | 널문자              |  *
'*       | 널문자|유효문자열 | > 전체길이 | = 0        | 널문자              |  *
'*       | 널문자|유효문자열 | > 전체길이 | > 0        | 널문자              |  *
'*       | 널문자|유효문자열 | > 전체길이 | > 전체길이 | 널문자              |  *
'*       +-------------------+------------+------------+---------------------+  *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     MidH()는 명시된 문자열의 부분문자열을 돌려주는 문자함수이다. MidH()는 지 *
'*     정된 위치에 있는 문자부터, 지정된 개수만큼의 부분문자열을 돌려준다.      *
'*                                                                              *
'*     MidH()는, 음수로 시작하는 <시작위치>를 허용한다. 즉, <시작위치>값이 음수 *
'*     로 시작하면, 오른쪽에서부터 계산된 위치를 시작위치로 정한다. 정해진 위치 *
'*     에서부터 <개수>만큼의 부분문자열을 추출하게 된다.                        *
'*                                                                              *
'*     RightH(), LeftH() 그리고 MidH()는 종종 AT()와 RAT()함수와 같이 사용된다. *
'*                                                                              *
'*     MidH()는 2바이트문자가 잘리는 경우, 해당 문자를 포함시키지 않는다. 즉,   *
'*     <개수>만큼의 문자열을 추출할 때, 마지막부분이 2바이트문자이면서 절반으로 *
'*     나뉘어지게 되면, 마지막부분의 2바이트문자를 제외시킨 문자열을 돌려준다.  *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     $ 다음 예제들은, 여러가지 매개변수에 따른 결과를 보여준다:               *
'*                                                                              *
'*       sName = "James, William"                                               *
'*       MidH("",    -1,-1) '(조건1) error           , (조건2) ""               *
'*       MidH("",    -1, 0) '(조건1) error           , (조건2) ""               *
'*       MidH("",    -1, 1) '(조건1) error           , (조건2) ""               *
'*       MidH("",    -1, 2) '(조건1) error           , (조건2) ""               *
'*       MidH("",     0,-1) '(조건1) error           , (조건2) ""               *
'*       MidH("",     0, 0) '(조건1) error           , (조건2) ""               *
'*       MidH("",     0, 1) '(조건1) error           , (조건2) ""               *
'*       MidH("",     0, 2) '(조건1) error           , (조건2) ""               *
'*       MidH("",     1,-1) '(조건1) error           , (조건2) ""               *
'*       MidH("",     1, 0) '(조건1) ""              , (조건2) ""               *
'*       MidH("",     1, 1) '(조건1) ""              , (조건2) ""               *
'*       MidH("",     1,15) '(조건1) ""              , (조건2) ""               *
'*       MidH("",    15,-1) '(조건1) error           , (조건2) ""               *
'*       MidH("",    15, 0) '(조건1) ""              , (조건2) ""               *
'*       MidH("",    15, 1) '(조건1) ""              , (조건2) ""               *
'*       MidH("",    15, 2) '(조건1) ""              , (조건2) ""               *
'*                                                                              *
'*       MidH(sName, -1,-1) '(조건1) error           , (조건2) ""               *
'*       MidH(sName, -1, 0) '(조건1) error           , (조건2) ""               *
'*       MidH(sName, -1, 1) '(조건1) error           , (조건2) "m"              *
'*       MidH(sName,-10.20) '(조건1) error           , (조건2) "s, William"     *
'*       MidH(sName,-15, 2) '(조건1) error           , (조건2) ""               *
'*       MidH(sName,  0,-1) '(조건1) error           , (조건2) ""               *
'*       MidH(sName,  0, 0) '(조건1) error           , (조건2) ""               *
'*       MidH(sName,  0, 1) '(조건1) error           , (조건2) ""               *
'*       MidH(sName,  0, 2) '(조건1) error           , (조건2) ""               *
'*       MidH(sName,  1,-1) '(조건1) error           , (조건2) ""               *
'*       MidH(sName,  1, 0) '(조건1) ""              , (조건2) ""               *
'*       MidH(sName,  1, 1) '(조건1) "J"             , (조건2) "J"              *
'*       MidH(sName,  1,15) '(조건1) "James, William", (조건2) "James, William" *
'*       MidH(sName, 15,-1) '(조건1) error           , (조건2) ""               *
'*       MidH(sName, 15, 0) '(조건1) ""              , (조건2) ""               *
'*       MidH(sName, 15, 1) '(조건1) ""              , (조건2) ""               *
'*       MidH(sName, 15, 2) '(조건1) ""              , (조건2) ""               *
'*     > MidH("가나", 2, 3) '(조건1) ""              , (조건2) ""               *
'*     > MidH("가나", 3, 3) '(조건1) "나"            , (조건2) "나"             *
'*                                                                              *
'*     $ 이 예제는, RAT()함수를 사용한 경우이다:                                *
'*                                                                              *
'*       sName = "James, William"                                               *
'*       MidH(sName, LenH(sName) - Rat(",", sName) - 1) '", William"            *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     RightH()는 유니코드를 2바이트로 계산한다.                                *
'*     *****************************************                                *
'*                                                                              *
'********************************************************************************
 Public Function MidH(ByVal sStr As String, ByVal nStart As Long, Optional ByVal nCount As Long = 0) As String

    '매개변수
     nStart = MinH(MaxH(nStart, -65536), 65536)                '문자열의 시작을 정의
     nStart = IIf(nStart < 0, LenH(sStr) + nStart + 1, nStart) '문자열의 역순을 정의
     nStart = MaxH(nStart, 0)                                  '문자열의 시작을 재정의
     nCount = IIf(nCount = 0, LenH(sStr) - nStart + 1, nCount) '문자열의 길이를 정의

    '추출
     If sStr = "" Then
        MidH = ""
     Else
        MidH = StrConv(MidB(StrConv(sStr, vbFromUnicode), nStart, nCount), vbUnicode)
     End If

 End Function


'********************************************************************************
'* Pad-()                                                                       *
'* 채울문자로 여백을 메운다.                                                    *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     PadR(<문자열>, <길이>, [<채울문자>]) --> 채워진문자열                    *
'*     PadL(<문자열>, <길이>, [<채울문자>]) --> 채워진문자열                    *
'*     PadC(<문자열>, <길이>, [<채울문자>]) --> 채워진문자열                    *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <문자열>은 채울 문자로 메울 문자열.                                      *
'*                                                                              *
'*     <길이>는 돌려받을 문자열의 길이.                                         *
'*      범위는 0 에서 65536 까지이다.                                           *
'*                                                                              *
'*     <채울문자>는 여백을 채울문자. 생략되면, 공백문자를 사용한다.             *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     PadR(), PadL() 그리고 PadC()는 <문자열>을 <길이>만큼의 문자열로 바꾼다.  *
'*     <문자열>의 길이보다 <길이>가 큰 경우에는, 여백을 <채울문자>로 메워서 돌  *
'*     려준다.                                                                  *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     PadR(), PadL() 그리고 PadC()는, 지정된 길이만큼의 새로운 문자열을 생성하 *
'*     면서 <채울문자>로 여백을 메워서 돌려주는 문자함수이다. PadR()함수는 오른 *
'*     쪽 여백에 <채울문자>를 채운다. PadL()함수는 왼쪽 여백에 채운다. PadC()함 *
'*     수는 <채울문자>를 왼쪽과 오른쪽 여백에 채운다. 문자열의 길이가, <길이>보 *
'*     다 크면, 생성된 문자열은 <길이>만큼의 크기로 잘려진다.                   *
'*                                                                              *
'*     Pad-()함수는 -Trim()함수와는 반대 역할을 한다. -Trim()함수는 문자열의 왼 *
'*     쪽, 오른쪽 혹은 좌우공백문자를 모두 자른다.                              *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     $ 다음 예제들은, 여러가지 매개변수에 따른 결과를 보여준다:               *
'*                                                                              *
'*       sName = "This is a Sample."                                            *
'*       sAmount = "1234567890"                                                 *
'*       PadR(sName, 30)                '"This is a Sample.             "       *
'*       PadR(sName, 10)                '"This is a "                           *
'*       PadL(sAmount, 20, "*")         '"**********1234567890"                 *
'*       PadL(sAmount, 5)               '"67890"                                *
'*       PadC(sName, 30, "*")           '"******This is a Sample.*******"       *
'*       PadC(sName, 10)                '"This is a "                           *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     Pad()함수는 PadR()함수와 동일하다.                                       *
'*                                                                              *
'********************************************************************************
 Public Function PAD(ByVal sStr As String, ByVal nLength As Long, _
                     Optional ByVal sFillChar As String = "") As String

     PAD = PADR(sStr, nLength, sFillChar)

 End Function


'********************************************************************************
'* PadR()                                                                       *
'* 오른쪽공백문자를 채울문자로 메운다.                                          *
'********************************************************************************
 Public Function PADR(ByVal sStr As String, ByVal nLength As Long, _
                      Optional ByVal sFillChar As String = "") As String

    '선언
     Dim nWidth As Long

    '정의
     sStr = Trim(sStr)
     nWidth = LenH(sStr)
     sFillChar = IIf(sFillChar = "", " ", LeftH(sFillChar, 1))

    '돌림
     If nWidth >= nLength Then
        PADR = LeftH(sStr, nLength)
     Else
        PADR = sStr + String(nLength - nWidth, sFillChar)
     End If

 End Function


'********************************************************************************
'* PadL()                                                                       *
'* 왼쪽공백문자를 채울문자로 메운다.                                            *
'********************************************************************************
 Public Function PADL(ByVal sStr As String, ByVal nLength As Long, _
                      Optional ByVal sFillChar As String = "") As String

    '선언
     Dim nWidth As Long

    '정의
     sStr = Trim(sStr)
     nWidth = LenH(sStr)
     sFillChar = IIf(sFillChar = "", " ", LeftH(sFillChar, 1))

    '돌림
     If nWidth >= nLength Then
        PADL = RightH(sStr, nLength)
     Else
        PADL = String(nLength - nWidth, sFillChar) + sStr
     End If

 End Function


'********************************************************************************
'* PadC()                                                                       *
'* 좌우공백문자를 채울문자로 메운다.                                            *
'********************************************************************************
 Public Function PADC(ByVal sStr As String, ByVal nLength As Long, _
                      Optional ByVal sFillChar As String = "") As String

    '선언
     Dim nWidth  As Long
     Dim nSpace  As Long
     Dim nRemain As Long

    '정의
     sStr = Trim(sStr)
     nWidth = LenH(sStr)
     sFillChar = IIf(sFillChar = "", " ", LeftH(sFillChar, 1))

    '돌림
     If nWidth >= nLength Then
        PADC = LeftH(sStr, nLength)
     Else
        nSpace = Int((nLength - nWidth) / 2)
        nRemain = nLength - (nWidth + nSpace)
        PADC = String(nSpace, sFillChar) + sStr + String(nRemain, sFillChar)
     End If

 End Function

'********************************************************************************
'* STUFF()                                                                      *
'* 문자열내의 특정 위치의 문자열을 지정한 문자열로 치환한다.                    *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     STUFF(<문자열>, <시작위치>, <치환길이>, [<치환할 문자열>]) --> 문자열    *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <문자열>은 치환의 대상이 되는 전체 문자열.                               *
'*                                                                              *
'*     <시작위치>는 치환을 시작할 문자열내의 열의 위치.                         *
'*      범위는 1 에서 65536 까지이다.                                           *
'*                                                                              *
'*     <치환길이>는 <시작위치>부터 <치환할 문자열>로 바꿀 문자열의 길이.        *
'*      범위는 0 에서 65536 까지이다.                                           *
'*                                                                              *
'*     <치환할 문자열>은 <시작위치>부터 <치환길이>만큼의 공간을 메꿀 문자열.    *
'*      생략되면, 공백문자를 사용한다.                                          *
'* Returns                                                                      *
'*                                                                              *
'*     <문자열>에서 <시작위치>부터 <치환길이>만큼의 문자열을 삭제한 후에, <치환 *
'*     할 문자열>로 메꾸어서 돌려준다.                                          *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     STUFF()는, 문자열 내부의 특정위치로 부터, 특정 길이만큼의 공간에 있는 문 *
'*     자열을 삭제한 뒤에, 삭제된 공간에 지정된 문자열로 채운다.                *
'*     치환할 문자열이 공백문자라면, 해당 공간을 삭제하는 효과가 있다.          *
'*     시작위치가 전체 문자열의 길이보다 크다면, 전체 문자열의 끝에 치환할 문자 *
'*     열을 삽입하게 된다.                                                      *
'*     치환길이가 ZERO라면, 치환할 문자열을 시작위치에서부터 삽입하는 효과가 있 *
'*     다. 치환길이가 가지는 의미는 삭제이다. 그러므로, ZERO 값은 삭제를 발생시 *
'*     키지 않는다. 결국 삽입만이 이루어진다.
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     $ 다음 예제들은, 여러가지 매개변수에 따른 결과를 보여준다:               *
'*                                                                              *
'*       sString = "This is a Sample."                                          *
'*       sPartitial = "12345"                                                   *
'*       STUFF(sString, -1, 10, sPartitial)   '"12345Sample."                   *
'*       STUFF(sString, 0, 10, sPartitial)    '"12345Sample."                   *
'*       STUFF(sString, 1, 10, sPartitial)    '"12345Sample."                   *
'*       STUFF(sString, 9, 5, sPartitial)     '"This is 12345ple."              *
'*       STUFF(sString, 9, 2)                 '"This is Sample."                *
'*       STUFF(sString, 18, 10, sPartitial)   '"This is a Sample.12345"         *
'*                                                                              *
'********************************************************************************
 Public Function STUFF(ByVal sSrc As String, ByVal nStart As Long, _
                       ByVal nLen As Long, Optional ByVal sObj As String = "") As String

    '선언
     Dim nStringLen As Long
     Dim nLeftLen   As Long
     Dim nRightLen  As Long
     Dim sLeftStr   As String
     Dim sRightStr  As String

    '매개변수
     nStringLen = LenH(sSrc)
     nStart = MinH(MaxH(nStart, 1), nStringLen + 1)       'Min=1, Max=Len(String)+1
     nLen = MinH(MaxH(nLen, 0), nStringLen - nStart + 1)  'Min=0, Max=Len(String)

    '변환
     nLeftLen = MaxH(0, nStart - 1)
     nRightLen = MaxH(0, nStringLen - (nStart + nLen - 1))
     sLeftStr = LeftH(sSrc, nLeftLen)
     sRightStr = RightH(sSrc, nRightLen)

    '돌림
     STUFF = sLeftStr + sObj + sRightStr

 End Function


'********************************************************************************
'* STRZERO()                                                                    *
'* 수치의 왼쪽을 ZERO문자로 채운다 .                                            *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     STRZERO(<값>, <수치화문자열의 길이>) --> 변환된 문자열                   *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <값>은 수치로 구성된 문자열이거나 수치.                                  *
'*                                                                              *
'*     <수치화문자열의 길이>는 변환될 문자열의 길이.                            *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     STRZERO()는 문자나 수치값을 지정한 길이만큼의 문자열로 만들어 돌려준다.  *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     STRZERO()는 수치로 구성된 문자열이나 수치를 문자열로 변환한다. 이 과정에 *
'*     서 변환될 문자열의 길이만큼으로 조정하게 된다. 이 때, 변환된 문자열의 길 *
'*     이가 지정된 길이보다 작은 경우, 왼쪽에 공백이 생긴다. 이 공백을 ZERO문자 *
'*     로 채운다. 변환된 문자열의 길이가 지정된 길이보다 클 경우, 변환된 문자열 *
'*     전체를 돌려준다.                                                         *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     $ 다음 예제들은, 여러가지 매개변수에 따른 결과를 보여준다.               *
'*                                                                              *
'*       STRZERO(10, 10)               '"0000000010"                            *
'*       STRZERO(1000000, 5)           '"1000000"                               *
'*       MinH(0, 2)                    '"00"                                    *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     이 함수는, 비주얼베이직에서 다음의 사용법과 동일한 결과를 돌려준다.      *
'*                                                                              *
'*       STRZERO(10, 10)     = FORMAT(10, "0000000000") = "0000000010"          *
'*       STRZERO(1000000, 5) = FORMAT(1000000, "00000") = "1000000"             *
'*       STRZERO(0, 2)       = FORMAT(0, "00")          = "00"                  *
'*                                                                              *
'********************************************************************************
 Public Function STRZEROS(ByVal nVar As Double, ByVal nLen As Long) As String


    '선언
     Dim sStr As String

    '변환
     sStr = Trim(Str(nVar))
     nLen = MaxH(nLen, LenH(sStr))

    '돌림
     STRZEROS = String(nLen - LenH(sStr), "0") + sStr

 End Function


'********************************************************************************
'* DTOS()                                                                       *
'* 날짜형 값을 년월일형식의 문자열로 변환한다.                                  *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     DTOS(<값>>) --> 변환된 문자열                                            *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <값>은 날짜형 값.                                                        *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     DTOS()는 날자형 값을 년 4자리, 월 2자리, 일 2자리로 맞추어 돌려준다.     *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     DTOS()는 날짜형 인수값을 전달받아, 년 4자리, 월 2자리, 일 2자리의 문자열 *
'*     형식으로 변환하여 돌려준다.                                              *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*       ? DATE                        '98-06-11                                *
'*       DTOS(DATE)                    '"19980611"                              *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     이 함수는, 비주얼베이직에서 다음의 사용법과 동일한 결과를 돌려준다.      *
'*                                                                              *
'*       ?DATE                                  '98-06-11                       *
'*       DTOS(DATE) = FORMAT(DATE, "yyyymmdd) = "19980611"                      *
'*                                                                              *
'********************************************************************************
 Public Function DTOS(ByVal dDate As Date) As String

     DTOS = STRZERO(Year(dDate), 4) + STRZERO(Month(dDate), 2) + STRZERO(Day(dDate), 2)

 End Function


'********************************************************************************
'* UPPER() ... 문자열내의 알파벳을 모두 대문자로 바꾼다.                        *
'* LOWER() ... 문자열내의 알파벳을 모두 소문자로 바꾼다.                        *
'* REVER() ... 문자열내의 알파벳을 대소문자로 각각 치환한다.                    *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     UPPER(<값>) --> 변환된 문자열                                            *
'*     LOWER(<값>) --> 변환된 문자열                                            *
'*     REVER(<값>) --> 변환된 문자열                                            *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <값>은 문자열.                                                           *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     UPPER()은 문자열내에 존재하는 알파벳을 모두 대문자로 바꾸어 돌려준다.    *
'*     LOWER()은 문자열내에 존재하는 알파벳을 모두 소문자로 바꾸어 돌려준다.    *
'*     REVER()은 문자열내에 존재하는 알파벳을 대소문자로 치환해서 돌려준다.     *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     UPPER()는 문자열내에 존재하는 알파벳 소문자만을, 대문자로 바꾸어 준다.   *
'*     LOWER()는 문자열내에 존재하는 알파벳 대문자만을, 소문자로 바꾸어 준다.   *
'*     REVER()은 문자열내에 존재하는 알파벳 대문자와 소문자를 각각 교환한다. 즉 *
'*     대문자는 소문자로, 소문자는 대문자로 서로 바꾸어 준다.                   *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*       S = "AbcdEfGhijKLMnopQrStUvWxyZ123가나다"                              *
'*                                                                              *
'*       UPPER(S)    '"ABCDEFGHIJKLMNOPQRSTUVWXYZ123가나다"                     *
'*       LOWER(S)    '"abcdefghijklmnopqrstuvwxyz123가나다"                     *
'*       REVER(S)    '"aBCDeFgHIJklmNOPqRsTuVwXYz123가나다"                     *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     이 함수들은, 비주얼베이직에서 다음의 사용법과 동일한 결과를 돌려준다.    *
'*                                                                              *
'*       S = "Abc123가나다"                                                     *
'*                                                                              *
'*       UPPER(S) = UCase(S) = "ABC123가나다"                                   *
'*       LOWER(S) = LCase(S) = "abc123가나다"                                   *
'*                                                                              *
'********************************************************************************
'********************************************************************************
'* UPPER()                                                                      *
'* 문자열내의 소문자를 대문자로 바꾼다.                                         *
'********************************************************************************
 Public Function UPPER(ByVal s As String) As String

     UPPER = UCase(s)

 End Function


'********************************************************************************
'* LOWER()                                                                      *
'* 문자열내의 대문자를 소문자로 바꾼다.                                         *
'********************************************************************************
 Public Function LOWER(ByVal s As String) As String

     LOWER = LCase(s)

 End Function


'********************************************************************************
'* REVER()                                                                      *
'* 문자열내의 대문자와 소문자를 각각 소문자와 대문자로 바꾼다.                  *
'********************************************************************************
 Public Function REVER(ByVal s As String) As String

    '선언
     Dim n, i As Long
     Dim c, r As String

    '정의
     n = Len(s)
     r = ""

    '변환
     If n > 0 Then

        For i = 1 To n
            c = Mid(s, i, 1)
            If Asc(c) >= Asc("A") And Asc(c) <= Asc("Z") Then
               c = Chr(Asc(c) + 32)
            ElseIf Asc(c) >= Asc("a") And Asc(c) <= Asc("z") Then
               c = Chr(Asc(c) - 32)
            End If
            r = r + c
        Next i

     End If

    '돌림
     REVER = r

 End Function



'********************************************************************************
'*  AT() ... 문자열내에서 특정 문자열의 위치를 앞에서부터 찾는다.               *
'* RAT() ... 문자열내에서 특정 문자열의 위치를 뒤에서부터 찾는다.               *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*      AT(<찾을 문자열>, <문자열>) --> 찾을 문자열의 앞에서부터의 위치         *
'*     RAT(<찾을 문자열>, <문자열>) --> 찾을 문자열의 앞에서부터의 위치         *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <찾을 문자열>은 검색을 위한 부분 문자열.                                 *
'*                                                                              *
'*     <문자열>은 검색의 대상이 되는 전체 문자열.                               *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*      AT()은 문자열내에 찾을 문자열이 시작되는 위치값을 돌려준다.             *
'*     RAT()은 문자열내에 찾을 문자열이 시작되는 위치값을 돌려준다.             *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     AT()나 RAT()는 전체 문자열에서 검색의 대상이 되는 부분문자열이 존재할 때 *
'*     에, 부분 문자열이 시작되는 위치를, 전체 문자열의 앞에서 부터의 위치를 돌 *
'*     려준다. 두 함수의 차이점은, AT()는 앞에서 부터 검색을 하고, RAT()는 뒤에 *
'*     서 부터 검색을 한다는 것이다.                                            *
'*                                                                              *
'*     비주얼베이직에서 INSTR()함수가 있으나, 유니코드문자를 1BYTE로 처리하므로 *
'*     2BYTES단위의 검색위치를 찾을 수 없다.                                    *
'*                                                                              *
'*     AT()나 RAT()를 사용할 경우, 한 가지 문제점이 있다. 즉, 2BYTES 그래픽이나 *
'*     한글이 연속적으로 사용되어질 경우에는, 연속되는 앞문자의 뒷쪽 1BYTES문자 *
'*     와 이어지는 뒷문자의 앞쪽 1BYTES문자가 충돌하면서,  임의의 문자를 만들어 *
'*     내는 것이다. 이 임의의 문자가 입력가능한 문자SET인 경우에, 검색할 문자값 *
'*     과 같을 경우가 있다. 그러나 실제 전체 문자열에는 포함되지 않는 문자이다. *
'*     이런 경우, 실제는 존재하지 않으나, 존재하는 문자로 처리되어, 해당 위치의 *
'*     값이 돌려지게 된다.                                                      *
'*     이런 경우를 막기 위해, 비주얼베이직의 INSTR()함수로 검사된 결과값이 ZERO *
'*     보다 큰 경우에만, 위치를 검색한다.                                       *
'*     아래의 예제에서 확인하기 바란다.                                         *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*       S = "각ㄴㄷㄹㅁabdㄱCDㄻㅂㄱ낙ㄴ"                                      *
'*       T = "ㄴ"                                                               *
'*       U = "깽"                                                               *
'*                                                                              *
'*       INSTR(1, T, S)        '2                                               *
'*       AT(T, S) only         '3                                               *
'*       AT(T, S) and INSTR()  '3                                               *
'*                                                                              *
'*       INSTR(1, U, S)        '0                                               *
'*       AT(U, S) only         '21                                              *
'*       AT(U, S) and INSTR()  '0                                               *
'*                                                                              *
'* Notice                                                                       *
'*                                                                              *
'*     주의!                                                                    *
'*                                                                              *
'*     비주얼베이직의 INSTRB()함수는, 찾은 위치를 바이트형식의 값으로 돌려준다. *
'*     그래서, 유니코드문자열을 검색할 때 사용할 수 있다. 그러나, 어떤 경우에는 *
'*     정확한 값을 돌려주지 못한다.                                             *
'*     다음의 예제를 참조한다.                                                  *
'*                                                                              *
'*       S = "각ㄴㄷㄹㅁabdㄱCDㄻㅂㄱ낙ㄴ"                                      *
'*       T = "ㄱ"                                                               *
'*                                                                              *
'*       INSTRB(1, S, T)       '17                                              *
'*       AT(T, S)              '14                                              *
'*                                                                              *
'********************************************************************************
'********************************************************************************
'* AT()                                                                         *
'* 문자열내에서 특정 문자열의 위치를 앞에서부터 찾는다.                         *
'********************************************************************************
 Public Function AT(ByVal sSearch As String, ByVal sString As String) As Long

    '선언
     Dim i, n, k, X As Long
     Dim c          As String

    '정의
     n = MaxH(LenH(sString), 0)
     k = MaxH(LenH(sSearch), 0)
     X = 0

    '검색
     If n > 0 And InStr(1, sString, sSearch) > 0 Then

        For i = 1 To n
            c = MidH(sString, i, k)
            If c = sSearch Then
               X = i
               Exit For
            End If

        Next i

     End If

    '돌림
     AT = X

 End Function


'********************************************************************************
'* RAT()                                                                        *
'* 문자열내에서 특정 문자열의 위치를 뒤에서부터 찾는다.                         *
'********************************************************************************
 Public Function RAT(ByVal sSearch As String, ByVal sString As String) As Long

    '선언
     Dim i, n, k, X As Long
     Dim c          As String

    '정의
     n = MaxH(LenH(sString), 0)
     k = MaxH(LenH(sSearch), 0)
     X = 0

    '검색
     If n > 0 And InStr(1, sString, sSearch) > 0 Then

        For i = n To 1 Step -1
            c = MidH(sString, i, k)
            If c = sSearch Then
               X = i
               Exit For
            End If

        Next i

     End If

    '돌림
     RAT = X

 End Function


'********************************************************************************
'* DRIVE()                                                                      *
'* 특정 드라이브에 매체가 사용가능한 상태인가를 돌려준다.                       *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     DRIVE(<드라이브>) --> 사용가능한 상태인가를 정수값으로 돌려준다.         *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <드라이브>는 검사할 드라이브명을 나타내는 문자열.                        *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     DRIVE()는 드라이브의 상태에 따라 다음과 같은 정수를 돌려준다.            *
'*                                                                              *
'*     드라이브    기록매체     돌림값                                          *
'*     --------    ---------    ------                                          *
'*     있음        있음          0 (intResDriveReady)                           *
'*     있음        없음|오류     1 (intResDiskNotReady)                         *
'*     없음        -            -1 (intResDriveNotReady)                        *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     DRIVE()는 지정한 드라이브가 존재하는 지, 혹은 존재하는 드라이브를 읽거나 *
'*     기록할 수 있는 지를 검사한다.  드라이브의 상태에 따라 각각에 해당하는 정 *
'*     수값을 돌려준다.                                                         *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*       DRIVE("A")            ' 1 (DISK NOT READY)                             *
'*       DRIVE("B")            '-1 (DRIVE NOT READY)                            *
'*       DRIVE("C")            ' 0                                              *
'*       DRIVE("F")            ' 1 (UNWRITABLE CD-ROM DRIVE)                    *
'*                                                                              *
'********************************************************************************
 Public Function DRIVE(Optional ByVal sDrive As String = "") As Integer

    '선언
     Dim sPath As String

    '매개변수
     sDrive = UCase(Left(IIf(sDrive = "", CurDir(), sDrive), 1)) + ":\"

    '검사
     On Error Resume Next
     sPath = CurDir(sDrive)
     On Error GoTo 0
   
    '돌림
     DRIVE = IIf(sPath = "", intResDriveNotReady, intResDriveReady)
     If DRIVE = intResDriveReady Then
        On Error Resume Next
        sPath = ""
        sPath = Dir(sDrive, vbDirectory)
        On Error GoTo 0
        DRIVE = IIf(sPath = "", intResDiskNotReady, intResDriveReady)
     End If
     
 End Function


'********************************************************************************
'* AFILL()                                                                      *
'* 배열을 특정한 값으로 채운다.                                                 *
'********************************************************************************
'* Syntax                                                                       *
'*                                                                              *
'*     AFILL(<배열>, <채울값> [, <시작위치> [, <마침위치>]]) --> 채워진 배열    *
'*                                                                              *
'* Arguments                                                                    *
'*                                                                              *
'*     <배열>은 채울값으로 채울 1차원배열.                                      *
'*                                                                              *
'*     <채울값>은 배열의 각 요소에 넣을 값.                                     *
'*                                                                              *
'*     <시작위치>은 채우기 시작할 배열의 시작 요소번호.                         *
'*      생략되면, 배열의 첫 요소번호를 사용한다.                                *
'*                                                                              *
'*     <마침위치>은 채우기를 마칠 배열의 마침 요소번호.                         *
'*      생략되면, 배열의 끝 요소번호를 사용한다.                                *
'*                                                                              *
'* Returns                                                                      *
'*                                                                              *
'*     채울값으로 배열의 각 요소를 채워서 돌려준다.                             *
'*                                                                              *
'* Description                                                                  *
'*                                                                              *
'*     배열은 1차원만을 허용한다. AFILL()은 일종의 배열 초기화 함수이다. 지정된 *
'*     요소범위를 하나의 지정된 값으로 채우기 때문이다.                         *
'*                                                                              *
'* Examples                                                                     *
'*                                                                              *
'*     DIM arr1(1 TO 10)                                                        *
'*                                                                              *
'*     AFILL(arr1, 10)        'arr1(1) = 10, ..., arr1(10) = 10                 *
'*     AFILL(arr1, 5, 5, 7)   'arr1(5) = 5, arr1(6) = 5, arr1(7) = 5            *
'*     AFILL(arr1, 8, 9)      'arr1(9) = 8, arr1(10) = 8                        *
'*                                                                              *
'********************************************************************************
 Public Function AFILL(ByVal aArr As Variant, ByVal vVar As Variant, _
                       Optional ByVal nLower As Long = -2147483647#, _
                       Optional ByVal nUpper As Long = 2147483647) As Variant

    '선언
     Dim i As Long

    '매개변수
     nLower = MinH(MaxH(LBound(aArr), nLower), UBound(aArr))
     nUpper = MinH(MaxH(nLower, nUpper), UBound(aArr))

    '변환
     For i = nLower To nUpper
         aArr(i) = vVar
     Next i

 End Function

'********************************************************************************
'* 돌림 ... 인수가 숫자로만 구성되었졌다면 참값을; 그렇지 않다면, 거짓값을 돌려 *
'*          준다. 빈문자열이라면, 거짓값을 돌려준다.                            *
'* 설명 ... 좌우공백문자는 제거한다.                                            *
'*          인수가 숫자로만 구성되어졌는가를 평가한다.                          *
'*          [숫자] 0,1,2,3,4,5,6,7,8,9                                          *
'********************************************************************************
 Public Function EsNumber(ByVal strValue As String) As Boolean

       '선언
        Dim i As Long '순환변수
        Dim n As Long '문자열길이
        Dim c As String  '문자

       '초기화
        EsNumber = False

       '정의
        strValue = Trim(strValue)
        n = LenH(strValue)

       '유효성
        If n > 0 Then

          '검사
           For i = 1 To n

              '문자추출
               c = MidH(strValue, i, 1)

              '검사
               If Not SUBSET(c, "0123456789") Then
                  Exit Function
               End If

           Next i

          '성공
           EsNumber = True

        End If

 End Function

'********************************************************************************
'* 기능 ... 날짜를 생성한다.                                                    *
'* 인수 ... 문자열. 년(4),월(2),일(2)로 구성된 숫자형문자열.                    *
'* 돌림 ... 인수가 유효하면, 날짜형식문자열을; 그렇지 않다면, 널문자열을 돌려준 *
'*          다.                                                                 *
'* 설명 ... yyyymmdd 형식의 문자열이 유효한 날짜인가를 평가한다.                *
'*          평가된 날짜가 유효하다면, 다음 형식의 문자열을 돌려준다.  그렇지 않 *
'*          다면, 널문자열을 돌려준다.                                          *
'*          [형식] yyyy-mm-dd                                                   *
'*          [숫자] 0,1,2,3,4,5,6,7,8,9                                          *
'********************************************************************************
 Public Function EsDate(ByVal strDate As String, _
                        Optional strDmt As String) As String

       '초기화
        EsDate = ""

       '유효성 (8자리의 숫자형문자열)
        strDate = Trim(strDate)
        If LenH(strDate) = 8 And EsNumber(strDate) Then

          '유효한 날짜인가?
           If IsDate(Mid$(strDate, 7, 2) + "/" + _
                     Mid$(strDate, 5, 2) + "/" + _
                     Mid$(strDate, 1, 4)) Then

             '구분자
              'strDmt = IIf(IsMissing(strDmt), "-", strDmt)
              strDmt = "-"
              strDmt = "0000" + strDmt + "00" + strDmt + "00"

             '돌림
              EsDate = Format(Val(strDate), strDmt)

           End If

        End If

 End Function

'********************************************************************************
'* 기능 ... 주민번호를 생성한다.                                                *
'* 인수 ... 문자열. 생년월일(6),일련번호(7)로 구성된 숫자형문자열.              *
'* 돌림 ... 인수가 유효하면, 주민번호형식문자열을; 그렇지 않다면, 널문자열을 돌 *
'*          려준다.                                                             *
'* 설명 ... yymmddsnnnnnn 형식의 문자열이 유효한 주민번호인가를 평가한다.       *
'*          평가된 주민번호가 유효하다면, 다음 형식의 문자열을 돌려준다. 그렇지 *
'*          않다면, 널문자열을 돌려준다.                                        *
'*          [형식] yymmdd-snnnnnn                                               *
'*          [숫자] 0,1,2,3,4,5,6,7,8,9                                          *
'********************************************************************************
 Public Function EsID(ByVal strID As String) As String

       '초기화
        EsID = ""

       '유효성 (13자리의 숫자형문자열)
        strID = Trim(strID)
        If LenH(strID) = 13 And EsNumber(strID) Then

          '돌림
           EsID = Format(Val(strID), "000000-0000000")

        End If

 End Function

'********************************************************************************
'* 기능 ... 우편번호를 생성한다.                                                *
'* 인수 ... 문자열. 우편번호1(3),우편번호2(3)로 구성된 숫자형문자열.            *
'* 돌림 ... 인수가 유효하면, 우편번호형식문자열을; 그렇지 않다면, 널문자열을 돌 *
'*          려준다.                                                             *
'* 설명 ... mmmnnn 형식의 문자열이 유효한 우편번호인가를 평가한다.              *
'*          평가된 우편번호가 유효하다면, 다음 형식의 문자열을 돌려준다. 그렇지 *
'*          않다면, 널문자열을 돌려준다.                                        *
'*          [형식] mmm-nnn                                                      *
'*          [숫자] 0,1,2,3,4,5,6,7,8,9                                          *
'********************************************************************************
 Public Function EsPost(ByVal strPost As String) As String

       '초기화
        EsPost = ""

       '유효성 (6자리의 숫자형문자열)
        strPost = Trim(strPost)
        If LenH(strPost) = 6 And EsNumber(strPost) Then

          '돌림
           EsPost = Format(Val(strPost), "000-000")

        End If

 End Function

'********************************************************************************
'* 기능 ... 좌측에 ZERO문자를 포함한 숫자형문자열을 생성한다.                   *
'* 인수 ... 문자열 또는 숫자.                                                   *
'* 돌림 ... 문자열이나 숫자를, 좌측공백문자를 ZERO문자로 바꾸어 돌려준다.       *
'* 설명 ... 문자열이면, 숫자로 변환한다. 숫자를 지정된 길이의 문자열로 변환하면 *
'*          서, 변환된 문자열이 지정된 길이보다 작을 경우, 좌측의 공백을 ZERO문 *
'*          자로 바꾸어 돌려준다.                                               *
'*          문자열로 전달된 인수는, 숫자만을 허용한다. 숫자가 아닐 경우, 널문자 *
'*          열을 돌려준다.                                                      *
'*          [형식] FORMAT(값,"000...")                                          *
'*          [숫자] 0,1,2,3,4,5,6,7,8,9                                          *
'********************************************************************************
 'Public Function EsZERO(ByVal varValue As Variant, ByVal intSize As Long) As String
 Public Function STRZERO(ByVal varValue As Variant, ByVal intSize As Long) As String

       '선언
        Dim strValue As String
        Dim intLen   As Long

       '초기화
        STRZERO = ""
        intLen = intSize

       '변환
        If VarType(varValue) = vbString Or _
           VarType(varValue) = vbCurrency Or _
           VarType(varValue) = vbDecimal Or _
           VarType(varValue) = vbDouble Or _
           VarType(varValue) = vbInteger Or _
           VarType(varValue) = vbLong Or _
           VarType(varValue) = vbSingle Then

          '유효성 (숫자형문자열)
           If VarType(varValue) = vbString Then
              If Not EsNumber(varValue) Then
                 Exit Function
              End If
              intLen = MaxH(intLen, Len(Trim(varValue)))
              varValue = Val(varValue)
           End If

          '돌림
           strValue = Format(varValue, String$(intLen, "0"))
           STRZERO = Right$(strValue, intSize)

        End If

 End Function

'********************************************************************************
'* 기능 ... 계좌번호를 생성한다.                                                *
'* 인수 ... 문자열.                                                             *
'* 돌림 ... 인수가 유효하면, 지정형식의 계좌번호문자열을; 그렇지 않다면, 널문자 *
'*          열을 돌려준다.                                                      *
'* 설명 ... 문자열이 숫자이며, 지정된 길이를 가지는지를 평가한다. 평가된 계좌번 *
'*          호가 유효하다면, 지정 형식의 문자열을 돌려준다. 그렇지 않다면, 널문 *
'*          자열을 돌려준다.                                                    *
'*          [숫자] 0,1,2,3,4,5,6,7,8,9                                          *
'********************************************************************************
 Public Function EsAccount(ByVal strAccount As String, _
                           ByVal intLen As Long, _
                           ByVal strFormat As String) As String

       '초기화
        EsAccount = ""

       '유효성 (6자리의 숫자형문자열)
        strAccount = Trim(strAccount)
        If LenH(strAccount) = intLen And EsNumber(strAccount) Then

          '돌림
           EsAccount = Format(Val(strAccount), strFormat)

        End If

 End Function

'********************************************************************************
'********************************************************************************
 Public Function InNumber(ByVal strValue As String) As String

       '선언
        Dim i As Long '순환변수
        Dim n As Long '문자열길이
        Dim c As String  '문자

       '초기화
        InNumber = ""

       '정의
        strValue = Trim(strValue)
        n = LenH(strValue)

       '유효성
        If n > 0 Then

          '검사
           For i = 1 To n

              '문자추출
               c = MidH(strValue, i, 1)

              '검사
               If SUBSET(c, "0123456789") Then
                  InNumber = InNumber + c
               End If

           Next i

        End If

 End Function


 Public Function FormatNumber(ByVal string_cNumber As String, _
                              ByVal string_cDelimiter As String) As String
    '
     Dim long_nLen As Long
     Dim string_cChr As String
     Dim string_cStr As String
     Dim i As Long

     string_cNumber = Trim(string_cNumber)
     long_nLen = LenH(string_cNumber)
     string_cStr = ""
     If long_nLen > 0 Then
        For i = 1 To long_nLen
            string_cChr = MidH(string_cNumber, i, 1)
            If EsNumber(string_cChr) Then
               string_cStr = string_cStr + string_cChr
            End If
        Next i
        string_cStr = Format(Val(string_cStr), "###,###,###,###,###")
     End If

     FormatNumber = string_cStr

 End Function

 Public Function CutAmount(ByVal nAmount As Currency, _
                           ByVal nUnit As Currency) As Currency

     CutAmount = Int(Int(nAmount / nUnit) * nUnit)

 End Function

'********************************************************************************
'* Extract the Original Numeric Character from the Character String to be given *
'********************************************************************************
 Public Function GetNumber(ByVal sString As String) As String

    'Declare the Variables
     Dim i, n As Long
     Dim c    As String

    'Define the Variables
     sString = Trim(LTrim(sString))
     n = Len(sString)
     GetNumber = ""

    'Extract the Value
     If n > 0 Then
        For i = 1 To n
            c = Mid(sString, i, 1)
            If Asc(c) >= Asc("0") And Asc(c) <= Asc("9") Then
               GetNumber = GetNumber + c
            End If
        Next i
     End If

 End Function

'********************************************************************************
'* Get TextFile                                                                 *
'********************************************************************************
 Public Function MemoRead(ByVal sFileName As String) As Variant

    'Declare
     Dim nFileLine   As Integer
     Dim nFreeFile   As Integer
     Dim aLineList() As String
     Dim s           As String

    'Check
     If Dir(sFileName) = "" Then
        ReDim aLineList(0) As String
        aLineList(0) = "00000"
        MemoRead = aLineList
        Exit Function
     End If

    'Open
     nFileLine = 0
     nFreeFile = FreeFile
     Open sFileName For Input As #nFreeFile
     While Not EOF(nFreeFile)
           nFileLine = nFileLine + 1
           Line Input #nFreeFile, s
           DoEvents
     Wend
     Close #nFreeFile

    'Exist ?
     If nFileLine > 0 Then
        ReDim aLineList(nFileLine) As String
        aLineList(0) = Format(nFileLine, "00000")
        nFreeFile = FreeFile
        nFileLine = 0
        Open sFileName For Input As #nFreeFile
        While Not EOF(nFreeFile)
              nFileLine = nFileLine + 1
              DoEvents
              Line Input #nFreeFile, aLineList(nFileLine)
        Wend
        Close #nFreeFile
     Else
        ReDim aLineList(0) As String
        aLineList(0) = "00000"
     End If

    'Return
     MemoRead = aLineList

 End Function



'********************************************************************************
'* 문자열내부의 지정 문자열을 다른 문자열로 바꾼다.                             *
'********************************************************************************
 Public Function STRTRAN(ByVal sStr As String, ByVal sSrc As String, ByVal sObj As String)
    '변수선언
     Dim nWid As Long
     Dim nLen As Long
     Dim nPos As Long
     Dim sNew As String
     Dim sChk As String

    '변수정의
     nWid = Len(sStr)
     nLen = Len(sSrc)
     nPos = 0
     sNew = ""

    '변환
     Do While True
        If nPos > nWid Then
           Exit Do
        End If
        nPos = nPos + 1
        sChk = Mid$(sStr, nPos, nLen)
        sNew = sNew + IIf(sChk = sSrc, sObj, Mid$(sStr, nPos, 1))
     Loop
     STRTRAN = sNew
 End Function


'********************************************************************************
'* 해당문자가 존재하는가?                                                       *
'********************************************************************************
 Public Function SUBSET(ByVal sChr As String, ByVal sStr As String) As Boolean
    '변수선언
     Dim nCount As Long                                                         '문자열위치
     Dim nTotal As Long                                                         '문자열길이

    '변수정의
     nTotal = Len(sStr)

    '문자검색
     For nCount = 1 To nTotal
         If Mid$(sStr, nCount, 1) = sChr Then
            SUBSET = True
            Exit Function
         End If
     Next

    '존재않음
     SUBSET = False
 End Function


'********************************************************************************
'* 24시간제를 12시간제로 (hh:mm:ss -> hh:mm:ss AM/PM)                           *
'********************************************************************************
 Public Function AMPM(ByVal sTime) As String
    '변수선언
     Dim nHH As Integer
     Dim sAP As String
     Dim sHH As String
     Dim sMM As String
     Dim sSS As String

    '변환
     nHH = Val(Mid$(sTime, 1, 2))
     sAP = IIf(nHH > 12, "오후 ", "오전 ")
     sHH = STRZERO(IIf(nHH > 12, nHH - 12, nHH), 2) + "시 "
     sMM = Mid$(sTime, 4, 2) + "분 "
     sSS = Mid$(sTime, 7, 2) + "초"
     AMPM = sAP + sHH + sMM + sSS
 End Function


'********************************************************************************
'* 배열을 지정한 값으로 채운다.                                                 *
'********************************************************************************
 Public Sub AFILLs(Arr As Variant, v As Variant, st As Long, ed As Long)
     Dim i As Long
     For i = st To ed
         Arr(i) = v
     Next i
 End Sub


 Public Function TRANSFORM(ByVal v As Double, ByVal F As String) As String
     Dim w As Long
     Dim s As String

     w = Len(F)
     s = Format(v, F)
     If w > Len(s) Then
        s = Space$(w - Len(s)) + s
     End If
     TRANSFORM = s
 End Function

 Public Function SKIPDATE(ByVal sDate As String, ByVal nInt As Long) As String
        Dim nYY As Long
        Dim nMM As Long
        Dim nDD As Long

        nYY = Val(Mid$(sDate, 1, 4))
        nMM = Val(Mid$(sDate, 5, 2))
        nDD = Val(Mid$(sDate, 7, 2))
        SKIPDATE = DTOS(DateSerial(nYY, nMM, nDD + nInt))
 End Function

 Public Function Percent(ByVal n As Double, ByVal T As Double) As Integer
        Percent = Int((100 * n) / T)
 End Function


 Public Function SVal(ByVal s As String) As String

    Dim i As Long
    Dim n As Long
    Dim r As String
    Dim c As String

    s = Trim(s)
    n = LenH(s)
    r = ""
    For i = 1 To n
        c = Mid$(s, i, 1)
        If SUBSET(c, "+-0123456789") Then
           If SUBSET(c, "+-") Then
              If i = 1 Then
                 r = r + c
              Else
                 SVal = r
                 Exit Function
              End If
           Else
              r = r + c
           End If
        End If
    Next i
    SVal = r

 End Function



 Public Function Time2() As String

        Time2 = Hour(Time) & Minute(Time) & Second(Time)

 End Function



 Public Function IsTrue(ByVal strLogicalValue As Boolean) As Boolean
     IsTrue = (strLogicalValue = misTRUE)
 End Function

 Public Function GetLines(ByVal strFileName As String) As Long 'GetTextFileLines

    '변수선언
     Dim lngLines    As Long
     Dim intFreeFile As Integer
     Dim strLine     As String

    '변수정의
     lngLines = -1
     strFileName = Trim(strFileName)

    '파일이 존재하는가를 검사한다.
     If (Not (strFileName = "")) And (Not (Dir(strFileName) = "")) Then

       '초기화
        intFreeFile = FreeFile
        lngLines = 0

       '파일개방
        Open strFileName For Input As #intFreeFile

       '파일읽기
        Do Until EOF(intFreeFile)

           Line Input #intFreeFile, strLine
           lngLines = lngLines + 1

        Loop

       '파일폐쇄
        Close #intFreeFile

     End If

    '돌림
     GetLines = lngLines

 End Function

 Public Function IsFile(ByVal strFileName As String) As Boolean

     strFileName = Trim(strFileName)
     IsFile = ((strFileName <> "") And (Dir(strFileName) <> ""))

 End Function

'********************************************************************************
'* 글꼴을 설정한다.                                                             *
'********************************************************************************
 Public Sub putFont(mm As Form, arrFont As Variant)

     Dim X As Control
     Dim n As Integer

     n = mm.FontSize
     mm.FontName = arrFont(1)
     mm.FontSize = IIf(arrFont(2) = 0, n, arrFont(2))
     mm.FontBold = arrFont(3)
     mm.FontItalic = arrFont(4)
     mm.FontStrikethru = arrFont(5)
     mm.FontUnderline = arrFont(6)

     For Each X In mm.Controls
         If TypeOf X Is CheckBox Or _
            TypeOf X Is ComboBox Or _
            TypeOf X Is CommandButton Or _
            TypeOf X Is Data Or _
            TypeOf X Is Frame Or _
            TypeOf X Is Label Or _
            TypeOf X Is ListBox Or _
            TypeOf X Is OptionButton Or _
            TypeOf X Is PictureBox Or _
            TypeOf X Is TextBox Then
                   n = X.FontSize
                   X.FontName = arrFont(1)
                   X.FontSize = IIf(arrFont(2) = 0, n, arrFont(2))
                   X.FontBold = arrFont(3)
                   X.FontItalic = arrFont(4)
                   X.FontStrikethru = arrFont(5)
                   X.FontUnderline = arrFont(6)
         'ElseIf TypeOf X Is DBGrid Then
         '          N = X.Font.Size
         '          X.Font.Name = arrFont(1)
         '          X.Font.Size = IIf(arrFont(2) = 0, N, arrFont(2))
         '          X.Font.Bold = arrFont(3)
         '          X.Font.Italic = arrFont(4)
         '          X.Font.Strikethrough = arrFont(5)
         '          X.Font.Underline = arrFont(6)
         '          N = X.HeadFont.Size
         '          X.HeadFont.Name = arrFont(1)
         '          X.HeadFont.Size = IIf(arrFont(2) = 0, N, arrFont(2))
         '          X.HeadFont.Bold = arrFont(3)
         '          X.HeadFont.Italic = arrFont(4)
         '          X.HeadFont.Strikethrough = arrFont(5)
         '          X.HeadFont.Underline = arrFont(6)
         'ElseIf TypeOf X Is RichTextBox Or _
         '       TypeOf X Is ListView Then
         '          N = X.Font.Size
         '          X.Font.Name = arrFont(1)
         '          X.Font.Size = IIf(arrFont(2) = 0, N, arrFont(2))
         '          X.Font.Bold = arrFont(3)
         '          X.Font.Italic = arrFont(4)
         '          X.Font.Strikethrough = arrFont(5)
         '          X.Font.Underline = arrFont(6)
         End If
     Next X

 End Sub


'********************************************************************************
'* 글꼴을 구한다.                                                               *
'********************************************************************************
 Public Function getFont() As Variant

     Dim arrFont(6) As Variant

     arrFont(1) = "굴림체" 'varDefaultFontName
     arrFont(2) = 12 'varDefaultFontSize
     arrFont(3) = False
     arrFont(4) = False
     arrFont(5) = False
     arrFont(6) = False

     getFont = arrFont

 End Function




'********************************************************************************
'* 문자열을 수치로 변환                                                         *
'*                                                                              *
'*   수치로 취급되는 조건:                                                      *
'*                                                                              *
'*     [1] 처음으로 시작하는 +, - 문자                                          *
'*     [2] 세 자리 단위로 끊어지는 , 문자                                       *
'*     [3] . 문자                                                               *
'*     [4] 일반 숫자                                                            *
'********************************************************************************
 Public Function Vals(ByVal sz As String) As Currency

    '변수선언
     Const udWON    As Integer = 92 '\
     Const udPLUS   As Integer = 43 '+
     Const udMINUS  As Integer = 45 '-
     Const udPOINT  As Integer = 46 '.
     Const udCOMMA  As Integer = 44 ',
     Const udPARENTHESISOPEN   As Integer = 40 '(
     Const udPARENTHESISCLOSE  As Integer = 41 ')
     Const udBRACKETOPEN       As Integer = 91 '[
     Const udBRACKETCLOSE      As Integer = 93 ']
     Const udBRACEOPEN         As Integer = 123 '{
     Const udBRACECLOSE        As Integer = 125 '}
     Dim blnPoint   As Boolean '소숫점부호
     Dim intComma   As Integer '콤마단위 숫자개수
     Dim intCount   As Integer '문자열개수
     Dim strValue   As String  '변환된 문자열
     Dim i          As Integer '순환변수
     Dim cc         As String  '문자
     Dim ca         As Integer '아스키값


    '돌림초기화

    '변수초기화
     sz = Trim(sz)
     blnPoint = False
     intComma = 0
     intCount = LenH(sz)
     strValue = ""

    '파싱
     If intCount > 0 Then

       '순환
        For i = 1 To intCount

            cc = MidH(sz, i, 1)
            ca = Asc(cc)
            Select Case ca
                   Case udWON
                        If i > 1 Then '처음 시작이면, OK.
                           If Not ((MidH(sz, 1, 1) = "+" Or MidH(sz, 1, 1) = "-") And (i = 2)) Then
                              strValue = ""
                              Exit For 'Error
                           End If
                        End If

                   Case udPLUS, udMINUS
                        If i = 1 Then '처음 시작이면, OK.
                           strValue = strValue + cc
                        Else
                           strValue = ""
                           Exit For 'Error
                        End If

                   Case udPOINT
                        If Not blnPoint Then
                           'If intComma = 0 Or intComma = 3 Then
                              strValue = strValue + cc
                              blnPoint = True
                           'Else
                           '   strValue = ""
                           '   Exit For 'Error
                           'End If
                        Else
                           strValue = ""
                           Exit For 'Error
                        End If

                   Case udCOMMA
                        'If Not blnPoint Then '소숫점이상에서만 콤마허용
                        '   If intComma = 0 Or intComma = 3 Then
                        '      intComma = 0
                           'Else
                           '   strValue = ""
                           '   Exit For 'Error
                        '   End If
                        'Else
                        '   strValue = ""
                        '   Exit For 'Error
                        'End If

                   Case udPARENTHESISOPEN, udPARENTHESISCLOSE, _
                        udBRACKETOPEN, udBRACKETCLOSE, _
                        udBRACEOPEN, udBRACECLOSE
                       'Ignore

                   Case Asc("0") To Asc("9")
                        strValue = strValue + cc
                        intComma = intComma + 1

            End Select

        Next i

     End If

    '돌림
     Vals = Val(strValue)

 End Function


'********************************************************************************
'* 변수형에 맞는 초기화                                                         *
'********************************************************************************
 Public Function InitializeValue(ByVal sz As Variant) As Variant

         Select Case VarType(sz)
                Case vbEmpty, vbNull
                     InitializeValue = ""
                Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                     InitializeValue = 0
                Case vbDate
                     InitializeValue = Date
                Case vbString
                     InitializeValue = ""
                Case vbBoolean
                     InitializeValue = True
                Case vbVariant
                     InitializeValue = ""
                Case vbByte, vbObject, vbError, vbDataObject, vbArray
                     InitializeValue = ""
         End Select
 End Function

 Public Function hPAD(ByVal sStr As String, ByVal nLength As Long, _
                      Optional ByVal sFillChar As String = "") As String
     hPAD = PAD(sStr, nLength, sFillChar)
 End Function
 Public Function hPADL(ByVal sStr As String, ByVal nLength As Long, _
                       Optional ByVal sFillChar As String = "") As String
     hPADL = PADL(sStr, nLength, sFillChar)
 End Function
 Public Function hPADR(ByVal sStr As String, ByVal nLength As Long, _
                       Optional ByVal sFillChar As String = "") As String
     hPADR = PADR(sStr, nLength, sFillChar)
 End Function
 Public Function hPADC(ByVal sStr As String, ByVal nLength As Long, _
                       Optional ByVal sFillChar As String = "") As String
     hPADC = PADC(sStr, nLength, sFillChar)
 End Function

