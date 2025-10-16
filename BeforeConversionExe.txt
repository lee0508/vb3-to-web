USE YmhDB
GO
BEGIN TRAN
GO
SET QUOTED_IDENTIFIER OFF
PRINT '--- 0. 특수문자(`)변경 ---'
UPDATE ymhprg.dbo.PART0CO SET 
       PCODE = REPLACE(PCODE, "'", '`'), PARTNAME = REPLACE(PARTNAME, "'", '`'), PARTSIZE = REPLACE(PARTSIZE, "'", '`'),
       REMK = REPLACE(REMK, "'", '`')      
 WHERE PCODE LIKE "%'%" OR PARTNAME LIKE "%'%" OR PARTSIZE LIKE "%'%" OR REMK LIKE "%'%"
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE ymhprg.dbo.INTR SET 
       INPART = REPLACE(INPART, "'", '`'), INPARTNM = REPLACE(INPARTNM, "'", '`'), INPARTSZ = REPLACE(INPARTSZ, "'", '`')     
 WHERE INPART LIKE "%'%" OR INPARTNM LIKE "%'%" OR INPARTSZ LIKE "%'%" 
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE ymhprg.dbo.OUTTR SET 
       OTPART = REPLACE(OTPART, "'", '`'), OTPARTNM = REPLACE(OTPARTNM, "'", '`'), OTPARTSZ = REPLACE(OTPARTSZ, "'", '`')
 WHERE OTPART LIKE "%'%" OR OTPARTNM LIKE "%'%" OR OTPARTSZ LIKE "%'%" 
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE ymhprg.dbo.OFFER SET 
       PCODE = REPLACE(PCODE, "'", '`'),
       SPARTNM = REPLACE(SPARTNM, "'", '`') 
 WHERE PCODE LIKE "%'%" OR SPARTNM LIKE "%'%"
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END             

UPDATE ymhprg.dbo.JANGVD SET VDRMK = REPLACE(VDRMK, "'", '`') WHERE VDRMK LIKE "%'%"
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE ymhprg.dbo.JANGCU SET CURMK = REPLACE(CURMK, "'", '`') WHERE CURMK LIKE "%'%"
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE ymhprg.dbo.ACCT SET ACCASE = REPLACE(ACCASE, "'", '`') WHERE ACCASE LIKE "%'%"
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE ymhprg.dbo.VENDOR SET VRMK = REPLACE(VRMK, "'", '`') WHERE VRMK LIKE "%'%"
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE ymhprg.dbo.CUSTOM SET CRMK = REPLACE(CRMK, "'", '`') WHERE CRMK LIKE "%'%"
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

SET QUOTED_IDENTIFIER ON

PRINT '--- 1. 사업장 ---'
--DELETE FROM 사업장
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO 사업장
SELECT ('0' + SETNO) AS 사업장코드, LTRIM(ISNULL(CUSTNAME, '')) AS 사업장명, LTRIM(ISNULL(CSANO, '')) AS 사업자번호, '' AS 법인번호, 
       LTRIM(ISNULL(CCONT, '')) AS 대표자명, '' AS 대표자주민번호, '' AS 개업일자, LTRIM(ISNULL(CZIP, '')) AS 우편번호,
       LTRIM(ISNULL(CADDR, '')) AS 주소, '' AS 번지, LTRIM(ISNULL(CUP, '')) AS 업태, LTRIM(ISNULL(CJONG, '')) AS 업종, 
       LTRIM(ISNULL(CTEL, '')) AS 전화번호, '' AS 팩스번호, CONVERT(MONEY, LTRIM(ISNULL(EVALUE, 0))) AS 부가세율, 
       2 AS 미지급금발생구분, 2 AS 미수금발생구분, 0 AS 사용구분, '' AS 수정일자, '8080' AS 사용자코드,
       '' AS 이메일주소, ''AS 홈페이지주소, 'C:/ymhprgw/Data' AS 백업폴더, 
       '200000' AS 자재기초마감년월, '200000' AS 미지급금기초마감년월, '200000' AS 미수금기초마감년월, '200000' AS 회계기초마감년월,
       1 AS 출력타입구분, 5 AS 거래명세서상단마진, 13 AS 거래명세서왼쪽마진, 2 AS 세금계산서상단마진, 20 AS 세금계산서왼쪽마진,
       0 AS 최종입고단가자동갱신구분, 0 AS 최종출고단가자동갱신구분
  FROM ymhprg.dbo.CODESET
 ORDER BY SETNO
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 2. 사용자 ---'
--DELETE FROM 사용자
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO 사용자
SELECT '8080' AS 사용자코드, '컨버젼담당' AS 사용자명, 99 AS 사용자권한, 
       'N' AS 로그인여부, '0808' AS 로그인비밀번호, '0808' AS 결재비밀번호, '01' AS 사업장코드,
       0 AS 사용구분, '' AS 수정일자, '' AS 시작일시, '' AS 종료일시
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 3. 제조처 ---'
--DELETE FROM 제조처
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 4. 매입처 ---'
--DELETE FROM 매입처
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO 매입처
SELECT '01' AS 사업장코드, LTRIM(ISNULL(VCODE,'')) AS 매입처코드, LTRIM(ISNULL(VNAME,'')) AS 매입처명, 
       LTRIM(ISNULL(VSANO,'')) AS 사업자번호, '' AS 법인번호, 
       LTRIM(ISNULL(VCONT,'')) AS 대표자명, '' AS 대표자주민번호, '' AS 개업일자, 
       LTRIM(ISNULL(VZIP,'')) AS 우편번호, LTRIM(ISNULL(VADDR,'')) AS 주소, '' AS 번지,
       LTRIM(ISNULL('','')) AS 업태, LTRIM(ISNULL('','')) AS 업종, 
       LTRIM(ISNULL(VTEL,'')) AS 전화번호, LTRIM(ISNULL(VFAX,'')) AS 팩스번호,        
       '' AS 은행코드, '' AS 계좌번호, 계산서발행여부 = CASE WHEN VSE = 'Y' THEN 1 ELSE 0 END, 100.00 AS 계산서발행율,                                                                                
       '' AS 담당자명, 0 AS 사용구분, '' AS 수정일자, '8080' AS 사용자코드,
       RTRIM(LTRIM(ISNULL(VRMK, ''))) AS 비고란, ISNULL(VKUBUN, 1) AS 단가구분  
  FROM ymhprg.dbo.VENDOR
 ORDER BY VCODE
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 5. 미지급금내역 ---'
--DELETE FROM 미지급금내역
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 6.미지급금원장마감 ---'
--DELETE FROM 미지급금원장마감
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

--INSERT INTO 미지급금원장마감
--SELECT '01' AS 사업장코드, LTRIM(ISNULL(VCODE,'')) AS 매입처코드, '200409' AS 마감년월,
--       ISNULL(VMONEY,0) AS 미지급금누계금액,  0 AS 미지급금지급금액, '' AS 수정일자, '8080' AS 사용자코드
--  FROM ymhprg.dbo.VENDOR
-- ORDER BY VCODE
--IF (@@ERROR <> 0)       
--    BEGIN       
--           GOTO ERR_RTN      
--    END      

PRINT '--- 7. 매출처 ---'
--DELETE FROM 매출처
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
INSERT INTO 매출처
SELECT '01' AS 사업장코드, LTRIM(ISNULL(CCODE,'')) AS 매출처코드, LTRIM(ISNULL(CNAME,'')) AS 매출처명, 
       LTRIM(ISNULL(CSANO,'')) AS 사업자번호, '' AS 법인번호, 
       LTRIM(ISNULL(CCONT,'')) AS 대표자명, '' AS 대표자주민번호, '' AS 개업일자, 
       LTRIM(ISNULL(CZIP,'')) AS 우편번호, LTRIM(ISNULL(CADDR,'')) AS 주소, '' AS 번지,
       LTRIM(ISNULL(CUP,'')) AS 업태, LTRIM(ISNULL(CJONG,'')) AS 업종, 
       LTRIM(ISNULL(CTEL,'')) AS 전화번호, LTRIM(ISNULL(CFAX,'')) AS 팩스번호,        
       '' AS 은행코드, '' AS 계좌번호, 계산서발행여부 = CASE WHEN CSE = 'Y' THEN 1 ELSE 0 END, ISNULL(CSEYUL,0) AS 계산서발행율,
       '' AS 담당자명, 0 AS 사용구분, '' AS 수정일자, '8080' AS 사용자코드,
       RTRIM(LTRIM(ISNULL(CRMK, ''))) AS 비고란, ISNULL(CKUBUN, 1) AS 단가구분         
  FROM ymhprg.dbo.CUSTOM
 ORDER BY CCODE
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 8. 미수금내역 ---'
--DELETE FROM 미수금내역
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 9. 미수금원장마감 ---'
--DELETE FROM 미수금원장마감
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

--INSERT INTO 미수금원장마감
--SELECT '01' AS 사업장코드, LTRIM(ISNULL(CCODE,'')) AS 매출처코드, '200409' AS 마감년월,
--       CONVERT(MONEY, ISNULL(CMONEY,0)) AS 미수금누계금액,  0 AS 미수금입금금액, '' AS 수정일자, '8080' AS 사용자코드
--  FROM ymhprg.dbo.CUSTOM
-- ORDER BY CCODE
--IF (@@ERROR <> 0)       
--    BEGIN       
--           GOTO ERR_RTN      
--    END      

PRINT '--- 10. 자재분류---'
--DELETE FROM 자재분류
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
INSERT INTO 자재분류
SELECT '01' AS 분류코드 , '자재' AS 분류명, '' AS 적요, 0 AS 사용구분, '' AS 수정일자, '8080' AS 사용자코드
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 11. 자재 ---'
--DELETE FROM 자재
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO 자재
SELECT '01' AS 분류코드, LTRIM(ISNULL(PCODE,'')) AS 세부코드, LTRIM(ISNULL(PARTNAME,'')) AS 자재명, '' AS 바코드,
       LTRIM(ISNULL(PARTSIZE,'')) AS 규격, LTRIM(ISNULL(UNIT,'')) AS 단위, 0 AS 폐기율, 1 AS 과세구분,
       LTRIM(ISNULL(REMK,'')) AS 적요, 0 AS 사용구분, '' AS 수정일자, '8080' AS 사용자코드
  FROM ymhprg.dbo.PART0CO
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 12. 자재시세(Not Used) ---'
--DELETE FROM 자재시세
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 13. 자재원장 ---'
--DELETE FROM 자재원장
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO 자재원장
SELECT '01' AS 사업장코드, '01' AS 분류코드, LTRIM(ISNULL(PCODE,'')) AS 세부코드, 0 AS 적정재고, ISNULL(MINQTY, 0) AS 최저재고,
       '' AS 최종입고일자, '' AS 최종출고일자,
       0 AS 사용구분, '' AS 수정일자, '8080' AS 사용자코드, 비고란 = RTRIM(LTRIM(ISNULL(REMK, ''))), ISNULL(VCODE, '') AS 주매입처코드, 
       ISNULL(COST1, 0) AS 입고단가1, 0 AS 입고단가2, 0 AS 입고단가3, 
       ISNULL(COST, 0) AS 출고단가1, ISNULL(COST2, 0) AS 출고단가2, ISNULL(COST3, 0) AS 출고단가3
  FROM ymhprg.dbo.PART0CO
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
PRINT '--- 14. 자재원장마감 ---'
--DELETE FROM 자재원장마감
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 15. 자재입출내역 ---'
--DELETE FROM 자재입출내역
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 16. 로그 ---'
--DELETE FROM 로그
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 17. 발주 ---'
--DELETE FROM 발주
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
PRINT '--- 18. 발주내역 ---'
--DELETE FROM 발주내역
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
PRINT '--- 19. 견적 ---'
--DELETE FROM 견적
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
PRINT '--- 20. 견적내역 ---'
--DELETE FROM 견적내역
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 21. 계정과목 ---'
--DELETE FROM 계정과목
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT 계정과목       
SELECT ACCODE AS 계정코드, ISNULL(ACNAME, '') AS 계정명, 
       합계시산표연결여부 = CASE WHEN ACOPT = 'Y' THEN 'Y' ELSE 'N' END, 
       '' AS 적요, 0 AS 사용구분, '' AS 수정일자, '8080' AS 사용자코드
  FROM ymhprg.dbo.ACCC
 WHERE ACCODE IS NOT NULL
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 22. 회계전표내역 ---'
--DELETE FROM 회계전표내역
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 23. 회계전표내역마감(월마감) ---'
--DELETE FROM 회계전표내역마감
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO 회계전표내역마감
SELECT '01', T1.계정코드, SUBSTRING(T1.작성일자, 1, 6),
       SUM(T1.입금금액), SUM(T1.출금금액),
       '', '8080'
  FROM 회계전표내역 T1 
 WHERE T1.사업장코드 = '01'
   AND T1.사용구분 = 0
 GROUP BY T1.사업장코드, T1.계정코드, SUBSTRING(T1.작성일자, 1, 6)
 ORDER BY T1.사업장코드, T1.계정코드, SUBSTRING(T1.작성일자, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 24. 세금계산서 ---'
--DELETE FROM 세금계산서
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO 세금계산서
SELECT '01' AS 사업장번호, SUBSTRING(CONVERT(VARCHAR(8), CUTRDAT, 112), 1,4) AS 작성년도, 
       ISNULL(CUK1, 0) AS 책번호, ISNULL(CUK2, 0) AS 일련번호, CUSTC AS 매출처코드, 
       CONVERT(VARCHAR(8), CUTRDAT, 112) AS 작성일자, ISNULL(CUPTNM, '') AS 품목및규격, 
       ISNULL(CUOPT3, 0) AS 수량, CONVERT(MONEY, ISNULL(CUCOST1, 0)) AS 공급가액, 
       (CONVERT(MONEY, ISNULL(CUCOST, 0)) - CONVERT(MONEY, ISNULL(CUCOST1, 0))) AS 세액,
       금액구분 = CASE WHEN CUOPT = 'A' THEN 0 
                       WHEN CUOPT = 'B' THEN 1
                       WHEN CUOPT = 'C' THEN 2
                       WHEN CUOPT = 'D' THEN 3
                       ELSE 0 END,
       영청구분 = CASE WHEN CUOPT1 = 'A' THEN 0 
                       WHEN CUOPT1 = 'B' THEN 1
                       ELSE 2 END,
       발행여부 = CASE WHEN CUOPT2 = 'Y' THEN 1 
                       ELSE 0 END,
       0 AS 작성구분, 1 AS 미수구분, '' AS 적요, 0 AS 사용구분, '' AS 수정일자, '8080' AS 사용자코드 
  FROM ymhprg.dbo.JANGCUSE
 WHERE CUTRDAT IS NOT NULL AND CUSTC IS NOT NULL
   AND ISNULL(CUK1 , 0) <> 0 AND ISNULL(CUK2, 0) <> 0 
 ORDER BY CUTRDAT, CUK1, CUK2
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO 로그
SELECT '세금계산서', (T1.사업장코드 + T1.작성년도),  MAX(T1.책번호), 
  (SELECT MAX(일련번호) FROM 세금계산서 
    WHERE (사업장코드 = T1.사업장코드) AND (사업장코드 + 작성년도) = (T1.사업장코드 + T1.작성년도) 
      AND 책번호 = MAX(T1.책번호)), '', '8080' 
  FROM 세금계산서 T1
 WHERE T1.사업장코드 = '01' 
 GROUP BY T1.사업장코드, T1.작성년도
 ORDER BY T1.사업장코드, T1.작성년도
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
PRINT '--- 25. 은행 ---'
IF (SELECT COUNT(*) FROM ymhprg.dbo.BANKC) > 0
   BEGIN
      DELETE FROM 은행
      IF (@@ERROR <> 0)       
          BEGIN       
                 GOTO ERR_RTN      
          END      
      INSERT INTO 은행
      SELECT BANKCD1 AS 은행코드, (ISNULL(BANKNAME, '') + SPACE(1) + ISNULL(BANKJIJM, '')) AS 은행이름, 
            (ISNULL(BANKTEL, '') + sPACE(1) + ISNULL(BANKFIL, '')) AS 적요, 0 AS 사용구분, '' AS 수정일자, '8080' AS 사용자코드
       FROM ymhprg.dbo.BANKC
      WHERE BANKCD1 IS NOT NULL
      IF (@@ERROR <> 0)       
          BEGIN       
                 GOTO ERR_RTN      
          END      
   END

/*------*/            
  pEXIT:
/*------*/      
-- CLOSE CUR11      
-- DEALLOCATE CUR11      
-- SET  @R_value = '1'             
PRINT '--- 정상적으로 자료가 저장 되었습니다. (OK) ---'
COMMIT TRAN
RETURN             
      
/*--------*/            
  ERR_RTN:
/*--------*/      
-- CLOSE CUR11      
-- DEALLOCATE CUR11      
-- SET  @R_value = '0'      
PRINT '--- 오류 발생으로 자료가 원위치 되었습니다. (ERROR) ---'
ROLLBACK TRAN
RETURN      
