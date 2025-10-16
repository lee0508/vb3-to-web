USE YmhDB
GO
BEGIN TRAN
GO
--SELECT SUBSTRING(LTRIM(적요), 1, 2) + SUBSTRING(적요, 3, 8), LTRIM(STUFF(LTRIM(적요), 1, 10, '')), *
--  FROM 미수금내역 
-- WHERE 결제방법 = 2 AND 적요 IS NOT NULL
--   AND ISNUMERIC(SUBSTRING(LTRIM(적요), 3, 8)) > 0
--   AND ASCII(SUBSTRING(LTRIM('적요'), 1, 1)) >= 172
--   AND ASCII(SUBSTRING(LTRIM('적요'), 2, 1)) >= 172
--   AND SUBSTRING(LTRIM(STUFF(LTRIM(적요), 3, 10, '')), 10, 1) <> ''
--
PRINT '--- 1. 적요에서 어음번호만 빼기(미지급금내역) ---' -- 172('가')
UPDATE 미지급금내역 SET 어음번호 = (SUBSTRING(LTRIM(적요), 1, 2) + SUBSTRING(적요, 3, 8)), 적요 = LTRIM(STUFF(LTRIM(적요), 1, 10, ''))
 WHERE 결제방법 = 2 AND 적요 IS NOT NULL
   AND ISNUMERIC(SUBSTRING(LTRIM(적요), 3, 8)) > 0
   AND ASCII(SUBSTRING(LTRIM('적요'), 1, 1)) >= 172
   AND ASCII(SUBSTRING(LTRIM('적요'), 2, 1)) >= 172
   AND SUBSTRING(LTRIM(STUFF(LTRIM(적요), 3, 10, '')), 10, 1) <> ''
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 2. 적요에서 어음번호만 빼기(미수금내역) ---'
UPDATE 미수금내역 SET 어음번호 = (SUBSTRING(LTRIM(적요), 1, 2) + SUBSTRING(적요, 3, 8)), 적요 = LTRIM(STUFF(LTRIM(적요), 1, 10, ''))
 WHERE 결제방법 = 2 AND 적요 IS NOT NULL
   AND ISNUMERIC(SUBSTRING(LTRIM(적요), 3, 8)) > 0
   AND ASCII(SUBSTRING(LTRIM('적요'), 1, 1)) >= 172
   AND ASCII(SUBSTRING(LTRIM('적요'), 2, 1)) >= 172
   AND SUBSTRING(LTRIM(STUFF(LTRIM(적요), 3, 10, '')), 10, 1) <> ''
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 3. 미지급금기초이월(미지급금원장마감) ---'
DELETE FROM 미지급금원장마감
INSERT INTO 미지급금원장마감
SELECT '01' AS 사업장코드, T1.VCODE AS 매입처코드, '200000' AS 마감년월, 
       (ISNULL(T1.VMONEY, 0) -
       ISNULL((SELECT SUM(CONVERT(MONEY, ISNULL(VDMONEY, 0)) + (CONVERT(MONEY, ISNULL(VDMONEY1, 0)) - CONVERT(MONEY, ISNULL(VDMONEY, 0))))
          FROM ymhprg.dbo.JANGVD 
         WHERE VENDCO = T1.VCODE AND CONVERT(VARCHAR(8), VDTRDT, 112) BETWEEN '20000101' AND '20051231'), 0) +
       ISNULL((SELECT SUM(ISNULL(VDCOST, 0))
          FROM ymhprg.dbo.JANGVDTR
         WHERE VENDC = T1.VCODE AND CONVERT(VARCHAR(8), VDTRDAT, 112) BETWEEN '20000101' AND '20051231'), 0)) AS 미지급금누계금액,
       0 AS 미지급금지급누계금액, '' AS 수정일자, '8080' AS 사용자코드
  FROM ymhprg.dbo.vendor T1 
 WHERE (ISNULL(T1.VMONEY, 0) -
       ISNULL((SELECT SUM(CONVERT(MONEY, ISNULL(VDMONEY, 0)) + (CONVERT(MONEY, ISNULL(VDMONEY1, 0)) - CONVERT(MONEY, ISNULL(VDMONEY, 0))))
          FROM ymhprg.dbo.JANGVD 
         WHERE VENDCO = T1.VCODE AND CONVERT(VARCHAR(8), VDTRDT, 112) BETWEEN '20000101' AND '20051231'), 0) +
       ISNULL((SELECT SUM(ISNULL(VDCOST, 0))
          FROM ymhprg.dbo.JANGVDTR
         WHERE VENDC = T1.VCODE AND CONVERT(VARCHAR(8), VDTRDAT, 112) BETWEEN '20000101' AND '20051231'), 0)) <> 0
 ORDER BY T1.VCODE
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 4. 미지급금원장 월마감 ---'
INSERT INTO 미지급금원장마감
SELECT '01', T1.매입처코드, SUBSTRING(T1.작성일자, 1, 6),
       SUM(T1.공급가액 + T1.세액), 0, '', '8080'       
  FROM 매입세금계산서장부 T1
 WHERE T1.사업장코드 = '01' AND T1.사용구분 = 0 AND T1.매입처코드 <> ''
   AND T1.미지급구분 = 1 AND SUBSTRING(T1.작성일자, 1, 6) > '200000'
 GROUP BY T1.사업장코드, T1.매입처코드, SUBSTRING(T1.작성일자, 1, 6)
 ORDER BY T1.사업장코드, T1.매입처코드, SUBSTRING(T1.작성일자, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE 미지급금원장마감 SET 미지급금지급누계금액 =
       ISNULL((SELECT SUM(T2.미지급금지급금액) FROM 미지급금내역 T2 
                WHERE T2.사업장코드 = T1.사업장코드 AND T2.매입처코드 = T1.매입처코드
                  AND T2.매입처코드 <> '' AND SUBSTRING(T2.미지급금지급일자, 1, 6) = T1.마감년월), 0)
  FROM 미지급금원장마감 T1
 WHERE T1.사업장코드 = '01' AND SUBSTRING(T1.마감년월, 5, 2) <> '00'
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT 미지급금원장마감 
SELECT '01', T1.매입처코드, SUBSTRING(T1.미지급금지급일자, 1, 6),
       0, SUM(T1.미지급금지급금액), '', '8080'
  FROM 미지급금내역 T1
 WHERE T1.사업장코드 = '01' AND T1.매입처코드 <> ''
   AND SUBSTRING(T1.미지급금지급일자, 1, 6) > '200000' 
   AND NOT EXISTS 
      (SELECT T2.마감년월
         FROM 미지급금원장마감 T2 
        WHERE T2.사업장코드 = T1.사업장코드 AND T2.매입처코드 = T1.매입처코드 
          AND T2.마감년월 = SUBSTRING(T1.미지급금지급일자, 1, 6))
 GROUP BY T1.사업장코드, T1.매입처코드, SUBSTRING(T1.미지급금지급일자, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 5. 미수금기초이월(미수금원장마감) ---' 
DELETE FROM 미수금원장마감
INSERT INTO 미수금원장마감
SELECT '01' AS 사업장코드, T1.CCODE AS 매입처코드, '200000' AS 마감년월, 
       (ISNULL(T1.CMONEY, 0) -
       ISNULL((SELECT SUM(CONVERT(MONEY, ISNULL(CUMONEY, 0)) + (CONVERT(MONEY, ISNULL(CUMONEY1, 0)) - CONVERT(MONEY, ISNULL(CUMONEY, 0))))
          FROM ymhprg.dbo.JANGCU 
         WHERE CUSTCO = T1.CCODE AND CONVERT(VARCHAR(8), CUTRDT, 112) BETWEEN '20000101' AND '20051231'), 0) +
       ISNULL((SELECT SUM(ISNULL(CUCOST, 0))
          FROM ymhprg.dbo.JANGCUTR
         WHERE CUSTC = T1.CCODE AND CONVERT(VARCHAR(8), CUTRDAT, 112) BETWEEN '20000101' AND '20051231'), 0)) AS 미수금누계금액,
       0 AS 미수금입금누계금액, '' AS 수정일자, '8080' AS 사용자코드
  FROM ymhprg.dbo.custom T1 
 WHERE (ISNULL(T1.CMONEY, 0) -
       ISNULL((SELECT SUM(CONVERT(MONEY, ISNULL(CUMONEY, 0)) + (CONVERT(MONEY, ISNULL(CUMONEY1, 0)) - CONVERT(MONEY, ISNULL(CUMONEY, 0))))
          FROM ymhprg.dbo.JANGCU
         WHERE CUSTCO = T1.CCODE AND CONVERT(VARCHAR(8), CUTRDT, 112) BETWEEN '20000101' AND '20051231'), 0) +
       ISNULL((SELECT SUM(ISNULL(CUCOST, 0))
          FROM ymhprg.dbo.JANGCUTR
         WHERE CUSTC = T1.CCODE AND CONVERT(VARCHAR(8), CUTRDAT, 112) BETWEEN '20000101' AND '20051231'), 0)) <> 0
 ORDER BY T1.CCODE
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 6. 미수금원장 월마감 ---'
INSERT INTO 미수금원장마감
SELECT '01', T1.매출처코드, SUBSTRING(T1.작성일자, 1, 6),
       SUM(T1.공급가액 + T1.세액), 0, '', '8080'       
  FROM 매출세금계산서장부 T1
 WHERE T1.사업장코드 = '01' AND T1.사용구분 = 0 AND T1.매출처코드 <> ''
   AND T1.미수구분 = 1 AND SUBSTRING(T1.작성일자, 1, 6) > '200000'
 GROUP BY T1.사업장코드, T1.매출처코드, SUBSTRING(T1.작성일자, 1, 6)
 ORDER BY T1.사업장코드, T1.매출처코드, SUBSTRING(T1.작성일자, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE 미수금원장마감 SET 미수금입금누계금액 =
       ISNULL((SELECT SUM(T2.미수금입금금액) FROM 미수금내역 T2 
                WHERE T2.사업장코드 = T1.사업장코드 AND T2.매출처코드 = T1.매출처코드
                  AND T2.매출처코드 <> '' AND SUBSTRING(T2.미수금입금일자, 1, 6) = T1.마감년월), 0)
  FROM 미수금원장마감 T1
 WHERE T1.사업장코드 = '01' AND SUBSTRING(T1.마감년월, 5, 2) <> '00'
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT 미수금원장마감 
SELECT '01', T1.매출처코드, SUBSTRING(T1.미수금입금일자, 1, 6),
       0, SUM(T1.미수금입금금액), '', '8080'
  FROM 미수금내역 T1
 WHERE T1.사업장코드 = '01' AND T1.매출처코드 <> ''
   AND SUBSTRING(T1.미수금입금일자, 1, 6) > '200000' 
   AND NOT EXISTS 
      (SELECT T2.마감년월
         FROM 미수금원장마감 T2 
        WHERE T2.사업장코드 = T1.사업장코드 AND T2.매출처코드 = T1.매출처코드 
          AND T2.마감년월 = SUBSTRING(T1.미수금입금일자, 1, 6))
 GROUP BY T1.사업장코드, T1.매출처코드, SUBSTRING(T1.미수금입금일자, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
/*
PRINT '--- 7. 입고단가 일괄변경 ---'
UPDATE 자재원장 
   SET 입고단가1 = ISNULL(
      (SELECT TOP 1 입고단가 
         FROM 자재입출내역         
        WHERE 사업장코드 = T1.사업장코드 AND 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드                          
          AND 사용구분 = 0 AND 입출고구분 = 1 AND 입고단가 > 0 
        ORDER BY 사업장코드, 입출고일자 DESC, 입출고시간 DESC), 0)
  FROM 자재원장 T1 
 WHERE (T1.사업장코드 = '01' AND T1.입고단가1 = 0)
   AND ISNULL((SELECT TOP 1 입고단가 
         FROM 자재입출내역         
        WHERE 사업장코드 = T1.사업장코드 AND 분류코드 = T1.분류코드 AND 세부코드 = T1.세부코드                          
          AND 사용구분 = 0 AND 입출고구분 = 1 AND 입고단가 > 0 
        ORDER BY 입출고일자 DESC, 입출고시간 DESC), 0) > 0
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
*/
/*
PRINT '--- 8. 출고단가1 일괄갱신) ---'
BEGIN TRAN
UPDATE 자재원장 
   SET 출고단가1 = ISNULL(
      (SELECT TOP 1 출고단가 
         FROM 자재입출내역 T2
        INNER JOIN 매출처 T3 ON T3.사업장코드 = '01' AND T3.매출처코드 = T2.매출처코드 AND T3.단가구분 = 1
        WHERE T2.사업장코드 = T1.사업장코드 AND T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드                          
          AND T2.사용구분 = 0 AND T2.입출고구분 = 2 AND T2.출고단가 > 0 AND T3.단가구분 = 1
        ORDER BY T2.사업장코드, T2.입출고일자 DESC, T2.입출고시간 DESC), 0)
  FROM 자재원장 T1 
 WHERE (T1.사업장코드 = '01' AND T1.출고단가1 = 0)
   AND ISNULL((SELECT TOP 1 출고단가 
         FROM 자재입출내역 T2
        INNER JOIN 매출처 T3 ON T3.사업장코드 = '01' AND T3.매출처코드 = T2.매출처코드 AND T3.단가구분 = 1
        WHERE T2.사업장코드 = T1.사업장코드 AND T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드                          
          AND T2.사용구분 = 0 AND T2.입출고구분 = 2 AND T2.출고단가 > 0 AND T3.단가구분 = 1
        ORDER BY T2.사업장코드, T2.입출고일자 DESC, T2.입출고시간 DESC), 0) > 0
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 9. 출고단가2 일괄갱신) ---'
UPDATE 자재원장 
   SET 출고단가2 = ISNULL(
      (SELECT TOP 1 출고단가 
         FROM 자재입출내역 T2
        INNER JOIN 매출처 T3 ON T3.사업장코드 = '01' AND T3.매출처코드 = T2.매출처코드 AND T3.단가구분 = 2
        WHERE T2.사업장코드 = T1.사업장코드 AND T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드                          
          AND T2.사용구분 = 0 AND T2.입출고구분 = 2 AND T2.출고단가 > 0 AND T3.단가구분 = 2
        ORDER BY T2.사업장코드, T2.입출고일자 DESC, T2.입출고시간 DESC), 0)
  FROM 자재원장 T1 
 WHERE (T1.사업장코드 = '01' AND T1.출고단가2 = 0)
   AND ISNULL((SELECT TOP 1 출고단가 
         FROM 자재입출내역 T2
        INNER JOIN 매출처 T3 ON T3.사업장코드 = '01' AND T3.매출처코드 = T2.매출처코드 AND T3.단가구분 = 2
        WHERE T2.사업장코드 = T1.사업장코드 AND T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드                          
          AND T2.사용구분 = 0 AND T2.입출고구분 = 2 AND T2.출고단가 > 0 AND T3.단가구분 = 2
        ORDER BY T2.사업장코드, T2.입출고일자 DESC, T2.입출고시간 DESC), 0) > 0
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 10. 출고단가3 일괄갱신) ---'
UPDATE 자재원장 
   SET 출고단가3 = ISNULL(
      (SELECT TOP 1 출고단가 
         FROM 자재입출내역 T2
        INNER JOIN 매출처 T3 ON T3.사업장코드 = '01' AND T3.매출처코드 = T2.매출처코드 AND T3.단가구분 = 3
        WHERE T2.사업장코드 = T1.사업장코드 AND T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드                          
          AND T2.사용구분 = 0 AND T2.입출고구분 = 2 AND T2.출고단가 > 0 AND T3.단가구분 = 3
        ORDER BY T2.사업장코드, T2.입출고일자 DESC, T2.입출고시간 DESC), 0)
  FROM 자재원장 T1 
 WHERE (T1.사업장코드 = '01' AND T1.출고단가3 = 0)
   AND ISNULL((SELECT TOP 1 출고단가 
         FROM 자재입출내역 T2
        INNER JOIN 매출처 T3 ON T3.사업장코드 = '01' AND T3.매출처코드 = T2.매출처코드 AND T3.단가구분 = 3
        WHERE T2.사업장코드 = T1.사업장코드 AND T2.분류코드 = T1.분류코드 AND T2.세부코드 = T1.세부코드                          
          AND T2.사용구분 = 0 AND T2.입출고구분 = 2 AND T2.출고단가 > 0 AND T3.단가구분 = 3
        ORDER BY T2.사업장코드, T2.입출고일자 DESC, T2.입출고시간 DESC), 0) > 0
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

COMMIT TRAN
*/

/*------*/            
  pEXIT:
/*------*/      
-- CLOSE CUR11      
-- DEALLOCATE CUR11      
-- SET  @R_value = '1'             
PRINT '--- 정상적으로 자료가 처리가 되었습니다. (OK) ---'
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
