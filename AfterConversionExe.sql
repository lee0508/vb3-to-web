USE YmhDB
GO
BEGIN TRAN
GO
--SELECT SUBSTRING(LTRIM(����), 1, 2) + SUBSTRING(����, 3, 8), LTRIM(STUFF(LTRIM(����), 1, 10, '')), *
--  FROM �̼��ݳ��� 
-- WHERE ������� = 2 AND ���� IS NOT NULL
--   AND ISNUMERIC(SUBSTRING(LTRIM(����), 3, 8)) > 0
--   AND ASCII(SUBSTRING(LTRIM('����'), 1, 1)) >= 172
--   AND ASCII(SUBSTRING(LTRIM('����'), 2, 1)) >= 172
--   AND SUBSTRING(LTRIM(STUFF(LTRIM(����), 3, 10, '')), 10, 1) <> ''
--
PRINT '--- 1. ���信�� ������ȣ�� ����(�����ޱݳ���) ---' -- 172('��')
UPDATE �����ޱݳ��� SET ������ȣ = (SUBSTRING(LTRIM(����), 1, 2) + SUBSTRING(����, 3, 8)), ���� = LTRIM(STUFF(LTRIM(����), 1, 10, ''))
 WHERE ������� = 2 AND ���� IS NOT NULL
   AND ISNUMERIC(SUBSTRING(LTRIM(����), 3, 8)) > 0
   AND ASCII(SUBSTRING(LTRIM('����'), 1, 1)) >= 172
   AND ASCII(SUBSTRING(LTRIM('����'), 2, 1)) >= 172
   AND SUBSTRING(LTRIM(STUFF(LTRIM(����), 3, 10, '')), 10, 1) <> ''
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 2. ���信�� ������ȣ�� ����(�̼��ݳ���) ---'
UPDATE �̼��ݳ��� SET ������ȣ = (SUBSTRING(LTRIM(����), 1, 2) + SUBSTRING(����, 3, 8)), ���� = LTRIM(STUFF(LTRIM(����), 1, 10, ''))
 WHERE ������� = 2 AND ���� IS NOT NULL
   AND ISNUMERIC(SUBSTRING(LTRIM(����), 3, 8)) > 0
   AND ASCII(SUBSTRING(LTRIM('����'), 1, 1)) >= 172
   AND ASCII(SUBSTRING(LTRIM('����'), 2, 1)) >= 172
   AND SUBSTRING(LTRIM(STUFF(LTRIM(����), 3, 10, '')), 10, 1) <> ''
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 3. �����ޱݱ����̿�(�����ޱݿ��帶��) ---'
DELETE FROM �����ޱݿ��帶��
INSERT INTO �����ޱݿ��帶��
SELECT '01' AS ������ڵ�, T1.VCODE AS ����ó�ڵ�, '200000' AS �������, 
       (ISNULL(T1.VMONEY, 0) -
       ISNULL((SELECT SUM(CONVERT(MONEY, ISNULL(VDMONEY, 0)) + (CONVERT(MONEY, ISNULL(VDMONEY1, 0)) - CONVERT(MONEY, ISNULL(VDMONEY, 0))))
          FROM ymhprg.dbo.JANGVD 
         WHERE VENDCO = T1.VCODE AND CONVERT(VARCHAR(8), VDTRDT, 112) BETWEEN '20000101' AND '20051231'), 0) +
       ISNULL((SELECT SUM(ISNULL(VDCOST, 0))
          FROM ymhprg.dbo.JANGVDTR
         WHERE VENDC = T1.VCODE AND CONVERT(VARCHAR(8), VDTRDAT, 112) BETWEEN '20000101' AND '20051231'), 0)) AS �����ޱݴ���ݾ�,
       0 AS �����ޱ����޴���ݾ�, '' AS ��������, '8080' AS ������ڵ�
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

PRINT '--- 4. �����ޱݿ��� ������ ---'
INSERT INTO �����ޱݿ��帶��
SELECT '01', T1.����ó�ڵ�, SUBSTRING(T1.�ۼ�����, 1, 6),
       SUM(T1.���ް��� + T1.����), 0, '', '8080'       
  FROM ���Լ��ݰ�꼭��� T1
 WHERE T1.������ڵ� = '01' AND T1.��뱸�� = 0 AND T1.����ó�ڵ� <> ''
   AND T1.�����ޱ��� = 1 AND SUBSTRING(T1.�ۼ�����, 1, 6) > '200000'
 GROUP BY T1.������ڵ�, T1.����ó�ڵ�, SUBSTRING(T1.�ۼ�����, 1, 6)
 ORDER BY T1.������ڵ�, T1.����ó�ڵ�, SUBSTRING(T1.�ۼ�����, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE �����ޱݿ��帶�� SET �����ޱ����޴���ݾ� =
       ISNULL((SELECT SUM(T2.�����ޱ����ޱݾ�) FROM �����ޱݳ��� T2 
                WHERE T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ�
                  AND T2.����ó�ڵ� <> '' AND SUBSTRING(T2.�����ޱ���������, 1, 6) = T1.�������), 0)
  FROM �����ޱݿ��帶�� T1
 WHERE T1.������ڵ� = '01' AND SUBSTRING(T1.�������, 5, 2) <> '00'
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT �����ޱݿ��帶�� 
SELECT '01', T1.����ó�ڵ�, SUBSTRING(T1.�����ޱ���������, 1, 6),
       0, SUM(T1.�����ޱ����ޱݾ�), '', '8080'
  FROM �����ޱݳ��� T1
 WHERE T1.������ڵ� = '01' AND T1.����ó�ڵ� <> ''
   AND SUBSTRING(T1.�����ޱ���������, 1, 6) > '200000' 
   AND NOT EXISTS 
      (SELECT T2.�������
         FROM �����ޱݿ��帶�� T2 
        WHERE T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� 
          AND T2.������� = SUBSTRING(T1.�����ޱ���������, 1, 6))
 GROUP BY T1.������ڵ�, T1.����ó�ڵ�, SUBSTRING(T1.�����ޱ���������, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 5. �̼��ݱ����̿�(�̼��ݿ��帶��) ---' 
DELETE FROM �̼��ݿ��帶��
INSERT INTO �̼��ݿ��帶��
SELECT '01' AS ������ڵ�, T1.CCODE AS ����ó�ڵ�, '200000' AS �������, 
       (ISNULL(T1.CMONEY, 0) -
       ISNULL((SELECT SUM(CONVERT(MONEY, ISNULL(CUMONEY, 0)) + (CONVERT(MONEY, ISNULL(CUMONEY1, 0)) - CONVERT(MONEY, ISNULL(CUMONEY, 0))))
          FROM ymhprg.dbo.JANGCU 
         WHERE CUSTCO = T1.CCODE AND CONVERT(VARCHAR(8), CUTRDT, 112) BETWEEN '20000101' AND '20051231'), 0) +
       ISNULL((SELECT SUM(ISNULL(CUCOST, 0))
          FROM ymhprg.dbo.JANGCUTR
         WHERE CUSTC = T1.CCODE AND CONVERT(VARCHAR(8), CUTRDAT, 112) BETWEEN '20000101' AND '20051231'), 0)) AS �̼��ݴ���ݾ�,
       0 AS �̼����Աݴ���ݾ�, '' AS ��������, '8080' AS ������ڵ�
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

PRINT '--- 6. �̼��ݿ��� ������ ---'
INSERT INTO �̼��ݿ��帶��
SELECT '01', T1.����ó�ڵ�, SUBSTRING(T1.�ۼ�����, 1, 6),
       SUM(T1.���ް��� + T1.����), 0, '', '8080'       
  FROM ���⼼�ݰ�꼭��� T1
 WHERE T1.������ڵ� = '01' AND T1.��뱸�� = 0 AND T1.����ó�ڵ� <> ''
   AND T1.�̼����� = 1 AND SUBSTRING(T1.�ۼ�����, 1, 6) > '200000'
 GROUP BY T1.������ڵ�, T1.����ó�ڵ�, SUBSTRING(T1.�ۼ�����, 1, 6)
 ORDER BY T1.������ڵ�, T1.����ó�ڵ�, SUBSTRING(T1.�ۼ�����, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

UPDATE �̼��ݿ��帶�� SET �̼����Աݴ���ݾ� =
       ISNULL((SELECT SUM(T2.�̼����Աݱݾ�) FROM �̼��ݳ��� T2 
                WHERE T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ�
                  AND T2.����ó�ڵ� <> '' AND SUBSTRING(T2.�̼����Ա�����, 1, 6) = T1.�������), 0)
  FROM �̼��ݿ��帶�� T1
 WHERE T1.������ڵ� = '01' AND SUBSTRING(T1.�������, 5, 2) <> '00'
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT �̼��ݿ��帶�� 
SELECT '01', T1.����ó�ڵ�, SUBSTRING(T1.�̼����Ա�����, 1, 6),
       0, SUM(T1.�̼����Աݱݾ�), '', '8080'
  FROM �̼��ݳ��� T1
 WHERE T1.������ڵ� = '01' AND T1.����ó�ڵ� <> ''
   AND SUBSTRING(T1.�̼����Ա�����, 1, 6) > '200000' 
   AND NOT EXISTS 
      (SELECT T2.�������
         FROM �̼��ݿ��帶�� T2 
        WHERE T2.������ڵ� = T1.������ڵ� AND T2.����ó�ڵ� = T1.����ó�ڵ� 
          AND T2.������� = SUBSTRING(T1.�̼����Ա�����, 1, 6))
 GROUP BY T1.������ڵ�, T1.����ó�ڵ�, SUBSTRING(T1.�̼����Ա�����, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
/*
PRINT '--- 7. �԰�ܰ� �ϰ����� ---'
UPDATE ������� 
   SET �԰�ܰ�1 = ISNULL(
      (SELECT TOP 1 �԰�ܰ� 
         FROM �������⳻��         
        WHERE ������ڵ� = T1.������ڵ� AND �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ�                          
          AND ��뱸�� = 0 AND ������� = 1 AND �԰�ܰ� > 0 
        ORDER BY ������ڵ�, ��������� DESC, �����ð� DESC), 0)
  FROM ������� T1 
 WHERE (T1.������ڵ� = '01' AND T1.�԰�ܰ�1 = 0)
   AND ISNULL((SELECT TOP 1 �԰�ܰ� 
         FROM �������⳻��         
        WHERE ������ڵ� = T1.������ڵ� AND �з��ڵ� = T1.�з��ڵ� AND �����ڵ� = T1.�����ڵ�                          
          AND ��뱸�� = 0 AND ������� = 1 AND �԰�ܰ� > 0 
        ORDER BY ��������� DESC, �����ð� DESC), 0) > 0
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
*/
/*
PRINT '--- 8. ���ܰ�1 �ϰ�����) ---'
BEGIN TRAN
UPDATE ������� 
   SET ���ܰ�1 = ISNULL(
      (SELECT TOP 1 ���ܰ� 
         FROM �������⳻�� T2
        INNER JOIN ����ó T3 ON T3.������ڵ� = '01' AND T3.����ó�ڵ� = T2.����ó�ڵ� AND T3.�ܰ����� = 1
        WHERE T2.������ڵ� = T1.������ڵ� AND T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ�                          
          AND T2.��뱸�� = 0 AND T2.������� = 2 AND T2.���ܰ� > 0 AND T3.�ܰ����� = 1
        ORDER BY T2.������ڵ�, T2.��������� DESC, T2.�����ð� DESC), 0)
  FROM ������� T1 
 WHERE (T1.������ڵ� = '01' AND T1.���ܰ�1 = 0)
   AND ISNULL((SELECT TOP 1 ���ܰ� 
         FROM �������⳻�� T2
        INNER JOIN ����ó T3 ON T3.������ڵ� = '01' AND T3.����ó�ڵ� = T2.����ó�ڵ� AND T3.�ܰ����� = 1
        WHERE T2.������ڵ� = T1.������ڵ� AND T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ�                          
          AND T2.��뱸�� = 0 AND T2.������� = 2 AND T2.���ܰ� > 0 AND T3.�ܰ����� = 1
        ORDER BY T2.������ڵ�, T2.��������� DESC, T2.�����ð� DESC), 0) > 0
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 9. ���ܰ�2 �ϰ�����) ---'
UPDATE ������� 
   SET ���ܰ�2 = ISNULL(
      (SELECT TOP 1 ���ܰ� 
         FROM �������⳻�� T2
        INNER JOIN ����ó T3 ON T3.������ڵ� = '01' AND T3.����ó�ڵ� = T2.����ó�ڵ� AND T3.�ܰ����� = 2
        WHERE T2.������ڵ� = T1.������ڵ� AND T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ�                          
          AND T2.��뱸�� = 0 AND T2.������� = 2 AND T2.���ܰ� > 0 AND T3.�ܰ����� = 2
        ORDER BY T2.������ڵ�, T2.��������� DESC, T2.�����ð� DESC), 0)
  FROM ������� T1 
 WHERE (T1.������ڵ� = '01' AND T1.���ܰ�2 = 0)
   AND ISNULL((SELECT TOP 1 ���ܰ� 
         FROM �������⳻�� T2
        INNER JOIN ����ó T3 ON T3.������ڵ� = '01' AND T3.����ó�ڵ� = T2.����ó�ڵ� AND T3.�ܰ����� = 2
        WHERE T2.������ڵ� = T1.������ڵ� AND T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ�                          
          AND T2.��뱸�� = 0 AND T2.������� = 2 AND T2.���ܰ� > 0 AND T3.�ܰ����� = 2
        ORDER BY T2.������ڵ�, T2.��������� DESC, T2.�����ð� DESC), 0) > 0
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 10. ���ܰ�3 �ϰ�����) ---'
UPDATE ������� 
   SET ���ܰ�3 = ISNULL(
      (SELECT TOP 1 ���ܰ� 
         FROM �������⳻�� T2
        INNER JOIN ����ó T3 ON T3.������ڵ� = '01' AND T3.����ó�ڵ� = T2.����ó�ڵ� AND T3.�ܰ����� = 3
        WHERE T2.������ڵ� = T1.������ڵ� AND T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ�                          
          AND T2.��뱸�� = 0 AND T2.������� = 2 AND T2.���ܰ� > 0 AND T3.�ܰ����� = 3
        ORDER BY T2.������ڵ�, T2.��������� DESC, T2.�����ð� DESC), 0)
  FROM ������� T1 
 WHERE (T1.������ڵ� = '01' AND T1.���ܰ�3 = 0)
   AND ISNULL((SELECT TOP 1 ���ܰ� 
         FROM �������⳻�� T2
        INNER JOIN ����ó T3 ON T3.������ڵ� = '01' AND T3.����ó�ڵ� = T2.����ó�ڵ� AND T3.�ܰ����� = 3
        WHERE T2.������ڵ� = T1.������ڵ� AND T2.�з��ڵ� = T1.�з��ڵ� AND T2.�����ڵ� = T1.�����ڵ�                          
          AND T2.��뱸�� = 0 AND T2.������� = 2 AND T2.���ܰ� > 0 AND T3.�ܰ����� = 3
        ORDER BY T2.������ڵ�, T2.��������� DESC, T2.�����ð� DESC), 0) > 0
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
PRINT '--- ���������� �ڷᰡ ó���� �Ǿ����ϴ�. (OK) ---'
COMMIT TRAN
RETURN             
      
/*--------*/            
  ERR_RTN:
/*--------*/      
-- CLOSE CUR11      
-- DEALLOCATE CUR11      
-- SET  @R_value = '0'      
PRINT '--- ���� �߻����� �ڷᰡ ����ġ �Ǿ����ϴ�. (ERROR) ---'
ROLLBACK TRAN
RETURN      
