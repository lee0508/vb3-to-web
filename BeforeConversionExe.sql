USE YmhDB
GO
BEGIN TRAN
GO
SET QUOTED_IDENTIFIER OFF
PRINT '--- 0. Ư������(`)���� ---'
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

PRINT '--- 1. ����� ---'
--DELETE FROM �����
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO �����
SELECT ('0' + SETNO) AS ������ڵ�, LTRIM(ISNULL(CUSTNAME, '')) AS ������, LTRIM(ISNULL(CSANO, '')) AS ����ڹ�ȣ, '' AS ���ι�ȣ, 
       LTRIM(ISNULL(CCONT, '')) AS ��ǥ�ڸ�, '' AS ��ǥ���ֹι�ȣ, '' AS ��������, LTRIM(ISNULL(CZIP, '')) AS �����ȣ,
       LTRIM(ISNULL(CADDR, '')) AS �ּ�, '' AS ����, LTRIM(ISNULL(CUP, '')) AS ����, LTRIM(ISNULL(CJONG, '')) AS ����, 
       LTRIM(ISNULL(CTEL, '')) AS ��ȭ��ȣ, '' AS �ѽ���ȣ, CONVERT(MONEY, LTRIM(ISNULL(EVALUE, 0))) AS �ΰ�����, 
       2 AS �����ޱݹ߻�����, 2 AS �̼��ݹ߻�����, 0 AS ��뱸��, '' AS ��������, '8080' AS ������ڵ�,
       '' AS �̸����ּ�, ''AS Ȩ�������ּ�, 'C:/ymhprgw/Data' AS �������, 
       '200000' AS ������ʸ������, '200000' AS �����ޱݱ��ʸ������, '200000' AS �̼��ݱ��ʸ������, '200000' AS ȸ����ʸ������,
       1 AS ���Ÿ�Ա���, 5 AS �ŷ�������ܸ���, 13 AS �ŷ��������ʸ���, 2 AS ���ݰ�꼭��ܸ���, 20 AS ���ݰ�꼭���ʸ���,
       0 AS �����԰�ܰ��ڵ����ű���, 0 AS �������ܰ��ڵ����ű���
  FROM ymhprg.dbo.CODESET
 ORDER BY SETNO
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 2. ����� ---'
--DELETE FROM �����
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO �����
SELECT '8080' AS ������ڵ�, '���������' AS ����ڸ�, 99 AS ����ڱ���, 
       'N' AS �α��ο���, '0808' AS �α��κ�й�ȣ, '0808' AS �����й�ȣ, '01' AS ������ڵ�,
       0 AS ��뱸��, '' AS ��������, '' AS �����Ͻ�, '' AS �����Ͻ�
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 3. ����ó ---'
--DELETE FROM ����ó
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 4. ����ó ---'
--DELETE FROM ����ó
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO ����ó
SELECT '01' AS ������ڵ�, LTRIM(ISNULL(VCODE,'')) AS ����ó�ڵ�, LTRIM(ISNULL(VNAME,'')) AS ����ó��, 
       LTRIM(ISNULL(VSANO,'')) AS ����ڹ�ȣ, '' AS ���ι�ȣ, 
       LTRIM(ISNULL(VCONT,'')) AS ��ǥ�ڸ�, '' AS ��ǥ���ֹι�ȣ, '' AS ��������, 
       LTRIM(ISNULL(VZIP,'')) AS �����ȣ, LTRIM(ISNULL(VADDR,'')) AS �ּ�, '' AS ����,
       LTRIM(ISNULL('','')) AS ����, LTRIM(ISNULL('','')) AS ����, 
       LTRIM(ISNULL(VTEL,'')) AS ��ȭ��ȣ, LTRIM(ISNULL(VFAX,'')) AS �ѽ���ȣ,        
       '' AS �����ڵ�, '' AS ���¹�ȣ, ��꼭���࿩�� = CASE WHEN VSE = 'Y' THEN 1 ELSE 0 END, 100.00 AS ��꼭������,                                                                                
       '' AS ����ڸ�, 0 AS ��뱸��, '' AS ��������, '8080' AS ������ڵ�,
       RTRIM(LTRIM(ISNULL(VRMK, ''))) AS ����, ISNULL(VKUBUN, 1) AS �ܰ�����  
  FROM ymhprg.dbo.VENDOR
 ORDER BY VCODE
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 5. �����ޱݳ��� ---'
--DELETE FROM �����ޱݳ���
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 6.�����ޱݿ��帶�� ---'
--DELETE FROM �����ޱݿ��帶��
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

--INSERT INTO �����ޱݿ��帶��
--SELECT '01' AS ������ڵ�, LTRIM(ISNULL(VCODE,'')) AS ����ó�ڵ�, '200409' AS �������,
--       ISNULL(VMONEY,0) AS �����ޱݴ���ݾ�,  0 AS �����ޱ����ޱݾ�, '' AS ��������, '8080' AS ������ڵ�
--  FROM ymhprg.dbo.VENDOR
-- ORDER BY VCODE
--IF (@@ERROR <> 0)       
--    BEGIN       
--           GOTO ERR_RTN      
--    END      

PRINT '--- 7. ����ó ---'
--DELETE FROM ����ó
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
INSERT INTO ����ó
SELECT '01' AS ������ڵ�, LTRIM(ISNULL(CCODE,'')) AS ����ó�ڵ�, LTRIM(ISNULL(CNAME,'')) AS ����ó��, 
       LTRIM(ISNULL(CSANO,'')) AS ����ڹ�ȣ, '' AS ���ι�ȣ, 
       LTRIM(ISNULL(CCONT,'')) AS ��ǥ�ڸ�, '' AS ��ǥ���ֹι�ȣ, '' AS ��������, 
       LTRIM(ISNULL(CZIP,'')) AS �����ȣ, LTRIM(ISNULL(CADDR,'')) AS �ּ�, '' AS ����,
       LTRIM(ISNULL(CUP,'')) AS ����, LTRIM(ISNULL(CJONG,'')) AS ����, 
       LTRIM(ISNULL(CTEL,'')) AS ��ȭ��ȣ, LTRIM(ISNULL(CFAX,'')) AS �ѽ���ȣ,        
       '' AS �����ڵ�, '' AS ���¹�ȣ, ��꼭���࿩�� = CASE WHEN CSE = 'Y' THEN 1 ELSE 0 END, ISNULL(CSEYUL,0) AS ��꼭������,
       '' AS ����ڸ�, 0 AS ��뱸��, '' AS ��������, '8080' AS ������ڵ�,
       RTRIM(LTRIM(ISNULL(CRMK, ''))) AS ����, ISNULL(CKUBUN, 1) AS �ܰ�����         
  FROM ymhprg.dbo.CUSTOM
 ORDER BY CCODE
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 8. �̼��ݳ��� ---'
--DELETE FROM �̼��ݳ���
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 9. �̼��ݿ��帶�� ---'
--DELETE FROM �̼��ݿ��帶��
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

--INSERT INTO �̼��ݿ��帶��
--SELECT '01' AS ������ڵ�, LTRIM(ISNULL(CCODE,'')) AS ����ó�ڵ�, '200409' AS �������,
--       CONVERT(MONEY, ISNULL(CMONEY,0)) AS �̼��ݴ���ݾ�,  0 AS �̼����Աݱݾ�, '' AS ��������, '8080' AS ������ڵ�
--  FROM ymhprg.dbo.CUSTOM
-- ORDER BY CCODE
--IF (@@ERROR <> 0)       
--    BEGIN       
--           GOTO ERR_RTN      
--    END      

PRINT '--- 10. ����з�---'
--DELETE FROM ����з�
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
INSERT INTO ����з�
SELECT '01' AS �з��ڵ� , '����' AS �з���, '' AS ����, 0 AS ��뱸��, '' AS ��������, '8080' AS ������ڵ�
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 11. ���� ---'
--DELETE FROM ����
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO ����
SELECT '01' AS �з��ڵ�, LTRIM(ISNULL(PCODE,'')) AS �����ڵ�, LTRIM(ISNULL(PARTNAME,'')) AS �����, '' AS ���ڵ�,
       LTRIM(ISNULL(PARTSIZE,'')) AS �԰�, LTRIM(ISNULL(UNIT,'')) AS ����, 0 AS �����, 1 AS ��������,
       LTRIM(ISNULL(REMK,'')) AS ����, 0 AS ��뱸��, '' AS ��������, '8080' AS ������ڵ�
  FROM ymhprg.dbo.PART0CO
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 12. ����ü�(Not Used) ---'
--DELETE FROM ����ü�
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 13. ������� ---'
--DELETE FROM �������
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO �������
SELECT '01' AS ������ڵ�, '01' AS �з��ڵ�, LTRIM(ISNULL(PCODE,'')) AS �����ڵ�, 0 AS �������, ISNULL(MINQTY, 0) AS �������,
       '' AS �����԰�����, '' AS �����������,
       0 AS ��뱸��, '' AS ��������, '8080' AS ������ڵ�, ���� = RTRIM(LTRIM(ISNULL(REMK, ''))), ISNULL(VCODE, '') AS �ָ���ó�ڵ�, 
       ISNULL(COST1, 0) AS �԰�ܰ�1, 0 AS �԰�ܰ�2, 0 AS �԰�ܰ�3, 
       ISNULL(COST, 0) AS ���ܰ�1, ISNULL(COST2, 0) AS ���ܰ�2, ISNULL(COST3, 0) AS ���ܰ�3
  FROM ymhprg.dbo.PART0CO
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
PRINT '--- 14. ������帶�� ---'
--DELETE FROM ������帶��
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 15. �������⳻�� ---'
--DELETE FROM �������⳻��
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 16. �α� ---'
--DELETE FROM �α�
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 17. ���� ---'
--DELETE FROM ����
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
PRINT '--- 18. ���ֳ��� ---'
--DELETE FROM ���ֳ���
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
PRINT '--- 19. ���� ---'
--DELETE FROM ����
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
 
PRINT '--- 20. �������� ---'
--DELETE FROM ��������
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 21. �������� ---'
--DELETE FROM ��������
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT ��������       
SELECT ACCODE AS �����ڵ�, ISNULL(ACNAME, '') AS ������, 
       �հ�û�ǥ���Ῡ�� = CASE WHEN ACOPT = 'Y' THEN 'Y' ELSE 'N' END, 
       '' AS ����, 0 AS ��뱸��, '' AS ��������, '8080' AS ������ڵ�
  FROM ymhprg.dbo.ACCC
 WHERE ACCODE IS NOT NULL
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 22. ȸ����ǥ���� ---'
--DELETE FROM ȸ����ǥ����
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 23. ȸ����ǥ��������(������) ---'
--DELETE FROM ȸ����ǥ��������
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO ȸ����ǥ��������
SELECT '01', T1.�����ڵ�, SUBSTRING(T1.�ۼ�����, 1, 6),
       SUM(T1.�Աݱݾ�), SUM(T1.��ݱݾ�),
       '', '8080'
  FROM ȸ����ǥ���� T1 
 WHERE T1.������ڵ� = '01'
   AND T1.��뱸�� = 0
 GROUP BY T1.������ڵ�, T1.�����ڵ�, SUBSTRING(T1.�ۼ�����, 1, 6)
 ORDER BY T1.������ڵ�, T1.�����ڵ�, SUBSTRING(T1.�ۼ�����, 1, 6)
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

PRINT '--- 24. ���ݰ�꼭 ---'
--DELETE FROM ���ݰ�꼭
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO ���ݰ�꼭
SELECT '01' AS ������ȣ, SUBSTRING(CONVERT(VARCHAR(8), CUTRDAT, 112), 1,4) AS �ۼ��⵵, 
       ISNULL(CUK1, 0) AS å��ȣ, ISNULL(CUK2, 0) AS �Ϸù�ȣ, CUSTC AS ����ó�ڵ�, 
       CONVERT(VARCHAR(8), CUTRDAT, 112) AS �ۼ�����, ISNULL(CUPTNM, '') AS ǰ��ױ԰�, 
       ISNULL(CUOPT3, 0) AS ����, CONVERT(MONEY, ISNULL(CUCOST1, 0)) AS ���ް���, 
       (CONVERT(MONEY, ISNULL(CUCOST, 0)) - CONVERT(MONEY, ISNULL(CUCOST1, 0))) AS ����,
       �ݾױ��� = CASE WHEN CUOPT = 'A' THEN 0 
                       WHEN CUOPT = 'B' THEN 1
                       WHEN CUOPT = 'C' THEN 2
                       WHEN CUOPT = 'D' THEN 3
                       ELSE 0 END,
       ��û���� = CASE WHEN CUOPT1 = 'A' THEN 0 
                       WHEN CUOPT1 = 'B' THEN 1
                       ELSE 2 END,
       ���࿩�� = CASE WHEN CUOPT2 = 'Y' THEN 1 
                       ELSE 0 END,
       0 AS �ۼ�����, 1 AS �̼�����, '' AS ����, 0 AS ��뱸��, '' AS ��������, '8080' AS ������ڵ� 
  FROM ymhprg.dbo.JANGCUSE
 WHERE CUTRDAT IS NOT NULL AND CUSTC IS NOT NULL
   AND ISNULL(CUK1 , 0) <> 0 AND ISNULL(CUK2, 0) <> 0 
 ORDER BY CUTRDAT, CUK1, CUK2
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      

INSERT INTO �α�
SELECT '���ݰ�꼭', (T1.������ڵ� + T1.�ۼ��⵵),  MAX(T1.å��ȣ), 
  (SELECT MAX(�Ϸù�ȣ) FROM ���ݰ�꼭 
    WHERE (������ڵ� = T1.������ڵ�) AND (������ڵ� + �ۼ��⵵) = (T1.������ڵ� + T1.�ۼ��⵵) 
      AND å��ȣ = MAX(T1.å��ȣ)), '', '8080' 
  FROM ���ݰ�꼭 T1
 WHERE T1.������ڵ� = '01' 
 GROUP BY T1.������ڵ�, T1.�ۼ��⵵
 ORDER BY T1.������ڵ�, T1.�ۼ��⵵
IF (@@ERROR <> 0)       
    BEGIN       
           GOTO ERR_RTN      
    END      
PRINT '--- 25. ���� ---'
IF (SELECT COUNT(*) FROM ymhprg.dbo.BANKC) > 0
   BEGIN
      DELETE FROM ����
      IF (@@ERROR <> 0)       
          BEGIN       
                 GOTO ERR_RTN      
          END      
      INSERT INTO ����
      SELECT BANKCD1 AS �����ڵ�, (ISNULL(BANKNAME, '') + SPACE(1) + ISNULL(BANKJIJM, '')) AS �����̸�, 
            (ISNULL(BANKTEL, '') + sPACE(1) + ISNULL(BANKFIL, '')) AS ����, 0 AS ��뱸��, '' AS ��������, '8080' AS ������ڵ�
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
PRINT '--- ���������� �ڷᰡ ���� �Ǿ����ϴ�. (OK) ---'
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
