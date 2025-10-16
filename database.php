<?php

/**
 * 데이터베이스 연결 설정
 * VB의 ADODB Connection을 PHP PDO로 구현
 * 
 * PDO 사용 이유:
 * 1. SQL Injection 방지 (Prepared Statements)
 * 2. 다양한 DB 지원 (MySQL, SQL Server, PostgreSQL 등)
 * 3. 에러 핸들링 향상
 * 4. 객체 지향적 접근
 */

// 데이터베이스 설정 상수
define('DB_TYPE', 'sqlsrv');           // MS SQL Server 사용 (VB와 동일)
define('DB_HOST', 'localhost');        // 서버 주소
define('DB_PORT', '1433');             // SQL Server 기본 포트
define('DB_NAME', 'SalesManagement');  // 데이터베이스 이름
define('DB_USER', 'sa');               // 사용자명
define('DB_PASS', 'your_password');    // 비밀번호
define('DB_CHARSET', 'UTF-8');         // 문자셋

// 에러 출력 설정
// 기본값: 운영환경 안전을 위해 false
// 개발 중에는 local 환경에서만 true로 변경하거나 별도 config/dev.php를 사용하세요.
define('DISPLAY_ERRORS', false);

if (DISPLAY_ERRORS) {
  error_reporting(E_ALL);
  ini_set('display_errors', 1);
} else {
  error_reporting(0);
  ini_set('display_errors', 0);
  ini_set('log_errors', 1);
  ini_set('error_log', __DIR__ . '/../logs/error.log');
}

// 타임존 설정
date_default_timezone_set('Asia/Seoul');

try {
  /**
   * MS SQL Server 연결 (VB의 ADODB.Connection과 동일 역할)
   * 
   * PDO DSN 형식:
   * - SQL Server: sqlsrv:Server=서버주소;Database=DB명
   * - MySQL: mysql:host=서버주소;dbname=DB명;charset=utf8mb4
   */

  if (DB_TYPE === 'sqlsrv') {
    // MS SQL Server 연결 (VB와 동일한 DB 사용)
    $dsn = "sqlsrv:Server=" . DB_HOST . "," . DB_PORT . ";Database=" . DB_NAME;
    $options = [
      PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,  // 에러 발생 시 예외 처리
      PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,  // 연관 배열로 결과 반환
      PDO::ATTR_EMULATE_PREPARES => false,  // 실제 Prepared Statement 사용
      PDO::SQLSRV_ATTR_ENCODING => PDO::SQLSRV_ENCODING_UTF8  // UTF-8 인코딩
    ];
  } elseif (DB_TYPE === 'mysql') {
    // MySQL 연결 (선택적 - 새로 구축할 경우)
    $dsn = "mysql:host=" . DB_HOST . ";dbname=" . DB_NAME . ";charset=utf8mb4";
    $options = [
      PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
      PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
      PDO::ATTR_EMULATE_PREPARES => false,
      PDO::MYSQL_ATTR_INIT_COMMAND => "SET NAMES utf8mb4"
    ];
  } else {
    throw new Exception('지원하지 않는 데이터베이스 타입입니다.');
  }

  // PDO 객체 생성 (전역 변수로 사용)
  $pdo = new PDO($dsn, DB_USER, DB_PASS, $options);

  // 연결 성공 로그 (개발 환경에서만)
  if (DISPLAY_ERRORS) {
    // echo "데이터베이스 연결 성공<br>";
  }
} catch (PDOException $e) {
  // 데이터베이스 연결 실패 처리
  $error_message = "데이터베이스 연결 실패: " . $e->getMessage();

  // 에러 로그 기록
  error_log($error_message);

  // 사용자에게 표시할 에러 메시지 (보안을 위해 상세 정보 숨김)
  if (DISPLAY_ERRORS) {
    die("<h3>데이터베이스 연결 오류</h3><p>" . htmlspecialchars($error_message) . "</p>");
  } else {
    die("<h3>시스템 오류</h3><p>데이터베이스에 연결할 수 없습니다. 관리자에게 문의하세요.</p>");
  }
}

/**
 * 데이터베이스 유틸리티 함수들
 */

/**
 * 안전한 쿼리 실행 함수
 * 
 * @param PDO $pdo PDO 객체
 * @param string $sql SQL 쿼리
 * @param array $params 바인딩할 파라미터
 * @return PDOStatement|false
 */
function executeQuery($pdo, $sql, $params = [])
{
  try {
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    return $stmt;
  } catch (PDOException $e) {
    error_log("Query Error: " . $e->getMessage() . "\nSQL: " . $sql);
    return false;
  }
}

/**
 * 단일 레코드 조회
 * 
 * @param PDO $pdo PDO 객체
 * @param string $sql SQL 쿼리
 * @param array $params 바인딩할 파라미터
 * @return array|false
 */
function fetchOne($pdo, $sql, $params = [])
{
  $stmt = executeQuery($pdo, $sql, $params);
  return $stmt ? $stmt->fetch() : false;
}

/**
 * 다중 레코드 조회
 * 
 * @param PDO $pdo PDO 객체
 * @param string $sql SQL 쿼리
 * @param array $params 바인딩할 파라미터
 * @return array
 */
function fetchAll($pdo, $sql, $params = [])
{
  $stmt = executeQuery($pdo, $sql, $params);
  return $stmt ? $stmt->fetchAll() : [];
}

/**
 * INSERT/UPDATE/DELETE 실행 및 영향받은 행 수 반환
 * 
 * @param PDO $pdo PDO 객체
 * @param string $sql SQL 쿼리
 * @param array $params 바인딩할 파라미터
 * @return int|false 영향받은 행 수
 */
function executeUpdate($pdo, $sql, $params = [])
{
  $stmt = executeQuery($pdo, $sql, $params);
  return $stmt ? $stmt->rowCount() : false;
}

/**
 * 트랜잭션 시작
 * VB의 BeginTrans와 동일
 */
function beginTransaction($pdo)
{
  return $pdo->beginTransaction();
}

/**
 * 트랜잭션 커밋
 * VB의 CommitTrans와 동일
 */
function commitTransaction($pdo)
{
  return $pdo->commit();
}

/**
 * 트랜잭션 롤백
 * VB의 RollbackTrans와 동일
 */
function rollbackTransaction($pdo)
{
  return $pdo->rollBack();
}

/**
 * 현재 날짜/시간 가져오기 (SQL Server 형식)
 * 
 * @return string YYYYMMDD 형식의 날짜
 */
function getCurrentDate()
{
  return date('Ymd');
}

/**
 * 현재 시간 가져오기 (SQL Server 형식)
 * 
 * @return string HHMMSSmmm 형식의 시간
 */
function getCurrentTime()
{
  return date('Hisu');  // HH:mm:ss 형식
}

/**
 * SQL Server GETDATE() 함수와 동일한 결과 반환
 * 
 * @return string YYYY-MM-DD HH:MM:SS 형식
 */
function getDateTime()
{
  return date('Y-m-d H:i:s');
}

/**
 * 데이터베이스 연결 종료
 */
function closeConnection($pdo)
{
  $pdo = null;
}

// VB 코드