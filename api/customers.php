<?php

/**
 * 거래처 API
 * AJAX 요청을 처리하는 백엔드 API
 * 
 * 지원 메서드:
 * - GET: 거래처 조회
 * - POST: 거래처 등록/수정
 * - DELETE: 거래처 삭제
 */

session_start();

// 로그인 체크
if (!isset($_SESSION['user_code'])) {
  header('Content-Type: application/json; charset=utf-8');
  echo json_encode(['success' => false, 'message' => '로그인이 필요합니다.']);
  exit;
}

require_once '../config/database.php';

// JSON 응답 헤더 설정
header('Content-Type: application/json; charset=utf-8');

// 요청 메서드에 따라 처리
$method = $_SERVER['REQUEST_METHOD'];

try {
  switch ($method) {
    case 'GET':
      handleGet($pdo);
      break;
    case 'POST':
      handlePost($pdo);
      break;
    case 'DELETE':
      handleDelete($pdo);
      break;
    default:
      echo json_encode(['success' => false, 'message' => '지원하지 않는 메서드입니다.']);
  }
} catch (Exception $e) {
  error_log('Customer API Error: ' . $e->getMessage());
  echo json_encode(['success' => false, 'message' => '서버 오류가 발생했습니다.']);
}

/**
 * GET 요청 처리 (조회)
 */
function handleGet($pdo)
{
  // 엑셀 출력 요청인지 확인
  if (isset($_GET['export']) && $_GET['export'] === 'excel') {
    exportToExcel($pdo);
    return;
  }

  // 상세 조회 요청인지 확인
  if (isset($_GET['detail']) && $_GET['detail'] === 'true') {
    getCustomerDetail($pdo);
    return;
  }

  // 일반 목록 조회
  getCustomerList($pdo);
}

/**
 * 거래처 목록 조회
 */
function getCustomerList($pdo)
{
  $code = isset($_GET['code']) ? trim($_GET['code']) : '';
  $name = isset($_GET['name']) ? trim($_GET['name']) : '';
  $type = isset($_GET['type']) ? trim($_GET['type']) : '';

  // 동적 쿼리 생성
  $sql = "SELECT 
                거래처코드,
                거래처명,
                거래처구분,
                대표자,
                연락처,
                주소,
                사용여부
            FROM 거래처마스터
            WHERE 1=1";

  $params = [];

  // 거래처코드 검색 조건
  if (!empty($code)) {
    $sql .= " AND 거래처코드 LIKE :code";
    $params[':code'] = '%' . $code . '%';
  }

  // 거래처명 검색 조건
  if (!empty($name)) {
    $sql .= " AND 거래처명 LIKE :name";
    $params[':name'] = '%' . $name . '%';
  }

  // 거래처구분 검색 조건
  if (!empty($type)) {
    $sql .= " AND 거래처구분 = :type";
    $params[':type'] = $type;
  }

  // 정렬
  $sql .= " ORDER BY 거래처코드";

  try {
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    $customers = $stmt->fetchAll(PDO::FETCH_ASSOC);

    echo json_encode([
      'success' => true,
      'data' => $customers,
      'count' => count($customers)
    ], JSON_UNESCAPED_UNICODE);
  } catch (PDOException $e) {
    error_log('Customer List Query Error: ' . $e->getMessage());
    echo json_encode([
      'success' => false,
      'message' => '거래처 목록 조회 중 오류가 발생했습니다.'
    ], JSON_UNESCAPED_UNICODE);
  }
}

/**
 * 거래처 상세 조회
 */
function getCustomerDetail($pdo)
{
  $code = isset($_GET['code']) ? trim($_GET['code']) : '';

  if (empty($code)) {
    echo json_encode([
      'success' => false,
      'message' => '거래처코드가 필요합니다.'
    ], JSON_UNESCAPED_UNICODE);
    return;
  }

  $sql = "SELECT * FROM 거래처마스터 WHERE 거래처코드 = :code";

  try {
    $stmt = $pdo->prepare($sql);
    $stmt->bindParam(':code', $code, PDO::PARAM_STR);
    $stmt->execute();
    $customer = $stmt->fetch(PDO::FETCH_ASSOC);

    if ($customer) {
      echo json_encode([
        'success' => true,
        'data' => $customer
      ], JSON_UNESCAPED_UNICODE);
    } else {
      echo json_encode([
        'success' => false,
        'message' => '거래처를 찾을 수 없습니다.'
      ], JSON_UNESCAPED_UNICODE);
    }
  } catch (PDOException $e) {
    error_log('Customer Detail Query Error: ' . $e->getMessage());
    echo json_encode([
      'success' => false,
      'message' => '거래처 상세 조회 중 오류가 발생했습니다.'
    ], JSON_UNESCAPED_UNICODE);
  }
}

/**
 * POST 요청 처리 (등록/수정)
 */
function handlePost($pdo)
{
  $mode = isset($_POST['mode']) ? $_POST['mode'] : '';

  // CSRF token check
  if (!isset($_POST['csrf_token']) || !hash_equals($_SESSION['csrf_token'] ?? '', $_POST['csrf_token'])) {
    echo json_encode([
      'success' => false,
      'message' => 'Invalid CSRF token.'
    ], JSON_UNESCAPED_UNICODE);
    return;
  }
  if ($mode === 'add') {
    insertCustomer($pdo);
  } elseif ($mode === 'edit') {
    updateCustomer($pdo);
  } else {
    echo json_encode([
      'success' => false,
      'message' => '처리 모드가 올바르지 않습니다.'
    ], JSON_UNESCAPED_UNICODE);
  }
}

/**
 * 거래처 등록
 */
function insertCustomer($pdo)
{
  // 입력값 받기 및 검증
  $customer_code = trim($_POST['customer_code'] ?? '');
  $customer_name = trim($_POST['customer_name'] ?? '');
  $customer_type = trim($_POST['customer_type'] ?? '');

  // 필수 입력 체크
  if (empty($customer_code) || empty($customer_name) || empty($customer_type)) {
    echo json_encode([
      'success' => false,
      'message' => '필수 항목을 모두 입력해주세요.'
    ], JSON_UNESCAPED_UNICODE);
    return;
  }

  // 거래처코드 중복 체크
  $check_sql = "SELECT COUNT(*) FROM 거래처마스터 WHERE 거래처코드 = :code";
  $check_stmt = $pdo->prepare($check_sql);
  $check_stmt->bindParam(':code', $customer_code, PDO::PARAM_STR);
  $check_stmt->execute();

  if ($check_stmt->fetchColumn() > 0) {
    echo json_encode([
      'success' => false,
      'message' => '이미 사용 중인 거래처코드입니다.'
    ], JSON_UNESCAPED_UNICODE);
    return;
  }

  try {
    // 트랜잭션 시작
    $pdo->beginTransaction();

    // INSERT 쿼리
    $sql = "INSERT INTO 거래처마스터 (
                    거래처코드, 거래처명, 거래처구분, 사업자번호,
                    대표자, 연락처, 팩스, 이메일, 주소,
                    거래시작일, 담당자, 비고, 사용여부,
                    등록일자, 등록자
                ) VALUES (
                    :code, :name, :type, :business_number,
                    :representative, :phone, :fax, :email, :address,
                    :start_date, :manager, :memo, :use_yn,
                    CONVERT(VARCHAR(8), GETDATE(), 112), :user_code
                )";

    $stmt = $pdo->prepare($sql);

    // 파라미터 바인딩
    $stmt->bindParam(':code', $customer_code, PDO::PARAM_STR);
    $stmt->bindParam(':name', $customer_name, PDO::PARAM_STR);
    $stmt->bindParam(':type', $customer_type, PDO::PARAM_STR);
    $stmt->bindValue(':business_number', $_POST['business_number'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':representative', $_POST['representative'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':phone', $_POST['phone'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':fax', $_POST['fax'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':email', $_POST['email'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':address', $_POST['address'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':start_date', $_POST['start_date'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':manager', $_POST['manager'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':memo', $_POST['memo'] ?? null, PDO::PARAM_STR);
    $stmt->bindParam(':use_yn', $_POST['use_yn'], PDO::PARAM_STR);
    $stmt->bindParam(':user_code', $_SESSION['user_code'], PDO::PARAM_STR);

    $stmt->execute();

    // 트랜잭션 커밋
    $pdo->commit();

    echo json_encode([
      'success' => true,
      'message' => '거래처가 성공적으로 등록되었습니다.'
    ], JSON_UNESCAPED_UNICODE);
  } catch (PDOException $e) {
    // 트랜잭션 롤백
    $pdo->rollBack();

    error_log('Customer Insert Error: ' . $e->getMessage());
    echo json_encode([
      'success' => false,
      'message' => '거래처 등록 중 오류가 발생했습니다.'
    ], JSON_UNESCAPED_UNICODE);
  }
}

/**
 * 거래처 수정
 */
function updateCustomer($pdo)
{
  $original_code = trim($_POST['original_code'] ?? '');
  $customer_code = trim($_POST['customer_code'] ?? '');
  $customer_name = trim($_POST['customer_name'] ?? '');
  $customer_type = trim($_POST['customer_type'] ?? '');

  // 필수 입력 체크
  if (
    empty($original_code) || empty($customer_code) ||
    empty($customer_name) || empty($customer_type)
  ) {
    echo json_encode([
      'success' => false,
      'message' => '필수 항목을 모두 입력해주세요.'
    ], JSON_UNESCAPED_UNICODE);
    return;
  }

  // 거래처코드 변경 시 중복 체크
  if ($original_code !== $customer_code) {
    $check_sql = "SELECT COUNT(*) FROM 거래처마스터 WHERE 거래처코드 = :code";
    $check_stmt = $pdo->prepare($check_sql);
    $check_stmt->bindParam(':code', $customer_code, PDO::PARAM_STR);
    $check_stmt->execute();

    if ($check_stmt->fetchColumn() > 0) {
      echo json_encode([
        'success' => false,
        'message' => '이미 사용 중인 거래처코드입니다.'
      ], JSON_UNESCAPED_UNICODE);
      return;
    }
  }

  try {
    // 트랜잭션 시작
    $pdo->beginTransaction();

    // UPDATE 쿼리
    $sql = "UPDATE 거래처마스터 SET
                    거래처코드 = :code,
                    거래처명 = :name,
                    거래처구분 = :type,
                    사업자번호 = :business_number,
                    대표자 = :representative,
                    연락처 = :phone,
                    팩스 = :fax,
                    이메일 = :email,
                    주소 = :address,
                    거래시작일 = :start_date,
                    담당자 = :manager,
                    비고 = :memo,
                    사용여부 = :use_yn,
                    수정일자 = CONVERT(VARCHAR(8), GETDATE(), 112),
                    수정자 = :user_code
                WHERE 거래처코드 = :original_code";

    $stmt = $pdo->prepare($sql);

    // 파라미터 바인딩
    $stmt->bindParam(':code', $customer_code, PDO::PARAM_STR);
    $stmt->bindParam(':name', $customer_name, PDO::PARAM_STR);
    $stmt->bindParam(':type', $customer_type, PDO::PARAM_STR);
    $stmt->bindValue(':business_number', $_POST['business_number'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':representative', $_POST['representative'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':phone', $_POST['phone'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':fax', $_POST['fax'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':email', $_POST['email'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':address', $_POST['address'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':start_date', $_POST['start_date'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':manager', $_POST['manager'] ?? null, PDO::PARAM_STR);
    $stmt->bindValue(':memo', $_POST['memo'] ?? null, PDO::PARAM_STR);
    $stmt->bindParam(':use_yn', $_POST['use_yn'], PDO::PARAM_STR);
    $stmt->bindParam(':user_code', $_SESSION['user_code'], PDO::PARAM_STR);
    $stmt->bindParam(':original_code', $original_code, PDO::PARAM_STR);

    $stmt->execute();

    // 트랜잭션 커밋
    $pdo->commit();

    echo json_encode([
      'success' => true,
      'message' => '거래처가 성공적으로 수정되었습니다.'
    ], JSON_UNESCAPED_UNICODE);
  } catch (PDOException $e) {
    // 트랜잭션 롤백
    $pdo->rollBack();

    error_log('Customer Update Error: ' . $e->getMessage());
    echo json_encode([
      'success' => false,
      'message' => '거래처 수정 중 오류가 발생했습니다.'
    ], JSON_UNESCAPED_UNICODE);
  }
}

/**
 * DELETE 요청 처리 (삭제)
 */
function handleDelete($pdo)
{
  // DELETE 요청은 body에 데이터가 있으므로 파싱
  parse_str(file_get_contents("php://input"), $_DELETE);

  $customer_code = trim($_DELETE['customer_code'] ?? '');

  if (empty($customer_code)) {
    echo json_encode([
      'success' => false,
      'message' => '거래처코드가 필요합니다.'
    ], JSON_UNESCAPED_UNICODE);
    return;
  }

  try {
    // 거래 내역 체크 (매출전표)
    $check_sales_sql = "SELECT COUNT(*) FROM 매출전표 WHERE 거래처코드 = :code";
    $check_sales_stmt = $pdo->prepare($check_sales_sql);
    $check_sales_stmt->bindParam(':code', $customer_code, PDO::PARAM_STR);
  
    // CSRF token check for DELETE
    if (!isset($_DELETE['csrf_token']) || !hash_equals($_SESSION['csrf_token'] ?? '', $_DELETE['csrf_token'])) {
      echo json_encode([
        'success' => false,
        'message' => 'Invalid CSRF token.'
      ], JSON_UNESCAPED_UNICODE);
      return;
    }
    $check_sales_stmt->execute();

    if ($check_sales_stmt->fetchColumn() > 0) {
      echo json_encode([
        'success' => false,
        'message' => '매출 거래 내역이 있는 거래처는 삭제할 수 없습니다.'
      ], JSON_UNESCAPED_UNICODE);
      return;
    }

    // 거래 내역 체크 (매입전표)
    $check_purchase_sql = "SELECT COUNT(*) FROM 매입전표 WHERE 거래처코드 = :code";
    $check_purchase_stmt = $pdo->prepare($check_purchase_sql);
    $check_purchase_stmt->bindParam(':code', $customer_code, PDO::PARAM_STR);
    $check_purchase_stmt->execute();

    if ($check_purchase_stmt->fetchColumn() > 0) {
      echo json_encode([
        'success' => false,
        'message' => '매입 거래 내역이 있는 거래처는 삭제할 수 없습니다.'
      ], JSON_UNESCAPED_UNICODE);
      return;
    }

    // 트랜잭션 시작
    $pdo->beginTransaction();

    // DELETE 쿼리
    $sql = "DELETE FROM 거래처마스터 WHERE 거래처코드 = :code";
    $stmt = $pdo->prepare($sql);
    $stmt->bindParam(':code', $customer_code, PDO::PARAM_STR);
    $stmt->execute();

    if ($stmt->rowCount() > 0) {
      // 트랜잭션 커밋
      $pdo->commit();

      echo json_encode([
        'success' => true,
        'message' => '거래처가 성공적으로 삭제되었습니다.'
      ], JSON_UNESCAPED_UNICODE);
    } else {
      // 트랜잭션 롤백
      $pdo->rollBack();

      echo json_encode([
        'success' => false,
        'message' => '삭제할 거래처를 찾을 수 없습니다.'
      ], JSON_UNESCAPED_UNICODE);
    }
  } catch (PDOException $e) {
    // 트랜잭션 롤백
    if ($pdo->inTransaction()) {
      $pdo->rollBack();
    }

    error_log('Customer Delete Error: ' . $e->getMessage());
    echo json_encode([
      'success' => false,
      'message' => '거래처 삭제 중 오류가 발생했습니다.'
    ], JSON_UNESCAPED_UNICODE);
  }
}

/**
 * 엑셀 출력
 */
function exportToExcel($pdo)
{
  $code = isset($_GET['code']) ? trim($_GET['code']) : '';
  $name = isset($_GET['name']) ? trim($_GET['name']) : '';
  $type = isset($_GET['type']) ? trim($_GET['type']) : '';

  // 쿼리 생성
  $sql = "SELECT 
                거래처코드, 거래처명, 거래처구분, 사업자번호,
                대표자, 연락처, 팩스, 이메일, 주소,
                거래시작일, 담당자, 사용여부
            FROM 거래처마스터
            WHERE 1=1";

  $params = [];

  if (!empty($code)) {
    $sql .= " AND 거래처코드 LIKE :code";
    $params[':code'] = '%' . $code . '%';
  }

  if (!empty($name)) {
    $sql .= " AND 거래처명 LIKE :name";
    $params[':name'] = '%' . $name . '%';
  }

  if (!empty($type)) {
    $sql .= " AND 거래처구분 = :type";
    $params[':type'] = $type;
  }

  $sql .= " ORDER BY 거래처코드";

  try {
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    $customers = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // CSV 헤더 설정
    header('Content-Type: text/csv; charset=UTF-8');
    header('Content-Disposition: attachment; filename="거래처목록_' . date('Ymd') . '.csv"');

    // UTF-8 BOM 추가 (엑셀에서 한글 깨짐 방지)
    echo "\xEF\xBB\xBF";

    // CSV 출력
    $output = fopen('php://output', 'w');

    // 헤더 행
    fputcsv($output, [
      '거래처코드',
      '거래처명',
      '거래처구분',
      '사업자번호',
      '대표자',
      '연락처',
      '팩스',
      '이메일',
      '주소',
      '거래시작일',
      '담당자',
      '사용여부'
    ]);

    // 데이터 행
    foreach ($customers as $customer) {
      // Sanitize fields to prevent CSV formula injection (Excel)
      $safeRow = array_map(function ($value) {
        if ($value === null) return '';
        $value = (string)$value;
        if (preg_match('/^[=+\-@]/', $value)) {
          return "'" . $value; // prefix with single quote to neutralize formulas
        }
        return $value;
      }, $customer);

      fputcsv($output, $safeRow);
    }

    fclose($output);
  } catch (PDOException $e) {
    error_log('Excel Export Error: ' . $e->getMessage());
    echo '엑셀 출력 중 오류가 발생했습니다.';
  }
}
