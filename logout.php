<?php

/**
 * 판매관리 시스템 - 로그아웃 처리
 * VB3 Main.frm의 Form_Unload 이벤트와 동일한 기능
 * 
 * 주요 기능:
 * - 세션 종료
 * - DB에 로그아웃 정보 업데이트
 * - 로그아웃 로그 기록
 */

session_start();

require_once '../../config/database.php';

// 로그인 체크
if (!isset($_SESSION['user_code'])) {
  header('Location: ../../login.php');
  exit;
}

// 세션에서 사용자 정보 가져오기
$user_code = $_SESSION['user_code'];

try {
  // 트랜잭션 시작 (VB의 BeginTrans와 동일)
  $pdo->beginTransaction();

  // 현재 날짜와 시간 가져오기 (VB의 GETDATE()와 동일)
  $sql_date = "SELECT CONVERT(VARCHAR(8), GETDATE(), 112) AS 로그아웃일자,
                        RIGHT(CONVERT(VARCHAR(23), GETDATE(), 121), 12) AS 로그아웃시각";

  $stmt = $pdo->query($sql_date);
  $datetime = $stmt->fetch(PDO::FETCH_ASSOC);

  // 날짜와 시간 포맷 변환 (VB 코드와 동일)
  $logout_date = $datetime['로그아웃일자'];
  $logout_time = str_replace(':', '', substr($datetime['로그아웃시각'], 0, 8)) .
    substr($datetime['로그아웃시각'], 9);
  $logout_datetime = $logout_date . $logout_time;

  // 사원 테이블 업데이트 (VB의 UPDATE 쿼리와 동일)
  $update_sql = "UPDATE 사원 
                   SET 로그인여부 = 'N',
                       로그아웃시각 = :logout_datetime
                   WHERE 사용자코드 = :user_code";

  $update_stmt = $pdo->prepare($update_sql);
  $update_stmt->bindParam(':logout_datetime', $logout_datetime, PDO::PARAM_STR);
  $update_stmt->bindParam(':user_code', $user_code, PDO::PARAM_STR);
  $update_stmt->execute();

  // 트랜잭션 커밋 (VB의 CommitTrans와 동일)
  $pdo->commit();

  // 세션 파괴
  session_unset();
  session_destroy();

  // 쿠키 삭제
  if (isset($_COOKIE[session_name()])) {
    setcookie(session_name(), '', time() - 3600, '/');
  }

  // 로그인 페이지로 리다이렉트
  header('Location: ../../login.php?logout=success');
  exit;
} catch (PDOException $e) {
  // 트랜잭션 롤백 (VB의 RollbackTrans와 동일)
  $pdo->rollBack();

  // 에러 로그 기록
  error_log('Logout Error: ' . $e->getMessage());

  // 세션은 어쨌든 종료
  session_unset();
  session_destroy();

  // 에러와 함께 로그인 페이지로
  header('Location: ../../login.php?error=logout_failed');
  exit;
}
