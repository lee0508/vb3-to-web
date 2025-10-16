<?php
/**
 * 판매관리 시스템 - 로그인 페이지
 *
 * 주요 기능:
 * - 사용자 인증
 * - 로그인 정보 DB 저장 (사원 테이블)
 * - 작업일자 설정
 */

session_start();

// 이미 로그인되어 있으면 메인 페이지로 리다이렉트
if (isset($_SESSION['user_code'])) {
  header('Location: index.php');
  exit;
}

// 데이터베이스 설정 파일 포함
require_once 'config/database.php';

// 로그인 처리
$error_message = '';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
  $user_code = trim($_POST['user_code'] ?? '');
  $password = trim($_POST['password'] ?? '');

  // 입력값 검증
  if (empty($user_code) || empty($password)) {
    $error_message = '사용자코드와 비밀번호를 입력해주세요.';
  } else {
    try {
      // PDO를 사용한 안전한 쿼리 (SQL Injection 방지)
      $sql = "SELECT 
                        사용자코드,
                        사용자명,
                        비밀번호,
                        권한,
                        지점명,
                        작업일자,
                        로그인여부
                    FROM 사원
                    WHERE 사용자코드 = :user_code
                    AND 사용여부 = 'Y'";

      $stmt = $pdo->prepare($sql);
      $stmt->bindParam(':user_code', $user_code, PDO::PARAM_STR);
      $stmt->execute();

      $user = $stmt->fetch(PDO::FETCH_ASSOC);

      // 사용자 존재 여부 및 비밀번호 확인
      if ($user && password_verify($password, $user['비밀번호'])) {

        // 이미 로그인된 사용자 체크
        if ($user['로그인여부'] === 'Y') {
          $error_message = '이미 다른 곳에서 로그인되어 있습니다.';
        } else {
          // 로그인 성공 - 세션 생성
            // Regenerate session id to prevent session fixation
            if (function_exists('session_regenerate_id')) {
              session_regenerate_id(true);
            }

            // CSRF token for subsequent AJAX/state-changing requests
            if (!isset($_SESSION['csrf_token'])) {
              try {
                $_SESSION['csrf_token'] = bin2hex(random_bytes(32));
              } catch (Exception $e) {
                $_SESSION['csrf_token'] = bin2hex(openssl_random_pseudo_bytes(32));
              }
            }

            $_SESSION['user_code'] = $user['사용자코드'];
            $_SESSION['user_name'] = $user['사용자명'];
            $_SESSION['user_authority'] = $user['권한'];
            $_SESSION['user_branch'] = $user['지점명'];
            $_SESSION['work_date'] = $user['작업일자'];
            $_SESSION['login_time'] = date('Y-m-d H:i:s');

          // DB에 로그인 정보 업데이트
          $update_sql = "UPDATE 사원 
                                   SET 로그인여부 = 'Y',
                                       로그인시각 = GETDATE()
                                   WHERE 사용자코드 = :user_code";

          $update_stmt = $pdo->prepare($update_sql);
          $update_stmt->bindParam(':user_code', $user_code, PDO::PARAM_STR);
          $update_stmt->execute();

          // 로그인 로그 기록 (선택사항)
          $log_sql = "INSERT INTO 로그인로그 (사용자코드, 로그인시각, IP주소)
                                VALUES (:user_code, GETDATE(), :ip_address)";

          $log_stmt = $pdo->prepare($log_sql);
          $log_stmt->bindParam(':user_code', $user_code, PDO::PARAM_STR);
          $log_stmt->bindParam(':ip_address', $_SERVER['REMOTE_ADDR'], PDO::PARAM_STR);
          $log_stmt->execute();

          // 메인 페이지로 리다이렉트
          header('Location: index.php');
          exit;
        }
      } else {
        $error_message = '사용자코드 또는 비밀번호가 일치하지 않습니다.';
      }
    } catch (PDOException $e) {
      // 에러 로그 기록 (실제 운영시에는 파일이나 DB에 기록)
      error_log('Login Error: ' . $e->getMessage());
      $error_message = '시스템 오류가 발생했습니다. 관리자에게 문의하세요.';
    }
  }
}
?>
<!DOCTYPE html>
<html lang="ko">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>판매관리 시스템 - 로그인</title>

  <!-- Bootstrap 5 CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">

  <style>
    /* 로그인 페이지 스타일 */
    body {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
    }

    .login-container {
      background: white;
      border-radius: 15px;
      box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2);
      padding: 40px;
      width: 100%;
      max-width: 400px;
    }

    .login-header {
      text-align: center;
      margin-bottom: 30px;
    }

    .login-header h2 {
      color: #333;
      font-weight: 700;
      margin-bottom: 10px;
    }

    .login-header p {
      color: #666;
      font-size: 14px;
    }

    .form-control:focus {
      border-color: #667eea;
      box-shadow: 0 0 0 0.2rem rgba(102, 126, 234, 0.25);
    }

    .btn-login {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      border: none;
      padding: 12px;
      font-weight: 600;
      transition: transform 0.2s;
    }

    .btn-login:hover {
      transform: translateY(-2px);
      box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
    }

    .alert {
      font-size: 14px;
    }
  </style>
</head>

<body>
  <div class="login-container">
    <!-- 로그인 헤더 -->
    <div class="login-header">
      <h2>판매관리 시스템</h2>
      <p>Ver 5.1.26a (Web Edition)</p>
    </div>

    <!-- 에러 메시지 표시 -->
    <?php if (!empty($error_message)): ?>
      <div class="alert alert-danger alert-dismissible fade show" role="alert">
        <i class="bi bi-exclamation-triangle-fill"></i>
        <?php echo htmlspecialchars($error_message); ?>
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
      </div>
    <?php endif; ?>

    <!-- 로그인 폼 -->
    <form method="POST" action="login.php" id="loginForm">
      <!-- 사용자코드 입력 -->
      <div class="mb-3">
        <label for="user_code" class="form-label">사용자코드</label>
        <input type="text"
          class="form-control"
          id="user_code"
          name="user_code"
          placeholder="사용자코드를 입력하세요"
          required
          autofocus
          value="<?php echo htmlspecialchars($_POST['user_code'] ?? ''); ?>">
      </div>

      <!-- 비밀번호 입력 -->
      <div class="mb-3">
        <label for="password" class="form-label">비밀번호</label>
        <input type="password"
          class="form-control"
          id="password"
          name="password"
          placeholder="비밀번호를 입력하세요"
          required>
      </div>

      <!-- 로그인 유지 체크박스 (선택사항) -->
      <div class="mb-3 form-check">
        <input type="checkbox" class="form-check-input" id="remember" name="remember">
        <label class="form-check-label" for="remember">
          로그인 상태 유지
        </label>
      </div>

      <!-- 로그인 버튼 -->
      <button type="submit" class="btn btn-primary btn-login w-100">
        로그인
      </button>
    </form>

    <!-- 추가 링크 -->
    <div class="mt-3 text-center">
      <small class="text-muted">
        비밀번호를 잊으셨나요?
        <a href="password_reset.php" class="text-decoration-none">비밀번호 재설정</a>
      </small>
    </div>
  </div>

  <!-- Bootstrap 5 JS -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

  <script>
    // 폼 유효성 검사 (클라이언트 측)
    document.getElementById('loginForm').addEventListener('submit', function(e) {
      const userCode = document.getElementById('user_code').value.trim();
      const password = document.getElementById('password').value.trim();

      if (userCode === '' || password === '') {
        e.preventDefault();
        alert('사용자코드와 비밀번호를 모두 입력해주세요.');
        return false;
      }
    });

    // Enter 키로 로그인
    document.getElementById('password').addEventListener('keypress', function(e) {
      if (e.key === 'Enter') {
        document.getElementById('loginForm').submit();
      }
    });
  </script>
</body>

</html>