<?php

/**
 * 판매관리 시스템 - 메인 대시보드
 * VB3 Main.frm의 메인 화면을 웹으로 구현
 * 
 * 주요 기능:
 * - 메뉴 네비게이션
 * - 사용자 정보 표시
 * - 작업일자 표시
 * - 권한별 메뉴 제어
 */

session_start();

// 로그인 체크
if (!isset($_SESSION['user_code'])) {
  header('Location: login.php');
  exit;
}

require_once 'config/database.php';

// 세션에서 사용자 정보 가져오기
$user_code = $_SESSION['user_code'];
$user_name = $_SESSION['user_name'];
$user_authority = $_SESSION['user_authority'];
$user_branch = $_SESSION['user_branch'];
$work_date = $_SESSION['work_date'];

// 작업일자 형식 변환 (YYYYMMDD -> YYYY-MM-DD)
$formatted_work_date = substr($work_date, 0, 4) . '-' .
  substr($work_date, 4, 2) . '-' .
  substr($work_date, 6, 2);
?>
<!DOCTYPE html>
<html lang="ko">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>판매관리 시스템 Ver 5.1.26a</title>

  <!-- Bootstrap 5 CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">

  <!-- Bootstrap Icons -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css" rel="stylesheet">

  <style>
    /* 전체 레이아웃 스타일 */
    body {
      font-family: 'Malgun Gothic', sans-serif;
      background-color: #f5f5f5;
    }

    /* 상단 네비게이션 바 */
    .navbar {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
    }

    .navbar-brand {
      font-weight: 700;
      color: white !important;
      font-size: 1.3rem;
    }

    /* 사이드바 메뉴 */
    .sidebar {
      min-height: calc(100vh - 56px);
      background: white;
      box-shadow: 2px 0 10px rgba(0, 0, 0, 0.05);
      padding: 20px 0;
    }

    .sidebar .nav-link {
      color: #333;
      padding: 12px 20px;
      border-left: 3px solid transparent;
      transition: all 0.3s;
    }

    .sidebar .nav-link:hover {
      background-color: #f8f9fa;
      border-left-color: #667eea;
      color: #667eea;
    }

    .sidebar .nav-link.active {
      background-color: #e7f1ff;
      border-left-color: #667eea;
      color: #667eea;
      font-weight: 600;
    }

    .sidebar .nav-link i {
      margin-right: 10px;
      width: 20px;
    }

    /* 메인 컨텐츠 영역 */
    .main-content {
      padding: 30px;
    }

    /* 대시보드 카드 */
    .dashboard-card {
      background: white;
      border-radius: 10px;
      padding: 25px;
      margin-bottom: 20px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
      transition: transform 0.2s;
    }

    .dashboard-card:hover {
      transform: translateY(-5px);
      box-shadow: 0 5px 20px rgba(0, 0, 0, 0.1);
    }

    .dashboard-card .icon {
      font-size: 3rem;
      margin-bottom: 15px;
    }

    .dashboard-card h5 {
      color: #333;
      font-weight: 600;
      margin-bottom: 10px;
    }

    .dashboard-card .value {
      font-size: 2rem;
      font-weight: 700;
      color: #667eea;
    }

    /* 상태바 (VB의 StatusBar와 동일) */
    .status-bar {
      background-color: #f8f9fa;
      border-top: 1px solid #dee2e6;
      padding: 10px 20px;
      position: fixed;
      bottom: 0;
      left: 0;
      right: 0;
      display: flex;
      justify-content: space-between;
      font-size: 0.9rem;
      z-index: 1000;
    }

    .status-item {
      padding: 0 15px;
      border-right: 1px solid #dee2e6;
    }

    .status-item:last-child {
      border-right: none;
    }

    /* 드롭다운 메뉴 스타일 */
    .dropdown-menu {
      box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
    }

    /* 메인 컨텐츠 하단 여백 (상태바 공간 확보) */
    .content-wrapper {
      margin-bottom: 60px;
    }

    /* 퀵 액세스 버튼 */
    .quick-access-btn {
      width: 100%;
      padding: 15px;
      text-align: left;
      border-radius: 8px;
      border: 1px solid #e0e0e0;
      background: white;
      margin-bottom: 10px;
      transition: all 0.2s;
    }

    .quick-access-btn:hover {
      border-color: #667eea;
      background-color: #f8f9ff;
      transform: translateX(5px);
    }
  </style>
</head>

<body>
  <!-- 상단 네비게이션 바 -->
  <nav class="navbar navbar-expand-lg navbar-dark">
    <div class="container-fluid">
      <a class="navbar-brand" href="index.php">
        <i class="bi bi-shop"></i> 판매관리 시스템
      </a>

      <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav">
        <span class="navbar-toggler-icon"></span>
      </button>

      <div class="collapse navbar-collapse" id="navbarNav">
        <ul class="navbar-nav ms-auto">
          <!-- 사용자 정보 드롭다운 -->
          <li class="nav-item dropdown">
            <a class="nav-link dropdown-toggle text-white" href="#" id="userDropdown"
              data-bs-toggle="dropdown">
              <i class="bi bi-person-circle"></i>
              <?php echo htmlspecialchars($user_name); ?>
            </a>
            <ul class="dropdown-menu dropdown-menu-end">
              <li>
                <span class="dropdown-item-text">
                  <strong>사용자:</strong> <?php echo htmlspecialchars($user_code); ?>
                </span>
              </li>
              <li>
                <span class="dropdown-item-text">
                  <strong>지점:</strong> <?php echo htmlspecialchars($user_branch); ?>
                </span>
              </li>
              <li>
                <hr class="dropdown-divider">
              </li>
              <li><a class="dropdown-item" href="modules/auth/change_password.php">
                  <i class="bi bi-key"></i> 비밀번호 변경
                </a></li>
              <li><a class="dropdown-item" href="modules/auth/logout.php">
                  <i class="bi bi-box-arrow-right"></i> 로그아웃
                </a></li>
            </ul>
          </li>
        </ul>
      </div>
    </div>
  </nav>

  <!-- 메인 컨테이너 -->
  <div class="container-fluid">
    <div class="row">
      <!-- 사이드바 메뉴 (VB의 메뉴바와 동일) -->
      <div class="col-md-2 sidebar">
        <nav class="nav flex-column">
          <!-- 1. 기초자료관리 -->
          <a class="nav-link" data-bs-toggle="collapse" href="#menu1">
            <i class="bi bi-database"></i> 기초자료관리
            <i class="bi bi-chevron-down float-end"></i>
          </a>
          <div class="collapse" id="menu1">
            <a class="nav-link ps-5" href="modules/master/customer.php">거래처관리</a>
            <a class="nav-link ps-5" href="modules/master/employee.php">사원관리</a>
            <a class="nav-link ps-5" href="modules/master/product.php">제품관리</a>
            <a class="nav-link ps-5" href="modules/master/category.php">제품분류</a>
            <a class="nav-link ps-5" href="modules/master/code.php">코드관리</a>
          </div>

          <!-- 2. 거래처리(입금관리) -->
          <a class="nav-link" data-bs-toggle="collapse" href="#menu2">
            <i class="bi bi-cash-coin"></i> 거래처리(입금관리)
            <i class="bi bi-chevron-down float-end"></i>
          </a>
          <div class="collapse" id="menu2">
            <a class="nav-link ps-5" href="modules/payment/receive_bill.php">매출입금관리서 입력</a>
            <a class="nav-link ps-5" href="modules/payment/receive_list.php">매출입금관리서 조회</a>
            <a class="nav-link ps-5" href="modules/payment/receive_detail.php">거래처 입금처리</a>
            <a class="nav-link ps-5" href="modules/payment/customer_status.php">거래처 현황</a>
            <hr>
            <a class="nav-link ps-5" href="modules/payment/pay_bill.php">매입세금관리서 입력</a>
            <a class="nav-link ps-5" href="modules/payment/pay_list.php">매입세금관리서 조회</a>
            <a class="nav-link ps-5" href="modules/payment/pay_detail.php">거래처 출금처리</a>
            <a class="nav-link ps-5" href="modules/payment/bill_classify.php">관리서분류처리</a>
            <a class="nav-link ps-5" href="modules/payment/bill_register.php">전자관리서등록(NO)</a>
            <a class="nav-link ps-5" href="modules/payment/bill_inquiry.php">관리서조회</a>
          </div>

          <!-- 3. 매출관리 -->
          <a class="nav-link" data-bs-toggle="collapse" href="#menu3">
            <i class="bi bi-cart-check"></i> 매출관리
            <i class="bi bi-chevron-down float-end"></i>
          </a>
          <div class="collapse" id="menu3">
            <a class="nav-link ps-5" href="modules/sales/order_form.php">매출전표입력</a>
            <a class="nav-link ps-5" href="modules/sales/order_cancel.php">거래처리취소</a>
            <a class="nav-link ps-5" href="modules/sales/order_list.php">매출전표조회</a>
            <hr>
            <a class="nav-link ps-5" href="modules/sales/quote_form.php">견적서작성</a>
            <a class="nav-link ps-5" href="modules/sales/quote_list.php">견적서조회</a>
            <a class="nav-link ps-5" href="modules/sales/quote_to_order.php">견적서 매출전표처리</a>
            <hr>
            <a class="nav-link ps-5" href="modules/sales/product_manage.php">매출전표 제품처리</a>
            <hr>
            <a class="nav-link ps-5" href="modules/sales/customer_balance.php">거래처 잔액조회</a>
            <a class="nav-link ps-5" href="modules/sales/product_sales.php">품목별 매출조회</a>
          </div>

          <!-- 4. 매입관리 -->
          <a class="nav-link" data-bs-toggle="collapse" href="#menu4">
            <i class="bi bi-box-seam"></i> 매입관리
            <i class="bi bi-chevron-down float-end"></i>
          </a>
          <div class="collapse" id="menu4">
            <a class="nav-link ps-5" href="modules/purchase/purchase_form.php">매입전표입력</a>
            <a class="nav-link ps-5" href="modules/purchase/purchase_cancel.php">거래처리취소</a>
            <a class="nav-link ps-5" href="modules/purchase/purchase_list.php">매입전표조회</a>
            <hr>
            <a class="nav-link ps-5" href="modules/purchase/order_form.php">발주서작성</a>
            <a class="nav-link ps-5" href="modules/purchase/order_list.php">발주서조회</a>
            <a class="nav-link ps-5" href="modules/purchase/order_to_purchase.php">발주서 매입전표처리</a>
            <hr>
            <a class="nav-link ps-5" href="modules/purchase/product_manage.php">매입전표 제품처리</a>
            <hr>
            <a class="nav-link ps-5" href="modules/purchase/supplier_balance.php">거래처 잔액조회</a>
            <a class="nav-link ps-5" href="modules/purchase/product_purchase.php">품목별 매입조회</a>
          </div>

          <!-- 5. 회계관리 -->
          <a class="nav-link" data-bs-toggle="collapse" href="#menu5">
            <i class="bi bi-calculator"></i> 회계관리
            <i class="bi bi-chevron-down float-end"></i>
          </a>
          <div class="collapse" id="menu5">
            <a class="nav-link ps-5" href="modules/accounting/year_end.php">연말정산관리</a>
            <a class="nav-link ps-5" href="modules/accounting/fund_mgmt.php">자금관리</a>
          </div>

          <!-- 6. 재고관리 -->
          <a class="nav-link" data-bs-toggle="collapse" href="#menu6">
            <i class="bi bi-boxes"></i> 재고관리
            <i class="bi bi-chevron-down float-end"></i>
          </a>
          <div class="collapse" id="menu6">
            <a class="nav-link ps-5" href="modules/inventory/stock_list.php">재고조회</a>
            <hr>
            <a class="nav-link ps-5" href="modules/inventory/product_move.php">이동상품</a>
            <a class="nav-link ps-5" href="modules/inventory/product_loss.php">폐기손</a>
            <hr>
            <a class="nav-link ps-5" href="modules/inventory/adjustment.php">재고조정</a>
          </div>

          <!-- 8. 통계분석 -->
          <a class="nav-link" data-bs-toggle="collapse" href="#menu8">
            <i class="bi bi-graph-up"></i> 통계분석
            <i class="bi bi-chevron-down float-end"></i>
          </a>
          <div class="collapse" id="menu8">
            <a class="nav-link ps-5" href="modules/statistics/sales_report.php">매출통계분석</a>
            <a class="nav-link ps-5" href="modules/statistics/purchase_report.php">매입/발주통계</a>
            <hr>
            <a class="nav-link ps-5" href="modules/statistics/monthly.php">월별거래현황</a>
            <hr>
            <a class="nav-link ps-5" href="modules/statistics/closing.php">마감거래작업</a>
            <hr>
            <a class="nav-link ps-5" href="modules/statistics/backup.php">자료백업</a>
          </div>

          <!-- 9. 환경설정 -->
          <a class="nav-link" href="modules/settings/system.php">
            <i class="bi bi-gear"></i> 환경설정
          </a>
        </nav>
      </div>

      <!-- 메인 컨텐츠 영역 -->
      <div class="col-md-10 main-content content-wrapper">
        <h2 class="mb-4">
          <i class="bi bi-house-door"></i> 대시보드
        </h2>

        <!-- 통계 카드 -->
        <div class="row">
          <!-- 오늘의 매출 -->
          <div class="col-md-3">
            <div class="dashboard-card text-center">
              <i class="bi bi-currency-dollar icon text-success"></i>
              <h5>오늘의 매출</h5>
              <div class="value" id="todaySales">-</div>
              <small class="text-muted">원</small>
            </div>
          </div>

          <!-- 오늘의 매입 -->
          <div class="col-md-3">
            <div class="dashboard-card text-center">
              <i class="bi bi-cart icon text-primary"></i>
              <h5>오늘의 매입</h5>
              <div class="value" id="todayPurchase">-</div>
              <small class="text-muted">원</small>
            </div>
          </div>

          <!-- 미수금 -->
          <div class="col-md-3">
            <div class="dashboard-card text-center">
              <i class="bi bi-exclamation-circle icon text-warning"></i>
              <h5>미수금</h5>
              <div class="value" id="receivable">-</div>
              <small class="text-muted">원</small>
            </div>
          </div>

          <!-- 재고금액 -->
          <div class="col-md-3">
            <div class="dashboard-card text-center">
              <i class="bi bi-box icon text-info"></i>
              <h5>재고금액</h5>
              <div class="value" id="stockValue">-</div>
              <small class="text-muted">원</small>
            </div>
          </div>
        </div>

        <!-- 빠른 실행 메뉴 -->
        <div class="row mt-4">
          <div class="col-md-6">
            <div class="dashboard-card">
              <h5 class="mb-3">
                <i class="bi bi-lightning"></i> 빠른 실행
              </h5>
              <button class="quick-access-btn" onclick="location.href='modules/sales/order_form.php'">
                <i class="bi bi-plus-circle text-success"></i>
                <strong>매출전표 입력</strong>
              </button>
              <button class="quick-access-btn" onclick="location.href='modules/purchase/purchase_form.php'">
                <i class="bi bi-plus-circle text-primary"></i>
                <strong>매입전표 입력</strong>
              </button>
              <button class="quick-access-btn" onclick="location.href='modules/inventory/stock_list.php'">
                <i class="bi bi-search text-info"></i>
                <strong>재고 조회</strong>
              </button>
              <button class="quick-access-btn" onclick="location.href='modules/master/customer.php'">
                <i class="bi bi-people text-warning"></i>
                <strong>거래처 관리</strong>
              </button>
            </div>
          </div>

          <!-- 최근 거래 내역 -->
          <div class="col-md-6">
            <div class="dashboard-card">
              <h5 class="mb-3">
                <i class="bi bi-clock-history"></i> 최근 거래 내역
              </h5>
              <div id="recentTransactions">
                <div class="text-center text-muted py-4">
                  <i class="bi bi-inbox" style="font-size: 2rem;"></i>
                  <p>최근 거래 내역을 불러오는 중...</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- 상태바 (VB의 StatusBar와 동일 역할) -->
  <div class="status-bar">
    <div class="status-item">
      <i class="bi bi-calendar3"></i>
      <strong>작업일자:</strong> <?php echo $formatted_work_date; ?>
    </div>
    <div class="status-item">
      <i class="bi bi-person"></i>
      <strong>사용자:</strong> <?php echo htmlspecialchars($user_name); ?>
    </div>
    <div class="status-item">
      <i class="bi bi-building"></i>
      <strong>지점:</strong> <?php echo htmlspecialchars($user_branch); ?>
    </div>
    <div class="status-item flex-fill text-end">
      <i class="bi bi-clock"></i>
      <span id="currentTime"></span>
    </div>
  </div>

  <!-- Bootstrap 5 JS -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

  <!-- jQuery (AJAX 통신용) -->
  <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>

  <script>
    /**
     * 페이지 로드 시 실행
     */
    $(document).ready(function() {
      // 대시보드 데이터 로드
      loadDashboardData();

      // 최근 거래 내역 로드
      loadRecentTransactions();

      // 현재 시간 업데이트 (1초마다)
      updateCurrentTime();
      setInterval(updateCurrentTime, 1000);
    });

    /**
     * 대시보드 통계 데이터 로드
     */
    function loadDashboardData() {
      $.ajax({
        url: 'api/dashboard_stats.php',
        method: 'GET',
        dataType: 'json',
        success: function(data) {
          if (data.success) {
            // 숫자를 천단위 콤마로 포맷팅
            $('#todaySales').text(formatNumber(data.todaySales));
            $('#todayPurchase').text(formatNumber(data.todayPurchase));
            $('#receivable').text(formatNumber(data.receivable));
            $('#stockValue').text(formatNumber(data.stockValue));
          } else {
            console.error('데이터 로드 실패:', data.message);
          }
        },
        error: function(xhr, status, error) {
          console.error('AJAX 에러:', error);
          // 에러 시 기본값 표시
          $('.value').text('0');
        }
      });
    }

    /**
     * 최근 거래 내역 로드
     */
    function loadRecentTransactions() {
      $.ajax({
        url: 'api/recent_transactions.php',
        method: 'GET',
        dataType: 'json',
        success: function(data) {
          if (data.success && data.transactions.length > 0) {
            let html = '<div class="list-group list-group-flush">';
            data.transactions.forEach(function(trans) {
              html += `
                                <div class="list-group-item">
                                    <div class="d-flex justify-content-between">
                                        <div>
                                            <strong>${trans.customerName}</strong>
                                            <br>
                                            <small class="text-muted">${trans.type}</small>
                                        </div>
                                        <div class="text-end">
                                            <strong class="${trans.type === '매출' ? 'text-success' : 'text-primary'}">
                                                ${formatNumber(trans.amount)}원
                                            </strong>
                                            <br>
                                            <small class="text-muted">${trans.date}</small>
                                        </div>
                                    </div>
                                </div>
                            `;
            });
            html += '</div>';
            $('#recentTransactions').html(html);
          } else {
            $('#recentTransactions').html(`
                            <div class="text-center text-muted py-4">
                                <i class="bi bi-inbox" style="font-size: 2rem;"></i>
                                <p>최근 거래 내역이 없습니다.</p>
                            </div>
                        `);
          }
        },
        error: function() {
          $('#recentTransactions').html(`
                        <div class="alert alert-warning">
                            거래 내역을 불러올 수 없습니다.
                        </div>
                    `);
        }
      });
    }

    /**
     * 숫자를 천단위 콤마 형식으로 변환
     */
    function formatNumber(num) {
      return Number(num).toLocaleString('ko-KR');
    }

    /**
     * 현재 시간 업데이트
     */
    function updateCurrentTime() {
      const now = new Date();
      const hours = String(now.getHours()).padStart(2, '0');
      const minutes = String(now.getMinutes()).padStart(2, '0');
      const seconds = String(now.getSeconds()).padStart(2, '0');
      $('#currentTime').text(`${hours}:${minutes}:${seconds}`);
    }
  </script>
</body>

</html>