<?php

/**
 * 판매관리 시스템 - 거래처 관리
 * VB3의 거래처마스터 폼을 웹으로 구현
 * 
 * 주요 기능:
 * - 거래처 목록 조회
 * - 거래처 등록/수정/삭제
 * - 검색 기능
 * - AJAX를 이용한 실시간 처리
 */

session_start();

// 로그인 체크
if (!isset($_SESSION['user_code'])) {
  header('Location: ../../login.php');
  exit;
}

require_once '../../config/database.php';

// 세션 정보
$user_name = $_SESSION['user_name'];
$user_branch = $_SESSION['user_branch'];
$work_date = $_SESSION['work_date'];

// 작업일자 포맷팅
$formatted_work_date = substr($work_date, 0, 4) . '-' .
  substr($work_date, 4, 2) . '-' .
  substr($work_date, 6, 2);
?>
<!DOCTYPE html>
<html lang="ko">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>거래처 관리 - 판매관리 시스템</title>

  <!-- Bootstrap 5 CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">

  <!-- Bootstrap Icons -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css" rel="stylesheet">

  <style>
    body {
      font-family: 'Malgun Gothic', sans-serif;
      background-color: #f5f5f5;
    }

    /* 페이지 헤더 */
    .page-header {
      background: white;
      padding: 20px;
      margin-bottom: 20px;
      border-radius: 10px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
    }

    .page-header h3 {
      color: #333;
      font-weight: 700;
      margin: 0;
    }

    /* 검색 영역 */
    .search-area {
      background: white;
      padding: 20px;
      margin-bottom: 20px;
      border-radius: 10px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
    }

    /* 버튼 그룹 */
    .btn-group-custom {
      gap: 10px;
    }

    .btn-custom {
      padding: 10px 20px;
      border-radius: 5px;
      font-weight: 600;
    }

    /* 테이블 영역 */
    .table-area {
      background: white;
      padding: 20px;
      border-radius: 10px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
    }

    .table thead {
      background-color: #667eea;
      color: white;
    }

    .table tbody tr {
      cursor: pointer;
      transition: background-color 0.2s;
    }

    .table tbody tr:hover {
      background-color: #f8f9ff;
    }

    .table tbody tr.selected {
      background-color: #e7f1ff;
    }

    /* 모달 스타일 */
    .modal-header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
    }

    /* 상태바 */
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

    .content-wrapper {
      margin-bottom: 60px;
    }

    /* 필수 입력 표시 */
    .required::after {
      content: " *";
      color: red;
    }
  </style>
</head>

<body>
  <div class="container-fluid content-wrapper">
    <!-- 페이지 헤더 -->
    <div class="page-header d-flex justify-content-between align-items-center">
      <div>
        <h3>
          <i class="bi bi-people-fill"></i> 거래처 관리
        </h3>
        <small class="text-muted">거래처 정보를 등록, 수정, 삭제할 수 있습니다.</small>
      </div>
      <div>
        <button type="button" class="btn btn-secondary" onclick="location.href='../../index.php'">
          <i class="bi bi-arrow-left"></i> 메인으로
        </button>
      </div>
    </div>

    <!-- 검색 영역 -->
    <div class="search-area">
      <div class="row g-3">
        <div class="col-md-3">
          <label class="form-label">거래처코드</label>
          <input type="text" class="form-control" id="search_code" placeholder="거래처코드">
        </div>
        <div class="col-md-3">
          <label class="form-label">거래처명</label>
          <input type="text" class="form-control" id="search_name" placeholder="거래처명">
        </div>
        <div class="col-md-3">
          <label class="form-label">거래처구분</label>
          <select class="form-select" id="search_type">
            <option value="">전체</option>
            <option value="매출처">매출처</option>
            <option value="매입처">매입처</option>
            <option value="양쪽">양쪽</option>
          </select>
        </div>
        <div class="col-md-3">
          <label class="form-label">&nbsp;</label>
          <div class="d-grid">
            <button type="button" class="btn btn-primary" onclick="searchCustomers()">
              <i class="bi bi-search"></i> 검색
            </button>
          </div>
        </div>
      </div>
    </div>

    <!-- 버튼 그룹 -->
    <div class="mb-3 d-flex justify-content-between">
      <div class="btn-group-custom d-flex">
        <button type="button" class="btn btn-success btn-custom" onclick="showAddModal()">
          <i class="bi bi-plus-circle"></i> 신규등록
        </button>
        <button type="button" class="btn btn-warning btn-custom" onclick="showEditModal()">
          <i class="bi bi-pencil-square"></i> 수정
        </button>
        <button type="button" class="btn btn-danger btn-custom" onclick="deleteCustomer()">
          <i class="bi bi-trash"></i> 삭제
        </button>
      </div>
      <div>
        <button type="button" class="btn btn-info btn-custom" onclick="exportToExcel()">
          <i class="bi bi-file-earmark-excel"></i> 엑셀 출력
        </button>
      </div>
    </div>

    <!-- 거래처 목록 테이블 -->
    <div class="table-area">
      <div class="mb-3">
        <strong>전체 거래처: <span id="total_count">0</span>건</strong>
      </div>
      <div class="table-responsive">
        <table class="table table-hover table-bordered" id="customerTable">
          <thead>
            <tr>
              <th width="10%">거래처코드</th>
              <th width="20%">거래처명</th>
              <th width="10%">구분</th>
              <th width="15%">대표자</th>
              <th width="15%">연락처</th>
              <th width="20%">주소</th>
              <th width="10%">사용여부</th>
            </tr>
          </thead>
          <tbody id="customerTableBody">
            <tr>
              <td colspan="7" class="text-center py-5">
                <i class="bi bi-inbox" style="font-size: 3rem; color: #ccc;"></i>
                <p class="text-muted mt-3">검색 버튼을 클릭하여 거래처를 조회하세요.</p>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- 거래처 등록/수정 모달 -->
  <div class="modal fade" id="customerModal" tabindex="-1">
    <div class="modal-dialog modal-lg">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="modalTitle">
            <i class="bi bi-person-plus"></i> 거래처 등록
          </h5>
          <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
        </div>
        <div class="modal-body">
          <form id="customerForm">
            <input type="hidden" id="mode" value="add">
            <input type="hidden" id="original_code">

            <div class="row g-3">
              <!-- 거래처코드 -->
              <div class="col-md-6">
                <label class="form-label required">거래처코드</label>
                <input type="text" class="form-control" id="customer_code"
                  placeholder="거래처코드 입력" required maxlength="10">
                <small class="text-muted">최대 10자리</small>
              </div>

              <!-- 거래처명 -->
              <div class="col-md-6">
                <label class="form-label required">거래처명</label>
                <input type="text" class="form-control" id="customer_name"
                  placeholder="거래처명 입력" required maxlength="50">
              </div>

              <!-- 거래처구분 -->
              <div class="col-md-6">
                <label class="form-label required">거래처구분</label>
                <select class="form-select" id="customer_type" required>
                  <option value="">선택하세요</option>
                  <option value="매출처">매출처</option>
                  <option value="매입처">매입처</option>
                  <option value="양쪽">양쪽</option>
                </select>
              </div>

              <!-- 사업자번호 -->
              <div class="col-md-6">
                <label class="form-label">사업자번호</label>
                <input type="text" class="form-control" id="business_number"
                  placeholder="000-00-00000" maxlength="12">
              </div>

              <!-- 대표자 -->
              <div class="col-md-6">
                <label class="form-label">대표자</label>
                <input type="text" class="form-control" id="representative"
                  placeholder="대표자명" maxlength="30">
              </div>

              <!-- 연락처 -->
              <div class="col-md-6">
                <label class="form-label">연락처</label>
                <input type="text" class="form-control" id="phone"
                  placeholder="000-0000-0000" maxlength="15">
              </div>

              <!-- 팩스 -->
              <div class="col-md-6">
                <label class="form-label">팩스</label>
                <input type="text" class="form-control" id="fax"
                  placeholder="000-0000-0000" maxlength="15">
              </div>

              <!-- 이메일 -->
              <div class="col-md-6">
                <label class="form-label">이메일</label>
                <input type="email" class="form-control" id="email"
                  placeholder="email@example.com" maxlength="50">
              </div>

              <!-- 주소 -->
              <div class="col-md-12">
                <label class="form-label">주소</label>
                <input type="text" class="form-control" id="address"
                  placeholder="주소 입력" maxlength="100">
              </div>

              <!-- 거래시작일 -->
              <div class="col-md-6">
                <label class="form-label">거래시작일</label>
                <input type="date" class="form-control" id="start_date">
              </div>

              <!-- 담당자 -->
              <div class="col-md-6">
                <label class="form-label">담당자</label>
                <input type="text" class="form-control" id="manager"
                  placeholder="담당자명" maxlength="30">
              </div>

              <!-- 비고 -->
              <div class="col-md-12">
                <label class="form-label">비고</label>
                <textarea class="form-control" id="memo" rows="3"
                  placeholder="비고사항" maxlength="200"></textarea>
              </div>

              <!-- 사용여부 -->
              <div class="col-md-12">
                <div class="form-check">
                  <input class="form-check-input" type="checkbox"
                    id="use_yn" checked>
                  <label class="form-check-label" for="use_yn">
                    사용함
                  </label>
                </div>
              </div>
            </div>
          </form>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">
            <i class="bi bi-x-circle"></i> 취소
          </button>
          <button type="button" class="btn btn-primary" onclick="saveCustomer()">
            <i class="bi bi-check-circle"></i> 저장
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- 상태바 -->
  <div class="status-bar">
    <div>
      <i class="bi bi-calendar3"></i>
      <strong>작업일자:</strong> <?php echo $formatted_work_date; ?>
    </div>
    <div>
      <i class="bi bi-person"></i>
      <strong>사용자:</strong> <?php echo htmlspecialchars($user_name); ?>
    </div>
    <div>
      <i class="bi bi-building"></i>
      <strong>지점:</strong> <?php echo htmlspecialchars($user_branch); ?>
    </div>
  </div>

  <!-- Bootstrap 5 JS -->
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

  <!-- jQuery -->
  <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>

  <!-- CSRF token for AJAX -->
  <script>
    const CSRF_TOKEN = '<?php echo htmlspecialchars($_SESSION['csrf_token'] ?? ''); ?>';
  </script>

  <script>
    // 전역 변수
    let selectedRow = null; // 선택된 행
    let customerModal; // Bootstrap 모달 객체

    /**
     * 페이지 로드 시 실행
     */
    $(document).ready(function() {
      // Bootstrap 모달 객체 초기화
      customerModal = new bootstrap.Modal(document.getElementById('customerModal'));

      // 테이블 행 클릭 이벤트
      $('#customerTable tbody').on('click', 'tr', function() {
        if ($(this).find('td').length === 1) return; // 빈 행 제외

        // 기존 선택 해제
        $('#customerTable tbody tr').removeClass('selected');
        // 현재 행 선택
        $(this).addClass('selected');
        selectedRow = $(this);
      });

      // 테이블 행 더블클릭 이벤트 (수정 모달 열기)
      $('#customerTable tbody').on('dblclick', 'tr', function() {
        if ($(this).find('td').length === 1) return;
        selectedRow = $(this);
        showEditModal();
      });

      // Enter 키로 검색
      $('#search_code, #search_name').on('keypress', function(e) {
        if (e.key === 'Enter') {
          searchCustomers();
        }
      });
    });

    /**
     * 거래처 검색
     */
    function searchCustomers() {
      const searchData = {
        code: $('#search_code').val().trim(),
        name: $('#search_name').val().trim(),
        type: $('#search_type').val()
      };

      $.ajax({
        url: '../../api/customers.php',
        method: 'GET',
        data: searchData,
        dataType: 'json',
        success: function(response) {
          if (response.success) {
            displayCustomers(response.data);
            $('#total_count').text(response.data.length);
          } else {
            alert('조회 중 오류가 발생했습니다: ' + response.message);
          }
        },
        error: function(xhr, status, error) {
          alert('서버 통신 중 오류가 발생했습니다.');
          console.error(error);
        }
      });
    }

    /**
     * 거래처 목록 표시
     */
    function displayCustomers(customers) {
      const tbody = $('#customerTableBody');
      tbody.empty();

      if (customers.length === 0) {
        tbody.html(`
                    <tr>
                        <td colspan="7" class="text-center py-5">
                            <i class="bi bi-inbox" style="font-size: 3rem; color: #ccc;"></i>
                            <p class="text-muted mt-3">검색 결과가 없습니다.</p>
                        </td>
                    </tr>
                `);
        return;
      }

      customers.forEach(function(customer) {
        const useYn = customer.사용여부 === 'Y' ? '사용' : '미사용';
        const useClass = customer.사용여부 === 'Y' ? 'text-success' : 'text-danger';

        const row = `
                    <tr data-code="${customer.거래처코드}">
                        <td>${customer.거래처코드}</td>
                        <td><strong>${customer.거래처명}</strong></td>
                        <td>${customer.거래처구분}</td>
                        <td>${customer.대표자 || '-'}</td>
                        <td>${customer.연락처 || '-'}</td>
                        <td>${customer.주소 || '-'}</td>
                        <td class="${useClass}"><strong>${useYn}</strong></td>
                    </tr>
                `;
        tbody.append(row);
      });
    }

    /**
     * 신규 등록 모달 표시
     */
    function showAddModal() {
      $('#mode').val('add');
      $('#modalTitle').html('<i class="bi bi-person-plus"></i> 거래처 등록');
      $('#customerForm')[0].reset();
      $('#customer_code').prop('readonly', false);
      $('#use_yn').prop('checked', true);
      customerModal.show();
    }

    /**
     * 수정 모달 표시
     */
    function showEditModal() {
      if (!selectedRow) {
        alert('수정할 거래처를 선택하세요.');
        return;
      }

      const customerCode = selectedRow.data('code');

      // 거래처 상세 정보 조회
      $.ajax({
        url: '../../api/customers.php',
        method: 'GET',
        data: {
          code: customerCode,
          detail: true
        },
        dataType: 'json',
        success: function(response) {
          if (response.success && response.data) {
            const customer = response.data;

            $('#mode').val('edit');
            $('#modalTitle').html('<i class="bi bi-pencil-square"></i> 거래처 수정');
            $('#original_code').val(customer.거래처코드);
            $('#customer_code').val(customer.거래처코드).prop('readonly', true);
            $('#customer_name').val(customer.거래처명);
            $('#customer_type').val(customer.거래처구분);
            $('#business_number').val(customer.사업자번호 || '');
            $('#representative').val(customer.대표자 || '');
            $('#phone').val(customer.연락처 || '');
            $('#fax').val(customer.팩스 || '');
            $('#email').val(customer.이메일 || '');
            $('#address').val(customer.주소 || '');
            $('#start_date').val(customer.거래시작일 || '');
            $('#manager').val(customer.담당자 || '');
            $('#memo').val(customer.비고 || '');
            $('#use_yn').prop('checked', customer.사용여부 === 'Y');

            customerModal.show();
          } else {
            alert('거래처 정보를 불러올 수 없습니다.');
          }
        },
        error: function() {
          alert('서버 통신 중 오류가 발생했습니다.');
        }
      });
    }

    /**
     * 거래처 저장 (등록/수정)
     */
    function saveCustomer() {
      // 필수 입력 체크
      if (!$('#customer_code').val().trim()) {
        alert('거래처코드를 입력하세요.');
        $('#customer_code').focus();
        return;
      }

      if (!$('#customer_name').val().trim()) {
        alert('거래처명을 입력하세요.');
        $('#customer_name').focus();
        return;
      }

      if (!$('#customer_type').val()) {
        alert('거래처구분을 선택하세요.');
        $('#customer_type').focus();
        return;
      }

      // 데이터 수집
      const formData = {
        mode: $('#mode').val(),
        original_code: $('#original_code').val(),
        customer_code: $('#customer_code').val().trim(),
        customer_name: $('#customer_name').val().trim(),
        customer_type: $('#customer_type').val(),
        business_number: $('#business_number').val().trim(),
        representative: $('#representative').val().trim(),
        phone: $('#phone').val().trim(),
        fax: $('#fax').val().trim(),
        email: $('#email').val().trim(),
        address: $('#address').val().trim(),
        start_date: $('#start_date').val(),
        manager: $('#manager').val().trim(),
        memo: $('#memo').val().trim(),
        use_yn: $('#use_yn').is(':checked') ? 'Y' : 'N'
      };

      // AJAX로 저장 요청
      $.ajax({
        url: '../../api/customers.php',
        method: 'POST',
        data: Object.assign({}, formData, { csrf_token: CSRF_TOKEN }),
        dataType: 'json',
        success: function(response) {
          if (response.success) {
            alert(response.message);
            customerModal.hide();
            searchCustomers(); // 목록 새로고침
          } else {
            alert('저장 실패: ' + response.message);
          }
        },
        error: function(xhr, status, error) {
          alert('서버 통신 중 오류가 발생했습니다.');
          console.error(error);
        }
      });
    }

    /**
     * 거래처 삭제
     */
    function deleteCustomer() {
      if (!selectedRow) {
        alert('삭제할 거래처를 선택하세요.');
        return;
      }

      const customerCode = selectedRow.data('code');
      const customerName = selectedRow.find('td:eq(1)').text();

      if (!confirm(`거래처 [${customerName}]을(를) 삭제하시겠습니까?\n\n※ 주의: 거래 내역이 있는 거래처는 삭제할 수 없습니다.`)) {
        return;
      }

      $.ajax({
        url: '../../api/customers.php',
        method: 'DELETE',
        data: {
          customer_code: customerCode,
          csrf_token: CSRF_TOKEN
        },
        dataType: 'json',
        success: function(response) {
          if (response.success) {
            alert(response.message);
            selectedRow = null;
            searchCustomers(); // 목록 새로고침
          } else {
            alert('삭제 실패: ' + response.message);
          }
        },
        error: function(xhr, status, error) {
          alert('서버 통신 중 오류가 발생했습니다.');
          console.error(error);
        }
      });
    }

    /**
     * 엑셀 출력
     */
    function exportToExcel() {
      const searchData = {
        code: $('#search_code').val().trim(),
        name: $('#search_name').val().trim(),
        type: $('#search_type').val(),
        export: 'excel'
      };

      // 엑셀 다운로드 URL 생성
      const params = new URLSearchParams(searchData);
      window.open('../../api/customers.php?' + params.toString(), '_blank');
    }
  </script>
</body>

</html>