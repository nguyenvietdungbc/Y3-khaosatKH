// ================================================================
// Y3 — VietinBank · Khảo sát Khách hàng
// Google Apps Script Backend
// Deploy: Extensions > Apps Script > Deploy > New deployment
//         Type: Web app | Execute as: Me | Access: Anyone
// ================================================================

const EMAIL_TO   = 'nguyenvietdung.bc@gmail.com';
const SHEET_ID   = '1WytwKAtL-wU7S0lWmHnj3Eo7A-tFzVnWLzD-14ikylc'; // Google Sheet cụ thể
const SHEET_NAME = 'Khảo sát KH';

// ----------------------------------------------------------------
// doPost — nhận dữ liệu từ landing page
// ----------------------------------------------------------------
function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    saveToSheet(data);
    sendEmailNotification(data);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'Đã lưu và gửi email.' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// doGet — kiểm tra API còn sống
function doGet() {
  return ContentService.createTextOutput('✅ Y3 Survey API đang hoạt động.');
}

// ----------------------------------------------------------------
// Lưu vào Google Sheets
// ----------------------------------------------------------------
function saveToSheet(d) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  // Tạo sheet + header nếu chưa có
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      'Thời gian', 'Tên doanh nghiệp', 'Mã số thuế', 'Số điện thoại',
      'Email DN', 'Người liên hệ', 'Chức vụ', 'Ngành nghề', 'Quy mô', 'Địa chỉ',
      'Ngân hàng chính', 'Doanh thu/tháng', 'Số GD/tháng',
      'Kênh thanh toán', 'Khó khăn hiện tại',
      'Nhu cầu thu hộ', 'Nhu cầu chi hộ', 'Nhu cầu CMS',
      'Phần mềm ERP', 'Ưu tiên API (1-5)',
      'Thời điểm triển khai', 'Kỳ vọng từ VietinBank'
    ];
    sheet.appendRow(headers);

    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold')
               .setBackground('#005993')
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(22, 300);
  }

  // Append dữ liệu
  sheet.appendRow([
    d.timestamp    || new Date().toLocaleString('vi-VN'),
    d.company      || '',
    d.taxcode      || '',
    d.phone        || '',
    d.email        || '',
    d.contact      || '',
    d.title        || '',
    d.industry     || '',
    d.size         || '',
    d.address      || '',
    d.mainbank     || '',
    d.revenue      || '',
    d.txcount      || '',
    d.channels     || '',
    d.painpoints   || '',
    d.collection   || '',
    d.disbursement || '',
    d.cms          || '',
    d.erp          || '',
    d.apiscale     || '',
    d.timing       || '',
    d.expectation  || ''
  ]);

  // Tô màu xen kẽ cho dễ đọc
  const lastRow   = sheet.getLastRow();
  const lastCol   = 22;
  const rowRange  = sheet.getRange(lastRow, 1, 1, lastCol);
  rowRange.setBackground(lastRow % 2 === 0 ? '#f0f7ff' : '#ffffff');
}

// ----------------------------------------------------------------
// Gửi email thông báo
// ----------------------------------------------------------------
function sendEmailNotification(d) {
  const subject = 'Y3 - Khảo sát KH: ' + (d.company || '(Chưa có tên)');

  const row = (label, val) =>
    `<tr>
      <td style="padding:9px 14px;color:#6b7280;border-bottom:1px solid #f3f4f6;
                 white-space:nowrap;font-weight:500;width:220px">${label}</td>
      <td style="padding:9px 14px;color:#1f2937;border-bottom:1px solid #f3f4f6;
                 font-weight:600">${val || '—'}</td>
    </tr>`;

  const section = (icon, title) =>
    `<tr>
      <td colspan="2" style="padding:10px 14px 7px;background:#eff6ff;
                             color:#1e40af;font-weight:700;font-size:13px;
                             letter-spacing:0.5px">${icon} ${title}</td>
    </tr>`;

  const htmlBody = `
<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:20px;font-family:Arial,sans-serif;background:#f8f9fb">
<div style="max-width:700px;margin:0 auto;background:white;border-radius:14px;
            overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.09)">

  <!-- Header -->
  <div style="background:linear-gradient(135deg,#005993,#007DB7);padding:28px 32px">
    <div style="font-size:11px;color:rgba(255,255,255,0.7);letter-spacing:2px;
                text-transform:uppercase;margin-bottom:6px">VietinBank · Sáng kiến Y3</div>
    <div style="font-size:22px;font-weight:800;color:white">
      📋 Khảo sát mới từ khách hàng
    </div>
    <div style="font-size:14px;color:rgba(255,255,255,0.8);margin-top:6px">
      ${d.company || ''} &nbsp;·&nbsp; ${d.timestamp || ''}
    </div>
  </div>

  <!-- Alert box -->
  <div style="background:#fef3c7;border-left:4px solid #f59e0b;padding:12px 20px;
              font-size:13px;color:#92400e">
    <strong>⚡ Hành động cần thực hiện:</strong>
    Liên hệ <strong>${d.contact || ''}</strong> (${d.phone || ''})
    để tư vấn giải pháp thanh toán & quản lý dòng tiền phù hợp.
  </div>

  <!-- Data table -->
  <div style="padding:20px 24px">
    <table style="width:100%;border-collapse:collapse;font-size:14px">

      ${section('🏢', 'THÔNG TIN DOANH NGHIỆP')}
      ${row('Tên doanh nghiệp', d.company)}
      ${row('Mã số thuế', d.taxcode)}
      ${row('Số điện thoại', d.phone)}
      ${row('Email doanh nghiệp', d.email)}
      ${row('Người liên hệ', d.contact + (d.title ? ' — ' + d.title : ''))}
      ${row('Ngành nghề', d.industry)}
      ${row('Quy mô', d.size)}
      ${row('Địa chỉ', d.address)}

      ${section('💳', 'NHU CẦU THANH TOÁN')}
      ${row('Ngân hàng đang dùng chính', d.mainbank)}
      ${row('Doanh thu hàng tháng', d.revenue)}
      ${row('Số giao dịch/tháng', d.txcount)}
      ${row('Kênh thanh toán đang dùng', d.channels)}
      ${row('Khó khăn hiện tại', d.painpoints)}

      ${section('📊', 'QUẢN LÝ DÒNG TIỀN')}
      ${row('Nhu cầu thu hộ', d.collection)}
      ${row('Nhu cầu chi hộ', d.disbursement)}
      ${row('Nhu cầu CMS / Pooling', d.cms)}
      ${row('Phần mềm ERP đang dùng', d.erp)}
      ${row('Mức ưu tiên kết nối API', d.apiscale ? d.apiscale + ' / 5' : '—')}
      ${row('Thời điểm mong muốn TK', d.timing)}

      ${section('🎯', 'KỲ VỌNG TỪ VIETINBANK')}
      ${row('Nội dung kỳ vọng', d.expectation)}

    </table>
  </div>

  <!-- Footer -->
  <div style="background:#f8f9fb;padding:14px 24px;border-top:1px solid #e5e7eb;
              text-align:center;font-size:12px;color:#9ca3af">
    Email tự động từ Landing Page Y3 · VietinBank &nbsp;|&nbsp;
    Dữ liệu đã được lưu vào Google Sheets tự động.
  </div>
</div>
</body></html>`;

  GmailApp.sendEmail(EMAIL_TO, subject, '', { htmlBody: htmlBody });
}
