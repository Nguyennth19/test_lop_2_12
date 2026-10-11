/**
 * Lớp 2/12 — V7.2.1: fix include_ for JavaScript/CSS/HTML fragments.
 * Giữ nguyên GAS deployment, QR Base64, Google Sheets/Drive và auth.
 * Không khai báo hàm doGet() / include_() trùng ở file khác.
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  const output = template.evaluate();
  return output
    .setTitle('Lớp 2/12 | Chia sẻ và minh bạch')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include raw fragment content for inline <script>/<style>/<template>.
 * Important: createHtmlOutputFromFile(name).getContent() parses a JS-only file as HTML,
 * and may throw "Nội dung HTML không đúng định dạng" for JS syntax such as /</g.
 * getRawContent() reads template source without interpreting it as HTML.
 * Use only for trusted source files in this project, never for user content.
 */
function include_(name) {
  return HtmlService.createTemplateFromFile(name).getRawContent();
}

/** Diagnostic: verifies source fragments can be read, then renders Index. */
function checkIframeConfigV7() {
  const required = ['Styles', 'ClientCore', 'ClientExpense'];
  required.forEach(function(name) {
    const raw = include_(name);
    if (!raw || typeof raw !== 'string') {
      throw new Error('Không đọc được file: ' + name);
    }
    Logger.log('RAW OK: ' + name + ' (' + raw.length + ' ký tự)');
  });
  const html = doGet();
  Logger.log('V7.2.1: HtmlOutput created successfully.');
  Logger.log('Title: ' + html.getTitle());
  Logger.log('HTML size (characters): ' + html.getContent().length);
  return { ok: true, title: html.getTitle(), size: html.getContent().length };
}
