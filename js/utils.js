// =============================================
// 共用工具函式
// =============================================

// 將文字轉為 HTML 安全字串，放進 innerHTML 前一律先經過此函式
function escapeHtml(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
