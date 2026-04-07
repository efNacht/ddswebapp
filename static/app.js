// Global utility functions
function formatMXN(n) {
    if (n == null) return '—';
    return new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', maximumFractionDigits: 0 }).format(n);
}

function formatN(n) {
    return n ? n.toLocaleString('es-MX', { maximumFractionDigits: 0 }) : '';
}

function escHtml(s) {
    return (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}
