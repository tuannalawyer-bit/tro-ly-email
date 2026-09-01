// Utility Functions
window.Utils = {
    formatDate: (dateStr) => {
        if (!dateStr) return '';
        const date = new Date(dateStr);
        const now = new Date();
        const diffMs = now - date;
        const diffMins = Math.floor(diffMs / 60000);
        const diffHours = Math.floor(diffMins / 60);
        const diffDays = Math.floor(diffHours / 24);

        if (diffMins < 1) return 'Vừa xong';
        if (diffMins < 60) return `${diffMins} phút trước`;
        
        const isToday = date.getDate() === now.getDate() && date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear();
        if (isToday) {
            return `Hôm nay ${date.getHours().toString().padStart(2, '0')}:${date.getMinutes().toString().padStart(2, '0')}`;
        }
        
        if (diffDays === 1) return 'Hôm qua';
        if (diffDays < 7) return `${diffDays} ngày trước`;
        
        return `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}`;
    },

    getInitials: (name) => {
        if (!name) return '?';
        const parts = name.trim().split(' ');
        if (parts.length >= 2) {
            return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
        }
        return name.substring(0, 2).toUpperCase();
    },

    getAvatarColor: (name) => {
        if (!name) return 'var(--accent-gradient)';
        let hash = 0;
        for (let i = 0; i < name.length; i++) {
            hash = name.charCodeAt(i) + ((hash << 5) - hash);
        }
        const colors = [
            'linear-gradient(135deg, #FF6B6B, #FF8E53)',
            'linear-gradient(135deg, #4E65FF, #92EFFD)',
            'linear-gradient(135deg, #7c5cfc, #00d4aa)',
            'linear-gradient(135deg, #11998e, #38ef7d)',
            'linear-gradient(135deg, #8E2DE2, #4A00E0)'
        ];
        return colors[Math.abs(hash) % colors.length];
    },

    stripHtml: (html) => {
        const tmp = document.createElement('DIV');
        tmp.innerHTML = html;
        return tmp.textContent || tmp.innerText || '';
    },

    escapeHtml: (unsafe) => {
        return (unsafe || '').toString()
             .replace(/&/g, "&amp;")
             .replace(/</g, "&lt;")
             .replace(/>/g, "&gt;")
             .replace(/"/g, "&quot;")
             .replace(/'/g, "&#039;");
    },

    debounce: (func, wait) => {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    },

    showToast: (message, type = 'info') => {
        const container = document.getElementById('toast-container');
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.innerText = message;
        container.appendChild(toast);
        
        setTimeout(() => {
            toast.style.opacity = '0';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    },

    showLoading: (show) => {
        const loader = document.getElementById('global-loading');
        if (show) loader.classList.remove('hidden');
        else loader.classList.add('hidden');
    },

    truncate: (text, maxLen) => {
        if (!text) return '';
        if (text.length <= maxLen) return text;
        return text.substring(0, maxLen) + '...';
    },

    /**
     * Bọc HTML thư vào một tài liệu độc lập để render trong iframe sandbox.
     * CSP chặn mọi tài nguyên ngoài (kể cả tracking pixel); sandbox chặn script.
     */
    buildEmailDocument: (html) => {
        const csp = "default-src 'none'; img-src data: cid:; style-src 'unsafe-inline'; font-src data:";
        return `<!DOCTYPE html><html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<base target="_blank">
<style>
  html,body{margin:0;padding:16px;background:#12122a;color:#dcdce6;
    font-family:'Segoe UI Variable','Segoe UI',system-ui,sans-serif;
    font-size:14px;line-height:1.6;word-wrap:break-word}
  a{color:#7c5cfc}
  img{max-width:100%;height:auto}
  table{max-width:100%!important}
  blockquote{border-left:3px solid #7c5cfc55;margin:8px 0;padding-left:12px;color:#8888a8}
</style></head><body>${html || '<p style="color:#55556a">Không có nội dung</p>'}</body></html>`;
    }
};

