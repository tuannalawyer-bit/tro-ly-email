// ============================================================
// email-viewer.js — Email Detail & Conversation Thread Viewer
// Hiển thị chi tiết email, lịch sử hội thoại
// ============================================================

class EmailViewer {
    constructor() {
        this.emptyState = document.getElementById('email-empty-state');
        this.emailView = document.getElementById('email-view');
        this.headerContainer = document.getElementById('email-header');
        this.bodyContainer = document.getElementById('email-body');
        this.attachmentsContainer = document.getElementById('email-attachments');
        this.threadContainer = document.getElementById('conversation-thread');

        this.bindEvents();
    }

    bindEvents() {
        document.getElementById('btn-ai-reply')?.addEventListener('click',
            () => window.DraftEditorUI?.showDraftPanel(false));
        document.getElementById('btn-ai-reply-all')?.addEventListener('click',
            () => window.DraftEditorUI?.showDraftPanel(true));
    }

    showEmptyDetail() {
        if (this.emptyState) this.emptyState.classList.remove('hidden');
        if (this.emailView) this.emailView.classList.add('hidden');
    }

    renderEmailDetail(emailDetail) {
        if (!emailDetail) {
            this.showEmptyDetail();
            return;
        }

        if (this.emptyState) this.emptyState.classList.add('hidden');
        if (this.emailView) this.emailView.classList.remove('hidden');

        this.renderEmailHeader(emailDetail);
        this.renderEmailBody(emailDetail.html_body || emailDetail.body);

        if (emailDetail.attachments && emailDetail.attachments.length > 0) {
            this.renderAttachments(emailDetail.attachments);
        } else if (this.attachmentsContainer) {
            this.attachmentsContainer.innerHTML = '';
        }
    }

    renderEmailHeader(email) {
        const senderName = window.Utils.escapeHtml(email.sender_name || 'Không rõ');
        const senderEmail = window.Utils.escapeHtml(email.sender_email || '');
        const initials = window.Utils.getInitials(senderName);
        const avatarBg = window.Utils.getAvatarColor(senderName);
        const timeStr = window.Utils.formatDate(email.received_time);
        const subject = window.Utils.escapeHtml(email.subject || '(Không có chủ đề)');
        const toStr = window.Utils.escapeHtml(email.to || '');
        const ccStr = email.cc ? `<div class="email-meta-line">CC: ${window.Utils.escapeHtml(email.cc)}</div>` : '';

        this.headerContainer.innerHTML = `
            <div class="email-header-top">
                <div class="sender-info">
                    <div class="avatar large" style="background: ${avatarBg}">${initials}</div>
                    <div class="sender-details">
                        <div class="sender-name">${senderName}</div>
                        <div class="sender-email">&lt;${senderEmail}&gt;</div>
                    </div>
                </div>
                <div class="email-date">${timeStr}</div>
            </div>
            <div class="email-subject-large">${subject}</div>
            <div class="email-recipients">
                <div class="email-meta-line">Tới: ${toStr}</div>
                ${ccStr}
            </div>
        `;
    }

    renderEmailBody(htmlContent) {
        if (!this.bodyContainer) return;
        this.bodyContainer.innerHTML = '';

        const iframe = document.createElement('iframe');
        iframe.className = 'email-iframe';
        // sandbox KHÔNG có allow-scripts và KHÔNG có allow-same-origin
        // => nội dung thư không thể chạm tới window.pywebview.api của trang cha.
        // allow-popups để link trong thư vẫn mở được ra trình duyệt.
        iframe.setAttribute('sandbox', 'allow-popups allow-popups-to-escape-sandbox');
        iframe.setAttribute('referrerpolicy', 'no-referrer');
        iframe.srcdoc = window.Utils.buildEmailDocument(htmlContent);
        this.bodyContainer.appendChild(iframe);
    }

    renderAttachments(attachments) {
        if (!this.attachmentsContainer) return;

        const getFileIcon = (filename) => {
            const ext = (filename || '').split('.').pop().toLowerCase();
            const icons = {
                'pdf': '📄', 'doc': '📝', 'docx': '📝', 'xls': '📊', 'xlsx': '📊',
                'ppt': '📊', 'pptx': '📊', 'jpg': '🖼️', 'jpeg': '🖼️', 'png': '🖼️',
                'gif': '🖼️', 'zip': '📦', 'rar': '📦', 'txt': '📃', 'csv': '📊'
            };
            return icons[ext] || '📎';
        };

        const formatSize = (bytes) => {
            if (!bytes) return '';
            if (bytes < 1024) return bytes + ' B';
            if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
            return (bytes / 1048576).toFixed(1) + ' MB';
        };

        let html = '<div class="attachments-section">';
        html += '<h4 class="attachments-title">📎 Đính kèm</h4>';
        html += '<div class="attachments-grid">';

        attachments.forEach(att => {
            const icon = getFileIcon(att.filename);
            const size = formatSize(att.size);
            html += `
                <div class="attachment-item">
                    <span class="att-icon">${icon}</span>
                    <span class="att-name">${window.Utils.escapeHtml(att.filename)}</span>
                    ${size ? `<span class="att-size">${size}</span>` : ''}
                </div>
            `;
        });

        html += '</div></div>';
        this.attachmentsContainer.innerHTML = html;
    }

    renderConversationThread(messages) {
        if (!this.threadContainer) return;
        if (!messages || messages.length <= 1) {
            this.threadContainer.innerHTML = '';
            return;
        }

        const rows = messages.map((msg, index) => {
            const senderName = window.Utils.escapeHtml(msg.sender_name || 'Không rõ');
            const timeStr = window.Utils.formatDate(msg.sent_time);
            const body = window.Utils.escapeHtml(
                window.Utils.truncate((msg.body || '').replace(/\s+/g, ' '), 400));
            const icon = msg.direction === 'sent' ? '📤' : '📥';
            const expanded = index === messages.length - 1 ? 'expanded' : '';
            return `<div class="thread-item ${expanded}">
                <div class="thread-dot"></div>
                <div class="thread-card">
                    <div class="thread-card-header">
                        <div class="thread-sender">${icon} <strong>${senderName}</strong></div>
                        <span class="thread-time">${timeStr}</span>
                    </div>
                    <div class="thread-body">${body}</div>
                </div></div>`;
        }).join('');

        this.threadContainer.innerHTML =
            `<div class="thread-section"><h4 class="thread-title">💬 Lịch sử hội thoại</h4>
             <div class="thread-timeline">${rows}</div></div>`;

        this.threadContainer.onclick = (e) => {
            e.target.closest('.thread-item')?.classList.toggle('expanded');
        };
    }
}

window.EmailViewerUI = new EmailViewer();
