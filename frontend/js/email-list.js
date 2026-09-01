// ============================================================
// email-list.js — Email List Component
// Render danh sách email, tìm kiếm, filter, renderFolders
// ============================================================

class EmailList {
    constructor() {
        this.container = document.getElementById('email-list-container');
        this.btnLoadMore = document.getElementById('btn-load-more');
        this.searchInput = document.getElementById('search-input');
        this.classifications = {};
        this.activeFilter = 'all';

        this.bindEvents();
    }

    bindEvents() {
        if (this.searchInput) {
            this.searchInput.addEventListener('input', window.Utils.debounce((e) => {
                if (window.App) window.App.handleSearch(e.target.value.trim());
            }, 500));
        }

        if (this.btnLoadMore) {
            this.btnLoadMore.addEventListener('click', () => {
                if (window.App) window.App.loadMoreEmails();
            });
        }

        document.getElementById('filter-bar')?.addEventListener('click', (e) => {
            const chip = e.target.closest('.filter-chip');
            if (!chip) return;
            document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');
            this.activeFilter = chip.dataset.filter;
            window.App?.rerenderList();
        });

        document.getElementById('btn-classify')?.addEventListener('click',
            () => window.App?.classifyVisible());
    }

    renderSkeletonLoading(count = 6) {
        this.container.innerHTML = '';
        for (let i = 0; i < count; i++) {
            const skeleton = document.createElement('div');
            skeleton.className = 'email-item skeleton-item';
            skeleton.style.animationDelay = `${i * 0.08}s`;
            skeleton.innerHTML = `
                <div class="email-item-header">
                    <div class="sender-info">
                        <div class="skeleton skeleton-avatar"></div>
                        <div class="skeleton skeleton-text short"></div>
                    </div>
                </div>
                <div class="skeleton skeleton-text"></div>
                <div class="skeleton skeleton-text short"></div>
            `;
            this.container.appendChild(skeleton);
        }
        if (this.btnLoadMore) this.btnLoadMore.style.display = 'none';
    }

    renderEmptyState(isSearch = false) {
        this.container.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon">${isSearch ? '🔍' : '📭'}</div>
                <p>${isSearch ? 'Không tìm thấy kết quả nào' : 'Thư mục trống'}</p>
            </div>
        `;
        if (this.btnLoadMore) this.btnLoadMore.style.display = 'none';
    }

    renderEmailList(emails, selectedId = null, hasMore = false) {
        this.container.innerHTML = '';

        const visible = this.activeFilter === 'all'
            ? emails
            : emails.filter(e => this.classifications[e.entry_id]?.category === this.activeFilter);

        if (!visible.length) {
            const isSearch = this.searchInput?.value.trim() !== '';
            this.renderEmptyState(isSearch);
            return;
        }

        visible.forEach((email, index) => {
            const el = this.createEmailItem(email, email.entry_id === selectedId);
            el.style.animationDelay = `${Math.min(index, 12) * 0.03}s`;
            this.container.appendChild(el);
        });

        if (this.btnLoadMore) {
            const canPage = hasMore && this.activeFilter === 'all'
                            && !this.searchInput?.value.trim();
            this.btnLoadMore.style.display = canPage ? 'block' : 'none';
        }
    }

    createEmailItem(email, isSelected) {
        const el = document.createElement('div');
        const isUnread = email.is_unread;
        el.className = `email-item ${isSelected ? 'active' : ''} ${isUnread ? 'unread' : ''}`;
        el.dataset.id = email.entry_id;

        const senderName = window.Utils.escapeHtml(email.sender_name || email.sender_email || 'Không rõ');
        const initials = window.Utils.getInitials(email.sender_name || email.sender_email || '?');
        const avatarBg = window.Utils.getAvatarColor(email.sender_name || '');
        const timeStr = window.Utils.formatDate(email.received_time);
        const subject = window.Utils.escapeHtml(email.subject || '(Không có chủ đề)');
        
        const cls = this.classifications[email.entry_id];
        const summaryText = cls?.summary || email.preview || '';
        const preview = window.Utils.truncate(window.Utils.escapeHtml(summaryText), 120);

        const CAT_SLUG = {
            'Cần trả lời': 'reply', 'Việc cần làm': 'todo',
            'Chỉ để biết': 'info', 'Quảng cáo/Rác': 'spam'
        };
        const PRIO_SLUG = { 'Cao': 'high', 'Trung bình': 'mid', 'Thấp': 'low' };
        const chip = cls
            ? `<span class="cat-chip cat-${CAT_SLUG[cls.category] || 'info'}">
                 <i class="prio-dot prio-${PRIO_SLUG[cls.priority] || 'mid'}"></i>
                 ${window.Utils.escapeHtml(cls.category)}</span>`
            : '';

        const attachIcon = email.has_attachments ? '<span class="attach-icon" title="Có đính kèm">📎</span>' : '';
        const unreadDot = isUnread ? '<span class="unread-dot"></span>' : '';

        el.innerHTML = `
            <div class="email-item-header">
                <div class="sender-info">
                    <div class="avatar" style="background: ${avatarBg}">${initials}</div>
                    <span class="sender-name">${senderName}</span>
                    ${unreadDot}
                </div>
                <div class="email-item-meta">
                    ${attachIcon}
                    <span class="email-time">${timeStr}</span>
                </div>
            </div>
            <div class="email-subject">${subject}</div>
            <div class="email-preview">${preview}</div>
            ${chip}
        `;

        el.addEventListener('click', () => {
            document.querySelectorAll('.email-item').forEach(item => item.classList.remove('active'));
            el.classList.add('active');
            el.classList.remove('unread');

            if (window.App) window.App.selectEmail(email.entry_id);
        });

        return el;
    }

    renderFolders(folders, activeId) {
        const list = document.getElementById('folder-list');
        if (!list) return;
        const ICONS = { inbox: '📥', sent: '📤', drafts: '📝', other: '📁' };

        list.innerHTML = folders.map(f => `
            <div class="folder-item ${f.entry_id === activeId ? 'active' : ''}"
                 data-folder-id="${f.entry_id}" style="padding-left:${12 + f.depth * 14}px">
                <span class="folder-icon">${ICONS[f.kind] || ICONS.other}</span>
                <span class="folder-name">${window.Utils.escapeHtml(f.name)}</span>
                <span class="badge" ${f.unread_count > 0 ? '' : 'style="display:none"'}>${f.unread_count}</span>
            </div>`).join('');
    }
}

window.EmailListUI = new EmailList();
