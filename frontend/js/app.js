// ============================================================
// app.js — Main Application Controller for Email Assistant
// Quản lý state, kết nối Outlook, điều phối UI components
// ============================================================

class AppController {
    constructor() {
        this.state = {
            folders: [],
            currentFolderId: null,
            currentEmailId: null,
            emails: [],
            offset: 0,
            limit: 30,
            hasMore: false,
            searchQuery: '',
            styleProfileLoaded: false,
        };
        this.init();
    }

    init() {
        window.addEventListener('pywebviewready', () => this.onReady());
        this.bindEvents();
    }

    async onReady() {
        window.Utils.showLoading(true);
        try {
            const conn = await window.pywebview.api.connect_outlook();
            if (!conn?.success) {
                window.Utils.showToast(conn?.error || 'Không kết nối được Outlook', 'error');
                return;
            }
            window.Utils.showToast('Đã kết nối Outlook', 'success');
            await this.loadFolders();
            await this.loadEmails(true);
            await this.checkStyleProfile();
        } catch (e) {
            window.Utils.showToast('Lỗi khởi tạo: ' + (e.message || e), 'error');
        } finally {
            window.Utils.showLoading(false);
        }
    }

    bindEvents() {
        document.getElementById('folder-list')?.addEventListener('click', (e) => {
            const item = e.target.closest('.folder-item');
            if (item?.dataset.folderId) this.switchFolder(item.dataset.folderId);
        });

        document.getElementById('btn-settings')?.addEventListener('click', async () => {
            document.getElementById('settings-modal')?.classList.remove('hidden');
            const res = await window.pywebview.api.get_settings();
            if (res?.success) {
                document.getElementById('api-key-status').textContent =
                    res.data.api_key_configured
                        ? `✅ Đã cấu hình Gemini (model: ${res.data.model})`
                        : '⚠️ Chưa cấu hình Gemini API key';
                const serperStatus = document.getElementById('serper-key-status');
                if (serperStatus) {
                    serperStatus.textContent = res.data.serper_key_configured
                        ? '✅ Đã kích hoạt tìm kiếm Google Search'
                        : '⚠️ Chưa có khóa tìm kiếm (soạn thư sẽ không gắn link thực tế)';
                }
                document.getElementById('style-status').textContent =
                    res.data.style_analyzed_at
                        ? `Lần phân tích gần nhất: ${window.Utils.formatDate(res.data.style_analyzed_at)}`
                        : 'AI sẽ đọc thư đã gửi để học cách bạn viết.';
            }
        });

        document.getElementById('btn-close-settings')?.addEventListener('click', () => {
            document.getElementById('settings-modal')?.classList.add('hidden');
        });

        document.getElementById('btn-save-settings')?.addEventListener('click', async () => {
            const apiKey = document.getElementById('api-key-input')?.value;
            const serperKey = document.getElementById('serper-key-input')?.value;
            if ((apiKey || serperKey) && window.pywebview) {
                try {
                    const result = await window.pywebview.api.save_api_key(apiKey, serperKey);
                    if (result && result.success) {
                        window.Utils.showToast('Đã lưu cấu hình API thành công', 'success');
                    } else {
                        window.Utils.showToast('Lỗi lưu API: ' + (result?.error || ''), 'error');
                    }
                } catch (e) {
                    window.Utils.showToast('Lỗi: ' + e.message, 'error');
                }
            }
            document.getElementById('settings-modal')?.classList.add('hidden');
        });

        document.getElementById('btn-analyze-style')?.addEventListener('click', async () => {
            await this.analyzeStyle();
        });
    }

    async loadFolders() {
        const res = await window.pywebview.api.get_folders();
        if (!res?.success) return;
        this.state.folders = res.data || [];
        const inbox = this.state.folders.find(f => f.kind === 'inbox') || this.state.folders[0];
        this.state.currentFolderId = inbox?.entry_id || null;
        window.EmailListUI.renderFolders(this.state.folders, this.state.currentFolderId);
    }

    async switchFolder(folderId) {
        if (this.state.currentFolderId === folderId) return;
        this.state.currentFolderId = folderId;
        this.state.searchQuery = '';
        this.state.currentEmailId = null;
        const s = document.getElementById('search-input');
        if (s) s.value = '';
        window.EmailListUI.renderFolders(this.state.folders, folderId);
        window.EmailViewerUI?.showEmptyDetail();
        window.DraftEditorUI?.hideDraftPanel();
        await this.loadEmails(true);
    }

    async loadEmails(reset = false) {
        if (reset) { this.state.offset = 0; this.state.emails = []; this.state.hasMore = false; }
        window.EmailListUI.renderSkeletonLoading();
        try {
            let items = [], hasMore = false, nextOffset = this.state.offset;
            if (this.state.searchQuery) {
                const res = await window.pywebview.api.search_emails(
                    this.state.searchQuery, this.state.currentFolderId);
                if (res?.success) items = res.data || [];
            } else {
                const res = await window.pywebview.api.get_emails(
                    this.state.currentFolderId, this.state.limit, this.state.offset);
                if (res?.success) {
                    items = res.data.items || [];
                    hasMore = !!res.data.has_more;
                    nextOffset = res.data.next_offset;
                }
            }
            this.state.emails = reset ? items : this.state.emails.concat(items);
            this.state.hasMore = hasMore;
            this.state.offset = nextOffset;
            this.rerenderList();
        } catch (e) {
            window.Utils.showToast('Lỗi tải email: ' + (e.message || e), 'error');
            window.EmailListUI.renderEmptyState();
        }
    }

    rerenderList() {
        window.EmailListUI.renderEmailList(
            this.state.emails, this.state.currentEmailId, this.state.hasMore);
    }

    async loadMoreEmails() { await this.loadEmails(false); }

    async handleSearch(query) {
        this.state.searchQuery = query;
        await this.loadEmails(true);
    }

    async selectEmail(entryId) {
        this.state.currentEmailId = entryId;
        window.DraftEditorUI?.hideDraftPanel();
        try {
            const res = await window.pywebview.api.get_email_detail(entryId);
            if (!res?.success) {
                window.Utils.showToast(res?.error || 'Không tải được email', 'error');
                return;
            }
            window.EmailViewerUI.renderEmailDetail(res.data);

            // Đánh dấu đã đọc trong Outlook THẬT, không chỉ đổi state JS
            window.pywebview.api.mark_as_read(entryId);
            const item = this.state.emails.find(e => e.entry_id === entryId);
            if (item) item.is_unread = false;

            const th = await window.pywebview.api.get_conversation(entryId);
            if (th?.success) window.EmailViewerUI.renderConversationThread(th.data);
        } catch (e) {
            window.Utils.showToast('Lỗi tải email: ' + (e.message || e), 'error');
        }
    }

    async classifyVisible() {
        window.Utils.showToast('Đang quét và phân tích toàn bộ thư mục…', 'info');
        try {
            const res = await window.pywebview.api.classify_all_emails(this.state.currentFolderId);
            if (!res?.success) {
                window.Utils.showToast(res?.error || 'Lỗi phân loại', 'error');
                return;
            }
            Object.assign(window.EmailListUI.classifications, res.data?.classifications || {});
            this.rerenderList();
            window.Utils.showToast(`✅ Đã phân tích ${res.data?.scanned || 0} thư`, 'success');
        } catch (e) {
            window.Utils.showToast('Lỗi phân loại: ' + (e.message || e), 'error');
        }
    }

    async generateDraft(instruction, replyAll = true, emailType = '') {
        if (!this.state.currentEmailId) return;
        if (!window.pywebview) return;

        if (window.DraftEditorUI) window.DraftEditorUI.setLoading(true);

        try {
            const result = await window.pywebview.api.generate_reply(
                this.state.currentEmailId,
                instruction,
                replyAll !== false,
                emailType
            );

            if (result && result.success && result.data) {
                const d = result.data;
                window.DraftEditorUI?.showDraftPreview(d.html_body);
                const label = d.email_type_title || d.email_type;
                window.Utils.showToast(
                    label
                        ? (d.email_type_auto
                            ? `Đã tạo nháp — AI chọn mẫu "${label}"`
                            : `Đã tạo nháp theo mẫu "${label}"`)
                        : 'Đã tạo email nháp',
                    'success');
            } else {
                window.Utils.showToast('Lỗi tạo nháp: ' + (result?.error || 'Unknown'), 'error');
            }
        } catch (error) {
            console.error("Generate draft error:", error);
            window.Utils.showToast('Lỗi khi tạo nháp AI', 'error');
        } finally {
            if (window.DraftEditorUI) window.DraftEditorUI.setLoading(false);
        }
    }

    async refineDraft(currentHtml, feedback) {
        if (!window.pywebview) return;

        if (window.DraftEditorUI) window.DraftEditorUI.setLoading(true);

        try {
            const result = await window.pywebview.api.refine_draft(currentHtml, feedback);

            if (result && result.success && result.data) {
                window.DraftEditorUI?.showDraftPreview(result.data.html_body);
                window.Utils.showToast('Đã chỉnh sửa nháp', 'success');
                const refineInput = document.getElementById('refine-instruction');
                if (refineInput) refineInput.value = '';
            } else {
                window.Utils.showToast('Lỗi chỉnh sửa: ' + (result?.error || ''), 'error');
            }
        } catch (error) {
            console.error("Refine draft error:", error);
        } finally {
            if (window.DraftEditorUI) window.DraftEditorUI.setLoading(false);
        }
    }

    async saveDraft(htmlContent, isReplyAll) {
        if (!this.state.currentEmailId) return;
        if (!window.pywebview) return;

        try {
            const result = await window.pywebview.api.save_draft(
                this.state.currentEmailId,
                htmlContent,
                isReplyAll
            );

            if (result && result.success) {
                window.Utils.showToast('✅ Đã lưu vào Thư nháp Outlook', 'success');
                window.DraftEditorUI?.hideDraftPanel();
            } else {
                window.Utils.showToast('Lỗi lưu nháp: ' + (result?.error || ''), 'error');
            }
        } catch (error) {
            console.error("Save draft error:", error);
            window.Utils.showToast('Lỗi khi lưu thư nháp', 'error');
        }
    }

    async checkStyleProfile() {
        if (!window.pywebview) return;

        try {
            const result = await window.pywebview.api.get_style_profile();
            const statusEl = document.querySelector('#profile-status span');
            if (result && result.success) {
                this.state.styleProfileLoaded = true;
                if (statusEl) statusEl.textContent = '✅ Đã có hồ sơ văn phong';
            } else {
                if (statusEl) statusEl.textContent = '⚠️ Chưa phân tích văn phong';
            }
        } catch (e) {
            console.error("Check style profile error:", e);
        }
    }

    async analyzeStyle() {
        if (!window.pywebview) return;

        window.Utils.showToast('Đang quét mọi kho thư trong Outlook. Có thể mất 10–30 phút…', 'info');

        try {
            const result = await window.pywebview.api.export_sent_emails();
            if (result && result.success) {
                const d = result.data || {};
                this.state.styleProfileLoaded = true;
                window.Utils.showToast(
                    `✅ Đã xuất ${d.unique || 0} thư thành ${d.files || 0} tệp. ` +
                    `Mở ${d.dir || 'xuat_thu/'} bằng Antigravity để viết hướng dẫn văn phong.`,
                    'success');
                const statusEl = document.querySelector('#profile-status span');
                if (statusEl) statusEl.textContent = '✅ Đã có hồ sơ văn phong';
            } else {
                window.Utils.showToast('Lỗi: ' + (result?.error || 'Không xuất được'), 'error');
            }
        } catch (error) {
            console.error("Export sent emails error:", error);
            window.Utils.showToast('Lỗi khi xuất thư', 'error');
        }
    }
}

// Start the application
window.App = new AppController();
