class DraftEditor {
    constructor() {
        this.panel = document.getElementById('draft-panel');
        this.btnClose = document.getElementById('btn-close-draft');
        this.btnGenerate = document.getElementById('btn-generate-draft');
        this.btnSave = document.getElementById('btn-save-draft');
        this.btnRefine = document.getElementById('btn-refine-draft');
        
        this.instructionInput = document.getElementById('ai-instruction');
        this.refineInput = document.getElementById('refine-instruction');
        this.editorArea = document.getElementById('draft-editor');
        this.resultArea = document.getElementById('draft-result-area');
        this.typeSelect = document.getElementById('ai-email-type');
        this.typeHint = document.getElementById('ai-type-hint');

        this.isReplyAll = false;
        this.typesLoaded = false;

        this.bindEvents();
    }

    /* Danh sách mẫu thư đọc từ kien_thuc/loai_thu/. Nạp một lần khi mở bảng chứ
       không nạp lúc khởi động: pywebview chưa sẵn sàng ở thời điểm đó. */
    async loadTypes() {
        if (this.typesLoaded || !this.typeSelect || !window.pywebview) return;
        this.typesLoaded = true;
        try {
            const res = await window.pywebview.api.list_email_types();
            const types = res?.data?.types || [];
            if (!types.length) {
                this.typeSelect.disabled = true;
                this.typeHint.textContent =
                    'Chưa có mẫu thư. Chạy XUAT_THU.bat rồi để Antigravity viết vào kien_thuc/loai_thu/.';
                return;
            }
            for (const t of types) {
                const opt = document.createElement('option');
                opt.value = t.name;
                opt.textContent = t.title;
                this.typeSelect.appendChild(opt);
            }
            this.typeHint.textContent = `${types.length} mẫu thư sẵn có.`;
        } catch (e) {
            this.typesLoaded = false;      // cho phép thử lại lần mở sau
            this.typeSelect.disabled = true;
            this.typeHint.textContent = 'Không nạp được danh sách mẫu thư.';
        }
    }

    selectedType() {
        return this.typeSelect ? this.typeSelect.value : '';
    }

    bindEvents() {
        this.btnClose?.addEventListener('click', () => this.hideDraftPanel());
        
        this.btnGenerate?.addEventListener('click', () => {
            const instruction = this.instructionInput.value.trim();
            if (!instruction) {
                window.Utils.showToast('Vui lòng nhập chỉ dẫn cho AI', 'warning');
                return;
            }
            if (window.App) window.App.generateDraft(instruction, this.isReplyAll,
                                                     this.selectedType());
        });

        this.btnRefine?.addEventListener('click', () => {
            const feedback = this.refineInput.value.trim();
            const currentDraft = this.editorArea.innerHTML;
            if (!feedback) return;
            if (window.App) window.App.refineDraft(currentDraft, feedback);
        });

        this.btnSave?.addEventListener('click', () => {
            const finalContent = this.editorArea.innerHTML;
            if (window.App) window.App.saveDraft(finalContent, this.isReplyAll);
        });
    }

    showDraftPanel(isReplyAll = false) {
        this.isReplyAll = isReplyAll;
        this.panel.classList.remove('hidden');
        this.panel.style.animation = 'slideIn 0.3s ease forwards';
        this.loadTypes();
        this.instructionInput.focus();
        
        // Reset state
        this.resultArea.classList.add('hidden');
        this.editorArea.innerHTML = '';
        this.refineInput.value = '';
    }

    hideDraftPanel() {
        this.panel.classList.add('hidden');
    }

    setLoading(isLoading) {
        const btnText = this.btnGenerate.querySelector('.btn-text');
        const loader = this.btnGenerate.querySelector('.loader');
        
        this.btnSave.disabled = isLoading;
        this.btnRefine.disabled = isLoading;

        if (isLoading) {
            btnText.textContent = 'Đang tạo...';
            loader.classList.remove('hidden');
            this.btnGenerate.disabled = true;
            this.instructionInput.disabled = true;
        } else {
            btnText.textContent = 'Tạo nháp';
            loader.classList.add('hidden');
            this.btnGenerate.disabled = false;
            this.instructionInput.disabled = false;
        }
    }

    showDraftPreview(htmlContent) {
        this.resultArea.classList.remove('hidden');
        this.editorArea.innerHTML = htmlContent;
        // Scroll to result
        this.resultArea.scrollIntoView({ behavior: 'smooth' });
    }
}
window.DraftEditorUI = new DraftEditor();
