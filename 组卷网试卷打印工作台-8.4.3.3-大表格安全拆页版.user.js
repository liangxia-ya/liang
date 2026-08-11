// ==UserScript==
// @name         组卷网试卷打印工作台
// @version      8.4.3.3
// @description  基于8.4.3.2：保留答案行级连续切页，新增HTML大表格按安全<tr>边界拆页并避开rowspan，减少整表推页造成的大块空白
// @author       nuym, WorkingFishQ, xiaohuya
// @match        *://zujuan.xkw.com/*
// @icon         https://zujuan.xkw.com/favicon.ico
// @grant        GM_notification
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// @connect      xkw.com
// @connect      *.xkw.com
// @require      https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.js
// @run-at       document-end
// @license      GNU Affero General Public License v3.0
// ==/UserScript==

(function () {
    'use strict';

    const Config = {
        fontMode: 'original',
        customFontFamily: '"Microsoft YaHei", "PingFang SC", "Noto Sans SC", Arial, sans-serif',
        adSelectors: ['.aside-pop.activity-btn', '.ai-entry.fixed'],
    };

    // ==========================================
    // 全局样式
    // ==========================================
    GM_addStyle(`
        .zujuanjs-float-print-btn {
            position: fixed !important; bottom: 30px !important; right: 30px !important;
            width: 56px !important; height: 56px !important; border-radius: 50% !important;
            background: #1677ff !important; color: #fff !important; border: none !important;
            box-shadow: 0 4px 14px rgba(22, 119, 255, 0.4) !important; cursor: grab !important;
            z-index: 99999 !important; display: flex !important; align-items: center !important;
            justify-content: center !important; transition: background .2s ease, box-shadow .2s ease !important; font-size: 22px !important;
            touch-action: none !important; user-select: none !important; -webkit-user-select: none !important;
        }
        .zujuanjs-float-print-btn:hover { background: #4096ff !important; box-shadow: 0 6px 20px rgba(22, 119, 255, 0.5) !important; }
        .zujuanjs-float-print-btn:active, .zujuanjs-float-print-btn.zujuanjs-dragging { cursor: grabbing !important; box-shadow: 0 2px 8px rgba(22, 119, 255, 0.3) !important; }
        .zujuanjs-float-print-btn-text { display: none; }
        @media (min-width: 768px) {
            .zujuanjs-float-print-btn { width: auto !important; height: auto !important; border-radius: 28px !important; padding: 12px 24px !important; font-size: 15px !important; font-weight: 500 !important; letter-spacing: 0 !important; }
            .zujuanjs-float-print-btn-icon { margin-right: 6px; }
            .zujuanjs-float-print-btn-text { display: inline; }
        }

        /* ===== 打印设置对话框 - 紧凑布局 ===== */
        .print-dialog-container { text-align: left; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; padding: 0 2px; }
        .print-dialog-section { margin-bottom: 14px; }
        .print-dialog-section:last-child { margin-bottom: 0; }
        .print-dialog-label { display: block; font-size: 13px; font-weight: 600; color: #1f1f1f; margin-bottom: 8px; }
        .print-dialog-input {
            width: 100%; height: 36px; padding: 0 12px; border: 1px solid #d9d9d9;
            border-radius: 6px; font-size: 13px; outline: none; transition: all 0.2s; box-sizing: border-box;
        }
        .print-dialog-input:focus { border-color: #1677ff; box-shadow: 0 0 0 2px rgba(22,119,255,0.15); }

        .print-option-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; }
        .print-option-card {
            display: flex; align-items: center; gap: 6px; padding: 8px 12px;
            border: 1px solid #e8e8e8; border-radius: 6px; cursor: pointer; background: #fafafa; transition: all 0.2s ease;
        }
        .print-option-card:hover { background: #f0f5ff; border-color: #91caff; }
        .print-option-card.active { background: #e6f4ff; border-color: #1677ff; }
        .print-option-card input[type="radio"] { width: 14px; height: 14px; margin: 0; accent-color: #1677ff; flex-shrink: 0; }
        .print-option-card span { font-size: 12px; color: #434343; }

        .print-row-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; }
        .print-field { min-width: 0; }
        .print-field-label { font-size: 11px; color: #888; margin-bottom: 4px; line-height: 1.2; }

        .print-custom-select { position: relative; width: 100%; }
        .print-custom-select-trigger {
            display: flex; align-items: center; justify-content: space-between;
            padding: 0 10px; height: 36px; background: #fff;
            border: 1px solid #d9d9d9; border-radius: 6px; cursor: pointer;
            font-size: 12px; color: #333; transition: all 0.2s ease; user-select: none; box-sizing: border-box;
        }
        .print-custom-select-trigger:hover { border-color: #1677ff; }
        .print-custom-select.open .print-custom-select-trigger { border-color: #1677ff; box-shadow: 0 0 0 2px rgba(22,119,255,0.15); }
        .print-custom-select-trigger > span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; min-width: 0; }

        .print-custom-select-arrow {
            width: 7px; height: 7px; border-left: 2px solid #999; border-bottom: 2px solid #999;
            transform: rotate(-45deg); transition: transform 0.2s ease; flex-shrink: 0; margin-left: 6px;
        }
        .print-custom-select.open .print-custom-select-arrow { transform: rotate(135deg); }

        .print-custom-select-dropdown {
            position: absolute; top: calc(100% + 4px); left: 0; right: 0; background: #fff;
            border: 1px solid #e8e8e8; border-radius: 6px; box-shadow: 0 6px 16px rgba(0,0,0,0.08);
            opacity: 0; visibility: hidden; pointer-events: none; z-index: 99999;
            transition: all 0.2s ease; max-height: 200px; overflow-y: auto; padding: 4px 0;
        }
        .print-custom-select.open .print-custom-select-dropdown { opacity: 1; visibility: visible; pointer-events: auto; }

        .print-custom-select-option {
            padding: 7px 10px; font-size: 12px; color: #333; cursor: pointer; transition: background 0.15s;
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        }
        .print-custom-select-option:hover { background: #f5f5f5; }
        .print-custom-select-option.selected { background: #e6f4ff; color: #1677ff; font-weight: 500; }

        .print-preview-box {
            padding: 12px 14px; border: 1px solid #e8e8e8; border-radius: 6px;
            background: #f9f9f9; min-height: 40px; line-height: 1.6; font-size: 14px;
            color: #333; overflow: hidden;
        }

        .swal2-popup.print-dialog-popup { padding: 18px 22px 16px !important; }
        .swal2-popup.print-dialog-popup .swal2-title { font-size: 18px !important; margin: 0 0 12px !important; padding: 0 !important; }
        .swal2-popup.print-dialog-popup .swal2-html-container { margin: 0 !important; padding: 0 !important; font-size: inherit !important; }
        .swal2-popup.print-dialog-popup .swal2-actions { margin: 14px 0 0 !important; gap: 8px !important; }
        .swal2-popup.print-dialog-popup .swal2-confirm,
        .swal2-popup.print-dialog-popup .swal2-cancel { padding: 8px 22px !important; font-size: 13px !important; font-weight: 500 !important; border-radius: 6px !important; box-shadow: none !important; }

        /* ===== 5.4 响应式精简设置面板 ===== */
        .swal2-popup.print-dialog-popup {
            width: min(700px, calc(100vw - 24px)) !important;
            max-height: calc(100vh - 24px) !important;
            padding: 20px 22px 16px !important;
            border-radius: 8px !important;
        }
        .swal2-popup.print-dialog-popup .swal2-title {
            margin-bottom: 16px !important;
            text-align: left !important;
            color: #171717 !important;
            font-size: 19px !important;
        }
        .swal2-popup.print-dialog-popup .swal2-html-container { overflow: visible !important; }
        .print-dialog-container { padding: 0; }
        .print-dialog-section { margin: 0; padding: 12px 0; border-top: 1px solid #ededed; }
        .print-dialog-section:first-child { padding-top: 0; border-top: 0; }
        .print-dialog-label { margin-bottom: 7px; color: #262626; font-size: 12px; }
        .print-dialog-input,
        .print-custom-select-trigger { height: 34px; border-color: #d4d4d4; border-radius: 5px; }
        .print-option-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 2px;
            padding: 3px;
            border-radius: 6px;
            background: #f0f1f2;
        }
        .print-option-card {
            justify-content: center;
            min-height: 30px;
            padding: 3px 8px;
            border: 0;
            border-radius: 4px;
            background: transparent;
        }
        .print-option-card:hover { border: 0; background: #fff; }
        .print-option-card.active { border: 0; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.12); }
        .print-option-card input[type="radio"] { position: absolute; opacity: 0; pointer-events: none; }
        .print-option-card span { color: #525252; font-size: 12px; }
        .print-option-card.active span { color: #1268d3; font-weight: 600; }
        .print-row-3 { gap: 10px; }
        .print-row-4 { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; }
        .print-field-label { margin-bottom: 5px; color: #737373; font-size: 11px; }
        .print-layout-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); height: 34px; box-sizing: border-box; }
        .print-preview-wrap { display: grid; grid-template-columns: 46px minmax(0, 1fr); align-items: center; gap: 10px; }
        .print-preview-caption { color: #737373; font-size: 11px; }
        .print-preview-box {
            min-height: 0;
            padding: 8px 10px;
            border: 0;
            border-left: 2px solid #1677ff;
            border-radius: 0;
            background: #f7f8fa;
            font-size: 13px;
            line-height: 1.5;
        }
        .swal2-popup.print-dialog-popup .swal2-actions { width: 100%; justify-content: flex-end; margin-top: 14px !important; }
        .swal2-popup.print-dialog-popup .swal2-confirm,
        .swal2-popup.print-dialog-popup .swal2-cancel { min-height: 34px; padding: 6px 18px !important; border-radius: 5px !important; }
        .swal2-popup.print-dialog-popup .swal2-cancel { background: #eef0f2 !important; color: #333 !important; }
        body.swal2-shown .zujuanjs-float-print-btn { display: none !important; }

        @media (max-width: 620px) {
            .swal2-popup.print-dialog-popup { padding: 16px 14px 12px !important; }
            .print-row-4 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .print-row-3 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .print-row-3 .print-field:first-child { grid-column: 1 / -1; }
        }
        @media (max-width: 390px) {
            .print-option-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .print-preview-wrap { grid-template-columns: 1fr; gap: 5px; }
            .print-dialog-section { padding-top: 9px; padding-bottom: 9px; }
        }

        /* 预览遮罩层 */
        #zujuanjs-preview-overlay {
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; z-index: 100000; background: #000;
        }
        #zujuanjs-preview-overlay iframe { width: 100%; height: 100%; border: none; }
        #zujuanjs-busy-overlay {
            position: fixed; inset: 0; z-index: 100001; display: flex; align-items: center; justify-content: center;
            background: rgba(15, 23, 42, 0.46); backdrop-filter: blur(2px);
        }
        #zujuanjs-busy-overlay .zujuanjs-busy-card {
            min-width: 220px; max-width: calc(100vw - 40px); padding: 18px 22px; border-radius: 10px;
            background: #fff; color: #1f2937; box-shadow: 0 18px 50px rgba(0,0,0,0.28);
            font: 14px/1.6 -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
            text-align: center;
        }
        #zujuanjs-busy-overlay .zujuanjs-busy-spinner {
            width: 28px; height: 28px; margin: 0 auto 10px; border: 3px solid #dbeafe;
            border-top-color: #1677ff; border-radius: 50%; animation: zujuanjs-spin .75s linear infinite;
        }
        @keyframes zujuanjs-spin { to { transform: rotate(360deg); } }
    `);

    class PaperPrinter {
        constructor() {
            this.groupExamAnswerSnapshots = new Map();
            this.groupExamAnswerPreparation = null;
            this.migrateRefinedLayout830();
            this.migrateRefinedLayout831();
            this.migratePageNumberDefaults843();
            this.init();
        }

        migrateRefinedLayout830() {
            if (String(GM_getValue('zujuanjsLayoutMigration830', '')) === 'done') return;

            // 只迁移“仍然等于 8.2.x 默认值”的用户，已经手动调过排版的人保持原设置。
            const legacy = {
                questionSize: '16px', questionLineHeight: '1.5', answerSize: '16px', answerLineHeight: '1.5',
                titleSize: '24px', pageSize: '12px', pageBold: 'true', pageMargins: '18,15,22,15',
                questionSpacing: '10', paragraphSpacing: '8', numberGap: '0.55'
            };
            const current = {
                questionSize: String(GM_getValue('questionSize', legacy.questionSize)),
                questionLineHeight: String(GM_getValue('questionLineHeight', legacy.questionLineHeight)),
                answerSize: String(GM_getValue('answerSize', legacy.answerSize)),
                answerLineHeight: String(GM_getValue('answerLineHeight', legacy.answerLineHeight)),
                titleSize: String(GM_getValue('titleSize', legacy.titleSize)),
                pageSize: String(GM_getValue('pageSize', legacy.pageSize)),
                pageBold: String(GM_getValue('pageBold', true)),
                pageMargins: String(GM_getValue('pageMargins', legacy.pageMargins)),
                questionSpacing: String(GM_getValue('questionSpacing', legacy.questionSpacing)),
                paragraphSpacing: String(GM_getValue('paragraphSpacing', legacy.paragraphSpacing)),
                numberGap: String(GM_getValue('numberGap', legacy.numberGap))
            };
            const legacyLike = Object.keys(legacy).every(key => current[key] === legacy[key]);
            if (legacyLike) {
                const refined = {
                    questionSize: '14px', questionLineHeight: '1.4', answerSize: '14px', answerLineHeight: '1.35',
                    titleSize: '22px', pageSize: '10px', pageBold: false, pageMargins: '15,17,18,17',
                    questionSpacing: '6', paragraphSpacing: '3', numberGap: '0.45', layoutPreset: 'refined'
                };
                Object.entries(refined).forEach(([key, value]) => GM_setValue(key, value));
            }
            GM_setValue('zujuanjsLayoutMigration830', 'done');
        }

        migrateRefinedLayout831() {
            if (String(GM_getValue('zujuanjsLayoutMigration831', '')) === 'done') return;

            // 8.3.1：标准精排正文改为 14px。只迁移仍处于 8.3.0“标准精排”默认值的用户，
            // 已经手动选择其他字号/版式的人不强制覆盖。
            const preset = String(GM_getValue('layoutPreset', 'refined'));
            const questionSize = String(GM_getValue('questionSize', '14px'));
            const lineHeight = String(GM_getValue('questionLineHeight', '1.4'));
            if (preset === 'refined' && questionSize === '15px' && lineHeight === '1.4') {
                GM_setValue('questionSize', '14px');
            }
            GM_setValue('zujuanjsLayoutMigration831', 'done');
        }

        migratePageNumberDefaults843() {
            if (String(GM_getValue('zujuanjsPageNumberMigration843', '')) === 'done') return;
            // 8.4.3：默认使用题目/答案分段页码，且页脚只显示数字。
            // 这是一次性默认迁移；之后用户仍可在设置中切回整卷连续计数。
            GM_setValue('pageNumberScope', 'sectioned');
            GM_setValue('pageNumberFormat', 'current-total');
            GM_setValue('zujuanjsPageNumberMigration843', 'done');
        }

        init() {
            this.applyFont();
            this.autoCheckIn();
            this.createFloatingButton();
            this.startUiKeeper();
            this.installShortcuts();
            if (document.body) this.startAdRemover();
            else document.addEventListener('DOMContentLoaded', () => this.startAdRemover());

            window.addEventListener('message', (e) => {
                if (e.data && e.data.type === 'closeZujuanPreview') {
                    const overlay = document.getElementById('zujuanjs-preview-overlay');
                    if (overlay) overlay.remove();
                } else if (e.data && e.data.type === 'saveZujuanPreviewPreference') {
                    if (e.data.key === 'previewLayout' && ['single', 'double'].includes(e.data.value)) {
                        GM_setValue('previewLayout', e.data.value);
                    }
                    if (e.data.key === 'previewZoom' && (e.data.value === 'auto' || /^\d+(\.\d+)?$/.test(e.data.value))) {
                        GM_setValue('previewZoom', e.data.value);
                    }
                    if (e.data.key === 'editorPanelWidth' && /^\d+(\.\d+)?$/.test(e.data.value)) {
                        GM_setValue('editorPanelWidth', Math.max(280, Math.min(520, Number(e.data.value))));
                    }
                    if (e.data.key === 'editorPanelTab' && ['document', 'page'].includes(e.data.value)) {
                        GM_setValue('editorPanelTab', e.data.value);
                    }
                    if (e.data.key === 'editorOpen' && ['true', 'false'].includes(e.data.value)) {
                        GM_setValue('editorOpen', e.data.value === 'true');
                    }
                } else if (e.data && e.data.type === 'saveZujuanPrintSettings') {
                    this.savePreviewSettings(e.data.settings);
                } else if (e.data && e.data.type === 'rebuildZujuanPreview') {
                    this.openPreviewWithSettings(e.data.settings);
                } else if (e.data && e.data.type === 'exportZujuanWord') {
                    this.exportWordDocument(e.data, e.source);
                } else if (e.data && e.data.type === 'zujuanFetchResource') {
                    this.fetchResourceForPreview(e.data, e.source);
                } else if (e.data && e.data.type === 'setZujuanTitleOverride') {
                    this.savePaperTitleOverride(e.data.title);
                } else if (e.data && e.data.type === 'clearZujuanTitleOverride') {
                    this.clearPaperTitleOverride();
                }
            });
        }

        applyFont() {
            if (Config.fontMode === 'custom') GM_addStyle(`body, * { font-family: ${Config.customFontFamily} !important; }`);
        }

        getPaperIdentityKey() {
            const path = String(location.pathname || '/').replace(/\/+$/, '') || '/';
            const idMatch = path.match(/(?:paper|exam|test|zujuan|detail)[^0-9]{0,8}(\d{4,})/i) || path.match(/\/(\d{5,})(?:\/|$)/);
            return idMatch ? ('paper:' + idMatch[1]) : ('url:' + path);
        }

        normalizePaperTitleText(value) {
            return String(value || '')
                .replace(/[\u200b-\u200d\ufeff]/g, '')
                .replace(/\u00a0/g, ' ')
                .replace(/[ \t]+/g, ' ')
                .replace(/^\s+|\s+$/g, '');
        }

        isPaperMetaLine(line) {
            const text = this.normalizePaperTitleText(line);
            if (!text) return true;
            return /^(?:整体难度|考查范围|知识点|适用年级|适用地区|地区|年级|题型|来源|更新时间|上传时间|浏览量|下载量|组卷次数|试卷类型|难度)[：:]/.test(text)
                || /^\d{4}[./-]\d{1,2}[./-]\d{1,2}(?:\s+\d{1,2}:\d{2})?$/.test(text)
                || /^\d+(?:\.\d+)?\s*(?:次|人浏览|浏览|下载)$/.test(text)
                || /^(?:广东|北京|上海|天津|重庆|河北|河南|山东|山西|陕西|江苏|浙江|安徽|福建|江西|湖北|湖南|四川|贵州|云南|辽宁|吉林|黑龙江|甘肃|青海|海南|内蒙古|广西|西藏|宁夏|新疆|香港|澳门|台湾)(?:省|市|自治区)?\s*(?:七|八|九|高[一二三]|[一二三四五六])?年级\s*(?:中考|高考|期中|期末|月考|联考)?\s*(?:真题|模拟题|试题)?$/.test(text);
        }

        trimPaperTitleMetadata(value) {
            let text = this.normalizePaperTitleText(value);
            if (!text) return '';
            const hardStops = [
                /\s+整体难度[：:]/,
                /\s+考查范围[：:]/,
                /\s+适用年级[：:]/,
                /\s+适用地区[：:]/,
                /\s+\d{4}[./-]\d{1,2}[./-]\d{1,2}(?=\s|$)/,
                /\s+\d+(?:\.\d+)?\s*次(?=\s|$)/
            ];
            for (const pattern of hardStops) {
                const match = text.match(pattern);
                if (match && typeof match.index === 'number') text = text.slice(0, match.index).trim();
            }
            // 组卷网常把“地区 年级 考试类型”作为标题下方标签；若它们被拼到主标题后，优先截断。
            const tagTail = text.match(/^(.*?(?:试卷|真题|模拟题|测试题|练习题|中考|高考))\s+(?:广东|北京|上海|天津|重庆|河北|河南|山东|山西|陕西|江苏|浙江|安徽|福建|江西|湖北|湖南|四川|贵州|云南|辽宁|吉林|黑龙江|甘肃|青海|海南|内蒙古|广西|西藏|宁夏|新疆|香港|澳门|台湾)(?:省|市|自治区)?(?:\s+.*)?$/);
            if (tagTail && tagTail[1]) text = tagTail[1].trim();
            return text.replace(/^[\-—–_|·•\s]+|[\-—–_|·•\s]+$/g, '');
        }

        scorePaperTitleCandidate(value, selectorRank = 0, order = 0) {
            const text = this.trimPaperTitleMetadata(value);
            if (!text || this.isPaperMetaLine(text)) return -9999;
            let score = 0;
            const len = text.length;
            if (len >= 8 && len <= 60) score += 30;
            else if (len <= 90) score += 15;
            else score -= Math.min(80, len - 90);
            if (/20\d{2}年/.test(text)) score += 12;
            if (/(?:语文|数学|英语|物理|化学|生物|历史|地理|政治|道德与法治|科学)/.test(text)) score += 16;
            if (/(?:中考|高考|真题|试卷|模拟|月考|期中|期末|联考|摸底|测试|练习)/.test(text)) score += 24;
            if (/(?:整体难度|考查范围|\d{4}[./-]\d{1,2}[./-]\d{1,2}|\d+次)/.test(text)) score -= 70;
            score += Math.max(0, 24 - selectorRank * 3);
            score += Math.max(0, 8 - order * 0.1);
            return score;
        }

        getPaperTitleInfo() {
            const selectors = [
                '.paper-title .main-title', '.paper-title .title', '.paper-title h1', '.paper-title h2',
                '.exam-title .main-title', '.exam-title h1', '.exam-title', '.paper-name', '.paper-title',
                '.title-box h1', '.title-box .main-title', '.title-box', 'main h1', 'h1'
            ];
            const candidates = [];
            const seen = new Set();
            const push = (raw, selector, rank, order, source = 'text') => {
                const normalized = this.normalizePaperTitleText(raw);
                if (!normalized) return;
                const lines = normalized.split(/\r?\n/).map(line => this.normalizePaperTitleText(line)).filter(Boolean);
                const values = lines.length > 1 ? lines : [normalized];
                values.forEach((value, lineIndex) => {
                    const title = this.trimPaperTitleMetadata(value);
                    if (!title) return;
                    const key = title.toLowerCase();
                    if (seen.has(key)) return;
                    seen.add(key);
                    candidates.push({
                        title,
                        raw: value,
                        selector,
                        source,
                        score: this.scorePaperTitleCandidate(title, rank, order + lineIndex / 10),
                        order: order + lineIndex / 10
                    });
                });
            };

            let order = 0;
            selectors.forEach((selector, rank) => {
                document.querySelectorAll(selector).forEach(element => {
                    const rect = element.getBoundingClientRect?.();
                    if (rect && rect.width === 0 && rect.height === 0 && element.closest('[hidden], [aria-hidden="true"]')) return;
                    push(element.getAttribute('title'), selector, rank, order++, 'title-attr');
                    // innerText 保留可视换行，比 textContent 更不容易把标题和元数据粘在一起。
                    push(element.innerText || element.textContent, selector, rank, order++, 'visible-text');
                    const directText = Array.from(element.childNodes || [])
                        .filter(node => node.nodeType === Node.TEXT_NODE)
                        .map(node => node.textContent || '')
                        .join(' ');
                    push(directText, selector, rank, order++, 'direct-text');
                });
            });

            candidates.sort((a, b) => b.score - a.score || a.order - b.order || a.title.length - b.title.length);
            const best = candidates.find(item => item.score > -100) || { title: '试卷', selector: '', source: 'fallback', score: 0, raw: '' };
            return {
                title: best.title || '试卷',
                selector: best.selector || '',
                source: best.source || '',
                confidence: Math.max(0, Math.min(100, Math.round(best.score))),
                raw: best.raw || '',
                candidates: candidates.slice(0, 8)
            };
        }

        getPaperTitleOverride() {
            const map = GM_getValue('paperTitleOverrides', {}) || {};
            const key = this.getPaperIdentityKey();
            return Object.prototype.hasOwnProperty.call(map, key) ? String(map[key] ?? '') : null;
        }

        savePaperTitleOverride(title) {
            const map = GM_getValue('paperTitleOverrides', {}) || {};
            map[this.getPaperIdentityKey()] = String(title ?? '');
            GM_setValue('paperTitleOverrides', map);
        }

        clearPaperTitleOverride() {
            const map = GM_getValue('paperTitleOverrides', {}) || {};
            delete map[this.getPaperIdentityKey()];
            GM_setValue('paperTitleOverrides', map);
        }

        getPaperTitle() {
            const override = this.getPaperTitleOverride();
            return override !== null ? override : this.getPaperTitleInfo().title;
        }

        isGroupExamCenterPage() {
            return /^\/zujuan\/?$/i.test(location.pathname) || Boolean(document.querySelector('.group-exam-page .paper-body'));
        }

        getPaperQuestionNodes() {
            // 8.3：组卷中心同时存在 .ques-item 与内部 .tk-quest-item.quesroot。
            // 旧版把两层都当成独立题目，因此同一道题会打印两遍。组卷中心只认外层 .ques-item；
            // 普通试卷详情页才使用 .tk-quest-item.quesroot。最后再按真实 quesdiv/QID 做一次去重。
            const selectors = this.isGroupExamCenterPage()
                ? ['.group-exam-page .paper-body .ques-item', '.paper-body .ques-item']
                : ['.tk-quest-item.quesroot'];
            const raw = [];
            const rawSet = new Set();
            selectors.forEach(selector => document.querySelectorAll(selector).forEach(node => {
                if (!rawSet.has(node)) { rawSet.add(node); raw.push(node); }
            }));
            raw.sort((a, b) => {
                if (a === b) return 0;
                const pos = a.compareDocumentPosition(b);
                return pos & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
            });

            const result = [];
            const seenWraps = new Set();
            const seenIds = new Set();
            raw.forEach(node => {
                const wrap = node.matches?.('.wrapper.quesdiv') ? node : (node.querySelector?.('.wrapper.quesdiv') || node);
                const stableId = [
                    node.getAttribute?.('data-qid'), node.dataset?.qid, node.getAttribute?.('data-id'), node.dataset?.id,
                    wrap.getAttribute?.('data-qid'), wrap.dataset?.qid, wrap.getAttribute?.('data-id'), wrap.dataset?.id,
                    node.id, wrap.id
                ].map(value => String(value || '').trim()).find(Boolean) || '';
                if (seenWraps.has(wrap)) return;
                if (stableId && seenIds.has(stableId)) return;
                seenWraps.add(wrap);
                if (stableId) seenIds.add(stableId);
                result.push(node);
            });
            return result;
        }

        findSectionNodeForQuestion(questionNode) {
            if (!questionNode) return null;
            if (this.isGroupExamCenterPage()) {
                let current = questionNode;
                for (let depth = 0; depth < 6 && current && current !== document.body; depth++) {
                    let sibling = current.previousElementSibling;
                    while (sibling) {
                        if (sibling.matches?.('.questype-head')) return sibling;
                        const nested = sibling.querySelector?.('.questype-head:last-of-type');
                        if (nested) return nested;
                        if (sibling.querySelector?.('.ques-item, .tk-quest-item.quesroot')) break;
                        sibling = sibling.previousElementSibling;
                    }
                    current = current.parentElement;
                }
                return null;
            }

            // 普通试卷详情页不能再全局抓所有 .sec-title：页面中的目录/导航副本也会使用这个类，
            // 旧版因此会在“答案与解析”前重新打印一串“一、二、三……”。
            // 这里从每一道真实题目向上/向前寻找结构上真正隶属于它的大题标题，置信不足宁可不加标题。
            let current = questionNode;
            for (let depth = 0; depth < 5 && current && current !== document.body; depth++) {
                const parent = current.parentElement;
                if (!parent || parent === document.body) break;
                const children = Array.from(parent.children);
                const currentIndex = children.indexOf(current);
                if (currentIndex >= 0) {
                    for (let i = currentIndex - 1, hops = 0; i >= 0 && hops < 6; i--, hops++) {
                        const sibling = children[i];
                        if (sibling.matches?.('.sec-title')) return sibling;
                        // 允许标题被一层轻量容器包裹，但遇到另一道题就停止回溯，防止跨大题误归属。
                        if (sibling.matches?.('.tk-quest-item.quesroot') || sibling.querySelector?.('.tk-quest-item.quesroot')) break;
                        const nestedTitles = sibling.querySelectorAll?.('.sec-title') || [];
                        if (nestedTitles.length === 1) return nestedTitles[0];
                    }
                }
                // 常见结构：.ques-type > .sec-title + .ques-list > .tk-quest-item
                const directTitles = Array.from(parent.children).filter(child => child.matches?.('.sec-title'));
                if (directTitles.length === 1) {
                    const title = directTitles[0];
                    const relation = title.compareDocumentPosition(questionNode);
                    if (relation & Node.DOCUMENT_POSITION_FOLLOWING) return title;
                }
                current = parent;
            }
            return null;
        }

        getPaperSectionNodes() {
            if (this.isGroupExamCenterPage()) {
                const root = document.querySelector('.group-exam-page .paper-body, .paper-body') || document;
                return Array.from(root.querySelectorAll('.questype-head'));
            }
            const set = new Set();
            this.getPaperQuestionNodes().forEach(question => {
                const section = this.findSectionNodeForQuestion(question);
                if (section) set.add(section);
            });
            return Array.from(set).sort((a, b) => {
                if (a === b) return 0;
                const pos = a.compareDocumentPosition(b);
                return pos & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
            });
        }

        getSectionTitleText(node) {
            if (!(node instanceof Element)) return '';
            const clone = node.cloneNode(true);
            // 组卷中心的 questype-head 往往把“序号”和“标题文字”拆成多个节点。
            // 旧版只读取第一个 span，因而“一、积累运用”会只剩“一、”。
            clone.querySelectorAll('script, style, noscript, input, select, textarea, button, [hidden], [aria-hidden="true"], .add-sec-ques, [data-type*="add" i], [data-action*="add" i]').forEach(item => item.remove());
            clone.querySelectorAll('[class]').forEach(item => {
                const cls = String(item.className || '');
                if (/(?:^|[-_\s])(tool|toolbar|operate|operation|action|ctrl|control|edit|delete|drag|score|add-sec-ques|add-ques|add-question)(?:$|[-_\s])/i.test(cls)) item.remove();
            });
            // 题型标题区域常把“添加题型下试题”等编辑入口作为普通 span/a 文本渲染，
            // 不能只依赖 button 或 class 过滤。这里对叶子节点做永久操作词清洗。
            const sectionActionTextRe = /^(?:添加题型下试题|添加(?:本题型)?试题|新增(?:本题型)?试题|编辑题型|删除题型|题型设置|设置题型|移动题型|复制题型|编辑|删除|移除|上移|下移|拖动|换题|设置|收起|展开|更多|操作)$/;
            clone.querySelectorAll('*').forEach(item => {
                if (!item.isConnected && !clone.contains(item)) return;
                if (item.children.length) return;
                const value = String(item.textContent || '').replace(/\s+/g, ' ').trim();
                if (sectionActionTextRe.test(value)) item.remove();
            });
            let value = String(clone.textContent || '')
                .replace(/[\u00a0\t\r\n]+/g, ' ')
                .replace(/\s+/g, ' ')
                .trim();
            // 清理常见的非标题统计信息，但不破坏“（一）”“一、”等原卷编号。
            value = value
                // 双保险：即使网站把编辑入口直接作为 sec-title 的文本节点，也只清掉操作词，不伤及原卷标题。
                .replace(/\s*(?:添加题型下试题|添加(?:本题型)?试题|新增(?:本题型)?试题|编辑题型|删除题型|题型设置|设置题型|移动题型|复制题型)(?:\s*|$)/g, ' ')
                .replace(/\s*(?:共\s*\d+\s*题|满分\s*[:：]?\s*\d+(?:\.\d+)?\s*分|本大题\s*\d+(?:\.\d+)?\s*分)\s*$/g, '')
                .replace(/^([一二三四五六七八九十百零〇\d]+[、．.])\s+/, '$1')
                .replace(/\s+/g, ' ')
                .trim();
            return value;
        }

        getPrintableSourceNodes() {
            const questions = this.getPaperQuestionNodes();
            if (this.isGroupExamCenterPage()) {
                const nodes = [...this.getPaperSectionNodes(), ...questions];
                return Array.from(new Set(nodes)).sort((a, b) => {
                    if (a === b) return 0;
                    const pos = a.compareDocumentPosition(b);
                    return pos & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
                });
            }

            // 普通详情页按“题目 -> 所属大题”的关系重新组装流，而不是把全站 .sec-title 与题目混在一起排序。
            // 这同时保证同一大题标题只出现一次，也从结构上杜绝答案区前混入目录副本。
            const nodes = [];
            let lastSection = null;
            questions.forEach(question => {
                const section = this.findSectionNodeForQuestion(question);
                if (section && section !== lastSection) {
                    nodes.push(section);
                    lastSection = section;
                }
                nodes.push(question);
            });
            return nodes;
        }

        getQuestionWrap(questionNode) {
            if (!questionNode) return null;
            if (questionNode.matches?.('.wrapper.quesdiv')) return questionNode;
            return questionNode.querySelector?.('.wrapper.quesdiv') || questionNode;
        }

        getQuestionKey(questionNode, fallbackIndex = 0) {
            const wrap = this.getQuestionWrap(questionNode);
            const candidates = [
                questionNode?.getAttribute?.('data-qid'), questionNode?.dataset?.qid,
                questionNode?.getAttribute?.('data-id'), questionNode?.dataset?.id,
                wrap?.getAttribute?.('data-qid'), wrap?.getAttribute?.('data-id'),
                questionNode?.id, wrap?.id
            ];
            const value = candidates.find(item => String(item || '').trim());
            return value ? `qid:${String(value).trim()}` : `index:${fallbackIndex}`;
        }

        getAnswerNodeForQuestion(questionNode) {
            const wrap = this.getQuestionWrap(questionNode);
            if (!wrap) return null;
            return wrap.querySelector('.exam-item__opt')
                || questionNode.querySelector?.('.exam-item__opt')
                || null;
        }

        getAnswerOptionNodes() {
            const nodes = [];
            this.getPaperQuestionNodes().forEach(question => {
                const node = this.getAnswerNodeForQuestion(question);
                if (node && !nodes.includes(node)) nodes.push(node);
            });
            return nodes;
        }

        getQuestionAnswerSnapshot(questionNode, fallbackIndex = 0) {
            const key = this.getQuestionKey(questionNode, fallbackIndex);
            const snapshot = this.groupExamAnswerSnapshots.get(key);
            return snapshot ? snapshot.cloneNode(true) : null;
        }

        saveQuestionAnswerSnapshot(questionNode, fallbackIndex = 0, answerNode = null) {
            const node = answerNode || this.getAnswerNodeForQuestion(questionNode);
            if (!node || !this.isAnswerContentReady(node)) return false;
            const clone = node.cloneNode(true);
            clone.querySelectorAll('script, style, noscript').forEach(item => item.remove());

            // 8.2.2：组卷中心答案往往使用懒加载/srcset。快照时锁定“此刻真正显示的原图”，
            // 避免进入预览 iframe 后浏览器重新选择另一张密度图，导致 naturalWidth/naturalHeight 与切线坐标不一致。
            const liveImages = Array.from(node.querySelectorAll('img'));
            const clonedImages = Array.from(clone.querySelectorAll('img'));
            clonedImages.forEach((image, index) => {
                const live = liveImages[index];
                const exactSrc = String(live?.currentSrc || live?.src || image.getAttribute('src') || image.getAttribute('data-src') || '');
                if (exactSrc) image.setAttribute('src', exactSrc);
                image.removeAttribute('srcset');
                image.removeAttribute('sizes');
                image.removeAttribute('data-src');
                image.removeAttribute('data-original');
                image.setAttribute('loading', 'eager');
                image.setAttribute('decoding', 'sync');
            });

            this.groupExamAnswerSnapshots.set(this.getQuestionKey(questionNode, fallbackIndex), clone);
            return true;
        }

        isElementActuallyVisible(element) {
            if (!(element instanceof Element)) return false;
            const style = getComputedStyle(element);
            if (style.display === 'none' || style.visibility === 'hidden') return false;
            const rect = element.getBoundingClientRect();
            return rect.width > 0 && rect.height > 0;
        }

        findGroupExamAnswerControl(questionNode) {
            if (!questionNode) return null;
            const candidates = Array.from(questionNode.querySelectorAll('button, a, [role="button"], [data-type], .btn, .ctrl-btn, .tool-btn'));
            let best = null;
            let bestScore = -Infinity;
            candidates.forEach(element => {
                if (!this.isElementActuallyVisible(element)) return;
                const hint = [
                    element.innerText, element.textContent, element.getAttribute('title'),
                    element.getAttribute('aria-label'), element.getAttribute('data-type'), element.className
                ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
                if (!hint) return;
                if (/删除|移除|换题|编辑|上移|下移|收藏|纠错|反馈|下载/i.test(hint)) return;
                let score = 0;
                if (/查看答案|显示答案|答案解析|查看解析|展开解析|展开答案/i.test(hint)) score += 100;
                else if (/答案|解析/i.test(hint)) score += 70;
                if (/answer|parse|analysis|solution|show.*ans/i.test(hint)) score += 45;
                if (/data-type/i.test(hint)) score += 3;
                if (score > bestScore) { best = element; bestScore = score; }
            });
            return bestScore >= 45 ? best : null;
        }

        dispatchGroupExamQuestionClick(questionNode) {
            // 记录点击前所有答案节点的内容。若站点把答案渲染到共享详情面板，
            // 点击后也能通过“新增/内容变化的 .exam-item__opt”识别并保存到当前题快照。
            this._lastGroupExamAnswerBaseline = new Map(
                Array.from(document.querySelectorAll('.exam-item__opt')).map(node => [node, node.innerHTML || ''])
            );
            const control = this.findGroupExamAnswerControl(questionNode);
            const wrap = this.getQuestionWrap(questionNode);
            const target = control
                || wrap?.querySelector('.exam-item__cnt')
                || questionNode.querySelector?.('.exam-item__cnt')
                || wrap
                || questionNode;
            if (!(target instanceof Element)) return false;
            try {
                const opts = { bubbles: true, cancelable: true, view: window, clientX: 1, clientY: 1 };
                if (typeof PointerEvent === 'function') target.dispatchEvent(new PointerEvent('pointerdown', opts));
                target.dispatchEvent(new MouseEvent('mousedown', opts));
                if (typeof PointerEvent === 'function') target.dispatchEvent(new PointerEvent('pointerup', opts));
                target.dispatchEvent(new MouseEvent('mouseup', opts));
                target.dispatchEvent(new MouseEvent('click', opts));
                return true;
            } catch (error) {
                try { target.click(); return true; } catch (_) { return false; }
            }
        }

        waitForQuestionAnswerSnapshot(questionNode, fallbackIndex = 0, timeout = 2800) {
            return new Promise(resolve => {
                const start = Date.now();
                let lastHtml = '';
                const check = () => {
                    let answerNode = this.getAnswerNodeForQuestion(questionNode);
                    if (!answerNode || !this.isAnswerContentReady(answerNode)) {
                        const baseline = this._lastGroupExamAnswerBaseline || new Map();
                        const changedGlobal = Array.from(document.querySelectorAll('.exam-item__opt')).find(node => {
                            if (!this.isAnswerContentReady(node)) return false;
                            return !baseline.has(node) || baseline.get(node) !== (node.innerHTML || '');
                        });
                        if (changedGlobal) answerNode = changedGlobal;
                    }
                    if (answerNode) {
                        const html = answerNode.innerHTML || '';
                        if (this.isAnswerContentReady(answerNode)) {
                            this.saveQuestionAnswerSnapshot(questionNode, fallbackIndex, answerNode);
                            const answerImages = Array.from(answerNode.querySelectorAll('img'));
                            const pending = answerImages.some(img => !img.complete || (img.naturalWidth || 0) <= 0);
                            if (!pending) return resolve(true);
                        }
                        if (html !== lastHtml) lastHtml = html;
                    }
                    if (Date.now() - start >= timeout) {
                        return resolve(this.saveQuestionAnswerSnapshot(questionNode, fallbackIndex, answerNode));
                    }
                    requestAnimationFrame(check);
                };
                check();
            });
        }

        async ensureGroupExamAnswersReady() {
            if (!this.isGroupExamCenterPage()) return { attempted: 0, captured: 0 };
            if (this.groupExamAnswerPreparation) return this.groupExamAnswerPreparation;
            this.groupExamAnswerPreparation = (async () => {
                const questions = this.getPaperQuestionNodes();
                if (!questions.length) return { attempted: 0, captured: 0 };
                let captured = 0;
                for (let index = 0; index < questions.length; index++) {
                    const question = questions[index];
                    const existing = this.getAnswerNodeForQuestion(question);
                    if (existing && this.isAnswerContentReady(existing)) {
                        if (this.saveQuestionAnswerSnapshot(question, index + 1, existing)) captured++;
                        continue;
                    }
                    this.setBusy(true, `组卷中心：正在读取第 ${index + 1}/${questions.length} 题答案…`);
                    this.dispatchGroupExamQuestionClick(question);
                    const ok = await this.waitForQuestionAnswerSnapshot(question, index + 1, 3200);
                    if (ok) captured++;
                    await new Promise(resolve => setTimeout(resolve, 120));
                }
                console.log(`✅ 组卷中心逐题答案准备：${captured}/${questions.length}`);
                return { attempted: questions.length, captured };
            })();
            try {
                return await this.groupExamAnswerPreparation;
            } finally {
                this.groupExamAnswerPreparation = null;
            }
        }

        extractPaperModel({ includeHtml = false } = {}) {
            const titleInfo = this.getPaperTitleInfo();
            const sections = [];
            const questions = [];
            let currentSection = '';
            this.getPrintableSourceNodes().forEach(node => {
                if (node.classList.contains('sec-title') || node.classList.contains('questype-head')) {
                    const text = this.getSectionTitleText(node);
                    if (text) {
                        currentSection = text;
                        sections.push(text);
                    }
                    return;
                }
                const wrap = this.getQuestionWrap(node);
                if (!wrap) return;
                const cnt = wrap.querySelector('.exam-item__cnt');
                const opt = this.getAnswerNodeForQuestion(node) || this.getQuestionAnswerSnapshot(node, questions.length + 1);
                const answerImages = opt ? Array.from(opt.querySelectorAll('img')).filter(img => String(img.currentSrc || img.src || '').includes('getAnswerAndParse')) : [];
                const qid = node.getAttribute('data-qid') || node.dataset?.qid || node.id || wrap.getAttribute('data-qid') || '';
                const item = {
                    index: questions.length + 1,
                    qid: String(qid || ''),
                    section: currentSection,
                    questionTextLength: this.normalizePaperTitleText(cnt?.innerText || cnt?.textContent || '').length,
                    hasAnswerArea: Boolean(opt),
                    answerReady: Boolean(opt && this.isAnswerContentReady(opt)),
                    imageCount: cnt ? cnt.querySelectorAll('img').length : 0,
                    answerImageCount: answerImages.length,
                    formulaCount: wrap.querySelectorAll('math, mjx-container, .MathJax, .katex').length,
                    tableCount: wrap.querySelectorAll('table').length
                };
                if (includeHtml) {
                    item.questionHtml = cnt?.innerHTML || '';
                    item.answerHtml = opt?.innerHTML || '';
                }
                questions.push(item);
            });
            return {
                version: '8.4.3',
                identity: this.getPaperIdentityKey(),
                url: location.href,
                title: this.getPaperTitle(),
                detectedTitle: titleInfo.title,
                titleSource: titleInfo.selector,
                titleConfidence: titleInfo.confidence,
                sections,
                questions
            };
        }

        getPaperDiagnostics() {
            const model = this.extractPaperModel();
            const allImages = this.getPaperQuestionNodes().flatMap(node => Array.from(node.querySelectorAll('img')));
            const answerImages = allImages.filter(img => String(img.currentSrc || img.src || '').includes('getAnswerAndParse'));
            const failedOrPending = answerImages.filter(img => !img.complete || (img.naturalWidth || 0) <= 0).length;
            const answerReady = model.questions.filter(question => question.answerReady).length;
            return {
                identity: model.identity,
                title: model.title,
                detectedTitle: model.detectedTitle,
                titleSource: model.titleSource,
                titleConfidence: model.titleConfidence,
                sectionCount: model.sections.length,
                questionCount: model.questions.length,
                answerAreaCount: model.questions.filter(question => question.hasAnswerArea).length,
                answerReadyCount: answerReady,
                answerImageCount: answerImages.length,
                pendingAnswerImageCount: failedOrPending,
                formulaCount: model.questions.reduce((sum, item) => sum + item.formulaCount, 0),
                tableCount: model.questions.reduce((sum, item) => sum + item.tableCount, 0),
                generatedAt: new Date().toISOString()
            };
        }

        blobToDataUrl(blob) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => resolve(String(reader.result || ''));
                reader.onerror = () => reject(reader.error || new Error('资源读取失败'));
                reader.readAsDataURL(blob);
            });
        }

        fetchResourceAsDataUrl(url) {
            return new Promise((resolve, reject) => {
                if (typeof GM_xmlhttpRequest !== 'function') {
                    reject(new Error('GM_xmlhttpRequest 不可用'));
                    return;
                }
                GM_xmlhttpRequest({
                    method: 'GET',
                    url,
                    responseType: 'blob',
                    timeout: 20000,
                    anonymous: false,
                    onload: async response => {
                        try {
                            if (response.status < 200 || response.status >= 400) throw new Error('HTTP ' + response.status);
                            resolve(await this.blobToDataUrl(response.response));
                        } catch (error) { reject(error); }
                    },
                    ontimeout: () => reject(new Error('资源请求超时')),
                    onerror: () => reject(new Error('资源请求失败'))
                });
            });
        }

        async fetchResourceForPreview(payload = {}, targetWindow = null) {
            const requestId = String(payload.requestId || '');
            try {
                const url = new URL(String(payload.url || ''), location.href);
                if (!/^https?:$/.test(url.protocol)) throw new Error('不支持的资源协议');
                const dataUrl = await this.fetchResourceAsDataUrl(url.href);
                targetWindow?.postMessage({ type: 'zujuanFetchResourceResult', requestId, ok: true, dataUrl }, '*');
            } catch (error) {
                targetWindow?.postMessage({ type: 'zujuanFetchResourceResult', requestId, ok: false, message: error?.message || String(error) }, '*');
            }
        }

        autoCheckIn() {
            setTimeout(() => {
                const signedInLink = document.querySelector('.user-assets-box a.assets-method[href="/score_task/"]');
                if (signedInLink && signedInLink.textContent.trim() !== '已签到') {
                    document.querySelector('a.sign-in-btn')?.click();
                    document.querySelector('a.day-sign-in')?.click();
                }
            }, 2500);
        }

        createFloatingButton() {
            if (document.getElementById('zujuanjs-float-print-btn')) return;
            const btn = document.createElement('button');
            btn.id = 'zujuanjs-float-print-btn';
            btn.className = 'zujuanjs-float-print-btn';
            btn.type = 'button';
            btn.setAttribute('aria-label', '打开试卷打印预览；可拖动位置');
            btn.title = '打印试卷（可拖动；双击恢复默认位置）';
            btn.innerHTML = `
                <span class="zujuanjs-float-print-btn-icon" aria-hidden="true">
                    <svg viewBox="0 0 16 16" width="18" height="18" fill="currentColor"><path d="M5 1a2 2 0 0 0-2 2v2h2V3h6v2h2V3a2 2 0 0 0-2-2H5Z"/><path d="M4 10h8v5H4v-5Zm1 1v3h6v-3H5Z"/><path d="M2 5a2 2 0 0 0-2 2v4a2 2 0 0 0 2 2h1v-4h10v4h1a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2H2Zm11 2.25a.75.75 0 1 1 1.5 0 .75.75 0 0 1-1.5 0Z"/></svg>
                </span>
                <span class="zujuanjs-float-print-btn-text">打印试卷</span>`;

            const clampPosition = (left, top) => {
                const rect = btn.getBoundingClientRect();
                const margin = 8;
                const maxLeft = Math.max(margin, window.innerWidth - Math.max(rect.width, 44) - margin);
                const maxTop = Math.max(margin, window.innerHeight - Math.max(rect.height, 44) - margin);
                return {
                    left: Math.max(margin, Math.min(maxLeft, Number(left) || margin)),
                    top: Math.max(margin, Math.min(maxTop, Number(top) || margin))
                };
            };
            const applyPosition = (left, top, persist = false) => {
                const pos = clampPosition(left, top);
                btn.style.setProperty('left', `${Math.round(pos.left)}px`, 'important');
                btn.style.setProperty('top', `${Math.round(pos.top)}px`, 'important');
                btn.style.setProperty('right', 'auto', 'important');
                btn.style.setProperty('bottom', 'auto', 'important');
                if (persist) GM_setValue('floatingButtonPosition', { left: Math.round(pos.left), top: Math.round(pos.top) });
            };
            const restoreSavedPosition = () => {
                const saved = GM_getValue('floatingButtonPosition', null);
                if (!saved || !Number.isFinite(Number(saved.left)) || !Number.isFinite(Number(saved.top))) return;
                requestAnimationFrame(() => applyPosition(Number(saved.left), Number(saved.top), false));
            };

            let pointerId = null;
            let startX = 0;
            let startY = 0;
            let startLeft = 0;
            let startTop = 0;
            let dragged = false;
            let suppressClick = false;

            btn.addEventListener('pointerdown', event => {
                if (event.button !== undefined && event.button !== 0) return;
                const rect = btn.getBoundingClientRect();
                pointerId = event.pointerId;
                startX = event.clientX;
                startY = event.clientY;
                startLeft = rect.left;
                startTop = rect.top;
                dragged = false;
                btn.classList.add('zujuanjs-dragging');
                try { btn.setPointerCapture(pointerId); } catch (_) {}
            });
            btn.addEventListener('pointermove', event => {
                if (pointerId === null || event.pointerId !== pointerId) return;
                const dx = event.clientX - startX;
                const dy = event.clientY - startY;
                if (!dragged && Math.hypot(dx, dy) < 4) return;
                dragged = true;
                event.preventDefault();
                applyPosition(startLeft + dx, startTop + dy, false);
            });
            const finishDrag = event => {
                if (pointerId === null || event.pointerId !== pointerId) return;
                if (dragged) {
                    const rect = btn.getBoundingClientRect();
                    applyPosition(rect.left, rect.top, true);
                    suppressClick = true;
                    setTimeout(() => { suppressClick = false; }, 0);
                }
                try { btn.releasePointerCapture(pointerId); } catch (_) {}
                pointerId = null;
                btn.classList.remove('zujuanjs-dragging');
            };
            btn.addEventListener('pointerup', finishDrag);
            btn.addEventListener('pointercancel', finishDrag);
            btn.addEventListener('click', e => {
                e.preventDefault();
                if (suppressClick || dragged) return;
                this.showPrintDialog();
            });
            btn.addEventListener('dblclick', event => {
                event.preventDefault();
                event.stopPropagation();
                GM_setValue('floatingButtonPosition', null);
                btn.style.removeProperty('left');
                btn.style.removeProperty('top');
                btn.style.removeProperty('right');
                btn.style.removeProperty('bottom');
            });
            window.addEventListener('resize', () => {
                const saved = GM_getValue('floatingButtonPosition', null);
                if (!saved) return;
                const rect = btn.getBoundingClientRect();
                applyPosition(rect.left, rect.top, true);
            }, { passive: true });

            document.body.appendChild(btn);
            restoreSavedPosition();
        }

        startUiKeeper() {
            const ensure = () => {
                if (!document.body) return;
                if (!document.getElementById('zujuanjs-float-print-btn')) this.createFloatingButton();
            };
            const observer = new MutationObserver(ensure);
            if (document.body) observer.observe(document.body, { childList: true, subtree: true });
            else document.addEventListener('DOMContentLoaded', () => {
                ensure();
                observer.observe(document.body, { childList: true, subtree: true });
            }, { once: true });
            window.addEventListener('popstate', ensure);
            window.addEventListener('hashchange', ensure);
        }

        installShortcuts() {
            document.addEventListener('keydown', event => {
                if (!(event.ctrlKey || event.metaKey) || !event.shiftKey || event.altKey) return;
                if (String(event.key).toLowerCase() !== 'p') return;
                const typingTarget = event.target instanceof Element && event.target.closest('input, textarea, select, [contenteditable="true"]');
                if (typingTarget) return;
                event.preventDefault();
                this.showPrintDialog();
            });
        }

        setBusy(visible, message = '正在准备试卷…') {
            let overlay = document.getElementById('zujuanjs-busy-overlay');
            if (!visible) {
                overlay?.remove();
                return;
            }
            if (!overlay) {
                overlay = document.createElement('div');
                overlay.id = 'zujuanjs-busy-overlay';
                overlay.innerHTML = '<div class="zujuanjs-busy-card" role="status" aria-live="polite"><div class="zujuanjs-busy-spinner" aria-hidden="true"></div><div class="zujuanjs-busy-message"></div></div>';
                document.body.appendChild(overlay);
            }
            const label = overlay.querySelector('.zujuanjs-busy-message');
            if (label) label.textContent = message;
        }

        notify(message, title = '组卷网打印工作台') {
            try {
                if (typeof GM_notification === 'function') {
                    GM_notification({ title, text: String(message), timeout: 5000 });
                    return;
                }
            } catch (_) {}
            window.alert(String(message));
        }

        startAdRemover() {
            const observer = new MutationObserver(() => {
                Config.adSelectors.forEach(sel => document.querySelector(sel)?.remove());
            });
            observer.observe(document.body, { childList: true, subtree: true });
        }

        getFontOptions() {
            return [
                { value: '"Times New Roman", SimSun, "Songti SC", serif', text: '宋体 + 新罗马' },
                { value: 'SimSun, "Songti SC", serif', text: '宋体' },
                { value: '"Microsoft YaHei", "PingFang SC", sans-serif', text: '微软雅黑' },
                { value: 'SimHei, "PingFang SC", sans-serif', text: '黑体' },
                { value: 'KaiTi, "Songti SC", serif', text: '楷体' },
                { value: 'FangSong, "Songti SC", serif', text: '仿宋' },
                { value: '"Noto Serif SC", "Times New Roman", serif', text: '思源宋体' },
                { value: '"Noto Sans SC", "PingFang SC", sans-serif', text: '思源黑体' }
            ];
        }

        getSizeOptions() {
            return [
                { value: '14px', text: '14px · 精排' }, { value: '15px', text: '15px' },
                { value: '16px', text: '16px' }, { value: '17px', text: '17px' },
                { value: '18px', text: '18px' }, { value: '20px', text: '20px' }, { value: '22px', text: '22px' }
            ];
        }

        getLineHeightOptions() {
            return [
                { value: '1.2', text: '1.2 · 紧凑' }, { value: '1.35', text: '1.35' },
                { value: '1.4', text: '1.4 · 精排' }, { value: '1.5', text: '1.5 · 舒展' }, { value: '1.75', text: '1.75' }, { value: '2.0', text: '2.0 · 宽松' }
            ];
        }

        getPageSizeOptions() {
            return [
                { value: '10px', text: '10px' }, { value: '12px', text: '12px' },
                { value: '14px', text: '14px' }, { value: '16px', text: '16px' }, { value: '18px', text: '18px' }
            ];
        }

        getBoldOptions() {
            return [{ value: 'false', text: '不加粗' }, { value: 'true', text: '加粗' }];
        }

        escapeAttribute(value) {
            return String(value)
                .replace(/&/g, '&amp;')
                .replace(/"/g, '&quot;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');
        }

        renderSelect(id, options, defaultValue) {
            let allOptions = options.slice();
            const exists = allOptions.some(o => o.value === defaultValue);
            if (!exists && defaultValue !== undefined && defaultValue !== '') {
                allOptions.unshift({ value: defaultValue, text: '自定义' });
            }
            const selected = allOptions.find(o => o.value === defaultValue) || allOptions[0];
            const escapeAttr = (s) => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            const optsHtml = allOptions.map(o =>
                `<div class="print-custom-select-option${o.value === selected.value ? ' selected' : ''}" data-value="${escapeAttr(o.value)}">${o.text}</div>`
            ).join('');
            return `
                <div class="print-custom-select" id="${id}" data-value="${escapeAttr(selected.value)}">
                    <div class="print-custom-select-trigger"><span>${selected.text}</span><div class="print-custom-select-arrow"></div></div>
                    <div class="print-custom-select-dropdown">${optsHtml}</div>
                </div>`;
        }

        initSelect(id, onChange) {
            const select = document.getElementById(id);
            if (!select) return;
            const trigger = select.querySelector('.print-custom-select-trigger');
            const options = select.querySelectorAll('.print-custom-select-option');

            trigger.addEventListener('click', (e) => {
                e.stopPropagation();
                document.querySelectorAll('.print-custom-select.open').forEach(s => { if (s !== select) s.classList.remove('open'); });
                select.classList.toggle('open');
            });

            options.forEach(opt => {
                opt.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const value = opt.dataset.value;
                    options.forEach(o => o.classList.remove('selected'));
                    opt.classList.add('selected');
                    select.dataset.value = value;
                    trigger.querySelector('span').textContent = opt.textContent;
                    select.classList.remove('open');
                    if (onChange) onChange(value);
                });
            });
        }

        showPrintDialogLegacy() {
            const savedFont = GM_getValue('questionFont', '"Times New Roman", SimSun, "Songti SC", serif');
            const savedSize = GM_getValue('questionSize', '14px');
            const savedLineHeight = GM_getValue('questionLineHeight', '1.4');
            const savedPageFont = GM_getValue('pageFont', savedFont);
            const savedPageSize = GM_getValue('pageSize', '10px');
            const savedPageBold = String(GM_getValue('pageBold', false));
            const savedMode = GM_getValue('printMode', 'q');
            const savedMargins = GM_getValue('pageMargins', '15,17,18,17');
            const savedQuestionSpacing = String(GM_getValue('questionSpacing', '6'));
            const savedPreviewLayout = GM_getValue('previewLayout', 'double');
            const savedPreviewZoom = String(GM_getValue('previewZoom', 'auto'));
            const defaultTitle = this.getPaperTitle();

            const fontOptions = this.getFontOptions();
            const sizeOptions = this.getSizeOptions();
            const lineHeightOptions = this.getLineHeightOptions();
            const pageSizeOptions = this.getPageSizeOptions();
            const boldOptions = this.getBoldOptions();

            const modeOptions = [
                { value: 'q', text: '仅试题' },
                { value: 'qa', text: '试题与答案' },
                { value: 'qe', text: '答案移至末尾' },
                { value: 'a', text: '仅答案' }
            ];
            const marginOptions = [
                { value: '12,12,18,12', text: '紧凑' },
                { value: '15,17,18,17', text: '精排' },
                { value: '18,15,22,15', text: '原标准' },
                { value: '22,18,26,18', text: '宽松' }
            ];
            const spacingOptions = [
                { value: '4', text: '4px 紧凑' },
                { value: '6', text: '6px 精排' },
                { value: '10', text: '10px 原标准' },
                { value: '16', text: '16px' },
                { value: '22', text: '22px 宽松' }
            ];
            const zoomOptions = [
                { value: 'auto', text: '自动适应' },
                { value: '0.65', text: '65%' },
                { value: '0.8', text: '80%' },
                { value: '1', text: '100%' }
            ];

            const closeAllSelects = (e) => {
                if (!e.target.closest('.print-custom-select')) {
                    document.querySelectorAll('.print-custom-select.open').forEach(s => s.classList.remove('open'));
                }
            };

            Swal.fire({
                title: '打印设置',
                width: 700,
                customClass: { popup: 'print-dialog-popup' },
                confirmButtonColor: '#1677ff',
                cancelButtonColor: '#d9d9d9',
                html: `
                    <div class="print-dialog-container">
                        <div class="print-dialog-section">
                            <label class="print-dialog-label">试卷标题</label>
                            <input type="text" id="print-title-input" class="print-dialog-input" value="${this.escapeAttribute(defaultTitle)}" placeholder="留空则不显示标题">
                        </div>

                        <div class="print-dialog-section">
                            <label class="print-dialog-label">打印内容</label>
                            <div class="print-option-grid" id="print-mode-grid">
                                ${modeOptions.map(m => `
                                    <label class="print-option-card${m.value === savedMode ? ' active' : ''}">
                                        <input type="radio" name="printMode" value="${m.value}" ${m.value === savedMode ? 'checked' : ''}>
                                        <span>${m.text}</span>
                                    </label>`).join('')}
                            </div>
                        </div>

                        <div class="print-dialog-section">
                            <label class="print-dialog-label">正文样式</label>
                            <div class="print-row-3">
                                <div class="print-field">
                                    <div class="print-field-label">字体</div>
                                    ${this.renderSelect('print-font-select', fontOptions, savedFont)}
                                </div>
                                <div class="print-field">
                                    <div class="print-field-label">字号</div>
                                    ${this.renderSelect('print-size-select', sizeOptions, savedSize)}
                                </div>
                                <div class="print-field">
                                    <div class="print-field-label">行距</div>
                                    ${this.renderSelect('print-lineheight-select', lineHeightOptions, savedLineHeight)}
                                </div>
                            </div>
                        </div>

                        <div class="print-dialog-section">
                            <label class="print-dialog-label">页面排版</label>
                            <div class="print-row-4">
                                <div class="print-field">
                                    <div class="print-field-label">预览排布</div>
                                    <div class="print-option-grid print-layout-grid" id="print-layout-grid">
                                        <label class="print-option-card${savedPreviewLayout === 'single' ? ' active' : ''}">
                                            <input type="radio" name="previewLayout" value="single" ${savedPreviewLayout === 'single' ? 'checked' : ''}>
                                            <span>单页</span>
                                        </label>
                                        <label class="print-option-card${savedPreviewLayout === 'double' ? ' active' : ''}">
                                            <input type="radio" name="previewLayout" value="double" ${savedPreviewLayout === 'double' ? 'checked' : ''}>
                                            <span>双页</span>
                                        </label>
                                    </div>
                                </div>
                                <div class="print-field">
                                    <div class="print-field-label">预览缩放</div>
                                    ${this.renderSelect('print-zoom-select', zoomOptions, savedPreviewZoom)}
                                </div>
                                <div class="print-field">
                                    <div class="print-field-label">页边距</div>
                                    ${this.renderSelect('print-margin-select', marginOptions, savedMargins)}
                                </div>
                                <div class="print-field">
                                    <div class="print-field-label">题间距</div>
                                    ${this.renderSelect('print-spacing-select', spacingOptions, savedQuestionSpacing)}
                                </div>
                            </div>
                        </div>

                        <div class="print-dialog-section">
                            <label class="print-dialog-label">页码样式</label>
                            <div class="print-row-3">
                                <div class="print-field">
                                    <div class="print-field-label">字体</div>
                                    ${this.renderSelect('print-pagefont-select', fontOptions, savedPageFont)}
                                </div>
                                <div class="print-field">
                                    <div class="print-field-label">字号</div>
                                    ${this.renderSelect('print-pagesize-select', pageSizeOptions, savedPageSize)}
                                </div>
                                <div class="print-field">
                                    <div class="print-field-label">字重</div>
                                    ${this.renderSelect('print-pagebold-select', boldOptions, savedPageBold)}
                                </div>
                            </div>
                        </div>

                        <div class="print-dialog-section">
                            <div class="print-preview-wrap">
                                <span class="print-preview-caption">样例</span>
                                <div class="print-preview-box" id="print-preview">
                                    1.&emsp;已知函数 f(x) = ax² + bx + c (a ≠ 0)，在区间 [1, 5] 上单调递增。
                                </div>
                            </div>
                        </div>
                    </div>`,
                confirmButtonText: '进入A4预览',
                showCancelButton: true,
                cancelButtonText: '取消',
                didOpen: () => {
                    const preview = document.getElementById('print-preview');
                    const updatePreview = () => {
                        const fontSelect = document.getElementById('print-font-select');
                        const sizeSelect = document.getElementById('print-size-select');
                        const lhSelect = document.getElementById('print-lineheight-select');
                        const font = fontSelect ? fontSelect.dataset.value : savedFont;
                        const size = sizeSelect ? sizeSelect.dataset.value : savedSize;
                        const lh = lhSelect ? lhSelect.dataset.value : savedLineHeight;
                        preview.style.fontFamily = font;
                        preview.style.fontSize = size;
                        preview.style.lineHeight = lh;
                    };

                    this.initSelect('print-font-select', () => updatePreview());
                    this.initSelect('print-size-select', () => updatePreview());
                    this.initSelect('print-lineheight-select', () => updatePreview());
                    this.initSelect('print-pagefont-select');
                    this.initSelect('print-pagesize-select');
                    this.initSelect('print-pagebold-select');
                    this.initSelect('print-zoom-select');
                    this.initSelect('print-margin-select');
                    this.initSelect('print-spacing-select');

                    ['print-mode-grid', 'print-layout-grid'].forEach(gridId => {
                        const grid = document.getElementById(gridId);
                        grid?.querySelectorAll('.print-option-card').forEach(card => {
                            card.addEventListener('click', () => {
                                grid.querySelectorAll('.print-option-card').forEach(c => c.classList.remove('active'));
                                card.classList.add('active');
                            });
                        });
                    });

                    document.addEventListener('click', closeAllSelects);
                    updatePreview();
                },
                willClose: () => {
                    document.removeEventListener('click', closeAllSelects);
                },
                preConfirm: () => ({
                    mode: document.querySelector('input[name="printMode"]:checked')?.value || 'q',
                    font: document.getElementById('print-font-select')?.dataset.value || savedFont,
                    size: document.getElementById('print-size-select')?.dataset.value || savedSize,
                    lineHeight: document.getElementById('print-lineheight-select')?.dataset.value || savedLineHeight,
                    title: document.getElementById('print-title-input').value.trim(),
                    pageFont: document.getElementById('print-pagefont-select')?.dataset.value || savedPageFont,
                    pageSize: document.getElementById('print-pagesize-select')?.dataset.value || savedPageSize,
                    pageBold: document.getElementById('print-pagebold-select')?.dataset.value === 'true',
                    pageMargins: document.getElementById('print-margin-select')?.dataset.value || savedMargins,
                    questionSpacing: document.getElementById('print-spacing-select')?.dataset.value || savedQuestionSpacing,
                    previewLayout: document.querySelector('input[name="previewLayout"]:checked')?.value || savedPreviewLayout,
                    previewZoom: document.getElementById('print-zoom-select')?.dataset.value || savedPreviewZoom
                })
            }).then(async res => {
                if (!res.isConfirmed) return;
                const { mode, font, size, lineHeight, title, pageFont, pageSize, pageBold, pageMargins, questionSpacing, previewLayout, previewZoom } = res.value;
                GM_setValue('questionFont', font);
                GM_setValue('questionSize', size);
                GM_setValue('questionLineHeight', lineHeight);
                GM_setValue('pageFont', pageFont || font);
                GM_setValue('pageSize', pageSize || '12px');
                GM_setValue('pageBold', pageBold);
                GM_setValue('printMode', mode);
                GM_setValue('pageMargins', pageMargins);
                GM_setValue('questionSpacing', questionSpacing);
                GM_setValue('previewLayout', previewLayout);
                GM_setValue('previewZoom', previewZoom);

                const includeQuestions = mode !== 'a';
                const includeAnswers = mode === 'qa' || mode === 'a';
                const answersAtEnd = mode === 'qe';

                if (includeAnswers || answersAtEnd) {
                    Swal.fire({ title: '请稍候', text: '正在准备答案解析...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
                    await this.ensureAnswersReady();
                    Swal.close();
                }

                const htmlContent = this.generatePreviewHTML(
                    includeQuestions,
                    includeAnswers,
                    answersAtEnd,
                    font,
                    size,
                    lineHeight,
                    title,
                    pageFont || font,
                    pageSize || '12px',
                    pageBold,
                    { pageMargins, questionSpacing, previewLayout, previewZoom }
                );
                this.openPreview(htmlContent);
            });
        }

        getPreviewSettings(overrides = {}) {
            const has = key => Object.prototype.hasOwnProperty.call(overrides, key);
            const font = has('font') ? overrides.font : GM_getValue('questionFont', '"Times New Roman", SimSun, "Songti SC", serif');
            return {
                mode: has('mode') ? overrides.mode : GM_getValue('printMode', 'q'),
                font,
                size: has('size') ? overrides.size : GM_getValue('questionSize', '14px'),
                lineHeight: has('lineHeight') ? overrides.lineHeight : GM_getValue('questionLineHeight', '1.4'),
                answerFont: has('answerFont') ? overrides.answerFont : GM_getValue('answerFont', font),
                answerSize: has('answerSize') ? overrides.answerSize : GM_getValue('answerSize', has('size') ? overrides.size : GM_getValue('questionSize', '14px')),
                answerLineHeight: has('answerLineHeight') ? overrides.answerLineHeight : GM_getValue('answerLineHeight', has('lineHeight') ? overrides.lineHeight : GM_getValue('questionLineHeight', '1.4')),
                answerImageScale: has('answerImageScale') ? overrides.answerImageScale : GM_getValue('answerImageScale', '100'),
                answerLongImageMode: has('answerLongImageMode') ? overrides.answerLongImageMode : GM_getValue('answerLongImageMode', 'split'),
                answerCutAutoLead: has('answerCutAutoLead') ? overrides.answerCutAutoLead : GM_getValue('answerCutAutoLead', '8'),
                answerCutOverlap: has('answerCutOverlap') ? overrides.answerCutOverlap : GM_getValue('answerCutOverlap', '22'),
                answerCutOverlapEnabled: has('answerCutOverlapEnabled')
                    ? String(overrides.answerCutOverlapEnabled) !== 'false'
                    : String(GM_getValue('answerCutOverlapEnabled', true)) !== 'false',
                answerCutMinFill: has('answerCutMinFill') ? overrides.answerCutMinFill : GM_getValue('answerCutMinFill', '8'),
                answerCutSmartSnap: has('answerCutSmartSnap') ? (overrides.answerCutSmartSnap !== false && String(overrides.answerCutSmartSnap) !== 'false') : String(GM_getValue('answerCutSmartSnap', true)) !== 'false',
                answerCutSearchWindow: has('answerCutSearchWindow') ? overrides.answerCutSearchWindow : GM_getValue('answerCutSearchWindow', '72'),
                answerStartNewPage: has('answerStartNewPage') ? String(overrides.answerStartNewPage) !== 'false' : String(GM_getValue('answerStartNewPage', false)) === 'true',
                answerImageCutOffsets: has('answerImageCutOffsets') && overrides.answerImageCutOffsets && typeof overrides.answerImageCutOffsets === 'object'
                    ? overrides.answerImageCutOffsets
                    : (GM_getValue('answerImageCutOffsets', {}) || {}),
                answerImageCutPositions: has('answerImageCutPositions') && overrides.answerImageCutPositions && typeof overrides.answerImageCutPositions === 'object'
                    ? overrides.answerImageCutPositions
                    : (GM_getValue('answerImageCutPositions', {}) || {}),
                title: has('title') ? overrides.title : this.getPaperTitle(),
                detectedTitle: has('detectedTitle') ? overrides.detectedTitle : this.getPaperTitleInfo().title,
                titleSource: has('titleSource') ? overrides.titleSource : this.getPaperTitleInfo().selector,
                paperDiagnostics: has('paperDiagnostics') && overrides.paperDiagnostics && typeof overrides.paperDiagnostics === 'object'
                    ? overrides.paperDiagnostics : this.getPaperDiagnostics(),
                titleSize: has('titleSize') ? overrides.titleSize : GM_getValue('titleSize', '22px'),
                pageFont: has('pageFont') ? overrides.pageFont : GM_getValue('pageFont', font),
                pageSize: has('pageSize') ? overrides.pageSize : GM_getValue('pageSize', '10px'),
                pageBold: has('pageBold') ? String(overrides.pageBold) !== 'false' : String(GM_getValue('pageBold', false)) !== 'false',
                showPageNumber: has('showPageNumber') ? Boolean(overrides.showPageNumber) : Boolean(GM_getValue('showPageNumber', true)),
                pageNumberFormat: has('pageNumberFormat') ? overrides.pageNumberFormat : GM_getValue('pageNumberFormat', 'current-total'),
                pageNumberScope: has('pageNumberScope') ? overrides.pageNumberScope : GM_getValue('pageNumberScope', 'sectioned'),
                printImageMode: has('printImageMode') ? overrides.printImageMode : GM_getValue('printImageMode', 'color'),
                pageMargins: has('pageMargins') ? overrides.pageMargins : GM_getValue('pageMargins', '15,17,18,17'),
                layoutPreset: has('layoutPreset') ? overrides.layoutPreset : GM_getValue('layoutPreset', 'refined'),
                questionSpacing: has('questionSpacing') ? overrides.questionSpacing : GM_getValue('questionSpacing', '6'),
                previewLayout: has('previewLayout') ? overrides.previewLayout : GM_getValue('previewLayout', 'double'),
                previewZoom: has('previewZoom') ? overrides.previewZoom : GM_getValue('previewZoom', 'auto'),
                paragraphSpacing: has('paragraphSpacing') ? overrides.paragraphSpacing : GM_getValue('paragraphSpacing', '3'),
                contentAlign: has('contentAlign') ? overrides.contentAlign : GM_getValue('contentAlign', 'left'),
                numberGap: has('numberGap') ? overrides.numberGap : GM_getValue('numberGap', '0.45'),
                answerRowHeight: has('answerRowHeight') ? overrides.answerRowHeight : GM_getValue('answerRowHeight', '1.8'),
                pageGap: has('pageGap') ? overrides.pageGap : GM_getValue('pageGap', '20'),
                editorPanelWidth: has('editorPanelWidth') ? overrides.editorPanelWidth : GM_getValue('editorPanelWidth', '340'),
                editorPanelTab: has('editorPanelTab') ? overrides.editorPanelTab : GM_getValue('editorPanelTab', 'document'),
                editorOpen: has('editorOpen') ? Boolean(overrides.editorOpen) : String(GM_getValue('editorOpen', true)) !== 'false',
                documentEdits: has('documentEdits') && overrides.documentEdits && typeof overrides.documentEdits === 'object'
                    ? overrides.documentEdits
                    : {},
                readingAnchor: has('readingAnchor') && overrides.readingAnchor && typeof overrides.readingAnchor === 'object'
                    ? overrides.readingAnchor
                    : null
            };
        }

        savePreviewSettings(settings = {}) {
            const allowed = [
                'mode', 'font', 'size', 'lineHeight', 'answerFont', 'answerSize', 'answerLineHeight', 'answerImageScale', 'answerLongImageMode', 'answerCutAutoLead', 'answerCutOverlap', 'answerCutOverlapEnabled', 'answerCutMinFill', 'answerCutSmartSnap', 'answerCutSearchWindow', 'answerStartNewPage', 'answerImageCutOffsets', 'answerImageCutPositions', 'titleSize', 'pageFont', 'pageSize', 'pageBold',
                'showPageNumber', 'pageNumberFormat', 'pageNumberScope', 'printImageMode', 'pageMargins', 'layoutPreset', 'questionSpacing', 'previewLayout', 'previewZoom',
                'paragraphSpacing', 'contentAlign', 'numberGap', 'answerRowHeight', 'pageGap',
                'editorPanelWidth', 'editorPanelTab', 'editorOpen'
            ];
            allowed.forEach(key => {
                if (!Object.prototype.hasOwnProperty.call(settings, key)) return;
                GM_setValue(key === 'mode' ? 'printMode' : key, settings[key]);
            });
        }

        async openPreviewWithSettings(overrides = {}) {
            const questionCount = this.getPaperQuestionNodes().length;
            if (!questionCount) {
                this.notify('当前页面没有检测到可打印的试题。请先打开试卷详情页，或在组卷中心等待题目加载完成后再试。');
                return;
            }
            const settings = this.getPreviewSettings(overrides);
            this.savePreviewSettings(settings);
            const includeQuestions = settings.mode !== 'a';
            const includeAnswers = settings.mode === 'qa' || settings.mode === 'a';
            const answersAtEnd = settings.mode === 'qe';

            try {
                if (includeAnswers || answersAtEnd) {
                    this.setBusy(true, '正在展开并加载全部答案…');
                    await this.ensureAnswersReady();
                } else {
                    this.setBusy(true, '正在生成 A4 预览…');
                }
                settings.detectedTitle = this.getPaperTitleInfo().title;
                settings.titleSource = this.getPaperTitleInfo().selector;
                settings.paperDiagnostics = this.getPaperDiagnostics();

                const htmlContent = this.generatePreviewHTML(
                includeQuestions,
                includeAnswers,
                answersAtEnd,
                settings.font,
                settings.size,
                settings.lineHeight,
                settings.title,
                settings.pageFont || settings.font,
                settings.pageSize || '12px',
                    settings.pageBold,
                    settings
                );
                this.openPreview(htmlContent);
            } catch (error) {
                console.error('生成打印预览失败：', error);
                this.notify('生成打印预览失败：' + (error && error.message ? error.message : error));
            } finally {
                this.setBusy(false);
            }
        }

        showPrintDialog() {
            this.openPreviewWithSettings();
        }

        triggerShowAnswers() {
            // 组卷网的“显示全部答案”并不总是只监听 input.click()。
            // 优先点击 label，保持与网页真实用户操作一致；再用 change 事件兜底。
            const cb = document.querySelector('#isshowAnswer');
            if (cb && !cb.checked) {
                const label = document.querySelector('label[for="isshowAnswer"]');
                if (label) label.click();
                else {
                    cb.click();
                    cb.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }

            const old = document.querySelector('.tklabel-checkbox.show-answer input');
            if (old && !old.checked) {
                const label = old.closest('label');
                if (label) label.click();
                else {
                    old.click();
                    old.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }
        }

        // 答案展开采用“真实点击 + DOM稳定 + 图片完成”三重判断。
        // 图片型答案等待 getAnswerAndParse 完成；文字/表格型答案在内容稳定后即可继续，避免无图片题型白等超时。
        isAnswerContentReady(optionNode) {
            if (!optionNode) return false;
            const answerImage = Array.from(optionNode.querySelectorAll('img')).some(img => /getAnswerAndParse/i.test(String(img.currentSrc || img.src || '')));
            if (answerImage) return true;
            if (optionNode.querySelector('table, mjx-container, .MathJax, .katex, svg, canvas')) return true;
            const clone = optionNode.cloneNode(true);
            clone.querySelectorAll('.knowledge-box, button, input, script, style, noscript').forEach(el => el.remove());
            return clone.textContent.replace(/\s+/g, '').length >= 4;
        }

        waitForAnswersReady(timeout = 25000) {
            return new Promise(resolve => {
                const start = Date.now();
                let lastRetrigger = 0;
                let lastSignature = '';
                let stableSince = 0;

                const check = () => {
                    const optionNodes = this.getAnswerOptionNodes();
                    const readyCount = optionNodes.filter(node => this.isAnswerContentReady(node)).length;
                    const imgs = Array.from(document.querySelectorAll('img'))
                        .filter(img => /getAnswerAndParse/i.test(String(img.currentSrc || img.src || '')));
                    const loadedCount = imgs.filter(img => img.complete && img.naturalWidth > 0).length;
                    const signature = optionNodes.length + ':' + readyCount + ':' + imgs.length + ':' + loadedCount;
                    if (signature !== lastSignature) {
                        lastSignature = signature;
                        stableSince = Date.now();
                    }

                    const allOptionsReady = optionNodes.length > 0 && readyCount === optionNodes.length;
                    const imagesReady = imgs.length > 0 && loadedCount === imgs.length;
                    const htmlAnswersStable = imgs.length === 0 && allOptionsReady && Date.now() - stableSince >= 450;
                    const mixedAnswersStable = imagesReady && allOptionsReady && Date.now() - stableSince >= 250;
                    if (mixedAnswersStable) {
                        console.log('✅ 答案准备完成：' + readyCount + '/' + optionNodes.length + ' 个答案区域，' + loadedCount + ' 张答案图片');
                        return resolve(true);
                    }
                    if (htmlAnswersStable) {
                        console.log('✅ 文字/表格答案准备完成：' + readyCount + '/' + optionNodes.length + ' 个答案区域');
                        return resolve(true);
                    }

                    if (Date.now() - start >= timeout) {
                        console.warn('⚠️ 答案等待超时，按当前已加载内容继续：', signature);
                        return resolve(false);
                    }
                    if (Date.now() - lastRetrigger > 1200) {
                        this.triggerShowAnswers();
                        lastRetrigger = Date.now();
                    }
                    requestAnimationFrame(check);
                };

                this.triggerShowAnswers();
                check();
            });
        }

        async ensureAnswersReady() {
            if (this.isGroupExamCenterPage()) {
                const result = await this.ensureGroupExamAnswersReady();
                if (result.attempted > 0) {
                    // 组卷中心可能只允许同时展开一题，因此答案以逐题快照为准；
                    // 不再强求所有 .exam-item__opt 必须同时存在于当前 DOM。
                    return result;
                }
            }
            this.triggerShowAnswers();
            await this.waitForAnswersReady();
            return { attempted: this.getPaperQuestionNodes().length, captured: this.getAnswerOptionNodes().length };
        }

        removeLeadingNumber(container) {
            const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT, null, false);
            let node;
            while (node = walker.nextNode()) {
                if (node.textContent.match(/^\s*\d+\.\s*/)) {
                    node.textContent = node.textContent.replace(/^\s*\d+\.\s*/, '');
                    break;
                }
            }
        }

        isFormulaSvgImage(image) {
            const source = String(image.getAttribute('src') || '').toLowerCase();
            const hint = `${image.className || ''} ${image.getAttribute('alt') || ''}`.toLowerCase();
            return /(?:\.svg(?:[?#]|$)|^data:image\/svg\+xml)/.test(source)
                || /(?:math|formula|latex|katex|mathjax)/.test(hint);
        }

        preparePreviewTypography(root) {
            root.querySelectorAll('img').forEach(image => {
                if (!this.isFormulaSvgImage(image)) return;
                image.classList.add('zujuanjs-formula-svg');
                image.dataset.formulaBaseline = '14';
                const pixelStyle = value => /^\s*\d+(?:\.\d+)?px\s*$/.test(value || '') ? Number.parseFloat(value) : 0;
                const width = Number.parseFloat(image.getAttribute('width')) || pixelStyle(image.style.width);
                const height = Number.parseFloat(image.getAttribute('height')) || pixelStyle(image.style.height);
                if (Number.isFinite(width) && width > 0) image.dataset.formulaBaseWidth = String(width);
                if (Number.isFinite(height) && height > 0) image.dataset.formulaBaseHeight = String(height);
            });

            const chinese = '[\\u3400-\\u9fff\\uf900-\\ufaff]';
            const roots = root.querySelectorAll('.zujuanjs-question-body, .zujuanjs-answer-item, .zujuanjs-section-title');
            roots.forEach(container => {
                const nodes = [];
                const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT, {
                    acceptNode: node => {
                        const parent = node.parentElement;
                        if (!parent || !node.textContent || !/[A-Za-z]/.test(node.textContent) || !new RegExp(chinese).test(node.textContent)) {
                            return NodeFilter.FILTER_REJECT;
                        }
                        return parent.closest('script, style, svg, math, mjx-container, .MathJax, .katex, code, pre, .zh-latin-gap')
                            ? NodeFilter.FILTER_REJECT
                            : NodeFilter.FILTER_ACCEPT;
                    }
                });
                let node;
                while (node = walker.nextNode()) nodes.push(node);

                nodes.forEach(textNode => {
                    const text = textNode.textContent;
                    const separated = text
                        .replace(new RegExp(`(${chinese})(?=[A-Za-z])`, 'g'), '$1\\u0000')
                        .replace(new RegExp(`([A-Za-z])(?=${chinese})`, 'g'), '$1\\u0000');
                    if (separated === text) return;
                    const fragment = document.createDocumentFragment();
                    separated.split('\\u0000').forEach((part, index, parts) => {
                        if (part) fragment.appendChild(document.createTextNode(part));
                        if (index < parts.length - 1) {
                            const gap = document.createElement('span');
                            gap.className = 'zh-latin-gap';
                            gap.setAttribute('aria-hidden', 'true');
                            fragment.appendChild(gap);
                        }
                    });
                    textNode.replaceWith(fragment);
                });
            });
        }

        generatePreviewHTML(includeQ, includeA, atEnd, font, size, lineHeight, title, pageFont, pageSize, pageBold, layoutOptions = {}) {
            const marginValues = String(layoutOptions.pageMargins || '18,15,22,15')
                .split(',')
                .map(value => Number(value));
            const validMargins = marginValues.length === 4 && marginValues.every(value => Number.isFinite(value) && value >= 8 && value <= 55);
            const [marginTop, marginRight, marginBottom, marginLeft] = validMargins ? marginValues : [15, 17, 18, 17];
            const contentWidth = 210 - marginLeft - marginRight;
            const contentHeight = 297 - marginTop - marginBottom;
            const footerBottom = Math.max(5, Math.min(9, marginBottom / 3));
            const questionSpacing = Math.max(0, Math.min(32, Number(layoutOptions.questionSpacing) || 10));
            const previewLayout = layoutOptions.previewLayout === 'single' ? 'single' : 'double';
            const rawPreviewZoom = String(layoutOptions.previewZoom || 'auto');
            const numericPreviewZoom = Number(rawPreviewZoom);
            const previewZoom = rawPreviewZoom === 'auto'
                ? 'auto'
                : String(Math.max(0.25, Math.min(2, Number.isFinite(numericPreviewZoom) ? numericPreviewZoom : 1)));
            const parsedTitleSize = Number.parseFloat(String(layoutOptions.titleSize || ''));
            const titleSize = Math.max(18, Math.min(36, Number.isFinite(parsedTitleSize) ? parsedTitleSize : 22));
            const showPageNumber = layoutOptions.showPageNumber !== false;
            const paragraphSpacing = Math.max(0, Math.min(24, Number(layoutOptions.paragraphSpacing) || 8));
            const contentAlign = ['left', 'justify'].includes(layoutOptions.contentAlign) ? layoutOptions.contentAlign : 'left';
            const numberGap = Math.max(0.2, Math.min(2, Number(layoutOptions.numberGap) || 0.55));
            const answerFont = layoutOptions.answerFont || font;
            const answerSize = layoutOptions.answerSize || size;
            const answerLineHeight = layoutOptions.answerLineHeight || lineHeight;
            const answerCutAutoLead = Math.max(0, Math.min(80, Number(layoutOptions.answerCutAutoLead) || 8));
            const answerCutOverlap = Math.max(0, Math.min(80, Number(layoutOptions.answerCutOverlap) || 22));
            const answerCutOverlapEnabled = layoutOptions.answerCutOverlapEnabled !== false && String(layoutOptions.answerCutOverlapEnabled) !== 'false';
            const answerCutMinFill = Math.max(4, Math.min(160, Number(layoutOptions.answerCutMinFill) || 8));
            const answerCutSmartSnap = layoutOptions.answerCutSmartSnap !== false;
            const answerCutSearchWindow = Math.max(24, Math.min(180, Number(layoutOptions.answerCutSearchWindow) || 72));
            const answerStartNewPage = layoutOptions.answerStartNewPage === true || String(layoutOptions.answerStartNewPage) === 'true';
            const pageNumberFormat = ['current-total', 'current', 'chinese'].includes(layoutOptions.pageNumberFormat) ? layoutOptions.pageNumberFormat : 'current-total';
            const printImageMode = layoutOptions.printImageMode === 'grayscale' ? 'grayscale' : 'color';
            const answerRowHeight = Math.max(1, Math.min(6, Number(layoutOptions.answerRowHeight) || 1.8));
            const pageGap = Math.max(8, Math.min(48, Number(layoutOptions.pageGap) || 20));
            const editorPanelWidth = Math.max(280, Math.min(520, Number(layoutOptions.editorPanelWidth) || 340));
            const editorPanelTab = layoutOptions.editorPanelTab === 'page' ? 'page' : 'document';
            const editorOpen = layoutOptions.editorOpen !== false;
            const documentEdits = layoutOptions.documentEdits && typeof layoutOptions.documentEdits === 'object'
                ? layoutOptions.documentEdits
                : {};
            const readingAnchor = layoutOptions.readingAnchor && typeof layoutOptions.readingAnchor === 'object'
                ? layoutOptions.readingAnchor
                : null;
            const answerImageCutOffsets = layoutOptions.answerImageCutOffsets && typeof layoutOptions.answerImageCutOffsets === 'object'
                ? layoutOptions.answerImageCutOffsets
                : {};
            const answerImageCutPositions = layoutOptions.answerImageCutPositions && typeof layoutOptions.answerImageCutPositions === 'object'
                ? layoutOptions.answerImageCutPositions
                : {};
            const paperDiagnostics = layoutOptions.paperDiagnostics && typeof layoutOptions.paperDiagnostics === 'object'
                ? layoutOptions.paperDiagnostics : {};
            const previewSettingsJson = JSON.stringify({
                mode: layoutOptions.mode || 'q', font, size, lineHeight, answerFont, answerSize, answerLineHeight, answerImageScale: String(layoutOptions.answerImageScale || '100'), answerLongImageMode: layoutOptions.answerLongImageMode === 'fit' ? 'fit' : 'split', answerCutAutoLead: String(answerCutAutoLead), answerCutOverlap: String(answerCutOverlap), answerCutOverlapEnabled, answerCutMinFill: String(answerCutMinFill), answerCutSmartSnap, answerCutSearchWindow: String(answerCutSearchWindow), answerStartNewPage, answerImageCutOffsets, answerImageCutPositions, title, detectedTitle: layoutOptions.detectedTitle || title, titleSource: layoutOptions.titleSource || '', paperDiagnostics, titleSize: `${titleSize}px`,
                pageFont, pageSize, pageBold, showPageNumber, pageNumberFormat, printImageMode, pageMargins: `${marginTop},${marginRight},${marginBottom},${marginLeft}`,
                questionSpacing: String(questionSpacing), previewLayout, previewZoom,
                paragraphSpacing: String(paragraphSpacing), contentAlign, numberGap: String(numberGap),
                answerRowHeight: String(answerRowHeight), pageGap: String(pageGap),
                editorPanelWidth: String(editorPanelWidth), editorPanelTab, editorOpen, documentEdits, readingAnchor
            }).replace(/</g, '\\u003c');
            const tempDiv = document.createElement('div');

            const titleEl = document.createElement('div');
            titleEl.className = 'zujuanjs-print-title';
            titleEl.dataset.documentTitle = 'true';
            titleEl.style.fontFamily = 'SimHei, "Microsoft YaHei", "PingFang SC", sans-serif';
            titleEl.style.fontSize = `${titleSize}px`;
            titleEl.style.display = title ? '' : 'none';
            titleEl.textContent = title;
            tempDiv.appendChild(titleEl);

            const answersEndList = [];
            let questionIndex = 1;

            this.getPrintableSourceNodes().forEach(node => {
                if (node.classList.contains('sec-title') || node.classList.contains('questype-head')) {
                    const sectionText = this.getSectionTitleText(node);
                    if (sectionText) {
                        const section = document.createElement('div');
                        section.className = 'zujuanjs-section-title';
                        section.style.fontFamily = 'SimHei, "Microsoft YaHei", "PingFang SC", sans-serif';
                        section.textContent = sectionText;
                        tempDiv.appendChild(section);
                    }
                    return;
                }

                const wrap = this.getQuestionWrap(node);
                if (!wrap) return;

                const qWrapper = document.createElement('div');
                qWrapper.className = 'q-wrapper';
                qWrapper.dataset.blockId = `question-${questionIndex}`;
                qWrapper.dataset.blockLabel = `第 ${questionIndex} 题`;
                const sourceQid = node.getAttribute('data-qid') || node.dataset?.qid || wrap.getAttribute('data-qid') || node.id || '';
                if (sourceQid) qWrapper.dataset.qid = String(sourceQid);
                qWrapper.tabIndex = 0;
                qWrapper.setAttribute('aria-label', `第 ${questionIndex} 题`);

                const qDiv = document.createElement('div');
                qDiv.className = 'zujuanjs-question';
                qDiv.style.fontFamily = font;
                qDiv.style.fontSize = size;
                qDiv.style.lineHeight = lineHeight;
                let questionBody = null;

                if (includeQ) {
                    const cnt = wrap.querySelector('.exam-item__cnt');
                    if (cnt) {
                        // 原卷题号优先：完整保留网站已经渲染好的题干与题号。
                        // questionIndex 仅用于内部答案对应、编辑与诊断，不再生成可见题号。
                        const cntClone = cnt.cloneNode(true);
                        const questionLayout = document.createElement('div');
                        questionLayout.className = 'zujuanjs-question-layout zujuanjs-original-numbering';
                        questionBody = document.createElement('div');
                        questionBody.className = 'zujuanjs-question-body';
                        questionBody.appendChild(cntClone);
                        questionLayout.appendChild(questionBody);
                        qDiv.appendChild(questionLayout);
                    }
                }

                const liveOpt = this.getAnswerNodeForQuestion(node);
                const savedOpt = this.getQuestionAnswerSnapshot(node, questionIndex);
                // 组卷中心在点击下一题后可能折叠/复用上一题答案区域；优先使用逐题确认过的快照。
                const opt = this.isGroupExamCenterPage()
                    ? (savedOpt || (liveOpt && this.isAnswerContentReady(liveOpt) ? liveOpt : null))
                    : (liveOpt || savedOpt);
                if (opt) {
                    const optClone = opt.cloneNode(true);
                    optClone.classList.add('zujuanjs-answer-content');
                    optClone.dataset.answerIndex = String(questionIndex);
                    optClone.querySelector('.knowledge-box')?.remove();

                    if (includeA) {
                        (questionBody || qDiv).appendChild(optClone);
                    } else if (atEnd) {
                        // 按 1.txt：不再重新包装答案、不人为增加“第 X 题解析”标题，
                        // 直接保留网站已经渲染好的 .exam-item__opt 结构。
                        answersEndList.push(optClone);
                    }
                }

                qWrapper.appendChild(qDiv);
                tempDiv.appendChild(qWrapper);
                questionIndex++;
            });

            if (atEnd && answersEndList.length) {
                const answerPageBreak = document.createElement('div');
                answerPageBreak.className = 'zujuanjs-answer-page-break';
                answerPageBreak.setAttribute('aria-hidden', 'true');
                tempDiv.appendChild(answerPageBreak);

                const section = document.createElement('div');
                section.className = 'zujuanjs-section-title zujuanjs-answer-section-title';
                section.style.fontFamily = 'SimHei, "Microsoft YaHei", "PingFang SC", sans-serif';
                section.textContent = '答案与解析';
                tempDiv.appendChild(section);
                answersEndList.forEach(answer => tempDiv.appendChild(answer));
            }

            this.preparePreviewTypography(tempDiv);

            tempDiv.querySelectorAll('img').forEach(img => {
                if (img.src) img.setAttribute('src', img.src);
                // 答案分页必须围绕一份固定像素资源计算，禁止预览 iframe 再根据 srcset 选择不同密度版本。
                if (/getAnswerAndParse/i.test(String(img.getAttribute('src') || ''))) {
                    img.removeAttribute('srcset');
                    img.removeAttribute('sizes');
                    img.setAttribute('loading', 'eager');
                    img.setAttribute('decoding', 'sync');
                }
            });

            const contentHtml = tempDiv.innerHTML;
            const fontWeight = pageBold ? 'bold' : 'normal';

            return `
            <!DOCTYPE html>
            <html lang="zh-CN">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1">
                <title>A4 试卷预览</title>
                <style>
                    body { margin: 0; background: #525659; font-family: sans-serif; padding-bottom: 40px; }

                    .preview-toolbar {
                        position: fixed; top: 0; left: 0; right: 0; height: 50px;
                        background: #333; color: #fff; display: flex; align-items: center;
                        justify-content: center; gap: 15px; z-index: 1000;
                        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
                    }
                    .preview-toolbar button {
                        background: #1677ff; color: #fff; border: none; padding: 8px 22px;
                        border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;
                        transition: background 0.2s;
                    }
                    .preview-toolbar button:hover { background: #4096ff; }
                    .preview-toolbar button.close { background: #ff4d4f; }
                    .preview-toolbar button.close:hover { background: #ff7875; }
                    .preview-hint { font-size: 12px; color: #aaa; margin-right: 20px; }

                    /* A4 纸张渲染 */
                    .paper-container { padding-top: 70px; }
                    .paper {
                        width: 210mm;
                        min-height: 297mm;
                        background: #fff;
                        margin: 0 auto 20px;
                        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
                        box-sizing: border-box;
                        position: relative;
                        overflow: hidden;
                    }
                    .paper-inner {
                        padding: 18mm 15mm 22mm 15mm;
                        position: relative;
                        box-sizing: border-box;
                        min-height: 297mm;
                    }
                    .content-wrapper { position: relative; z-index: 2; }

                    .q-wrapper { position: relative; padding-top: 15px; margin-bottom: 10px; page-break-inside: avoid; }
                    .q-toolbar {
                        position: absolute; top: -12px; right: 0;
                        background: rgba(22, 119, 255, 0.9); border-radius: 4px; padding: 2px;
                        display: flex; gap: 2px; opacity: 0; transition: opacity 0.2s; z-index: 10;
                    }
                    .q-wrapper:hover .q-toolbar,
                    .q-wrapper:focus-within .q-toolbar { opacity: 1; }
                    .q-toolbar button {
                        background: transparent; color: #fff; border: none;
                        padding: 3px 8px; font-size: 11px; cursor: pointer; border-radius: 2px;
                    }
                    .q-toolbar button:hover { background: rgba(255,255,255,0.2); }

                    .answer-blank,
                    .answer-blank-large {
                        display: block;
                        width: 100%;
                        margin-top: 4px;
                        border: 0;
                        background: transparent;
                        page-break-inside: avoid;
                    }
                    .answer-blank { height: 2.5em; }
                    .answer-blank-large { height: 12em; }

                    /* JS 动态插入的自然分页占位符 */
                    .auto-page-spacer {
                        position: relative;
                        width: 100%;
                        pointer-events: none;
                    }
                    .page-break-line {
                        position: absolute;
                        bottom: 0;
                        left: 0;
                        right: 0;
                        border-bottom: 2px dashed #1677ff;
                        color: #1677ff;
                        font-size: 12px;
                        text-align: center;
                        padding-bottom: 5px;
                        background: #fff;
                    }

                    /* 用户手动插入的分页符 (红色) */
                    .page-break {
                        border-top: 2px dashed #ff4d4f; margin: 15px 0; position: relative; height: 0; z-index: 3; page-break-inside: avoid;
                    }
                    .page-break::after {
                        content: "↓ 强制分页符 (打印时在此处换页) ↓";
                        position: absolute; top: -10px; left: 50%; transform: translateX(-50%);
                        background: #fff; color: #ff4d4f; font-size: 12px; padding: 0 8px; white-space: nowrap;
                    }

                    .zujuanjs-section-title { font-family: SimHei, \"Microsoft YaHei\", \"PingFang SC\", sans-serif !important; font-size: 1.10em; font-weight: 700; line-height: 1.35; margin: 14px 0 7px; border: 0; padding: 0; letter-spacing: .01em; break-inside: avoid; break-after: avoid-page; page-break-inside: avoid; page-break-after: avoid; }
                    .zujuanjs-answer-page-break { display: none; width: 0; height: 0; margin: 0; padding: 0; border: 0; }
                    .zujuanjs-print-title { font-family: SimHei, \"Microsoft YaHei\", \"PingFang SC\", sans-serif !important; font-size: 22px; text-align: center; font-weight: 700; margin: 0 0 18px; line-height: 1.28; letter-spacing: .025em; page-break-inside: avoid; }
                    .zujuanjs-question { margin-bottom: 0; padding: 0; border-bottom: none; }
                    .zujuanjs-question-number { font-weight: 600; white-space: pre; }
                    .zujuanjs-answer-item { margin-bottom: 18px; padding: 4px 0; border: none; }
                    .zujuanjs-answer-title { font-weight: bold; margin-bottom: 6px; }
                    /* 6.1.3：答案末尾模式直接使用网站原始 .exam-item__opt 结构 */
                    #source-content > .zujuanjs-answer-content,
                    #paper-container .zujuanjs-answer-content { display: block; width: 100%; margin: 0 0 8px; padding: 0; }
                    body[data-answer-long-image-mode="split"] #source-content > .zujuanjs-answer-content,
                    body[data-answer-long-image-mode="split"] #paper-container .zujuanjs-answer-content { margin: 0 !important; padding-top: 0 !important; padding-bottom: 0 !important; }
                    body[data-answer-long-image-mode="split"] .zujuanjs-answer-content > :first-child { margin-top: 0 !important; }
                    body[data-answer-long-image-mode="split"] .zujuanjs-answer-content > :last-child { margin-bottom: 0 !important; }
                    img { max-width: 100%; }

                    @media print {
                        body { background: #fff; margin: 0; padding: 0; }
                        .preview-toolbar { display: none !important; }
                        .paper-container { padding: 0; }
                        .paper {
                            width: 100%; min-height: 0; margin: 0; box-shadow: none; padding: 0; overflow: visible;
                        }
                        .paper-inner { padding: 0; min-height: 0; }
                        .q-toolbar { display: none !important; }
                        .page-break-line { display: none !important; }
                        .page-break {
                            border: none; margin: 0; height: 0;
                        }
                        .page-break::after { display: none; }

                        .zujuanjs-print-title, .zujuanjs-section-title, .zujuanjs-question, .zujuanjs-answer-item, .zujuanjs-answer-title {
                            color: #000 !important;
                        }
                        img { filter: ${printImageMode === 'grayscale' ? 'grayscale(100%) contrast(120%)' : 'none'}; }
                        * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
                    }

                    /* ===== 5.3: 类 Word 的独立 A4 多页预览 ===== */
                    #source-content { display: none !important; }
                    .paper-container {
                        padding: 70px 24px 40px;
                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        gap: 20px;
                        overflow-x: auto;
                    }
                    .paper {
                        width: 210mm;
                        height: 297mm;
                        min-height: 297mm;
                        margin: 0;
                        flex: 0 0 auto;
                        overflow: hidden;
                        background: #fff;
                        box-shadow: 0 4px 14px rgba(0,0,0,0.38);
                        position: relative;
                        box-sizing: border-box;
                    }
                    .paper-content {
                        position: absolute;
                        top: 18mm;
                        right: 15mm;
                        bottom: 22mm;
                        left: 15mm;
                        width: 180mm;
                        height: 257mm;
                        overflow: hidden;
                        box-sizing: border-box;
                        display: flow-root;
                        font-family: ${font};
                        font-size: ${size};
                        line-height: ${lineHeight};
                    }
                    .paper-content img { max-height: 245mm; object-fit: contain; }
                    .page-footer {
                        position: absolute;
                        left: 15mm;
                        right: 15mm;
                        bottom: 7mm;
                        text-align: center;
                        color: #000;
                        font-family: ${pageFont};
                        font-size: ${pageSize};
                        font-weight: ${fontWeight};
                        line-height: 1;
                    }
                    .q-wrapper {
                        position: relative;
                        padding-top: 1px;
                        margin-bottom: 6px;
                        break-inside: auto;
                        page-break-inside: auto;
                    }
                    .q-wrapper.continued-from-previous { padding-top: 0; }
                    .q-wrapper.continues-on-next { margin-bottom: 0; }
                    .q-wrapper.continues-on-next .zujuanjs-question,
                    .q-wrapper.continues-on-next .zujuanjs-answer-item { margin-bottom: 0; }
                    .q-wrapper.continued-from-previous .zujuanjs-question,
                    .q-wrapper.continued-from-previous .zujuanjs-answer-item { padding-top: 0; }
                    .q-toolbar { top: -14px; }
                    .manual-break-indicator {
                        position: absolute;
                        left: 0;
                        right: 0;
                        bottom: 0;
                        border-bottom: 1px dashed #ff4d4f;
                        color: #ff4d4f;
                        font-size: 11px;
                        line-height: 18px;
                        text-align: center;
                        pointer-events: none;
                    }
                    .preview-page-count { min-width: 52px; font-size: 12px; color: #d0d0d0; }

                    @media screen and (max-width: 850px) {
                        .paper-container { align-items: flex-start; padding-left: 12px; padding-right: 12px; }
                        .preview-hint { display: none; }
                        .preview-toolbar { justify-content: flex-start; padding: 0 10px; box-sizing: border-box; }
                    }

                    @media print {
                        @page { size: A4; margin: 0 !important; }
                        html, body { width: 210mm; margin: 0 !important; padding: 0 !important; background: #fff; }
                        .paper-container { display: block; padding: 0; margin: 0; overflow: visible; }
                        .paper {
                            width: 210mm;
                            height: 297mm;
                            min-height: 297mm;
                            margin: 0;
                            box-shadow: none;
                            overflow: hidden;
                            break-after: page;
                            page-break-after: always;
                        }
                        .paper:last-child { break-after: auto; page-break-after: auto; }
                        .paper-content { top: 18mm; right: 15mm; bottom: 22mm; left: 15mm; width: 180mm; height: 257mm; }
                        .q-toolbar, .manual-break-indicator { display: none !important; }
                    }

                    /* ===== 5.4: 双页并排、自动缩放与精简工具栏 ===== */
                    :root {
                        --preview-scale: 1;
                        --paper-display-width: 793.7px;
                        --paper-display-height: 1122.52px;
                    }
                    html { scrollbar-gutter: stable; }
                    body { min-height: 100vh; padding: 0; background: #45484b; letter-spacing: 0; }
                    .preview-toolbar {
                        height: 56px;
                        padding: 0 14px;
                        justify-content: space-between;
                        gap: 12px;
                        background: #202224;
                        box-shadow: 0 1px 5px rgba(0,0,0,0.35);
                        box-sizing: border-box;
                    }
                    .toolbar-section { display: flex; align-items: center; gap: 8px; min-width: 0; }
                    .toolbar-segment,
                    .zoom-control {
                        display: flex;
                        align-items: center;
                        height: 34px;
                        padding: 2px;
                        border: 1px solid #45484b;
                        border-radius: 6px;
                        background: #2b2e31;
                        box-sizing: border-box;
                    }
                    .preview-toolbar button {
                        height: 28px;
                        min-width: 30px;
                        padding: 0 9px;
                        border: 0;
                        border-radius: 4px;
                        background: transparent;
                        color: #d6d9dc;
                        font-size: 12px;
                        font-weight: 500;
                        line-height: 28px;
                        letter-spacing: 0;
                        white-space: nowrap;
                    }
                    .preview-toolbar button:hover { background: #3b3f43; color: #fff; }
                    .preview-toolbar button.active { background: #fff; color: #202224; }
                    .preview-toolbar .print-action {
                        height: 34px;
                        padding: 0 15px;
                        border-radius: 5px;
                        background: #1677ff;
                        color: #fff;
                        line-height: 34px;
                    }
                    .preview-toolbar .print-action:hover { background: #0f68df; }
                    .preview-toolbar button.close {
                        width: 34px;
                        height: 34px;
                        min-width: 34px;
                        padding: 0;
                        border-radius: 5px;
                        background: transparent;
                        color: #d6d9dc;
                        font-size: 20px;
                        line-height: 32px;
                    }
                    .preview-toolbar button.close:hover { background: #c9363e; color: #fff; }
                    .zoom-control button { padding: 0; font-size: 18px; font-weight: 400; }
                    .zoom-value { min-width: 48px !important; padding: 0 4px !important; font-variant-numeric: tabular-nums; }
                    .preview-page-count {
                        min-width: 48px;
                        color: #b9bdc1;
                        font-size: 12px;
                        font-variant-numeric: tabular-nums;
                        white-space: nowrap;
                    }
                    .paper-container {
                        display: grid;
                        align-items: start;
                        justify-content: safe center;
                        gap: 20px;
                        padding: 76px 20px 40px;
                        overflow: visible;
                        box-sizing: border-box;
                    }
                    .paper-container.layout-single { grid-template-columns: var(--paper-display-width); }
                    .paper-container.layout-double { grid-template-columns: repeat(2, var(--paper-display-width)); }
                    .paper-shell {
                        width: var(--paper-display-width);
                        height: var(--paper-display-height);
                        position: relative;
                        overflow: visible;
                    }
                    .paper {
                        width: 210mm;
                        height: 297mm;
                        min-height: 297mm;
                        margin: 0;
                        transform: scale(var(--preview-scale));
                        transform-origin: top left;
                        box-shadow: 0 3px 12px rgba(0,0,0,0.4);
                    }
                    .paper-content {
                        top: ${marginTop}mm;
                        right: ${marginRight}mm;
                        bottom: ${marginBottom}mm;
                        left: ${marginLeft}mm;
                        width: ${contentWidth}mm;
                        height: ${contentHeight}mm;
                    }
                    .paper-content img { max-height: ${Math.max(20, contentHeight - 12)}mm; }
                    .page-footer {
                        left: ${marginLeft}mm;
                        right: ${marginRight}mm;
                        bottom: ${footerBottom}mm;
                    }
                    .q-wrapper { margin-bottom: ${questionSpacing}px; }
                    .zujuanjs-question, .zujuanjs-answer-item { margin-bottom: 0; }

                    @media screen and (max-width: 680px) {
                        .preview-toolbar { padding: 0 8px; gap: 7px; overflow-x: auto; }
                        .toolbar-section { gap: 5px; }
                        .preview-toolbar .print-action { width: 34px; min-width: 34px; padding: 0; font-size: 16px; }
                        .print-action-label { display: none; }
                        .paper-container { padding-left: 12px; padding-right: 12px; gap: 12px; }
                    }

                    @media print {
                        .paper-container { display: block !important; padding: 0 !important; }
                        .paper-shell {
                            width: 210mm !important;
                            height: 297mm !important;
                            margin: 0 !important;
                            break-after: page;
                            page-break-after: always;
                        }
                        .paper-shell:last-child { break-after: auto; page-break-after: auto; }
                        .paper {
                            transform: none !important;
                            break-after: auto !important;
                            page-break-after: auto !important;
                        }
                        .paper-content {
                            top: ${marginTop}mm;
                            right: ${marginRight}mm;
                            bottom: ${marginBottom}mm;
                            left: ${marginLeft}mm;
                            width: ${contentWidth}mm;
                            height: ${contentHeight}mm;
                        }
                    }

                    /* ===== 5.5: 预览内实时排版工具 ===== */
                    :root {
                        --question-font: ${font};
                        --question-size: ${size};
                        --question-line-height: ${lineHeight};
                        --answer-font: ${answerFont};
                        --answer-size: ${answerSize};
                        --answer-line-height: ${answerLineHeight};
                        --page-font: ${pageFont};
                        --page-size: ${pageSize};
                        --page-weight: ${fontWeight};
                        --page-margin-top: ${marginTop}mm;
                        --page-margin-right: ${marginRight}mm;
                        --page-margin-bottom: ${marginBottom}mm;
                        --page-margin-left: ${marginLeft}mm;
                        --page-content-width: ${contentWidth}mm;
                        --page-content-height: ${contentHeight}mm;
                        --page-image-max-height: ${Math.max(20, contentHeight - 12)}mm;
                        --page-footer-bottom: ${footerBottom}mm;
                        --question-spacing: ${questionSpacing}px;
                        --title-size: ${titleSize}px;
                    }
                    .editor-toggle { font-size: 17px !important; }
                    .editor-panel {
                        position: fixed;
                        top: 56px;
                        right: 0;
                        bottom: 0;
                        z-index: 900;
                        width: 318px;
                        border-left: 1px solid #d8dadd;
                        background: #fff;
                        box-shadow: -4px 0 16px rgba(0,0,0,0.14);
                        transform: translateX(100%);
                        transition: transform 0.18s ease;
                        overflow-y: auto;
                        overscroll-behavior: contain;
                    }
                    .editor-panel.editor-initializing { transition: none; }
                    body.editor-open .editor-panel { transform: translateX(0); }
                    body.editor-open .paper-container { padding-right: 338px; }
                    .editor-panel-header {
                        position: sticky;
                        top: 0;
                        z-index: 2;
                        display: flex;
                        align-items: center;
                        justify-content: space-between;
                        height: 48px;
                        padding: 0 12px 0 16px;
                        border-bottom: 1px solid #e5e7ea;
                        background: #fff;
                        color: #202124;
                        font-size: 14px;
                        font-weight: 650;
                        box-sizing: border-box;
                    }
                    .editor-panel-header button {
                        width: 28px;
                        height: 28px;
                        padding: 0;
                        border: 0;
                        border-radius: 4px;
                        background: transparent;
                        color: #555b61;
                        font-size: 20px;
                        line-height: 26px;
                        cursor: pointer;
                    }
                    .editor-panel-header button:hover { background: #eef0f2; color: #1e2328; }
                    .editor-section { padding: 13px 16px; border-bottom: 1px solid #eceef0; }
                    .editor-section-title { margin: 0 0 10px; color: #34383d; font-size: 12px; font-weight: 650; }
                    .editor-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
                    .editor-field { display: grid; gap: 5px; min-width: 0; color: #6d737a; font-size: 11px; }
                    .editor-field.wide { grid-column: 1 / -1; }
                    .editor-field input,
                    .editor-field select {
                        width: 100%;
                        height: 32px;
                        padding: 0 8px;
                        border: 1px solid #d3d7db;
                        border-radius: 4px;
                        background: #fff;
                        color: #25292d;
                        font: inherit;
                        font-size: 12px;
                        box-sizing: border-box;
                        outline: none;
                    }
                    .editor-field input:focus,
                    .editor-field select:focus { border-color: #1677ff; box-shadow: 0 0 0 2px rgba(22,119,255,0.14); }
                    .editor-margin-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 6px; }
                    .editor-margin-grid .editor-field { gap: 3px; font-size: 10px; }
                    .editor-margin-grid input { padding: 0 5px; text-align: center; }
                    .editor-check { display: flex; align-items: center; gap: 7px; min-height: 32px; color: #34383d; font-size: 12px; }
                    .editor-check input { width: 15px; height: 15px; margin: 0; accent-color: #1677ff; }
                    .editor-reset {
                        width: 100%;
                        height: 32px;
                        border: 1px solid #d3d7db;
                        border-radius: 4px;
                        background: #fff;
                        color: #40454a;
                        font-size: 12px;
                        cursor: pointer;
                    }
                    .editor-reset:hover { border-color: #9ba1a7; background: #f7f8f9; }
                    .paper-content {
                        top: var(--page-margin-top) !important;
                        right: var(--page-margin-right) !important;
                        bottom: var(--page-margin-bottom) !important;
                        left: var(--page-margin-left) !important;
                        width: var(--page-content-width) !important;
                        height: var(--page-content-height) !important;
                        font-family: var(--question-font) !important;
                        font-size: var(--question-size) !important;
                        line-height: var(--question-line-height) !important;
                    }
                    .paper-content img { max-height: var(--page-image-max-height) !important; }
                    .page-footer {
                        left: var(--page-margin-left) !important;
                        right: var(--page-margin-right) !important;
                        bottom: var(--page-footer-bottom) !important;
                        font-family: var(--page-font) !important;
                        font-size: var(--page-size) !important;
                        font-weight: var(--page-weight) !important;
                    }
                    .q-wrapper { margin-bottom: var(--question-spacing) !important; }
                    .zujuanjs-print-title { font-size: var(--title-size) !important; }
                    .zujuanjs-question-layout {
                        display: grid;
                        grid-template-columns: max-content minmax(0, 1fr);
                        column-gap: 0.55em;
                        align-items: start;
                        width: 100%;
                        min-width: 0;
                    }
                    .zujuanjs-question-number {
                        grid-column: 1;
                        font-weight: 600;
                        white-space: nowrap;
                    }
                    .zujuanjs-question-number.continuation-placeholder { visibility: hidden; }
                    .zujuanjs-question-body {
                        grid-column: 2;
                        min-width: 0;
                    }
                    .zujuanjs-question-body > .exam-item__cnt,
                    .zujuanjs-question-body > .exam-item__opt { min-width: 0; }
                    .zujuanjs-question-body > .exam-item__cnt > :first-child,
                    .zujuanjs-question-body > :first-child { margin-top: 0 !important; }
                    .zujuanjs-question-body > .exam-item__opt:last-child,
                    .zujuanjs-question-body > :last-child { margin-bottom: 0; }
                    @media screen and (max-width: 680px) {
                        .editor-panel {
                            top: auto;
                            left: 0;
                            width: auto;
                            max-height: min(68vh, 560px);
                            border-top: 1px solid #d8dadd;
                            border-left: 0;
                            transform: translateY(110%);
                            box-shadow: 0 -4px 16px rgba(0,0,0,0.18);
                        }
                        body.editor-open .paper-container {
                            padding-right: 12px;
                            padding-bottom: calc(min(68vh, 560px) + 24px);
                        }
                        body.editor-open .editor-panel { transform: translateY(0); }
                    }
                    @media print {
                        .editor-panel, .editor-toggle { display: none !important; }
                        .paper-container { padding-right: 0 !important; }
                    }

                    .preview-toolbar > .toolbar-section:first-child,
                    .preview-toolbar > .toolbar-section:last-child { flex: 1 1 0; }
                    .preview-toolbar > .toolbar-section:last-child { justify-content: flex-end; }
                    .preview-toolbar > .toolbar-section:nth-child(2) { flex: 0 0 auto; }
                    .preview-toolbar .btn-group {
                        display: inline-flex;
                        align-items: center;
                        height: 34px;
                        padding: 0;
                        border: 0;
                        border-radius: 6px;
                        background: #2b2e31;
                        overflow: hidden;
                        vertical-align: middle;
                    }
                    .preview-toolbar .btn,
                    .q-toolbar .btn,
                    .editor-panel .btn,
                    .editor-panel .form-control,
                    .editor-panel .form-select,
                    .editor-panel .form-check-input { box-sizing: border-box; }
                    .preview-toolbar .btn {
                        display: inline-flex;
                        align-items: center;
                        justify-content: center;
                        gap: 5px;
                        height: 34px;
                        min-height: 34px;
                        min-width: 34px;
                        padding: 0 10px;
                        border: 1px solid transparent;
                        border-radius: 5px;
                        font-size: 12px;
                        font-weight: 500;
                        line-height: 1;
                        letter-spacing: 0;
                        white-space: nowrap;
                        cursor: pointer;
                        touch-action: manipulation;
                        transition: background-color 0.15s ease, border-color 0.15s ease, color 0.15s ease, box-shadow 0.15s ease;
                    }
                    .preview-toolbar .btn-outline-light {
                        color: #d6d9dc;
                        border: 1px solid #45484b;
                        background: #2b2e31;
                    }
                    .preview-toolbar .btn-outline-light:hover,
                    .preview-toolbar .btn-outline-light:focus-visible {
                        color: #fff;
                        border-color: #6c737a;
                        background: #3b3f43;
                    }
                    .preview-toolbar .btn-outline-light.active,
                    .preview-toolbar .btn-outline-light[aria-pressed="true"] {
                        color: #202224;
                        border-color: #fff;
                        background: #fff;
                        box-shadow: 0 1px 2px rgba(0,0,0,0.18);
                    }
                    .preview-toolbar .btn-primary {
                        color: #fff;
                        border-color: #1677ff;
                        background: #1677ff;
                    }
                    .preview-toolbar .btn-primary:hover,
                    .preview-toolbar .btn-primary:focus-visible { border-color: #0f68df; background: #0f68df; }
                    .preview-toolbar .toolbar-segment,
                    .preview-toolbar .zoom-control {
                        display: inline-flex;
                        align-items: center;
                        flex: 0 0 auto;
                    }
                    .preview-toolbar .btn-group > .btn {
                        position: relative;
                        margin: 0;
                        border-radius: 0 !important;
                    }
                    .preview-toolbar .btn-group > .btn + .btn { margin-left: -1px; }
                    .preview-toolbar .btn-group > .btn:first-child { border-radius: 5px 0 0 5px !important; }
                    .preview-toolbar .btn-group > .btn:last-child { border-radius: 0 5px 5px 0 !important; }
                    .preview-toolbar .btn-group:focus-within { box-shadow: 0 0 0 2px #69b1ff; }
                    .preview-toolbar .toolbar-segment .btn,
                    .preview-toolbar .zoom-control .btn { flex: 0 0 auto; }
                    .preview-toolbar .zoom-control .btn { padding: 0 7px; font-size: 17px; font-weight: 400; }
                    .preview-toolbar .zoom-control .zoom-value {
                        min-width: 58px;
                        padding: 0 5px;
                        font-size: 12px;
                        font-variant-numeric: tabular-nums;
                    }
                    .preview-toolbar .print-action {
                        height: 34px;
                        min-height: 34px;
                        min-width: 34px;
                        padding: 0 15px;
                        border-radius: 5px;
                        line-height: 1;
                    }
                    .preview-toolbar .editor-toggle { font-size: 16px; }
                    .preview-toolbar button.close {
                        width: 34px;
                        height: 34px;
                        min-width: 34px;
                        min-height: 34px;
                        padding: 0;
                        border: 1px solid #45484b;
                        background: #2b2e31;
                        font-size: 16px;
                        line-height: 1;
                    }
                    .preview-toolbar button.close:hover,
                    .preview-toolbar button.close:focus-visible { border-color: #c9363e; background: #c9363e; color: #fff; }
                    .toolbar-icon {
                        display: block;
                        width: 15px;
                        height: 15px;
                        flex: 0 0 auto;
                        fill: currentColor;
                    }
                    .preview-page-count {
                        display: inline-flex;
                        align-items: center;
                        height: 34px;
                    }
                    .q-toolbar {
                        display: flex;
                        align-items: center;
                        flex-wrap: wrap;
                        gap: 3px;
                        padding: 4px;
                    }
                    .q-toolbar .btn {
                        display: inline-flex;
                        align-items: center;
                        justify-content: center;
                        height: 26px;
                        min-height: 26px;
                        min-width: 0;
                        padding: 0 8px;
                        border: 0;
                        border-radius: 3px;
                        color: #174f8c;
                        background: rgba(255,255,255,0.96);
                        font-size: 11px;
                        font-weight: 600;
                        line-height: 1;
                        white-space: nowrap;
                    }
                    .q-toolbar .btn:hover,
                    .q-toolbar .btn:focus-visible { color: #0e3b6a; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.18); }
                    #paper-container .q-wrapper:focus-within > .q-toolbar,
                    #source-content .q-wrapper:focus-within > .q-toolbar { opacity: 1 !important; }
                    .preview-toolbar .btn:focus-visible,
                    .q-toolbar .btn:focus-visible,
                    .editor-panel .btn:focus-visible {
                        position: relative;
                        z-index: 2;
                        outline: 2px solid #69b1ff;
                        outline-offset: 1px;
                    }
                    .editor-panel .form-control,
                    .editor-panel .form-select {
                        width: 100%;
                        height: 32px;
                        min-height: 32px;
                        padding: 0 8px;
                        border: 1px solid #d3d7db;
                        border-radius: 4px;
                        background-color: #fff;
                        color: #25292d;
                        font: inherit;
                        font-size: 12px;
                        line-height: 1.2;
                        outline: none;
                    }
                    .editor-panel .form-select { padding-right: 28px; }
                    .editor-panel .form-control:focus,
                    .editor-panel .form-select:focus {
                        border-color: #1677ff;
                        box-shadow: 0 0 0 2px rgba(22,119,255,0.14);
                    }
                    .editor-panel .form-check-input {
                        width: 15px;
                        height: 15px;
                        margin: 0;
                        accent-color: #1677ff;
                        flex: 0 0 auto;
                    }
                    .editor-panel-header .btn {
                        display: inline-flex;
                        align-items: center;
                        justify-content: center;
                        width: 28px;
                        height: 28px;
                        min-height: 28px;
                        padding: 0;
                        border: 0;
                        border-radius: 4px;
                        background: transparent;
                        color: #555b61;
                        font-size: 20px;
                        line-height: 1;
                    }
                    .editor-panel-header .btn:hover,
                    .editor-panel-header .btn:focus-visible { background: #eef0f2; color: #1e2328; }
                    .editor-panel .editor-reset {
                        display: inline-flex;
                        align-items: center;
                        justify-content: center;
                        width: 100%;
                        height: 32px;
                        min-height: 32px;
                        padding: 0 12px;
                        border: 1px solid #c8cdd2;
                        border-radius: 4px;
                        color: #40454a;
                        background: #fff;
                        font-size: 12px;
                        line-height: 1;
                    }
                    .editor-panel .editor-reset:hover,
                    .editor-panel .editor-reset:focus-visible { border-color: #9ba1a7; background: #f7f8f9; }
                    .editor-grid > .wide { grid-column: 1 / -1; }

                    #source-content *,
                    #paper-container * { box-sizing: content-box; }
                    #paper-container,
                    #paper-container .paper-shell,
                    #paper-container .paper,
                    #paper-container .paper-content { box-sizing: border-box; }
                    #source-content .q-toolbar *,
                    #paper-container .q-toolbar *,
                    .preview-toolbar *,
                    .editor-panel * { box-sizing: border-box; }
                    #source-content p,
                    #paper-container p { margin: 1em 0; }
                    #source-content h1,
                    #source-content h2,
                    #source-content h3,
                    #source-content h4,
                    #source-content h5,
                    #source-content h6,
                    #paper-container h1,
                    #paper-container h2,
                    #paper-container h3,
                    #paper-container h4,
                    #paper-container h5,
                    #paper-container h6 { margin: revert; font-size: revert; font-weight: revert; line-height: revert; }
                    #source-content ul,
                    #source-content ol,
                    #source-content dl,
                    #paper-container ul,
                    #paper-container ol,
                    #paper-container dl { margin: revert; padding: revert; }
                    #source-content table,
                    #paper-container table { border-collapse: revert; border-spacing: revert; }
                    #source-content blockquote,
                    #source-content figure,
                    #paper-container blockquote,
                    #paper-container figure { margin: revert; }
                    #source-content img,
                    #source-content svg,
                    #paper-container img,
                    #paper-container svg { vertical-align: revert; }
                    #source-content th,
                    #source-content td,
                    #paper-container th,
                    #paper-container td {
                        padding: revert;
                        border-color: revert;
                        text-align: revert;
                    }
                    #source-content a,
                    #paper-container a { color: revert; text-decoration: revert; }
                    #source-content pre,
                    #source-content code,
                    #source-content kbd,
                    #source-content samp,
                    #paper-container pre,
                    #paper-container code,
                    #paper-container kbd,
                    #paper-container samp { font-family: revert; font-size: revert; }

                    @media screen and (max-width: 680px) {
                        .paper-container.layout-double { grid-template-columns: var(--paper-display-width); }
                        .preview-toolbar {
                            justify-content: flex-start;
                            overflow-x: auto;
                            scrollbar-width: none;
                        }
                        .preview-toolbar::-webkit-scrollbar { display: none; }
                        .preview-toolbar > .toolbar-section { flex: 0 0 auto !important; }
                        .preview-page-count { display: none; }
                        .preview-toolbar .btn { padding-left: 8px; padding-right: 8px; }
                        .preview-toolbar .print-action { width: 34px; padding: 0; }
                    }
                    @media screen and (max-width: 480px) {
                        .layout-label, .print-action-label { display: none; }
                        .preview-toolbar .toolbar-segment .btn { width: 36px; padding: 0; }
                    }
                    @media screen and (max-width: 360px) {
                        .preview-toolbar { gap: 5px; padding-left: 6px; padding-right: 6px; }
                        .preview-toolbar .toolbar-section { gap: 4px; }
                        .preview-toolbar .zoom-control [data-zoom="out"],
                        .preview-toolbar .zoom-control [data-zoom="in"] { display: none; }
                        .preview-toolbar .zoom-control .zoom-value { min-width: 50px; }
                    }
                    @media (hover: none), (pointer: coarse) {
                        .q-toolbar { opacity: 1; }
                    }

                    /* ===== 5.8: Word 式工作区、可调宽侧栏与顶层题目工具条 ===== */
                    :root {
                        --toolbar-height: 56px;
                        --editor-panel-width: ${editorPanelWidth}px;
                        --page-gap: ${pageGap}px;
                        --paragraph-spacing: ${paragraphSpacing}px;
                        --number-gap: ${numberGap}em;
                        --answer-row-height: ${answerRowHeight}em;
                        --content-align: ${contentAlign};
                    }
                    html, body {
                        width: 100%;
                        height: 100%;
                        margin: 0;
                        overflow: hidden;
                        background: #45484b;
                    }
                    body {
                        min-height: 0;
                        padding: 0;
                        color: #202124;
                        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
                    }
                    .preview-toolbar {
                        z-index: 2000;
                        height: var(--toolbar-height);
                    }
                    #source-content { display: none !important; }
                    .preview-workspace {
                        position: fixed;
                        top: var(--toolbar-height);
                        right: 0;
                        bottom: 0;
                        left: 0;
                        display: grid;
                        grid-template-columns: minmax(0, 1fr) 0;
                        min-width: 0;
                        min-height: 0;
                        overflow: hidden;
                        background: #45484b;
                    }
                    body.question-tools-open .preview-workspace { top: 112px; }
                    body.editor-open .preview-workspace {
                        grid-template-columns: minmax(0, 1fr) var(--editor-panel-width);
                    }
                    .page-viewport {
                        position: relative;
                        min-width: 0;
                        min-height: 0;
                        overflow: auto;
                        overscroll-behavior: contain;
                        scrollbar-gutter: stable;
                        background: #45484b;
                    }
                    .page-viewport:focus-visible { outline: 2px solid #69b1ff; outline-offset: -2px; }
                    .preview-workspace .paper-container,
                    body.editor-open .preview-workspace .paper-container {
                        display: grid;
                        min-height: 100%;
                        align-items: start;
                        justify-content: safe center;
                        gap: var(--page-gap);
                        padding: 20px 24px 40px !important;
                        overflow: visible;
                        box-sizing: border-box;
                    }
                    .paper-container.layout-single { grid-template-columns: var(--paper-display-width); }
                    .paper-container.layout-double { grid-template-columns: repeat(2, var(--paper-display-width)); }
                    .paper-shell {
                        width: var(--paper-display-width);
                        height: var(--paper-display-height);
                        position: relative;
                        overflow: visible;
                    }
                    .paper {
                        width: 210mm;
                        height: 297mm;
                        min-height: 297mm;
                        margin: 0;
                        overflow: hidden;
                        transform: scale(var(--preview-scale));
                        transform-origin: top left;
                        background: #fff;
                        box-shadow: 0 3px 12px rgba(0,0,0,0.42);
                    }
                    .paper-content {
                        overflow: hidden;
                        text-align: var(--content-align);
                    }
                    .q-wrapper {
                        position: relative;
                        padding-top: 1px !important;
                        margin-bottom: var(--question-spacing) !important;
                        outline: 0 solid transparent;
                        outline-offset: 2px;
                    }
                    .q-wrapper.is-selected {
                        outline: 0;
                        box-shadow: inset 2px 0 0 rgba(22,119,255,0.78);
                    }
                    .q-wrapper:focus-visible {
                        outline: 0;
                        box-shadow: inset 3px 0 0 #1677ff;
                    }
                    .zujuanjs-question-layout { column-gap: var(--number-gap); }
                    .zujuanjs-question-layout.zujuanjs-original-numbering { display: block !important; grid-template-columns: none !important; column-gap: 0 !important; }
                    .zujuanjs-question-body,
                    .zujuanjs-answer-item { text-align: var(--content-align); }
                    .zujuanjs-answer-item,
                    .zujuanjs-answer-content { font-family: var(--answer-font) !important; font-size: var(--answer-size) !important; line-height: var(--answer-line-height) !important; }
                    /* 组卷网答案内部有大量自带字号/字体的嵌套标签；统一压过网站样式。 */
                    .zujuanjs-answer-item :not(img):not(svg):not(svg *):not(math):not(math *):not(mjx-container):not(mjx-container *),
                    .zujuanjs-answer-content :not(img):not(svg):not(svg *):not(math):not(math *):not(mjx-container):not(mjx-container *) {
                        font-family: inherit !important;
                        font-size: inherit !important;
                        line-height: inherit !important;
                    }
                    /* getAnswerAndParse 返回的整块答案是位图，无法改图片内部字体，只能按答案字号同步视觉缩放。 */
                    .zujuanjs-answer-render-image {
                        display: block !important;
                        height: auto !important;
                        object-fit: contain !important;
                        transform-origin: left top !important;
                    }
                    /* 8.1：原图虚拟视窗 + 当前A4剩余空间动态分页；智能行间吸附优先，低置信度再用重叠保护。 */
                    body[data-answer-long-image-mode="split"] .zujuanjs-answer-render-image {
                        max-height: none !important;
                    }
                    .zujuanjs-answer-content,
                    .zujuanjs-answer-flow-html,
                    .zujuanjs-answer-image-flow {
                        margin: 0 !important;
                        padding: 0 !important;
                        min-height: 0 !important;
                        height: auto !important;
                        border: 0 !important;
                    }
                    .zujuanjs-answer-image-flow { display: contents !important; }
                    .zujuanjs-answer-image-slice {
                        position: relative !important;
                        display: block !important;
                        overflow: hidden !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        border: 0 !important;
                        font-size: 0 !important;
                        line-height: 0 !important;
                        break-inside: avoid !important;
                        page-break-inside: avoid !important;
                    }
                    .zujuanjs-answer-image-slice > img {
                        position: absolute !important;
                        left: 0 !important;
                        display: block !important;
                        max-width: none !important;
                        max-height: none !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        border: 0 !important;
                        object-fit: fill !important;
                    }
                    /* 6.1.6：不再显示 1/3、2/3 之类的子页编号。切线仅在预览中显示，并可手动微调。 */
                    .zujuanjs-answer-cut-line {
                        position: absolute;
                        top: 0;
                        left: 0;
                        right: 0;
                        z-index: 4;
                        height: 8px;
                        margin-top: -4px;
                        border-top: 1px dashed rgba(22,119,255,0.58);
                        cursor: ns-resize;
                        pointer-events: auto;
                        touch-action: none;
                    }
                    .zujuanjs-answer-cut-control {
                        position: absolute;
                        top: 3px;
                        right: 4px;
                        z-index: 6;
                        display: inline-flex;
                        align-items: center;
                        gap: 2px;
                        padding: 2px;
                        border: 1px solid rgba(22,119,255,0.28);
                        border-radius: 5px;
                        background: rgba(255,255,255,0.94);
                        box-shadow: 0 1px 4px rgba(0,0,0,0.12);
                        font-size: 10px !important;
                        line-height: 1 !important;
                    }
                    .zujuanjs-answer-cut-control span { color: #5f6b76; padding: 0 3px; white-space: nowrap; }
                    .zujuanjs-answer-cut-control button {
                        min-width: 26px;
                        height: 22px;
                        padding: 0 5px;
                        border: 1px solid #ccd7e3;
                        border-radius: 3px;
                        background: #fff;
                        color: #145da8;
                        font-size: 10px !important;
                        line-height: 20px !important;
                        cursor: pointer;
                    }
                    .zujuanjs-answer-cut-control button:hover { background: #edf6ff; border-color: #91caff; }
                    .zujuanjs-answer-cut-numeric {
                        display: inline-flex;
                        align-items: center;
                        gap: 2px;
                        margin-left: 2px;
                        padding-left: 3px;
                        border-left: 1px solid #d9e2ec;
                        color: #5f6b76;
                        white-space: nowrap;
                        font-size: 10px !important;
                        line-height: 1 !important;
                    }
                    .zujuanjs-answer-cut-numeric input {
                        width: 54px;
                        height: 22px;
                        padding: 0 3px;
                        border: 1px solid #ccd7e3;
                        border-radius: 3px;
                        background: #fff;
                        color: #1f2328;
                        font: 10px/20px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
                        text-align: right;
                        box-sizing: border-box;
                    }
                    .zujuanjs-answer-cut-numeric input:focus { outline: 1px solid #4096ff; border-color: #4096ff; }
                    .zujuanjs-diagnostic-card {
                        margin-top: 8px; padding: 9px 10px; border: 1px solid #e5e7eb; border-radius: 6px;
                        background: #f8fafc; color: #475569; font-size: 11px; line-height: 1.55;
                    }
                    .zujuanjs-diagnostic-card strong { color: #1f2937; font-weight: 600; }
                    .zujuanjs-diagnostic-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:4px 10px; margin-top:5px; }
                    .zujuanjs-title-tools { display:flex; gap:5px; margin-top:5px; }
                    .zujuanjs-title-tools button { flex:1; min-height:26px; font-size:11px; }
                    .zujuanjs-answer-cut-control[data-cut-mode="manual"] { border-color: rgba(82,196,26,.48); background: rgba(246,255,237,.96); }
                    .zujuanjs-answer-cut-control[data-cut-mode="smart"] { border-color: rgba(19,194,194,.52); background: rgba(230,255,251,.97); }
                    .zujuanjs-answer-cut-control[data-cut-mode="protect"] { border-color: rgba(250,173,20,.52); background: rgba(255,251,230,.97); }
                    @media print {
                        .zujuanjs-answer-cut-line,
                        .zujuanjs-answer-cut-control { display: none !important; }
                    }
                    #source-content .zujuanjs-question-body p,
                    #source-content .zujuanjs-answer-item p,
                    #paper-container .zujuanjs-question-body p,
                    #paper-container .zujuanjs-answer-item p {
                        margin-top: 0;
                        margin-bottom: var(--paragraph-spacing);
                    }
                    #source-content .zujuanjs-question-body p:last-child,
                    #source-content .zujuanjs-answer-item p:last-child,
                    #paper-container .zujuanjs-question-body p:last-child,
                    #paper-container .zujuanjs-answer-item p:last-child { margin-bottom: 0; }
                    .answer-blank {
                        display: block;
                        width: 100%;
                        height: var(--answer-row-height);
                        margin: 0;
                        padding: 0;
                        border: 0 !important;
                        outline: 0;
                        background: transparent !important;
                        box-shadow: none !important;
                        break-inside: avoid;
                        page-break-inside: avoid;
                    }
                    .answer-blank::before,
                    .answer-blank::after { display: none !important; content: none !important; }
                    .q-toolbar { display: none !important; }

                    .editor-panel {
                        position: relative !important;
                        inset: auto !important;
                        z-index: 1500;
                        display: flex;
                        width: 100% !important;
                        min-width: 0;
                        height: 100%;
                        flex-direction: column;
                        overflow: hidden;
                        border: 0;
                        border-left: 1px solid #d8dadd;
                        background: #fff;
                        box-shadow: -4px 0 16px rgba(0,0,0,0.14);
                        box-sizing: border-box;
                        opacity: 0;
                        visibility: hidden;
                        transform: none !important;
                        transition: opacity 0.12s ease;
                    }
                    body.editor-open .editor-panel {
                        opacity: 1;
                        visibility: visible;
                        transform: none !important;
                    }
                    .editor-resize-handle {
                        position: absolute;
                        top: 0;
                        bottom: 0;
                        left: -5px;
                        z-index: 5;
                        width: 10px;
                        border: 0;
                        background: transparent;
                        cursor: col-resize;
                        touch-action: none;
                    }
                    .editor-resize-handle::after {
                        position: absolute;
                        top: 0;
                        bottom: 0;
                        left: 4px;
                        width: 2px;
                        background: transparent;
                        content: "";
                        transition: background-color 0.15s ease;
                    }
                    .editor-resize-handle:hover::after,
                    .editor-resize-handle:focus-visible::after,
                    body.editor-resizing .editor-resize-handle::after { background: #1677ff; }
                    body.editor-resizing,
                    body.editor-resizing * { cursor: col-resize !important; user-select: none !important; }
                    .editor-panel-header {
                        position: relative;
                        flex: 0 0 48px;
                        height: 48px;
                        padding: 0 12px 0 16px;
                    }
                    .editor-panel-title { display: flex; min-width: 0; align-items: center; gap: 8px; }
                    .editor-panel-title small { color: #858b91; font-size: 10px; font-weight: 500; }
                    .editor-tablist {
                        display: grid;
                        grid-template-columns: 1fr 1fr;
                        flex: 0 0 40px;
                        gap: 4px;
                        padding: 5px 12px;
                        border-bottom: 1px solid #e5e7ea;
                        background: #f8f9fa;
                        box-sizing: border-box;
                    }
                    .editor-tablist .btn {
                        display: inline-flex;
                        height: 30px;
                        min-height: 30px;
                        align-items: center;
                        justify-content: center;
                        padding: 0 10px;
                        border: 1px solid transparent;
                        border-radius: 4px;
                        color: #555b61;
                        background: transparent;
                        font-size: 12px;
                        line-height: 1;
                    }
                    .editor-tablist .btn.active {
                        border-color: #cfd4da;
                        color: #202124;
                        background: #fff;
                        box-shadow: 0 1px 2px rgba(0,0,0,0.06);
                    }
                    .editor-panel-body {
                        min-height: 0;
                        flex: 1 1 auto;
                        overflow-y: auto;
                        overscroll-behavior: contain;
                    }
                    .editor-pane[hidden] { display: none !important; }
                    .editor-section { padding: 14px 16px; }
                    .editor-section-title { margin-bottom: 10px; }
                    .editor-grid { gap: 10px; }
                    .editor-grid.three-columns { grid-template-columns: repeat(3, minmax(0, 1fr)); }
                    .editor-panel .form-control,
                    .editor-panel .form-select { height: 32px; min-height: 32px; }
                    .editor-field-unit { position: relative; }
                    .editor-field-unit .form-control { padding-right: 30px; }
                    .editor-field-unit > span {
                        position: absolute;
                        right: 8px;
                        bottom: 8px;
                        color: #8a9096;
                        font-size: 10px;
                        line-height: 1;
                        pointer-events: none;
                    }
                    .editor-panel-footer {
                        flex: 0 0 auto;
                        padding: 10px 16px 12px;
                        border-top: 1px solid #e5e7ea;
                        background: #fff;
                    }
                    .editor-footer-actions { display: grid; grid-template-columns: 1fr 1fr; gap: 7px; }
                    .editor-footer-actions .editor-reset { grid-column: 1 / -1; }
                    .editor-footer-actions .btn { min-height: 32px; font-size: 11px; }

                    /* ===== 5.9: 统一栅格、常用预设与自动保存状态 ===== */
                    .editor-panel {
                        border-left-color: #d6d9dd;
                        background: #f5f6f7;
                        box-shadow: -8px 0 24px rgba(0,0,0,0.13);
                    }
                    .editor-panel-header {
                        flex-basis: 52px;
                        height: 52px;
                        padding: 0 14px 0 16px;
                        border-bottom: 1px solid #dfe2e5;
                        background: #fff;
                    }
                    .editor-panel-title { gap: 7px; }
                    .editor-panel-title > span { color: #202428; font-size: 14px; font-weight: 650; }
                    .editor-save-status {
                        display: inline-flex;
                        align-items: center;
                        height: 20px;
                        padding: 0 6px;
                        border: 1px solid #dce7de;
                        border-radius: 3px;
                        color: #397447;
                        background: #f4faf4;
                        font-size: 10px;
                        font-weight: 500;
                        line-height: 1;
                        white-space: nowrap;
                    }
                    .editor-tablist {
                        flex-basis: 44px;
                        height: 44px;
                        gap: 5px;
                        padding: 6px 12px;
                        border-bottom-color: #dfe2e5;
                        background: #f5f6f7;
                    }
                    .editor-tablist .btn {
                        height: 32px;
                        min-height: 32px;
                        border-radius: 4px;
                        font-weight: 550;
                    }
                    .editor-panel-body { background: #fff; }
                    .editor-section {
                        padding: 16px;
                        border-bottom: 1px solid #e7e9eb;
                        background: #fff;
                    }
                    .editor-section-title {
                        display: flex;
                        align-items: center;
                        min-height: 16px;
                        margin: 0 0 12px;
                        color: #30353a;
                        font-size: 12px;
                        font-weight: 650;
                    }
                    .editor-section-note {
                        margin: -6px 0 12px;
                        color: #80868d;
                        font-size: 11px;
                        line-height: 1.45;
                    }
                    .editor-grid {
                        grid-auto-rows: min-content;
                        gap: 12px 10px;
                    }
                    .editor-field {
                        grid-template-rows: 15px 32px;
                        align-content: start;
                        gap: 5px;
                        color: #686f76;
                        font-size: 11px;
                        line-height: 15px;
                    }
                    .editor-margin-grid {
                        grid-template-columns: repeat(4, minmax(0, 1fr));
                        gap: 8px;
                    }
                    .editor-margin-grid .editor-field {
                        grid-template-rows: 15px 32px;
                        gap: 5px;
                        color: #686f76;
                        font-size: 11px;
                        line-height: 15px;
                        text-align: center;
                    }
                    .editor-margin-grid .form-control { min-width: 0; padding: 0 3px; text-align: center; }
                    .editor-panel .form-control,
                    .editor-panel .form-select {
                        height: 32px;
                        min-height: 32px;
                        border-color: #d1d5d9;
                        border-radius: 4px;
                        color: #252a2f;
                        background-color: #fff;
                    }
                    .editor-panel .form-control:hover,
                    .editor-panel .form-select:hover { border-color: #aeb5bb; }
                    .editor-check {
                        min-height: 32px;
                        margin-top: 20px;
                        padding: 0 9px;
                        border: 1px solid #d1d5d9;
                        border-radius: 4px;
                        background: #fff;
                    }
                    .editor-check.wide { margin-top: 0; }
                    .editor-check:has(.form-check-input:checked) { border-color: #91caff; background: #f0f7ff; }
                    .editor-presets {
                        display: grid;
                        grid-template-columns: repeat(2, minmax(0, 1fr));
                        gap: 8px;
                    }
                    .editor-preset {
                        display: grid;
                        grid-template-columns: minmax(0, 1fr);
                        grid-template-rows: 18px 15px;
                        min-width: 0;
                        min-height: 52px;
                        padding: 8px 9px;
                        border: 1px solid #d4d8dc;
                        border-radius: 4px;
                        color: #343a40;
                        background: #fff;
                        text-align: left;
                        cursor: pointer;
                    }
                    .editor-preset:hover,
                    .editor-preset:focus-visible { border-color: #69aaf5; background: #f4f9ff; outline: 0; }
                    .editor-preset.is-active {
                        border-color: #1677ff;
                        color: #0f5ebc;
                        background: #eef6ff;
                        box-shadow: inset 3px 0 0 #1677ff;
                    }
                    .editor-preset strong,
                    .editor-preset small {
                        display: block;
                        overflow: hidden;
                        text-overflow: ellipsis;
                        white-space: nowrap;
                    }
                    .editor-preset strong { font-size: 12px; font-weight: 600; line-height: 18px; }
                    .editor-preset small { color: #7a828a; font-size: 10px; line-height: 15px; }
                    .editor-preset.is-active small { color: #4b7fbd; }
                    .editor-panel-footer {
                        padding: 12px 16px;
                        border-top-color: #dfe2e5;
                        background: #f5f6f7;
                    }
                    .editor-panel .editor-reset {
                        height: 34px;
                        min-height: 34px;
                        border-color: #bfc5cb;
                        border-radius: 4px;
                        color: #30363c;
                        background: #fff;
                        font-weight: 550;
                    }
                    @media screen and (max-width: 680px) {
                        .editor-panel-header { flex-basis: 50px; height: 50px; }
                        .editor-section { padding: 14px 16px; }
                    }

                    /* ===== 6.0: 连续工作台层次、公式字号比例与中英微间距 ===== */
                    :root {
                        --ui-workspace: #e7edeb;
                        --ui-toolbar: #183332;
                        --ui-toolbar-raised: #244442;
                        --ui-panel: #fbfcfb;
                        --ui-panel-subtle: #f2f6f4;
                        --ui-border: #d8e2de;
                        --ui-text: #24312f;
                        --ui-muted: #6c7a76;
                        --ui-accent: #08756d;
                        --ui-accent-soft: #e2f3ee;
                        --ui-primary: #2f67d8;
                    }
                    html, body,
                    .preview-workspace,
                    .page-viewport { background: var(--ui-workspace); }
                    .preview-toolbar {
                        border-bottom: 1px solid rgba(255,255,255,0.12);
                        background: var(--ui-toolbar);
                        box-shadow: 0 2px 12px rgba(17,41,39,0.20);
                    }
                    .preview-toolbar .btn-group {
                        border-color: rgba(229,245,241,0.22);
                        background: var(--ui-toolbar-raised);
                    }
                    .preview-toolbar .btn-outline-light {
                        border-color: rgba(229,245,241,0.15);
                        color: #d9e8e4;
                        background: var(--ui-toolbar-raised);
                    }
                    .preview-toolbar .btn-outline-light:hover,
                    .preview-toolbar .btn-outline-light:focus-visible {
                        border-color: rgba(229,245,241,0.42);
                        color: #fff;
                        background: #315552;
                    }
                    .preview-toolbar .btn-outline-light.active,
                    .preview-toolbar .btn-outline-light[aria-pressed="true"] {
                        border-color: #cce9e2;
                        color: #075e58;
                        background: #e4f4ef;
                        box-shadow: none;
                    }
                    .preview-toolbar .btn-primary {
                        border-color: var(--ui-primary);
                        background: var(--ui-primary);
                    }
                    .preview-toolbar .btn-primary:hover,
                    .preview-toolbar .btn-primary:focus-visible { border-color: #2558c0; background: #2558c0; }
                    .preview-page-count { color: #b8cbc5; }
                    .paper { box-shadow: 0 10px 26px rgba(25,48,44,0.19), 0 1px 3px rgba(25,48,44,0.14); }
                    .paper-content { color: #1c2423; }
                    .zh-latin-gap {
                        display: inline-block;
                        width: 0.16em;
                        min-width: 0.16em;
                        vertical-align: baseline;
                    }
                    #source-content .zujuanjs-formula-svg,
                    #paper-container .zujuanjs-formula-svg {
                        display: inline-block;
                        max-width: 100%;
                        vertical-align: -0.16em;
                        object-fit: contain;
                    }
                    .editor-panel {
                        border-left-color: var(--ui-border);
                        background: var(--ui-panel);
                        box-shadow: -10px 0 28px rgba(26,50,46,0.12);
                    }
                    .editor-panel-header,
                    .editor-panel-body,
                    .editor-section { background: var(--ui-panel); }
                    .editor-panel-header,
                    .editor-tablist,
                    .editor-panel-footer { border-color: var(--ui-border); }
                    .editor-panel-title > span,
                    .editor-section-title { color: var(--ui-text); }
                    .editor-tablist,
                    .editor-panel-footer { background: var(--ui-panel-subtle); }
                    .editor-tablist .btn { color: var(--ui-muted); }
                    .editor-tablist .btn.active {
                        border-color: #c7ddd6;
                        color: #086b64;
                        background: #fff;
                        box-shadow: 0 1px 2px rgba(24,51,48,0.06);
                    }
                    .editor-section { border-bottom-color: #e4ebe8; }
                    .editor-section-note,
                    .editor-field { color: var(--ui-muted); }
                    .editor-panel .form-control,
                    .editor-panel .form-select,
                    .editor-check,
                    .editor-preset,
                    .editor-panel .editor-reset {
                        border-color: #d4dfdb;
                        color: var(--ui-text);
                        background: #fff;
                    }
                    .editor-panel .form-control:hover,
                    .editor-panel .form-select:hover,
                    .editor-check:hover,
                    .editor-preset:hover { border-color: #a9c6bd; }
                    .editor-panel .form-control:focus,
                    .editor-panel .form-select:focus {
                        border-color: var(--ui-accent);
                        box-shadow: 0 0 0 2px rgba(8,117,109,0.14);
                    }
                    .editor-check:has(.form-check-input:checked) {
                        border-color: #8ccdc0;
                        background: var(--ui-accent-soft);
                    }
                    .editor-preset:hover,
                    .editor-preset:focus-visible { border-color: #77b9ad; background: #f3faf7; }
                    .editor-preset.is-active {
                        border-color: var(--ui-accent);
                        color: #075e58;
                        background: var(--ui-accent-soft);
                        box-shadow: inset 3px 0 0 var(--ui-accent);
                    }
                    .editor-preset.is-active small { color: #3f7f74; }
                    .editor-save-status {
                        border-color: #c8e3da;
                        color: #176857;
                        background: #ecf8f3;
                    }
                    .editor-panel .editor-reset:hover,
                    .editor-panel .editor-reset:focus-visible { border-color: #93bbb0; background: #f4faf7; }
                    .question-float-toolbar {
                        border-color: #c9dbd5;
                        background: #fdfefd;
                        box-shadow: 0 10px 24px rgba(25,51,47,0.20);
                    }
                    .question-toolbar-label { color: var(--ui-text); }
                    .question-float-toolbar .btn[aria-pressed="true"] {
                        border-color: #75b9aa;
                        color: #075e58;
                        background: var(--ui-accent-soft);
                    }
                    .question-line-count {
                        border-color: #d4dfdb;
                        color: #53615d;
                        background: #f2f6f4;
                    }

                    .question-float-toolbar {
                        position: fixed;
                        top: calc(var(--toolbar-height) + 8px);
                        left: 8px;
                        z-index: 2147483000;
                        display: flex;
                        max-width: calc(100vw - 16px);
                        min-height: 40px;
                        align-items: center;
                        gap: 6px;
                        padding: 4px 6px;
                        overflow-x: auto;
                        border: 1px solid #cfd4da;
                        border-radius: 6px;
                        background: #fff;
                        box-shadow: 0 6px 20px rgba(0,0,0,0.22);
                        opacity: 0;
                        visibility: hidden;
                        pointer-events: none;
                        scrollbar-width: none;
                        transition: opacity 0.12s ease;
                    }
                    .question-float-toolbar::-webkit-scrollbar { display: none; }
                    .question-float-toolbar.is-visible {
                        opacity: 1;
                        visibility: visible;
                        pointer-events: auto;
                    }
                    .question-toolbar-label {
                        max-width: 94px;
                        overflow: hidden;
                        color: #34383d;
                        font-size: 12px;
                        font-weight: 650;
                        text-overflow: ellipsis;
                        white-space: nowrap;
                    }
                    .question-float-toolbar .btn-group { flex: 0 0 auto; }
                    .question-float-toolbar .btn {
                        display: inline-flex;
                        height: 32px;
                        min-height: 32px;
                        align-items: center;
                        justify-content: center;
                        padding: 0 9px;
                        border-radius: 4px;
                        font-size: 11px;
                        font-weight: 600;
                        line-height: 1;
                        letter-spacing: 0;
                        white-space: nowrap;
                    }
                    .question-float-toolbar .btn[aria-pressed="true"] {
                        border-color: #1677ff;
                        color: #0f5fc4;
                        background: #eaf3ff;
                    }
                    .question-float-toolbar .btn:focus-visible { outline: 2px solid #1677ff; outline-offset: 1px; }
                    .question-float-toolbar .btn-outline-danger > span[aria-hidden="true"] { display: none; }
                    .question-toolbar-short-label { display: none; }
                    .question-line-count {
                        display: inline-flex;
                        width: 52px;
                        height: 32px;
                        align-items: center;
                        justify-content: center;
                        border-top: 1px solid #dee2e6;
                        border-bottom: 1px solid #dee2e6;
                        color: #4f555b;
                        background: #f8f9fa;
                        font-size: 11px;
                        font-variant-numeric: tabular-nums;
                        white-space: nowrap;
                    }

                    @media screen and (max-width: 680px) {
                        .preview-workspace,
                        body.editor-open .preview-workspace { display: block; }
                        .page-viewport { position: absolute; inset: 0; }
                        .preview-workspace .paper-container,
                        body.editor-open .preview-workspace .paper-container {
                            gap: 12px;
                            padding: 12px 12px 32px !important;
                        }
                        .paper-container.layout-double { grid-template-columns: var(--paper-display-width); }
                        .editor-panel {
                            position: fixed !important;
                            top: auto !important;
                            right: 0 !important;
                            bottom: 0 !important;
                            left: 0 !important;
                            width: 100% !important;
                            height: min(72dvh, 620px);
                            max-height: none;
                            border-top: 1px solid #d8dadd;
                            border-left: 0;
                            opacity: 1;
                            visibility: visible;
                            transform: translateY(105%) !important;
                            transition: transform 0.18s ease !important;
                        }
                        body.editor-open .editor-panel { transform: translateY(0) !important; }
                        body:not(.editor-open) .editor-panel { visibility: hidden; }
                        .editor-resize-handle { display: none; }
                        body.editor-open .preview-workspace .paper-container {
                            padding-right: 12px !important;
                            padding-bottom: calc(min(72dvh, 620px) + 24px) !important;
                        }
                        .question-float-toolbar {
                            top: calc(var(--toolbar-height) + 8px) !important;
                            right: 8px;
                            left: 8px !important;
                            max-width: none;
                            min-height: 52px;
                            padding: 4px;
                            gap: 4px;
                        }
                        .question-float-toolbar .btn { height: 44px; min-height: 44px; padding: 0 10px; }
                        .question-line-count { height: 44px; }
                        .question-toolbar-label { display: none; }
                        body.editor-open .question-float-toolbar { display: none !important; }
                        body.question-tools-open .preview-workspace { top: 132px; }
                    }
                    @media screen and (max-width: 390px) {
                        .question-float-toolbar .btn { width: 40px; min-width: 40px; padding: 0; }
                        .question-line-count { width: 42px; }
                        .question-toolbar-wide-label { display: none; }
                        .question-toolbar-short-label { display: inline; }
                        .question-float-toolbar .btn-outline-danger > span[aria-hidden="true"] { display: inline; }
                    }
                    @media screen and (max-width: 340px) {
                        .question-float-toolbar [data-question-action="add-lines"] { display: none; }
                    }
                    @media print {
                        html, body {
                            width: 210mm !important;
                            height: auto !important;
                            overflow: visible !important;
                            background: #fff !important;
                        }
                        .preview-toolbar,
                        .editor-panel,
                        .question-float-toolbar { display: none !important; }
                        .preview-workspace {
                            position: static !important;
                            display: block !important;
                            overflow: visible !important;
                            background: #fff !important;
                        }
                        .page-viewport { position: static !important; overflow: visible !important; }
                        .preview-workspace .paper-container,
                        body.editor-open .preview-workspace .paper-container {
                            display: block !important;
                            min-height: 0 !important;
                            padding: 0 !important;
                            overflow: visible !important;
                        }
                        .paper-shell {
                            width: 210mm !important;
                            height: 297mm !important;
                            margin: 0 !important;
                            break-after: page;
                            page-break-after: always;
                        }
                        .paper-shell:last-child { break-after: auto; page-break-after: auto; }
                        .paper { transform: none !important; box-shadow: none !important; }
                        .q-wrapper,
                        .q-wrapper.is-selected,
                        .q-wrapper:focus,
                        .q-wrapper:focus-visible,
                        .q-wrapper:focus-within { outline: 0 !important; box-shadow: none !important; }
                    }
                </style>
            </head>
            <body>
                <div class="preview-toolbar" role="toolbar" aria-label="预览工具栏">
                    <div class="toolbar-section">
                        <div class="toolbar-segment btn-group btn-group-sm" role="group" aria-label="预览排布">
                            <button type="button" class="btn btn-outline-light" data-layout="single" title="单页预览" aria-label="单页预览"><svg class="toolbar-icon" viewBox="0 0 16 16" aria-hidden="true"><path d="M3 1.5A1.5 1.5 0 0 1 4.5 0h7A1.5 1.5 0 0 1 13 1.5v13a1.5 1.5 0 0 1-1.5 1.5h-7A1.5 1.5 0 0 1 3 14.5v-13ZM4.5 1a.5.5 0 0 0-.5.5v13a.5.5 0 0 0 .5.5h7a.5.5 0 0 0 .5-.5v-13a.5.5 0 0 0-.5-.5h-7Z"/></svg><span class="layout-label">单页</span></button>
                            <button type="button" class="btn btn-outline-light" data-layout="double" title="双页并排" aria-label="双页并排"><svg class="toolbar-icon" viewBox="0 0 16 16" aria-hidden="true"><path d="M1 2.5A1.5 1.5 0 0 1 2.5 1h4A1.5 1.5 0 0 1 8 2.5v11A1.5 1.5 0 0 1 6.5 15h-4A1.5 1.5 0 0 1 1 13.5v-11ZM2.5 2a.5.5 0 0 0-.5.5v11a.5.5 0 0 0 .5.5h4a.5.5 0 0 0 .5-.5v-11a.5.5 0 0 0-.5-.5h-4ZM8 2.5A1.5 1.5 0 0 1 9.5 1h4A1.5 1.5 0 0 1 15 2.5v11a1.5 1.5 0 0 1-1.5 1.5h-4A1.5 1.5 0 0 1 8 13.5v-11ZM9.5 2a.5.5 0 0 0-.5.5v11a.5.5 0 0 0 .5.5h4a.5.5 0 0 0 .5-.5v-11a.5.5 0 0 0-.5-.5h-4Z"/></svg><span class="layout-label">双页</span></button>
                        </div>
                        <span class="preview-page-count" id="preview-page-count" aria-live="polite" aria-atomic="true">排版中…</span>
                    </div>
                    <div class="toolbar-section">
                        <div class="zoom-control btn-group btn-group-sm" role="group" aria-label="预览缩放">
                            <button type="button" class="btn btn-outline-light" data-zoom="out" title="缩小" aria-label="缩小">−</button>
                            <button type="button" class="btn btn-outline-light zoom-value" id="zoom-value" data-zoom="auto" title="自动适应">自动</button>
                            <button type="button" class="btn btn-outline-light" data-zoom="in" title="放大" aria-label="放大">+</button>
                        </div>
                    </div>
                    <div class="toolbar-section">
                        <button type="button" class="btn btn-outline-light btn-sm editor-toggle" id="editor-toggle" title="排版工具" aria-label="排版工具" aria-controls="editor-panel" aria-expanded="false"><svg class="toolbar-icon" viewBox="0 0 16 16" aria-hidden="true"><path d="M9.405 1.05c-.413-1.4-2.397-1.4-2.81 0l-.1.34a1.47 1.47 0 0 1-2.105.872l-.31-.17c-1.283-.698-2.686.705-1.987 1.987l.169.311c.446.82.023 1.841-.872 2.105l-.34.1c-1.4.413-1.4 2.397 0 2.81l.34.1a1.47 1.47 0 0 1 .872 2.105l-.17.31c-.698 1.283.705 2.686 1.987 1.987l.311-.169a1.47 1.47 0 0 1 2.105.872l.1.34c.413 1.4 2.397 1.4 2.81 0l.1-.34a1.47 1.47 0 0 1 2.105-.872l.31.17c1.283.698 2.686-.705 1.987-1.987l-.169-.311a1.47 1.47 0 0 1 .872-2.105l.34-.1c1.4-.413 1.4-2.397 0-2.81l-.34-.1a1.47 1.47 0 0 1-.872-2.105l.17-.31c.698-1.283-.705-2.686-1.987-1.987l-.311.169a1.47 1.47 0 0 1-2.105-.872l-.1-.34ZM8 10.5a2.5 2.5 0 1 1 0-5 2.5 2.5 0 0 1 0 5Z"/></svg></button>
                        <button type="button" class="btn btn-outline-light btn-sm" id="quick-answer-end" title="切换：答案移至末尾" aria-label="答案移至末尾" aria-pressed="false">答末</button>
                        <button type="button" class="btn btn-outline-light btn-sm" id="export-word" title="导出可编辑 Word（DOCX）">Word</button>
                        <button type="button" class="btn btn-outline-light btn-sm" id="export-html" title="导出独立 HTML">HTML</button>
                        <button type="button" class="btn btn-outline-light btn-sm" id="export-json" title="导出结构化试卷 JSON">JSON</button>
                        <button type="button" class="btn btn-primary btn-sm print-action" id="print-now" title="打印（会先完成答案安全切线分析）"><svg class="toolbar-icon" viewBox="0 0 16 16" aria-hidden="true"><path d="M5 1a2 2 0 0 0-2 2v2h2V3h6v2h2V3a2 2 0 0 0-2-2H5Zm-1 9h8v5H4v-5Zm1 1v3h6v-3H5ZM2 5a2 2 0 0 0-2 2v4a2 2 0 0 0 2 2h1V9h10v4h1a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2H2Zm11 2.25a.75.75 0 1 1 1.5 0 .75.75 0 0 1-1.5 0Z"/></svg><span class="print-action-label">打印</span></button>
                        <button type="button" class="btn btn-outline-light btn-sm close" title="关闭预览" aria-label="关闭预览" onclick="window.parent.postMessage({type: 'closeZujuanPreview'}, '*')"><svg class="toolbar-icon" viewBox="0 0 16 16" aria-hidden="true"><path d="M4.146 4.146a.5.5 0 0 1 .708 0L8 7.293l3.146-3.147a.5.5 0 0 1 .708.708L8.707 8l3.147 3.146a.5.5 0 0 1-.708.708L8 8.707l-3.146 3.147a.5.5 0 0 1-.708-.708L7.293 8 4.146 4.854a.5.5 0 0 1 0-.708Z"/></svg></button>
                    </div>
                </div>
                <div id="source-content" style="font-family: ${font}; font-size: ${size}; line-height: ${lineHeight};">${contentHtml}</div>
                <div class="preview-workspace" id="preview-workspace">
                    <main class="page-viewport" id="page-viewport" tabindex="0" aria-label="试卷页面预览">
                        <div class="paper-container" id="paper-container" data-render-version="0"></div>
                    </main>
                    <aside class="editor-panel" id="editor-panel" aria-label="排版工具" aria-hidden="true">
                        <div class="editor-resize-handle" id="editor-resize-handle" role="separator" tabindex="0" aria-label="调整排版工具宽度" aria-orientation="vertical" aria-valuemin="280" aria-valuemax="520" aria-valuenow="${Math.round(editorPanelWidth)}"></div>
                        <div class="editor-panel-header">
                            <div class="editor-panel-title"><span>排版工具</span><small class="editor-save-status" id="editor-save-status" aria-live="polite">已自动保存</small></div>
                            <button type="button" class="btn btn-light btn-sm" id="editor-close" title="收起" aria-label="收起排版工具">×</button>
                        </div>
                        <div class="editor-tablist" role="tablist" aria-label="排版类别">
                            <button type="button" class="btn" id="editor-tab-document" role="tab" data-editor-tab="document" aria-controls="editor-pane-document">文档</button>
                            <button type="button" class="btn" id="editor-tab-page" role="tab" data-editor-tab="page" aria-controls="editor-pane-page">页面</button>
                        </div>
                        <div class="editor-panel-body">
                            <div class="editor-pane" id="editor-pane-document" role="tabpanel" aria-labelledby="editor-tab-document">
                                <section class="editor-section">
                                    <div class="editor-section-title">内容</div>
                                    <div class="editor-grid">
                                        <label class="editor-field wide">打印内容
                                            <select class="form-select form-select-sm" id="setting-mode"><option value="q">仅试题</option><option value="qa">试题与答案</option><option value="qe">答案移至末尾</option><option value="a">仅答案</option></select>
                                        </label>
                                        <label class="editor-field wide">试卷标题<input class="form-control form-control-sm" id="setting-title" type="text"><span class="zujuanjs-title-tools"><button type="button" class="btn btn-outline-secondary btn-sm" id="title-redetect">恢复智能识别</button><button type="button" class="btn btn-outline-secondary btn-sm" id="title-copy">复制标题</button></span></label>
                                        <label class="editor-field">标题字号<select class="form-select form-select-sm" id="setting-title-size"><option value="18px">18px</option><option value="20px">20px</option><option value="22px">22px · 精排</option><option value="24px">24px</option><option value="26px">26px</option><option value="28px">28px</option><option value="32px">32px</option></select></label>
                                        <label class="editor-field editor-field-unit">题间距<input class="form-control form-control-sm" id="setting-spacing" type="number" min="0" max="32" step="1"><span>px</span></label>
                                    </div>
                                </section>
                                <section class="editor-section">
                                    <div class="editor-section-title">试卷诊断</div>
                                    <div class="editor-section-note">用于发现标题误识别、答案未加载和站点结构变化；不会影响打印内容。</div>
                                    <div class="zujuanjs-diagnostic-card" id="paper-diagnostics">正在检测…</div>
                                </section>
                                <section class="editor-section">
                                    <div class="editor-section-title">正文</div>
                                    <div class="editor-grid">
                                        <label class="editor-field wide">字体<select class="form-select form-select-sm" id="setting-font"><option value='"Times New Roman", SimSun, "Songti SC", serif'>宋体 + 新罗马</option><option value='SimSun, "Songti SC", serif'>宋体</option><option value='"Microsoft YaHei", "PingFang SC", sans-serif'>微软雅黑</option><option value='SimHei, "PingFang SC", sans-serif'>黑体</option><option value='KaiTi, "Songti SC", serif'>楷体</option></select></label>
                                        <label class="editor-field">字号<select class="form-select form-select-sm" id="setting-size"><option value="14px">14px · 精排</option><option value="15px">15px</option><option value="16px">16px</option><option value="17px">17px</option><option value="18px">18px</option><option value="20px">20px</option><option value="22px">22px</option></select></label>
                                        <label class="editor-field">行距<select class="form-select form-select-sm" id="setting-line-height"><option value="1.2">1.2</option><option value="1.35">1.35</option><option value="1.4">1.4 · 精排</option><option value="1.5">1.5</option><option value="1.75">1.75</option><option value="2">2.0</option></select></label>
                                        <label class="editor-field editor-field-unit">段间距<input class="form-control form-control-sm" id="setting-paragraph-spacing" type="number" min="0" max="24" step="1"><span>px</span></label>
                                        <label class="editor-field">对齐<select class="form-select form-select-sm" id="setting-align"><option value="left">左对齐</option><option value="justify">两端对齐</option></select></label>
                                        <label class="editor-field editor-field-unit wide">题号间距<input class="form-control form-control-sm" id="setting-number-gap" type="number" min="0.2" max="2" step="0.05"><span>em</span></label>
                                    </div>
                                </section>
                                <section class="editor-section">
                                    <div class="editor-section-title">答案与解析</div>
                                    <div class="editor-section-note">文字型答案直接自然分页；图片型解析使用“虚拟视窗”显示原图，不重新编码。智能切线只在页底附近寻找真实行间空白；若页底落在文字行内，会把整行移到下一页。切线附近文字重叠保护可独立开关：开启时低置信度切线重复少量内容确保完整，关闭时不重复。</div>
                                    <div class="editor-grid">
                                        <label class="editor-field wide">答案字体（文字）<select class="form-select form-select-sm" id="setting-answer-font"><option value='"Times New Roman", SimSun, "Songti SC", serif'>宋体 + 新罗马</option><option value='SimSun, "Songti SC", serif'>宋体</option><option value='"Microsoft YaHei", "PingFang SC", sans-serif'>微软雅黑</option><option value='SimHei, "PingFang SC", sans-serif'>黑体</option><option value='KaiTi, "Songti SC", serif'>楷体</option><option value='FangSong, "Songti SC", serif'>仿宋</option></select></label>
                                        <label class="editor-field">答案字号（文字）<select class="form-select form-select-sm" id="setting-answer-size"><option value="12px">12px</option><option value="14px">14px</option><option value="16px">16px</option><option value="18px">18px</option><option value="20px">20px</option><option value="22px">22px</option></select></label>
                                        <label class="editor-field">答案行距（文字）<select class="form-select form-select-sm" id="setting-answer-line-height"><option value="1.2">1.2</option><option value="1.35">1.35</option><option value="1.5">1.5</option><option value="1.75">1.75</option><option value="2">2.0</option></select></label>
                                        <label class="editor-field">长答案图<select class="form-select form-select-sm" id="setting-answer-long-image-mode"><option value="split">连续填页 · 推荐</option><option value="fit">整图缩进一页</option></select></label>
                                        <label class="editor-field">图片宽度<select class="form-select form-select-sm" id="setting-answer-image-scale"><option value="100">100% · 推荐</option><option value="90">90%</option><option value="80">80%</option><option value="70">70%</option><option value="60">60%</option><option value="50">50%</option></select></label>
                                        <label class="editor-field editor-field-unit">自动提前<input class="form-control form-control-sm" id="setting-answer-cut-lead" type="number" min="0" max="80" step="1"><span>px</span></label>
                                        <label class="editor-field editor-field-unit">自动重叠<input class="form-control form-control-sm" id="setting-answer-cut-overlap" type="number" min="0" max="80" step="1"><span>px</span></label>
                                        <label class="editor-check wide"><input class="form-check-input" id="setting-answer-cut-overlap-enabled" type="checkbox">切线附近文字重叠保护（防丢字）</label>
                                        <label class="editor-field">智能切线<select class="form-select form-select-sm" id="setting-answer-cut-smart"><option value="true">回归保护行间 · 推荐</option><option value="false">纯填页</option></select></label>
                                        <label class="editor-field editor-field-unit">吸附搜索<input class="form-control form-control-sm" id="setting-answer-cut-search" type="number" min="24" max="180" step="1"><span>px</span></label>
                                        <label class="editor-field editor-field-unit wide">最小续排高度<input class="form-control form-control-sm" id="setting-answer-cut-min-fill" type="number" min="4" max="160" step="1"><span>px</span></label>
                                        <label class="editor-check wide"><input class="form-check-input" id="setting-answer-start-new-page" type="checkbox">答案移至末尾时另起一页</label>
                                    </div>
                                </section>
                            </div>
                            <div class="editor-pane" id="editor-pane-page" role="tabpanel" aria-labelledby="editor-tab-page" hidden>
                                <section class="editor-section">
                                    <div class="editor-section-title">常用版式</div>
                                    <div class="editor-section-note" id="layout-preset-note">选择预设后仍可继续微调，所有改动会自动记住。</div>
                                    <div class="editor-presets" role="group" aria-label="常用版式预设">
                                        <button type="button" class="editor-preset" data-layout-preset="refined"><strong>标准精排</strong><small>更像正式考试卷 · 推荐</small></button>
                                        <button type="button" class="editor-preset" data-layout-preset="exam"><strong>舒展大字</strong><small>保留 8.2.x 视觉</small></button>
                                        <button type="button" class="editor-preset" data-layout-preset="word-normal"><strong>Word 普通</strong><small>上下左右 25.4 mm</small></button>
                                        <button type="button" class="editor-preset" data-layout-preset="word-narrow"><strong>Word 窄</strong><small>上下左右 12.7 mm</small></button>
                                        <button type="button" class="editor-preset" data-layout-preset="word-moderate"><strong>Word 适中</strong><small>上下 25.4，左右 19.1 mm</small></button>
                                        <button type="button" class="editor-preset" data-layout-preset="compact"><strong>紧凑省纸</strong><small>更少留白与间距</small></button>
                                        <button type="button" class="editor-preset" data-layout-preset="word-wide"><strong>Word 宽</strong><small>上下 25.4，左右 50.8 mm</small></button>
                                    </div>
                                </section>
                                <section class="editor-section">
                                    <div class="editor-section-title">页边距</div>
                                    <div class="editor-margin-grid">
                                        <label class="editor-field">上 mm<input class="form-control form-control-sm" id="setting-margin-top" type="number" min="8" max="55" step="0.1"></label>
                                        <label class="editor-field">右 mm<input class="form-control form-control-sm" id="setting-margin-right" type="number" min="8" max="55" step="0.1"></label>
                                        <label class="editor-field">下 mm<input class="form-control form-control-sm" id="setting-margin-bottom" type="number" min="8" max="55" step="0.1"></label>
                                        <label class="editor-field">左 mm<input class="form-control form-control-sm" id="setting-margin-left" type="number" min="8" max="55" step="0.1"></label>
                                    </div>
                                </section>
                                <section class="editor-section">
                                    <div class="editor-section-title">页面显示</div>
                                    <div class="editor-grid">
                                        <label class="editor-field editor-field-unit">答题行高<input class="form-control form-control-sm" id="setting-answer-row-height" type="number" min="1" max="6" step="0.1"><span>em</span></label>
                                        <label class="editor-field editor-field-unit">页面间距<input class="form-control form-control-sm" id="setting-page-gap" type="number" min="8" max="48" step="1"><span>px</span></label>
                                    </div>
                                </section>
                                <section class="editor-section">
                                    <div class="editor-section-title">页码</div>
                                    <div class="editor-grid">
                                        <label class="editor-check wide"><input class="form-check-input" id="setting-page-number" type="checkbox">显示页码</label>
                                        <label class="editor-field wide">页码格式<select class="form-select form-select-sm" id="setting-page-number-format"><option value="current-total">1 / 10</option><option value="current">1</option><option value="chinese">第 1 页，共 10 页</option></select></label>
                                        <label class="editor-field wide">页码计数<select class="form-select form-select-sm" id="setting-page-number-scope"><option value="sectioned">题目 / 答案分段（默认）</option><option value="continuous">整卷连续计数</option></select><small>分段计数时页脚只显示数字，答案自动从 1 重新计数。</small></label>
                                        <label class="editor-field wide">图片打印<select class="form-select form-select-sm" id="setting-print-image-mode"><option value="color">保持彩色</option><option value="grayscale">黑白增强</option></select></label>
                                        <label class="editor-field wide">字体<select class="form-select form-select-sm" id="setting-page-font"><option value='"Times New Roman", SimSun, "Songti SC", serif'>宋体 + 新罗马</option><option value='SimSun, "Songti SC", serif'>宋体</option><option value='"Microsoft YaHei", "PingFang SC", sans-serif'>微软雅黑</option><option value='SimHei, "PingFang SC", sans-serif'>黑体</option></select></label>
                                        <label class="editor-field">字号<select class="form-select form-select-sm" id="setting-page-size"><option value="10px">10px</option><option value="12px">12px</option><option value="14px">14px</option><option value="16px">16px</option></select></label>
                                        <label class="editor-check"><input class="form-check-input" id="setting-page-bold" type="checkbox">加粗</label>
                                    </div>
                                </section>
                            </div>
                        </div>
                        <div class="editor-panel-footer">
                            <div class="editor-footer-actions">
                                <button type="button" class="btn btn-outline-secondary btn-sm" id="editor-export-settings">导出设置</button>
                                <button type="button" class="btn btn-outline-secondary btn-sm" id="editor-import-settings">导入设置</button>
                                <button type="button" class="editor-reset btn btn-outline-secondary btn-sm" id="editor-reset">恢复考试默认</button>
                                <input type="file" id="editor-import-settings-file" accept="application/json,.json" hidden>
                            </div>
                        </div>
                    </aside>
                </div>
                <div class="question-float-toolbar" id="question-float-toolbar" role="toolbar" aria-label="当前题目排版工具" aria-hidden="true" data-active-block-id="">
                    <span class="question-toolbar-label" id="question-toolbar-label">当前题目</span>
                    <div class="btn-group btn-group-sm" role="group" aria-label="答题行数">
                        <button type="button" class="btn btn-outline-secondary" data-question-action="remove-line" title="减少一行答题空间" aria-label="减少一行答题空间">−</button>
                        <span class="question-line-count" id="question-line-count" aria-live="polite">0 行</span>
                        <button type="button" class="btn btn-outline-secondary" data-question-action="add-line" title="增加一行答题空间" aria-label="增加一行答题空间">+</button>
                    </div>
                    <button type="button" class="btn btn-outline-secondary" data-question-action="add-lines" title="增加四行答题空间">+4行</button>
                    <div class="btn-group btn-group-sm" role="group" aria-label="手动分页">
                        <button type="button" class="btn btn-outline-secondary" data-question-action="break-before" aria-label="在本题前分页" aria-pressed="false"><span class="question-toolbar-wide-label">前分页</span><span class="question-toolbar-short-label" aria-hidden="true">前</span></button>
                        <button type="button" class="btn btn-outline-secondary" data-question-action="break-after" aria-label="在本题后分页" aria-pressed="false"><span class="question-toolbar-wide-label">后分页</span><span class="question-toolbar-short-label" aria-hidden="true">后</span></button>
                    </div>
                    <button type="button" class="btn btn-outline-danger" data-question-action="clear" title="清除本题留白和分页"><span class="question-toolbar-wide-label">清除</span><span aria-hidden="true">×</span></button>
                </div>
                <script>
                    const sourceContent = document.getElementById('source-content');
                    const paperContainer = document.getElementById('paper-container');
                    const pageViewport = document.getElementById('page-viewport');
                    const paperWidthPx = 210 * (96 / 25.4);
                    const paperHeightPx = 297 * (96 / 25.4);
                    const zoomSteps = [0.25, 0.33, 0.4, 0.5, 0.65, 0.8, 1, 1.25, 1.5, 2];
                    const overflowTolerance = 0.75;
                    const previewSettings = ${previewSettingsJson};
                    previewSettings.answerLongImageMode = previewSettings.answerLongImageMode === 'fit' ? 'fit' : 'split';
                    previewSettings.answerCutAutoLead = String(Math.max(0, Math.min(80, Number(previewSettings.answerCutAutoLead) || 8)));
                    previewSettings.answerCutOverlap = String(Math.max(0, Math.min(80, Number(previewSettings.answerCutOverlap) || 22)));
                    previewSettings.answerCutOverlapEnabled = previewSettings.answerCutOverlapEnabled !== false && String(previewSettings.answerCutOverlapEnabled) !== 'false';
                    previewSettings.answerCutMinFill = String(Math.max(4, Math.min(160, Number(previewSettings.answerCutMinFill) || 8)));
                    previewSettings.answerCutSmartSnap = previewSettings.answerCutSmartSnap !== false && String(previewSettings.answerCutSmartSnap) !== 'false';
                    previewSettings.answerCutSearchWindow = String(Math.max(24, Math.min(180, Number(previewSettings.answerCutSearchWindow) || 72)));
                    previewSettings.pageNumberFormat = ['current-total', 'current', 'chinese'].includes(previewSettings.pageNumberFormat) ? previewSettings.pageNumberFormat : 'current-total';
                    previewSettings.pageNumberScope = ['continuous', 'sectioned'].includes(previewSettings.pageNumberScope) ? previewSettings.pageNumberScope : 'sectioned';
                    if (previewSettings.pageNumberScope === 'sectioned' && previewSettings.mode === 'qe') previewSettings.answerStartNewPage = true;
                    previewSettings.printImageMode = previewSettings.printImageMode === 'grayscale' ? 'grayscale' : 'color';
                    if (!previewSettings.answerImageCutOffsets || typeof previewSettings.answerImageCutOffsets !== 'object') {
                        previewSettings.answerImageCutOffsets = {};
                    }
                    if (!previewSettings.answerImageCutPositions || typeof previewSettings.answerImageCutPositions !== 'object') previewSettings.answerImageCutPositions = {};
                    if (!previewSettings.paperDiagnostics || typeof previewSettings.paperDiagnostics !== 'object') previewSettings.paperDiagnostics = {};
                    // 8.2.1：答案图只做一次跨域解锁；真正断点在每个分页位置附近按原图 Y 像素局部扫描。
                    const answerImageAnalysisCache = new Map();
                    const answerImageAnalysisQueue = [];
                    let answerImageAnalysisActive = 0;

                    function enqueueAnswerImageAnalysis(task) {
                        return new Promise((resolve, reject) => {
                            answerImageAnalysisQueue.push({ task, resolve, reject });
                            pumpAnswerImageAnalysisQueue();
                        });
                    }

                    function pumpAnswerImageAnalysisQueue() {
                        while (answerImageAnalysisActive < 2 && answerImageAnalysisQueue.length) {
                            const job = answerImageAnalysisQueue.shift();
                            answerImageAnalysisActive++;
                            Promise.resolve()
                                .then(job.task)
                                .then(job.resolve, job.reject)
                                .finally(() => {
                                    answerImageAnalysisActive--;
                                    window.setTimeout(pumpAnswerImageAnalysisQueue, 40);
                                });
                        }
                    }
                    const editorPanel = document.getElementById('editor-panel');
                    const editorToggle = document.getElementById('editor-toggle');
                    const editorResizeHandle = document.getElementById('editor-resize-handle');
                    const questionToolbar = document.getElementById('question-float-toolbar');
                    const editorSaveStatus = document.getElementById('editor-save-status');
                    const layoutPresetNote = document.getElementById('layout-preset-note');
                    const layoutPresets = Object.freeze({
                        refined: {
                            label: '标准精排',
                            settings: {
                                font: '"Times New Roman", SimSun, "Songti SC", serif', size: '14px', lineHeight: '1.4',
                                answerFont: '"Times New Roman", SimSun, "Songti SC", serif', answerSize: '14px', answerLineHeight: '1.35', answerImageScale: '100', answerLongImageMode: 'split', answerCutAutoLead: '6', answerCutOverlap: '20', answerCutMinFill: '8', answerCutSmartSnap: true, answerCutSearchWindow: '72',
                                titleSize: '22px', pageFont: '"Times New Roman", SimSun, "Songti SC", serif', pageSize: '10px',
                                pageBold: false, showPageNumber: true, pageNumberFormat: 'current-total', printImageMode: 'color', pageMargins: '15,17,18,17', questionSpacing: '6',
                                paragraphSpacing: '3', contentAlign: 'left', numberGap: '0.45', answerRowHeight: '1.7', pageGap: '20',
                                previewLayout: 'double', previewZoom: 'auto'
                            }
                        },
                        exam: {
                            label: '舒展大字',
                            settings: {
                                font: '"Times New Roman", SimSun, "Songti SC", serif', size: '16px', lineHeight: '1.5',
                                answerFont: '"Times New Roman", SimSun, "Songti SC", serif', answerSize: '16px', answerLineHeight: '1.5', answerImageScale: '100', answerLongImageMode: 'split', answerCutAutoLead: '6', answerCutOverlap: '20', answerCutMinFill: '8', answerCutSmartSnap: true, answerCutSearchWindow: '72',
                                titleSize: '24px', pageFont: '"Times New Roman", SimSun, "Songti SC", serif', pageSize: '12px',
                                pageBold: true, showPageNumber: true, pageNumberFormat: 'current-total', printImageMode: 'color', pageMargins: '18,15,22,15', questionSpacing: '10',
                                paragraphSpacing: '8', contentAlign: 'left', numberGap: '0.55', answerRowHeight: '1.8', pageGap: '20',
                                previewLayout: 'double', previewZoom: 'auto'
                            }
                        },
                        'word-normal': { label: 'Word 普通', settings: { pageMargins: '25.4,25.4,25.4,25.4' } },
                        'word-narrow': { label: 'Word 窄', settings: { pageMargins: '12.7,12.7,12.7,12.7' } },
                        'word-moderate': { label: 'Word 适中', settings: { pageMargins: '25.4,19.05,25.4,19.05' } },
                        'word-wide': { label: 'Word 宽', settings: { pageMargins: '25.4,50.8,25.4,50.8' } },
                        compact: { label: '紧凑省纸', settings: { pageMargins: '12,13,15,13', size: '14px', titleSize: '20px', questionSpacing: '4', paragraphSpacing: '2', lineHeight: '1.35', numberGap: '0.4', pageSize: '10px', pageBold: false } }
                    });
                    let currentLayout = previewSettings.previewLayout;
                    let currentZoom = previewSettings.previewZoom;
                    let currentScale = 1;
                    let currentPanelWidth = clamp(previewSettings.editorPanelWidth, 280, 520, 340);
                    let currentEditorTab = previewSettings.editorPanelTab === 'page' ? 'page' : 'document';
                    let editorOpen = window.innerWidth > 900 && previewSettings.editorOpen !== false;
                    let activeBlockId = previewSettings.readingAnchor?.blockId || '';
                    let pendingReadingAnchor = previewSettings.readingAnchor || null;
                    const blockEdits = normalizeBlockEdits(previewSettings.documentEdits);
                    let renderFrame = 0;
                    let toolbarPositionFrame = 0;
                    let editorApplyTimer = 0;
                    let settingsSaveTimer = 0;
                    let zoomPreferenceTimer = 0;
                    let isRendering = false;
                    let renderAgain = false;
                    let renderVersion = 0;

                    function setSaveStatus(message) {
                        if (editorSaveStatus) editorSaveStatus.textContent = message;
                    }

                    function updateLayoutPresetUI() {
                        const requested = layoutPresets[previewSettings.layoutPreset] ? previewSettings.layoutPreset : 'custom';
                        const preset = layoutPresets[requested];
                        const matchesPreset = preset && Object.entries(preset.settings).every(([key, value]) => String(previewSettings[key]) === String(value));
                        const activePreset = matchesPreset ? requested : 'custom';
                        document.querySelectorAll('[data-layout-preset]').forEach(button => {
                            const active = button.dataset.layoutPreset === activePreset;
                            button.classList.toggle('is-active', active);
                            button.setAttribute('aria-pressed', String(active));
                        });
                        if (layoutPresetNote) {
                            layoutPresetNote.textContent = activePreset === 'custom'
                                ? '当前为自定义版式，所有改动会自动记住。'
                                : '当前使用“' + layoutPresets[activePreset].label + '”，仍可继续微调并自动保存。';
                        }
                    }

                    function savePreviewPreference(key, value) {
                        window.parent.postMessage({ type: 'saveZujuanPreviewPreference', key: key, value: String(value) }, '*');
                    }

                    function savePrintSettings() {
                        previewSettings.documentEdits = serializeBlockEdits();
                        previewSettings.editorPanelWidth = String(currentPanelWidth);
                        previewSettings.editorPanelTab = currentEditorTab;
                        previewSettings.editorOpen = editorOpen;
                        window.parent.postMessage({ type: 'saveZujuanPrintSettings', settings: previewSettings }, '*');
                        setSaveStatus('已自动保存');
                    }

                    function clamp(value, min, max, fallback) {
                        const number = typeof value === 'string' ? Number.parseFloat(value) : Number(value);
                        return Number.isFinite(number) ? Math.max(min, Math.min(max, number)) : fallback;
                    }

                    function normalizeBlockEdits(input) {
                        const normalized = {};
                        if (!input || typeof input !== 'object') return normalized;
                        Object.entries(input).forEach(([blockId, state]) => {
                            if (!state || typeof state !== 'object') return;
                            const extraLines = Math.round(clamp(state.extraLines, 0, 40, 0));
                            const breakBefore = Boolean(state.breakBefore);
                            const breakAfter = Boolean(state.breakAfter);
                            if (extraLines || breakBefore || breakAfter) {
                                normalized[blockId] = { extraLines, breakBefore, breakAfter };
                            }
                        });
                        return normalized;
                    }

                    function getBlockEdit(blockId) {
                        if (!blockEdits[blockId]) {
                            blockEdits[blockId] = { extraLines: 0, breakBefore: false, breakAfter: false };
                        }
                        return blockEdits[blockId];
                    }

                    function serializeBlockEdits() {
                        const serialized = {};
                        Object.entries(blockEdits).forEach(([blockId, state]) => {
                            const extraLines = Math.round(clamp(state.extraLines, 0, 40, 0));
                            const breakBefore = Boolean(state.breakBefore);
                            const breakAfter = Boolean(state.breakAfter);
                            if (extraLines || breakBefore || breakAfter) {
                                serialized[blockId] = { extraLines, breakBefore, breakAfter };
                            }
                        });
                        return serialized;
                    }

                    function applyBlockEditsToSource() {
                        Array.from(sourceContent.children)
                            .filter(element => element.classList.contains('page-break'))
                            .forEach(element => element.remove());
                        sourceContent.querySelectorAll('.q-wrapper > .answer-blank, .q-wrapper > .answer-blank-large')
                            .forEach(element => element.remove());
                        const wrappers = Array.from(sourceContent.children).filter(element => element.classList.contains('q-wrapper'));

                        wrappers.forEach(wrapper => {
                            const state = getBlockEdit(wrapper.dataset.blockId);
                            wrapper.dataset.extraLines = String(state.extraLines);
                            for (let index = 0; index < state.extraLines; index++) {
                                const blank = document.createElement('div');
                                blank.className = 'answer-blank';
                                blank.dataset.lineIndex = String(index + 1);
                                blank.setAttribute('aria-hidden', 'true');
                                wrapper.appendChild(blank);
                            }
                        });

                        wrappers.forEach(wrapper => {
                            const state = getBlockEdit(wrapper.dataset.blockId);
                            if (state.breakBefore && !wrapper.previousElementSibling?.classList.contains('page-break')) {
                                const marker = document.createElement('div');
                                marker.className = 'page-break';
                                marker.dataset.ownerBlockId = wrapper.dataset.blockId;
                                marker.dataset.breakSide = 'before';
                                wrapper.parentNode.insertBefore(marker, wrapper);
                            }
                            if (state.breakAfter && !wrapper.nextElementSibling?.classList.contains('page-break')) {
                                const marker = document.createElement('div');
                                marker.className = 'page-break';
                                marker.dataset.ownerBlockId = wrapper.dataset.blockId;
                                marker.dataset.breakSide = 'after';
                                wrapper.parentNode.insertBefore(marker, wrapper.nextSibling);
                            }
                        });
                        previewSettings.documentEdits = serializeBlockEdits();
                    }

                    function setSelectValue(id, value) {
                        const select = document.getElementById(id);
                        if (!select) return;
                        const matchingOption = Array.from(select.options).find(option => option.value === String(value));
                        if (matchingOption) {
                            select.value = matchingOption.value;
                            return;
                        }
                        const custom = new Option('自定义', String(value));
                        select.add(custom, 0);
                        select.value = custom.value;
                    }

                    function writeSettingsToEditor() {
                        const margins = String(previewSettings.pageMargins).split(',');
                        setSelectValue('setting-mode', previewSettings.mode);
                        document.getElementById('setting-title').value = previewSettings.title || '';
                        setSelectValue('setting-title-size', previewSettings.titleSize);
                        document.getElementById('setting-spacing').value = previewSettings.questionSpacing;
                        setSelectValue('setting-font', previewSettings.font);
                        setSelectValue('setting-size', previewSettings.size);
                        setSelectValue('setting-line-height', previewSettings.lineHeight);
                        setSelectValue('setting-answer-font', previewSettings.answerFont || previewSettings.font);
                        setSelectValue('setting-answer-size', previewSettings.answerSize || previewSettings.size);
                        setSelectValue('setting-answer-line-height', previewSettings.answerLineHeight || previewSettings.lineHeight);
                        setSelectValue('setting-answer-image-scale', previewSettings.answerImageScale || '100');
                        setSelectValue('setting-answer-long-image-mode', previewSettings.answerLongImageMode === 'fit' ? 'fit' : 'split');
                        document.getElementById('setting-answer-cut-lead').value = previewSettings.answerCutAutoLead || '8';
                        document.getElementById('setting-answer-cut-overlap').value = previewSettings.answerCutOverlap || '22';
                        document.getElementById('setting-answer-cut-overlap-enabled').checked = previewSettings.answerCutOverlapEnabled !== false;
                        document.getElementById('setting-answer-cut-overlap').disabled = previewSettings.answerCutOverlapEnabled === false;
                        document.getElementById('setting-answer-cut-min-fill').value = previewSettings.answerCutMinFill || '8';
                        setSelectValue('setting-answer-cut-smart', previewSettings.answerCutSmartSnap === false ? 'false' : 'true');
                        document.getElementById('setting-answer-cut-search').value = previewSettings.answerCutSearchWindow || '72';
                        document.getElementById('setting-answer-start-new-page').checked = Boolean(previewSettings.answerStartNewPage);
                        document.getElementById('setting-paragraph-spacing').value = previewSettings.paragraphSpacing;
                        setSelectValue('setting-align', previewSettings.contentAlign);
                        document.getElementById('setting-number-gap').value = previewSettings.numberGap;
                        ['top', 'right', 'bottom', 'left'].forEach((side, index) => {
                            document.getElementById('setting-margin-' + side).value = margins[index] || 15;
                        });
                        document.getElementById('setting-answer-row-height').value = previewSettings.answerRowHeight;
                        document.getElementById('setting-page-gap').value = previewSettings.pageGap;
                        document.getElementById('setting-page-number').checked = previewSettings.showPageNumber !== false;
                        setSelectValue('setting-page-number-format', previewSettings.pageNumberFormat || 'current-total');
                        setSelectValue('setting-page-number-scope', previewSettings.pageNumberScope || 'sectioned');
                        setSelectValue('setting-print-image-mode', previewSettings.printImageMode || 'color');
                        setSelectValue('setting-page-font', previewSettings.pageFont);
                        setSelectValue('setting-page-size', previewSettings.pageSize);
                        document.getElementById('setting-page-bold').checked = Boolean(previewSettings.pageBold);
                        updateLayoutPresetUI();
                    }

                    function renderPaperDiagnostics() {
                        const target = document.getElementById('paper-diagnostics');
                        if (!target) return;
                        const d = previewSettings.paperDiagnostics || {};
                        const q = Number(d.questionCount) || 0;
                        const ready = Number(d.answerReadyCount) || 0;
                        const areas = Number(d.answerAreaCount) || 0;
                        const pending = Number(d.pendingAnswerImageCount) || 0;
                        const titleSource = String(previewSettings.titleSource || d.titleSource || '未识别');
                        const confidence = Number(d.titleConfidence) || 0;
                        const status = pending > 0 || (areas > 0 && ready < areas) ? '需要留意' : '正常';
                        target.innerHTML = '<strong>状态：' + status + '</strong>'
                            + '<div class="zujuanjs-diagnostic-grid">'
                            + '<span>题目：' + q + '</span><span>章节：' + (Number(d.sectionCount) || 0) + '</span>'
                            + '<span>答案：' + ready + ' / ' + areas + '</span><span>答案图：' + (Number(d.answerImageCount) || 0) + '</span>'
                            + '<span>待加载图：' + pending + '</span><span>公式：' + (Number(d.formulaCount) || 0) + '</span>'
                            + '</div><div style="margin-top:6px;word-break:break-all">标题来源：' + titleSource.replace(/[<>]/g, '') + ' · 置信度 ' + confidence + '</div>';
                    }

                    function restoreDetectedTitle() {
                        const input = document.getElementById('setting-title');
                        if (!input) return;
                        input.value = String(previewSettings.detectedTitle || previewSettings.paperDiagnostics?.detectedTitle || '').trim();
                        previewSettings.title = input.value;
                        window.parent.postMessage({ type: 'clearZujuanTitleOverride' }, '*');
                        applyDocumentStyles();
                        scheduleRender();
                        scheduleSettingsSave();
                        setSaveStatus('已恢复智能识别');
                    }

                    function readSettingsFromEditor() {
                        const margin = side => clamp(document.getElementById('setting-margin-' + side).value, 8, 55, 15);
                        return {
                            mode: document.getElementById('setting-mode').value,
                            title: document.getElementById('setting-title').value.trim(),
                            titleSize: document.getElementById('setting-title-size').value,
                            questionSpacing: String(clamp(document.getElementById('setting-spacing').value, 0, 32, 10)),
                            font: document.getElementById('setting-font').value,
                            size: document.getElementById('setting-size').value,
                            lineHeight: document.getElementById('setting-line-height').value,
                            answerFont: document.getElementById('setting-answer-font').value,
                            answerSize: document.getElementById('setting-answer-size').value,
                            answerLineHeight: document.getElementById('setting-answer-line-height').value,
                            answerImageScale: String(clamp(document.getElementById('setting-answer-image-scale').value, 50, 100, 100)),
                            answerLongImageMode: document.getElementById('setting-answer-long-image-mode').value === 'fit' ? 'fit' : 'split',
                            answerCutAutoLead: String(clamp(document.getElementById('setting-answer-cut-lead').value, 0, 80, 8)),
                            answerCutOverlap: String(clamp(document.getElementById('setting-answer-cut-overlap').value, 0, 80, 22)),
                            answerCutOverlapEnabled: document.getElementById('setting-answer-cut-overlap-enabled').checked,
                            answerCutMinFill: String(clamp(document.getElementById('setting-answer-cut-min-fill').value, 4, 160, 8)),
                            answerCutSmartSnap: document.getElementById('setting-answer-cut-smart').value !== 'false',
                            answerCutSearchWindow: String(clamp(document.getElementById('setting-answer-cut-search').value, 24, 180, 72)),
                            answerStartNewPage: document.getElementById('setting-answer-start-new-page').checked,
                            answerImageCutOffsets: previewSettings.answerImageCutOffsets,
                            answerImageCutPositions: previewSettings.answerImageCutPositions,
                            paragraphSpacing: String(clamp(document.getElementById('setting-paragraph-spacing').value, 0, 24, 8)),
                            contentAlign: document.getElementById('setting-align').value === 'justify' ? 'justify' : 'left',
                            numberGap: String(clamp(document.getElementById('setting-number-gap').value, 0.2, 2, 0.55)),
                            pageMargins: [margin('top'), margin('right'), margin('bottom'), margin('left')].join(','),
                            layoutPreset: layoutPresets[previewSettings.layoutPreset] ? previewSettings.layoutPreset : 'custom',
                            answerRowHeight: String(clamp(document.getElementById('setting-answer-row-height').value, 1, 6, 1.8)),
                            pageGap: String(clamp(document.getElementById('setting-page-gap').value, 8, 48, 20)),
                            showPageNumber: document.getElementById('setting-page-number').checked,
                            pageNumberFormat: document.getElementById('setting-page-number-format').value,
                            pageNumberScope: document.getElementById('setting-page-number-scope').value === 'sectioned' ? 'sectioned' : 'continuous',
                            printImageMode: document.getElementById('setting-print-image-mode').value === 'grayscale' ? 'grayscale' : 'color',
                            pageFont: document.getElementById('setting-page-font').value,
                            pageSize: document.getElementById('setting-page-size').value,
                            pageBold: document.getElementById('setting-page-bold').checked,
                            previewLayout: currentLayout,
                            previewZoom: currentZoom,
                            editorPanelWidth: String(currentPanelWidth),
                            editorPanelTab: currentEditorTab,
                            editorOpen,
                            documentEdits: serializeBlockEdits(),
                            readingAnchor: captureReadingAnchor()
                        };
                    }

                    function applyLayoutPreset(name) {
                        const preset = layoutPresets[name];
                        if (!preset) return;
                        Object.assign(previewSettings, preset.settings, { layoutPreset: name });
                        currentLayout = previewSettings.previewLayout;
                        currentZoom = previewSettings.previewZoom;
                        writeSettingsToEditor();
                        applyDocumentStyles();
                        applyPreviewView();
                        scheduleRender();
                        setSaveStatus('正在保存…');
                        scheduleSettingsSave();
                    }

                    function applyFormulaScale() {
                        sourceContent.querySelectorAll('img.zujuanjs-formula-svg').forEach(image => {
                            const localFontSize = Number.parseFloat(getComputedStyle(image.parentElement || image).fontSize) || Number.parseFloat(previewSettings.size) || 14;
                            const ratio = clamp(localFontSize / 14, 0.65, 2.2, 1);
                            let baseWidth = Number.parseFloat(image.dataset.formulaBaseWidth);
                            let baseHeight = Number.parseFloat(image.dataset.formulaBaseHeight);
                            if (!Number.isFinite(baseWidth) || baseWidth <= 0) {
                                baseWidth = image.naturalWidth || 0;
                                if (baseWidth > 0) image.dataset.formulaBaseWidth = String(baseWidth);
                            }
                            if (!Number.isFinite(baseHeight) || baseHeight <= 0) {
                                baseHeight = image.naturalHeight || 0;
                                if (baseHeight > 0) image.dataset.formulaBaseHeight = String(baseHeight);
                            }
                            if (baseWidth > 0) {
                                image.style.width = (baseWidth * ratio) + 'px';
                                image.style.height = baseHeight > 0 ? (baseHeight * ratio) + 'px' : 'auto';
                            } else if (baseHeight > 0) {
                                image.style.height = (baseHeight * ratio) + 'px';
                                image.style.width = 'auto';
                            }
                        });
                    }

                    function isRenderedAnswerImage(image) {
                        const src = String(image.currentSrc || image.getAttribute('src') || '');
                        return /getAnswerAndParse/i.test(src);
                    }

                    function applyAnswerImageScale() {
                        const ratio = clamp(Number(previewSettings.answerImageScale || 100) / 100, 0.5, 1, 1);
                        sourceContent.querySelectorAll('.zujuanjs-answer-content').forEach((answerContent, answerOrder) => {
                            const answerId = answerContent.dataset.answerIndex || String(answerOrder + 1);
                            Array.from(answerContent.querySelectorAll('img')).forEach((image, imageOrder) => {
                                if (!isRenderedAnswerImage(image)) return;
                                image.dataset.answerImageKey = answerId + '-' + (imageOrder + 1);
                                image.classList.add('zujuanjs-answer-render-image');

                            const apply = () => {
                                /*
                                 * 答案正文本身是服务器生成的位图，不能真正改字体。
                                 * 图片内部字体不可重排；按独立的“图片答案缩放”比例缩放整张答案图。
                                 */
                                image.style.setProperty('width', (ratio * 100) + '%', 'important');
                                image.style.setProperty('max-width', (ratio * 100) + '%', 'important');
                                image.style.setProperty('height', 'auto', 'important');
                                image.style.setProperty('object-fit', 'contain', 'important');
                                image.style.setProperty('transform-origin', 'left top', 'important');
                            };
                                if (image.complete) apply();
                                else image.addEventListener('load', apply, { once: true });
                            });
                        });
                    }

                    function applyDocumentStyles() {
                        const margins = String(previewSettings.pageMargins).split(',').map(value => clamp(value, 8, 55, 15));
                        const top = margins[0], right = margins[1], bottom = margins[2], left = margins[3];
                        const contentWidth = 210 - left - right;
                        const contentHeight = 297 - top - bottom;
                        const root = document.documentElement;
                        document.body.dataset.answerLongImageMode = previewSettings.answerLongImageMode === 'fit' ? 'fit' : 'split';
                        document.body.dataset.printImageMode = previewSettings.printImageMode === 'grayscale' ? 'grayscale' : 'color';
                        root.style.setProperty('--question-font', previewSettings.font);
                        root.style.setProperty('--question-size', previewSettings.size);
                        root.style.setProperty('--question-line-height', previewSettings.lineHeight);
                        root.style.setProperty('--answer-font', previewSettings.answerFont || previewSettings.font);
                        root.style.setProperty('--answer-size', previewSettings.answerSize || previewSettings.size);
                        root.style.setProperty('--answer-line-height', previewSettings.answerLineHeight || previewSettings.lineHeight);
                        root.style.setProperty('--page-font', previewSettings.pageFont);
                        root.style.setProperty('--page-size', previewSettings.pageSize);
                        root.style.setProperty('--page-weight', previewSettings.pageBold ? 'bold' : 'normal');
                        root.style.setProperty('--page-margin-top', top + 'mm');
                        root.style.setProperty('--page-margin-right', right + 'mm');
                        root.style.setProperty('--page-margin-bottom', bottom + 'mm');
                        root.style.setProperty('--page-margin-left', left + 'mm');
                        root.style.setProperty('--page-content-width', contentWidth + 'mm');
                        root.style.setProperty('--page-content-height', contentHeight + 'mm');
                        root.style.setProperty('--page-image-max-height', Math.max(20, contentHeight - 12) + 'mm');
                        root.style.setProperty('--page-footer-bottom', Math.max(5, Math.min(9, bottom / 3)) + 'mm');
                        root.style.setProperty('--question-spacing', clamp(previewSettings.questionSpacing, 0, 32, 10) + 'px');
                        root.style.setProperty('--title-size', clamp(previewSettings.titleSize, 18, 36, 22) + 'px');
                        root.style.setProperty('--paragraph-spacing', clamp(previewSettings.paragraphSpacing, 0, 24, 8) + 'px');
                        root.style.setProperty('--number-gap', clamp(previewSettings.numberGap, 0.2, 2, 0.55) + 'em');
                        root.style.setProperty('--answer-row-height', clamp(previewSettings.answerRowHeight, 1, 6, 1.8) + 'em');
                        root.style.setProperty('--page-gap', clamp(previewSettings.pageGap, 8, 48, 20) + 'px');
                        root.style.setProperty('--content-align', previewSettings.contentAlign === 'justify' ? 'justify' : 'left');

                        const title = sourceContent.querySelector('[data-document-title]');
                        if (title) {
                            title.textContent = previewSettings.title || '';
                            title.style.display = previewSettings.title ? '' : 'none';
                            title.style.fontFamily = 'SimHei, "Microsoft YaHei", "PingFang SC", sans-serif';
                            title.style.setProperty('font-size', clamp(previewSettings.titleSize, 18, 36, 22) + 'px', 'important');
                        }
                        sourceContent.querySelectorAll('.zujuanjs-question').forEach(element => {
                            element.style.fontFamily = previewSettings.font;
                            element.style.fontSize = previewSettings.size;
                            element.style.lineHeight = previewSettings.lineHeight;
                        });
                        sourceContent.querySelectorAll('.zujuanjs-answer-item, .zujuanjs-answer-content').forEach(element => {
                            element.style.setProperty('font-family', previewSettings.answerFont || previewSettings.font, 'important');
                            element.style.setProperty('font-size', previewSettings.answerSize || previewSettings.size, 'important');
                            element.style.setProperty('line-height', previewSettings.answerLineHeight || previewSettings.lineHeight, 'important');
                        });
                        sourceContent.querySelectorAll('.zujuanjs-answer-item *, .zujuanjs-answer-content *').forEach(element => {
                            if (element.matches('img, svg, svg *, math, math *, mjx-container, mjx-container *')) return;
                            element.style.setProperty('font-family', 'inherit', 'important');
                            element.style.setProperty('font-size', 'inherit', 'important');
                            element.style.setProperty('line-height', 'inherit', 'important');
                        });
                        sourceContent.querySelectorAll('.zujuanjs-section-title').forEach(element => {
                            element.style.fontFamily = previewSettings.font;
                        });
                        sourceContent.querySelectorAll('.zujuanjs-answer-title').forEach(element => {
                            element.style.setProperty('font-family', previewSettings.answerFont || previewSettings.font, 'important');
                        });
                        applyFormulaScale();
                        applyAnswerImageScale();
                        primeAnswerImageAnalyses();
                    }

                    function getSafeExportName(extension) {
                        const raw = String(previewSettings.title || '试卷').trim() || '试卷';
                        const clean = raw.replace(/[^A-Za-z0-9\u3400-\u9fff._ -]+/g, '_').replace(/\s+/g, ' ').trim().slice(0, 90) || '试卷';
                        return clean + extension;
                    }

                    function downloadExportBlob(blob, filename) {
                        const url = URL.createObjectURL(blob);
                        const link = document.createElement('a');
                        link.href = url;
                        link.download = filename;
                        link.style.display = 'none';
                        document.body.appendChild(link);
                        link.click();
                        link.remove();
                        window.setTimeout(() => URL.revokeObjectURL(url), 1500);
                    }

                    function blobToDataUrl(blob) {
                        return new Promise((resolve, reject) => {
                            const reader = new FileReader();
                            reader.onload = () => resolve(String(reader.result || ''));
                            reader.onerror = () => reject(reader.error || new Error('图片读取失败'));
                            reader.readAsDataURL(blob);
                        });
                    }

                    function requestResourceFromParent(url, timeout = 22000) {
                        return new Promise((resolve, reject) => {
                            const requestId = 'res-' + Date.now() + '-' + Math.random().toString(36).slice(2);
                            const timer = setTimeout(() => {
                                window.removeEventListener('message', onMessage);
                                reject(new Error('跨域资源回退超时'));
                            }, timeout);
                            function onMessage(event) {
                                if (!event.data || event.data.type !== 'zujuanFetchResourceResult' || event.data.requestId !== requestId) return;
                                clearTimeout(timer);
                                window.removeEventListener('message', onMessage);
                                if (event.data.ok && event.data.dataUrl) resolve(event.data.dataUrl);
                                else reject(new Error(event.data.message || '跨域资源回退失败'));
                            }
                            window.addEventListener('message', onMessage);
                            window.parent.postMessage({ type: 'zujuanFetchResource', requestId, url }, '*');
                        });
                    }

                    async function inlineImagesForExport(root) {
                        let unresolved = 0;
                        const images = Array.from(root.querySelectorAll('img'));
                        for (const image of images) {
                            const source = image.getAttribute('src') || '';
                            if (!source || source.startsWith('data:') || source.startsWith('blob:')) continue;
                            const absoluteUrl = new URL(source, document.baseURI).href;
                            try {
                                const response = await fetch(absoluteUrl, { credentials: 'include', cache: 'force-cache' });
                                if (!response.ok) throw new Error('HTTP ' + response.status);
                                const blob = await response.blob();
                                image.setAttribute('src', await blobToDataUrl(blob));
                            } catch (firstError) {
                                try {
                                    image.setAttribute('src', await requestResourceFromParent(absoluteUrl));
                                } catch (fallbackError) {
                                    unresolved++;
                                    console.warn('Word/HTML 导出时图片未能内嵌：', source, firstError, fallbackError);
                                }
                            }
                        }
                        return unresolved;
                    }

                    function buildExportHtml(contentHtml) {
                        const margins = String(previewSettings.pageMargins || '18,15,22,15').split(',').map(value => clamp(value, 8, 55, 15));
                        const top = margins[0], right = margins[1], bottom = margins[2], left = margins[3];
                        const questionSpacing = clamp(previewSettings.questionSpacing, 0, 32, 10);
                        const paragraphSpacing = clamp(previewSettings.paragraphSpacing, 0, 24, 8);
                        const numberGap = clamp(previewSettings.numberGap, 0.2, 2, 0.55);
                        const rowHeight = clamp(previewSettings.answerRowHeight, 1, 6, 1.8);
                        const titleSize = clamp(previewSettings.titleSize, 18, 36, 22);
                        const align = previewSettings.contentAlign === 'justify' ? 'justify' : 'left';
                        const answerFont = previewSettings.answerFont || previewSettings.font;
                        const answerSize = previewSettings.answerSize || previewSettings.size;
                        const answerLineHeight = previewSettings.answerLineHeight || previewSettings.lineHeight;
                        const css = [
                            '@page { size: A4; margin: ' + top + 'mm ' + right + 'mm ' + bottom + 'mm ' + left + 'mm; }',
                            'html,body{background:#fff;color:#000;margin:0;padding:0;}',
                            'body{font-family:' + previewSettings.font + ';font-size:' + previewSettings.size + ';line-height:' + previewSettings.lineHeight + ';text-align:' + align + ';}',
                            '.zujuanjs-print-title{font-family:SimHei,Microsoft YaHei,PingFang SC,sans-serif;text-align:center;font-weight:700;font-size:' + titleSize + 'px;line-height:1.28;margin:0 0 18px;letter-spacing:.025em;}',
                            '.zujuanjs-section-title{font-family:SimHei,Microsoft YaHei,PingFang SC,sans-serif;font-weight:700;font-size:1.10em;line-height:1.35;margin:14px 0 7px;border:0;padding:0;break-after:avoid-page;page-break-after:avoid;}',
                            '.zujuanjs-answer-page-break{' + (previewSettings.answerStartNewPage ? 'display:block;height:0;break-before:page;page-break-before:always;' : 'display:none;') + '}',
                            '.q-wrapper{margin:0 0 ' + questionSpacing + 'px;page-break-inside:auto;}',
                            '.zujuanjs-question,.zujuanjs-answer-item{margin:0;padding:0;}',
                            '.zujuanjs-answer-item,.zujuanjs-answer-content{font-family:' + answerFont + ' !important;font-size:' + answerSize + ' !important;line-height:' + answerLineHeight + ' !important;}',
                            '.zujuanjs-answer-item *:not(img):not(svg),.zujuanjs-answer-content *:not(img):not(svg){font-family:inherit !important;font-size:inherit !important;line-height:inherit !important;}',
                            '.zujuanjs-answer-content,.zujuanjs-answer-flow-html,.zujuanjs-answer-image-flow{margin:0!important;padding:0!important;min-height:0!important;height:auto!important;border:0!important}.zujuanjs-answer-image-slice{position:relative;overflow:hidden;break-inside:avoid;page-break-inside:avoid}.zujuanjs-answer-image-slice>img{position:absolute;left:0;max-width:none!important;max-height:none!important}.zujuanjs-answer-cut-line,.zujuanjs-answer-cut-control{display:none!important}',
                            '.zujuanjs-answer-render-image{display:block !important;height:auto !important;object-fit:contain !important;transform-origin:left top !important;}',
                            '.zujuanjs-question-layout{display:table;width:100%;table-layout:fixed;}',
                            '.zujuanjs-question-number{display:none !important;}', '.zujuanjs-question-layout.zujuanjs-original-numbering{display:block!important;}',
                            '.zujuanjs-question-body{display:table-cell;vertical-align:top;width:auto;}', '.zujuanjs-question-layout.zujuanjs-original-numbering .zujuanjs-question-body{display:block!important;width:auto!important;}',
                            '.zujuanjs-answer-title{font-weight:700;margin:0 0 6px;}',
                            '.zujuanjs-question-body p,.zujuanjs-answer-item p{margin:0 0 ' + paragraphSpacing + 'px;}',
                            '.zujuanjs-question-body p:last-child,.zujuanjs-answer-item p:last-child{margin-bottom:0;}',
                            '.answer-blank,.answer-blank-large{display:block;width:100%;height:' + rowHeight + 'em;margin:0;border:0;}',
                            '.answer-blank-large{height:' + (rowHeight * 5) + 'em;}',
                            '.page-break{display:block;height:0;page-break-before:always;break-before:page;}',
                            'img{max-width:100%;height:auto;}',
                            'table{max-width:100%;border-collapse:collapse;}',
                            '.knowledge-box,.q-toolbar,.manual-break-indicator{display:none!important;}'
                        ].join('');
                        const titleNode = document.createElement('div');
                        titleNode.textContent = previewSettings.title || '试卷';
                        return '<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><title>' + titleNode.innerHTML + '</title><style>' + css + '</style></head><body><div class="Section1">' + contentHtml + '</div></body></html>';
                    }

                    async function prepareExportDocument() {
                        clearTimeout(editorApplyTimer);
                        Object.assign(previewSettings, readSettingsFromEditor());
                        applyBlockEditsToSource();
                        applyDocumentStyles();
                        savePrintSettings();
                        const clone = sourceContent.cloneNode(true);
                        clone.removeAttribute('id');
                        clone.removeAttribute('style');
                        clone.querySelectorAll('.q-toolbar,.manual-break-indicator').forEach(element => element.remove());
                        const unresolvedImages = await inlineImagesForExport(clone);
                        return { html: buildExportHtml(clone.innerHTML), unresolvedImages: unresolvedImages };
                    }

                    function getPortableSettings() {
                        const settings = readSettingsFromEditor();
                        settings.answerImageCutOffsets = previewSettings.answerImageCutOffsets || {};
                        settings.answerImageCutPositions = previewSettings.answerImageCutPositions || {};
                        settings.documentEdits = serializeBlockEdits();
                        return {
                            app: '组卷网试卷打印工作台',
                            version: '8.4.3',
                            exportedAt: new Date().toISOString(),
                            settings: settings
                        };
                    }

                    function exportPortableSettings() {
                        const data = JSON.stringify(getPortableSettings(), null, 2);
                        downloadExportBlob(new Blob([data], { type: 'application/json;charset=utf-8' }), getSafeExportName('-打印设置.json'));
                    }

                    function importPortableSettings(file) {
                        if (!file) return;
                        const reader = new FileReader();
                        reader.onload = () => {
                            try {
                                const parsed = JSON.parse(String(reader.result || '{}'));
                                const incoming = parsed && parsed.settings && typeof parsed.settings === 'object' ? parsed.settings : parsed;
                                if (!incoming || typeof incoming !== 'object') throw new Error('设置文件格式不正确');
                                const allowed = [
                                    'mode','font','size','lineHeight','answerFont','answerSize','answerLineHeight','answerImageScale','answerLongImageMode',
                                    'answerCutAutoLead','answerCutOverlap','answerCutOverlapEnabled','answerCutMinFill','answerCutSmartSnap','answerCutSearchWindow','answerStartNewPage','answerImageCutOffsets','answerImageCutPositions','title','titleSize','pageFont','pageSize','pageBold',
                                    'showPageNumber','pageNumberFormat','pageNumberScope','printImageMode','pageMargins','layoutPreset','questionSpacing','previewLayout','previewZoom',
                                    'paragraphSpacing','contentAlign','numberGap','answerRowHeight','pageGap','editorPanelWidth','editorPanelTab','editorOpen','documentEdits'
                                ];
                                const next = {};
                                allowed.forEach(key => {
                                    if (Object.prototype.hasOwnProperty.call(incoming, key)) next[key] = incoming[key];
                                });
                                window.parent.postMessage({ type: 'rebuildZujuanPreview', settings: next }, '*');
                            } catch (error) {
                                alert('导入设置失败：' + (error && error.message ? error.message : error));
                            }
                        };
                        reader.onerror = () => alert('读取设置文件失败');
                        reader.readAsText(file, 'utf-8');
                    }

                    function setExportButtonBusy(button, busy, label) {
                        if (!button) return;
                        if (!button.dataset.originalText) button.dataset.originalText = button.textContent;
                        button.disabled = Boolean(busy);
                        button.textContent = busy ? label : button.dataset.originalText;
                    }

                    function buildStructuredExportData() {
                        const sections = [];
                        const questionMap = new Map();
                        let currentSection = '';
                        Array.from(sourceContent.children).forEach(node => {
                            if (node.classList.contains('zujuanjs-section-title')) {
                                const text = String(node.textContent || '').trim();
                                if (text && text !== '答案与解析') {
                                    currentSection = text;
                                    sections.push(text);
                                }
                                return;
                            }
                            if (node.classList.contains('q-wrapper')) {
                                const questionBody = node.querySelector('.zujuanjs-question-body');
                                const answer = node.querySelector('.zujuanjs-answer-content');
                                const index = Number(String(node.dataset.blockId || '').replace(/\D+/g, '')) || questionMap.size + 1;
                                questionMap.set(index, {
                                    index,
                                    qid: node.dataset.qid || '',
                                    blockId: node.dataset.blockId || '',
                                    section: currentSection,
                                    questionText: String(questionBody?.innerText || questionBody?.textContent || '').trim(),
                                    questionHtml: questionBody?.innerHTML || '',
                                    answerText: String(answer?.innerText || answer?.textContent || '').trim(),
                                    answerHtml: answer?.innerHTML || '',
                                    imageUrls: Array.from(node.querySelectorAll('img')).map(img => img.getAttribute('src') || '').filter(Boolean)
                                });
                                return;
                            }
                            if (node.classList.contains('zujuanjs-answer-content')) {
                                const index = Number(node.dataset.answerIndex) || 0;
                                if (!index) return;
                                const existing = questionMap.get(index) || {
                                    index, qid: '', blockId: '', section: '', questionText: '', questionHtml: '', answerText: '', answerHtml: '', imageUrls: []
                                };
                                existing.answerText = String(node.innerText || node.textContent || '').trim();
                                existing.answerHtml = node.innerHTML || '';
                                existing.imageUrls = Array.from(new Set(existing.imageUrls.concat(Array.from(node.querySelectorAll('img')).map(img => img.getAttribute('src') || '').filter(Boolean))));
                                questionMap.set(index, existing);
                            }
                        });
                        return {
                            app: '组卷网试卷打印工作台',
                            version: '8.4.3',
                            exportedAt: new Date().toISOString(),
                            title: previewSettings.title || '',
                            detectedTitle: previewSettings.detectedTitle || '',
                            diagnostics: previewSettings.paperDiagnostics || {},
                            sections: Array.from(new Set(sections)),
                            questions: Array.from(questionMap.values()).sort((a, b) => a.index - b.index)
                        };
                    }

                    function exportStructuredJson() {
                        const data = JSON.stringify(buildStructuredExportData(), null, 2);
                        downloadExportBlob(new Blob([data], { type: 'application/json;charset=utf-8' }), getSafeExportName('-结构化.json'));
                    }

                    async function exportStandaloneHtml() {
                        const button = document.getElementById('export-html');
                        setExportButtonBusy(button, true, '处理中');
                        try {
                            const prepared = await prepareExportDocument();
                            const blob = new Blob(['\ufeff', prepared.html], { type: 'text/html;charset=utf-8' });
                            downloadExportBlob(blob, getSafeExportName('.html'));
                            if (prepared.unresolvedImages) console.warn('独立 HTML 中有 ' + prepared.unresolvedImages + ' 张图片保留为网络地址。');
                        } catch (error) {
                            console.error('HTML 导出失败：', error);
                            alert('HTML 导出失败：' + (error && error.message ? error.message : error));
                        } finally {
                            setExportButtonBusy(button, false, '处理中');
                        }
                    }

                    async function exportWordDocument() {
                        const button = document.getElementById('export-word');
                        setExportButtonBusy(button, true, '处理中');
                        try {
                            const prepared = await prepareExportDocument();
                            window.parent.postMessage({
                                type: 'exportZujuanWord',
                                html: prepared.html,
                                fileName: getSafeExportName('.docx'),
                                margins: String(previewSettings.pageMargins || '18,15,22,15').split(',').map(Number),
                                unresolvedImages: prepared.unresolvedImages
                            }, '*');
                        } catch (error) {
                            setExportButtonBusy(button, false, '处理中');
                            console.error('Word 导出准备失败：', error);
                            alert('Word 导出准备失败：' + (error && error.message ? error.message : error));
                        }
                    }

                    function updateQuickAnswerEndButton() {
                        const button = document.getElementById('quick-answer-end');
                        if (!button) return;
                        const active = previewSettings.mode === 'qe';
                        button.classList.toggle('active', active);
                        button.setAttribute('aria-pressed', String(active));
                        button.title = active ? '已启用：答案移至末尾（点击恢复仅试题）' : '切换：答案移至末尾';
                    }

                    function setEditorTab(tab, persist = true) {
                        currentEditorTab = tab === 'page' ? 'page' : 'document';
                        previewSettings.editorPanelTab = currentEditorTab;
                        document.querySelectorAll('[data-editor-tab]').forEach(button => {
                            const active = button.dataset.editorTab === currentEditorTab;
                            button.classList.toggle('active', active);
                            button.setAttribute('aria-selected', String(active));
                            button.tabIndex = active ? 0 : -1;
                        });
                        document.querySelectorAll('.editor-pane').forEach(pane => {
                            pane.hidden = pane.id !== 'editor-pane-' + currentEditorTab;
                        });
                        if (persist) savePreviewPreference('editorPanelTab', currentEditorTab);
                    }

                    function applyEditorPanelWidth(width, persist = false) {
                        currentPanelWidth = Math.round(clamp(width, 280, 520, 340));
                        previewSettings.editorPanelWidth = String(currentPanelWidth);
                        document.documentElement.style.setProperty('--editor-panel-width', currentPanelWidth + 'px');
                        editorResizeHandle.setAttribute('aria-valuenow', String(currentPanelWidth));
                        if (persist) savePreviewPreference('editorPanelWidth', currentPanelWidth);
                        applyPreviewView();
                    }

                    function toggleEditor(open, persist = true) {
                        editorOpen = typeof open === 'boolean' ? open : !editorOpen;
                        previewSettings.editorOpen = editorOpen;
                        if (!editorOpen && editorPanel.contains(document.activeElement)) editorToggle.focus();
                        if (editorOpen && window.innerWidth <= 680) hideQuestionToolbar();
                        document.body.classList.toggle('editor-open', editorOpen);
                        editorToggle.classList.toggle('active', editorOpen);
                        editorToggle.setAttribute('aria-pressed', String(editorOpen));
                        editorToggle.setAttribute('aria-expanded', String(editorOpen));
                        editorPanel.setAttribute('aria-hidden', String(!editorOpen));
                        editorPanel.inert = !editorOpen;
                        if (persist) savePreviewPreference('editorOpen', editorOpen);
                        requestAnimationFrame(() => {
                            applyPreviewView();
                            scheduleQuestionToolbarPosition();
                        });
                    }

                    function scheduleSettingsSave() {
                        clearTimeout(settingsSaveTimer);
                        setSaveStatus('正在保存…');
                        settingsSaveTimer = window.setTimeout(savePrintSettings, 180);
                    }

                    function applyEditorSettings() {
                        Object.assign(previewSettings, readSettingsFromEditor());
                        currentLayout = previewSettings.previewLayout;
                        currentZoom = previewSettings.previewZoom;
                        applyDocumentStyles();
                        applyPreviewView();
                        scheduleRender();
                        scheduleSettingsSave();
                    }

                    function queueEditorSettings() {
                        clearTimeout(editorApplyTimer);
                        editorApplyTimer = window.setTimeout(applyEditorSettings, 60);
                    }

                    function getAutoScale() {
                        const pageCount = currentLayout === 'double' && window.innerWidth > 680 ? 2 : 1;
                        const gap = window.innerWidth <= 680 ? 12 : clamp(previewSettings.pageGap, 8, 48, 20);
                        const padding = window.innerWidth <= 680 ? 24 : 48;
                        const available = Math.max(220, pageViewport.clientWidth - padding - gap * (pageCount - 1));
                        return Math.max(0.32, Math.min(1, available / (paperWidthPx * pageCount)));
                    }

                    function applyPreviewView(zoomFocus = null) {
                        const readingAnchor = zoomFocus ? null : captureReadingAnchor();
                        currentScale = currentZoom === 'auto'
                            ? getAutoScale()
                            : Math.max(0.25, Math.min(2, Number(currentZoom) || 1));
                        const root = document.documentElement;
                        root.style.setProperty('--preview-scale', String(currentScale));
                        root.style.setProperty('--paper-display-width', (paperWidthPx * currentScale) + 'px');
                        root.style.setProperty('--paper-display-height', (paperHeightPx * currentScale) + 'px');
                        paperContainer.classList.toggle('layout-single', currentLayout === 'single');
                        paperContainer.classList.toggle('layout-double', currentLayout === 'double');

                        document.querySelectorAll('[data-layout]').forEach(button => {
                            button.classList.toggle('active', button.dataset.layout === currentLayout);
                            button.setAttribute('aria-pressed', String(button.dataset.layout === currentLayout));
                        });
                        const zoomValue = document.getElementById('zoom-value');
                        zoomValue.textContent = (currentZoom === 'auto' ? '自动 ' : '') + Math.round(currentScale * 100) + '%';
                        requestAnimationFrame(() => {
                            if (zoomFocus) restoreZoomFocus(zoomFocus);
                            else restoreReadingAnchor(readingAnchor);
                            scheduleQuestionToolbarPosition();
                        });
                    }

                    function setPreviewLayout(layout) {
                        if (!['single', 'double'].includes(layout)) return;
                        currentLayout = layout;
                        previewSettings.previewLayout = layout;
                        applyPreviewView();
                        savePreviewPreference('previewLayout', layout);
                        savePrintSettings();
                    }

                    function setPreviewZoom(zoom, focusPoint = null) {
                        const zoomFocus = focusPoint && zoom !== 'auto'
                            ? captureZoomFocus(focusPoint.clientX, focusPoint.clientY)
                            : null;
                        currentZoom = zoom === 'auto' ? 'auto' : String(Math.max(0.25, Math.min(2, Number(zoom) || 1)));
                        previewSettings.previewZoom = currentZoom;
                        applyPreviewView(zoomFocus);
                        clearTimeout(zoomPreferenceTimer);
                        zoomPreferenceTimer = window.setTimeout(() => {
                            savePreviewPreference('previewZoom', currentZoom);
                            savePrintSettings();
                        }, 180);
                    }

                    function stepPreviewZoom(direction) {
                        const base = currentScale;
                        const candidates = direction > 0
                            ? zoomSteps.filter(value => value > base + 0.01)
                            : zoomSteps.filter(value => value < base - 0.01).reverse();
                        setPreviewZoom(candidates[0] || (direction > 0 ? zoomSteps[zoomSteps.length - 1] : zoomSteps[0]));
                    }

                    function getRenderedFragments(blockId) {
                        if (!blockId) return [];
                        return Array.from(paperContainer.querySelectorAll('.q-wrapper'))
                            .filter(fragment => fragment.dataset.blockId === blockId);
                    }

                    function captureZoomFocus(clientX, clientY) {
                        const hit = document.elementFromPoint(clientX, clientY);
                        const paper = hit?.closest?.('.paper');
                        if (!paper) return null;
                        const papers = Array.from(paperContainer.querySelectorAll('.paper'));
                        const paperIndex = papers.indexOf(paper);
                        const rect = paper.getBoundingClientRect();
                        if (paperIndex < 0 || !rect.width || !rect.height) return null;
                        return {
                            paperIndex,
                            xRatio: Math.max(0, Math.min(1, (clientX - rect.left) / rect.width)),
                            yRatio: Math.max(0, Math.min(1, (clientY - rect.top) / rect.height)),
                            clientX,
                            clientY
                        };
                    }

                    function restoreZoomFocus(focus) {
                        if (!focus) return;
                        const paper = paperContainer.querySelectorAll('.paper')[focus.paperIndex];
                        if (!paper) return;
                        const rect = paper.getBoundingClientRect();
                        const focusedX = rect.left + rect.width * focus.xRatio;
                        const focusedY = rect.top + rect.height * focus.yRatio;
                        pageViewport.scrollLeft += focusedX - focus.clientX;
                        pageViewport.scrollTop += focusedY - focus.clientY;
                    }

                    function zoomAtPoint(deltaY, clientX, clientY) {
                        const rawFactor = Math.exp(-deltaY * 0.002);
                        const factor = Math.max(0.85, Math.min(1.15, rawFactor));
                        const nextScale = Math.max(0.25, Math.min(2, currentScale * factor));
                        if (Math.abs(nextScale - currentScale) < 0.001) return;
                        setPreviewZoom(String(Math.round(nextScale * 1000) / 1000), { clientX, clientY });
                    }

                    function captureReadingAnchor() {
                        const viewportRect = pageViewport.getBoundingClientRect();
                        if (!viewportRect.height) return null;

                        if (activeBlockId) {
                            const fragments = getRenderedFragments(activeBlockId);
                            const visible = fragments
                                .map((fragment, index) => ({ fragment, index, rect: fragment.getBoundingClientRect() }))
                                .filter(item => item.rect.bottom > viewportRect.top + 2 && item.rect.top < viewportRect.bottom - 2)
                                .sort((a, b) => Math.abs(a.rect.top - viewportRect.top - 24) - Math.abs(b.rect.top - viewportRect.top - 24));
                            if (visible[0]) {
                                return {
                                    blockId: activeBlockId,
                                    fragmentIndex: visible[0].index,
                                    offset: visible[0].rect.top - viewportRect.top
                                };
                            }
                        }

                        const shells = Array.from(paperContainer.querySelectorAll('.paper-shell'));
                        const pageIndex = shells.findIndex(shell => shell.getBoundingClientRect().bottom > viewportRect.top + 8);
                        if (pageIndex < 0) return null;
                        return {
                            pageIndex,
                            offset: shells[pageIndex].getBoundingClientRect().top - viewportRect.top
                        };
                    }

                    function restoreReadingAnchor(anchor) {
                        if (!anchor) return;
                        const viewportRect = pageViewport.getBoundingClientRect();
                        let target = null;
                        if (anchor.blockId) {
                            const fragments = getRenderedFragments(anchor.blockId);
                            target = fragments[Math.min(Math.max(0, Number(anchor.fragmentIndex) || 0), Math.max(0, fragments.length - 1))] || null;
                        } else if (Number.isInteger(anchor.pageIndex)) {
                            target = paperContainer.querySelectorAll('.paper-shell')[anchor.pageIndex] || null;
                        }
                        if (!target) return;
                        const delta = target.getBoundingClientRect().top - viewportRect.top - (Number(anchor.offset) || 0);
                        if (Math.abs(delta) > 0.5) pageViewport.scrollTop += delta;
                    }

                    function setQuestionToolbarVisible(visible) {
                        questionToolbar.classList.toggle('is-visible', visible);
                        questionToolbar.setAttribute('aria-hidden', String(!visible));
                    }

                    function hideQuestionToolbar(clearSelection = true) {
                        setQuestionToolbarVisible(false);
                        if (clearSelection) {
                            activeBlockId = '';
                            questionToolbar.dataset.activeBlockId = '';
                            document.body.classList.remove('question-tools-open');
                            paperContainer.querySelectorAll('.q-wrapper.is-selected').forEach(fragment => fragment.classList.remove('is-selected'));
                        }
                    }

                    function updateQuestionToolbarState() {
                        if (!activeBlockId) return;
                        questionToolbar.dataset.activeBlockId = activeBlockId;
                        const sourceWrapper = Array.from(sourceContent.querySelectorAll('.q-wrapper'))
                            .find(wrapper => wrapper.dataset.blockId === activeBlockId);
                        if (!sourceWrapper) {
                            hideQuestionToolbar();
                            return;
                        }
                        const state = getBlockEdit(activeBlockId);
                        document.getElementById('question-toolbar-label').textContent = sourceWrapper.dataset.blockLabel || '当前题目';
                        document.getElementById('question-line-count').textContent = state.extraLines + ' 行';
                        const removeButton = questionToolbar.querySelector('[data-question-action="remove-line"]');
                        const clearButton = questionToolbar.querySelector('[data-question-action="clear"]');
                        removeButton.disabled = state.extraLines <= 0;
                        clearButton.disabled = state.extraLines <= 0 && !state.breakBefore && !state.breakAfter;
                        ['before', 'after'].forEach(side => {
                            const button = questionToolbar.querySelector('[data-question-action="break-' + side + '"]');
                            button.setAttribute('aria-pressed', String(Boolean(state['break' + side[0].toUpperCase() + side.slice(1)])));
                        });
                        const wrappers = Array.from(sourceContent.querySelectorAll('.q-wrapper'));
                        questionToolbar.querySelector('[data-question-action="break-before"]').disabled = wrappers[0] === sourceWrapper;
                    }

                    function scheduleQuestionToolbarPosition() {
                        if (toolbarPositionFrame) cancelAnimationFrame(toolbarPositionFrame);
                        toolbarPositionFrame = requestAnimationFrame(() => {
                            toolbarPositionFrame = 0;
                            positionQuestionToolbar();
                        });
                    }

                    function positionQuestionToolbar() {
                        if (!activeBlockId || (editorOpen && window.innerWidth <= 680)) {
                            setQuestionToolbarVisible(false);
                            return;
                        }
                        const viewportRect = pageViewport.getBoundingClientRect();
                        setQuestionToolbarVisible(true);
                        const toolbarRect = questionToolbar.getBoundingClientRect();
                        const margin = 8;
                        const left = Math.max(
                            viewportRect.left + margin,
                            Math.min(
                                viewportRect.left + (viewportRect.width - toolbarRect.width) / 2,
                                viewportRect.right - toolbarRect.width - margin
                            )
                        );
                        questionToolbar.style.top = '64px';
                        questionToolbar.style.left = Math.round(left) + 'px';
                    }

                    function activateQuestion(wrapper) {
                        const blockId = wrapper?.dataset.blockId;
                        if (!blockId) return;
                        activeBlockId = blockId;
                        document.body.classList.add('question-tools-open');
                        questionToolbar.dataset.activeBlockId = blockId;
                        paperContainer.querySelectorAll('.q-wrapper').forEach(fragment => {
                            fragment.classList.toggle('is-selected', fragment.dataset.blockId === blockId);
                        });
                        updateQuestionToolbarState();
                        scheduleQuestionToolbarPosition();
                    }

                    function bindRenderedQuestionInteractions() {
                        paperContainer.querySelectorAll('.q-wrapper').forEach(wrapper => {
                            wrapper.addEventListener('pointerenter', () => activateQuestion(wrapper));
                            wrapper.addEventListener('focusin', () => activateQuestion(wrapper));
                            wrapper.addEventListener('click', event => {
                                event.stopPropagation();
                                activateQuestion(wrapper);
                            });
                            wrapper.addEventListener('keydown', event => {
                                if (!['Enter', ' '].includes(event.key)) return;
                                event.preventDefault();
                                activateQuestion(wrapper);
                            });
                        });
                    }

                    function performQuestionAction(action) {
                        if (!activeBlockId) return;
                        const state = getBlockEdit(activeBlockId);
                        if (action === 'remove-line') state.extraLines = Math.max(0, state.extraLines - 1);
                        else if (action === 'add-line') state.extraLines = Math.min(40, state.extraLines + 1);
                        else if (action === 'add-lines') state.extraLines = Math.min(40, state.extraLines + 4);
                        else if (action === 'break-before') state.breakBefore = !state.breakBefore;
                        else if (action === 'break-after') state.breakAfter = !state.breakAfter;
                        else if (action === 'clear') {
                            state.extraLines = 0;
                            state.breakBefore = false;
                            state.breakAfter = false;
                        } else return;

                        if (!state.extraLines && !state.breakBefore && !state.breakAfter) delete blockEdits[activeBlockId];
                        applyBlockEditsToSource();
                        updateQuestionToolbarState();
                        scheduleRender();
                        scheduleSettingsSave();
                    }

                    function createPaper() {
                        const shell = document.createElement('div');
                        shell.className = 'paper-shell';
                        const paper = document.createElement('section');
                        paper.className = 'paper';
                        paper.innerHTML = '<div class="paper-content"></div><div class="page-footer"></div>';
                        shell.appendChild(paper);
                        paperContainer.appendChild(shell);
                        return { paper: paper, content: paper.querySelector('.paper-content') };
                    }

                    function pageHasContent(content) {
                        return Array.from(content.children).some(el => !el.classList.contains('manual-break-indicator'));
                    }

                    function hasPrintableContent(node) {
                        const copy = node.cloneNode(true);
                        copy.querySelectorAll('.q-toolbar').forEach(el => el.remove());
                        if (copy.textContent.replace(/\s/g, '')) return true;
                        return !!copy.querySelector('img, table, svg, canvas, mjx-container, .MathJax, .katex, .answer-blank, .answer-blank-large');
                    }

                    function nodeOverflowsPage(node, pageContent) {
                        const pageBottom = pageContent.getBoundingClientRect().bottom;
                        if (node.classList.contains('q-wrapper')) {
                            // 题目的段后距可以落在页边距中；只按真实文字行和不可拆元素判定溢出。
                            return !!findOverflowBoundary(node, pageBottom);
                        }
                        return node.getBoundingClientRect().bottom > pageBottom + overflowTolerance;
                    }

                    function boundaryBefore(element) {
                        const parent = element.parentNode;
                        if (!parent) return null;
                        return { container: parent, offset: Array.prototype.indexOf.call(parent.childNodes, element) };
                    }

                    function textOverflowBoundary(textNode, pageBottom) {
                        if (!textNode.data.length) return null;
                        const fullRange = document.createRange();
                        fullRange.selectNodeContents(textNode);
                        const fullRects = Array.from(fullRange.getClientRects());
                        const overflowLine = fullRects.find(rect => rect.bottom > pageBottom + overflowTolerance);
                        if (!overflowLine) return null;

                        // 先二分定位第一个进入溢出视觉行的字符，再向前回退到该视觉行的首字符。
                        // 这样即使浏览器对长文本节点的 Range 矩形计算有细微误差，也不会把同一行拆成上下两页。
                        let low = 0;
                        let high = textNode.data.length - 1;
                        while (low < high) {
                            const middle = Math.floor((low + high) / 2);
                            const range = document.createRange();
                            range.setStart(textNode, 0);
                            range.setEnd(textNode, middle + 1);
                            const prefixOverflows = Array.from(range.getClientRects()).some(rect => rect.bottom > pageBottom + overflowTolerance);
                            if (prefixOverflows) high = middle;
                            else low = middle + 1;
                        }

                        let offset = low;
                        const charRect = index => {
                            if (index < 0 || index >= textNode.data.length) return null;
                            const range = document.createRange();
                            range.setStart(textNode, index);
                            range.setEnd(textNode, index + 1);
                            const rects = Array.from(range.getClientRects());
                            return rects.length ? rects[0] : null;
                        };
                        const currentRect = charRect(offset);
                        if (currentRect) {
                            const lineTop = currentRect.top;
                            while (offset > 0) {
                                const previousRect = charRect(offset - 1);
                                if (!previousRect || Math.abs(previousRect.top - lineTop) > 1.5) break;
                                offset--;
                            }
                        }
                        return { container: textNode, offset };
                    }

                    // 8.4.3.3：HTML 大表格安全行级分页。
                    // 旧逻辑把 table 视为不可拆原子元素，只要剩余空间放不下整表，就把整张表推到下一页，
                    // 会在英语阅读等复杂版式题中制造半页甚至更大的空白。
                    // 新逻辑只允许在真实 <tr> 边界拆分，并拒绝穿过 rowspan；找不到安全边界时仍保持旧行为。
                    function tableRowGroupEndIndex(rows, rowIndex) {
                        if (!rows[rowIndex]) return rowIndex + 1;
                        const group = rows[rowIndex].parentElement;
                        let end = rowIndex + 1;
                        while (end < rows.length && rows[end].parentElement === group) end++;
                        return end;
                    }

                    function tableBoundaryCrossesRowspan(table, splitRowIndex) {
                        const rows = Array.from(table.rows || []);
                        if (splitRowIndex <= 0 || splitRowIndex >= rows.length) return true;
                        for (let rowIndex = 0; rowIndex < splitRowIndex; rowIndex++) {
                            const row = rows[rowIndex];
                            const groupEnd = tableRowGroupEndIndex(rows, rowIndex);
                            for (const cell of Array.from(row.cells || [])) {
                                const raw = cell.getAttribute('rowspan');
                                let span = raw == null || raw === '' ? 1 : Number.parseInt(raw, 10);
                                if (!Number.isFinite(span) || span < 0) span = 1;
                                const spanEnd = span === 0 ? groupEnd : Math.min(groupEnd, rowIndex + Math.max(1, span));
                                if (spanEnd > splitRowIndex) return true;
                            }
                        }
                        return false;
                    }

                    function findSafeTableSplitBoundary(table, pageBottom) {
                        const rows = Array.from(table.rows || []);
                        if (rows.length < 2) return null;
                        const tableRect = table.getBoundingClientRect();
                        if (tableRect.top >= pageBottom - overflowTolerance) return null;

                        // 找出本页能够完整容纳的最后一行。使用真实行矩形，不依据文字高度猜测。
                        let fullyFittingRows = 0;
                        for (let i = 0; i < rows.length; i++) {
                            const rect = rows[i].getBoundingClientRect();
                            if (rect.bottom <= pageBottom + overflowTolerance) fullyFittingRows = i + 1;
                            else break;
                        }
                        if (fullyFittingRows <= 0 || fullyFittingRows >= rows.length) return null;

                        // 从最靠近页底的位置向上找安全边界；rowspan 穿越的边界一律禁止。
                        // 同时避免只留下一个几乎没有高度的表格碎片。
                        for (let splitRowIndex = fullyFittingRows; splitRowIndex >= 1; splitRowIndex--) {
                            if (tableBoundaryCrossesRowspan(table, splitRowIndex)) continue;
                            const keptBottom = rows[splitRowIndex - 1].getBoundingClientRect().bottom;
                            const keptHeight = keptBottom - tableRect.top;
                            if (keptHeight < 18 && splitRowIndex < fullyFittingRows) continue;
                            return {
                                type: 'table-row',
                                table,
                                rowIndex: splitRowIndex,
                                totalRows: rows.length,
                                keptBottom,
                                unusedSpace: Math.max(0, pageBottom - keptBottom)
                            };
                        }
                        return null;
                    }

                    function findOverflowBoundary(root, pageBottom) {
                        const atomicSelector = 'img, table, svg, canvas, pre, br, mjx-container, .MathJax, .katex, .answer-blank, .answer-blank-large, .zujuanjs-answer-image-slice';

                        function visit(parent) {
                            for (const child of Array.from(parent.childNodes)) {
                                if (child.nodeType === Node.TEXT_NODE) {
                                    const boundary = textOverflowBoundary(child, pageBottom);
                                    if (boundary) return boundary;
                                    continue;
                                }
                                if (child.nodeType !== Node.ELEMENT_NODE || child.classList.contains('q-toolbar')) continue;

                                const rect = child.getBoundingClientRect();
                                if (rect.top >= pageBottom - overflowTolerance && hasPrintableContent(child)) {
                                    return boundaryBefore(child);
                                }
                                if (child.matches('table')) {
                                    if (rect.bottom > pageBottom + overflowTolerance) {
                                        const tableBoundary = findSafeTableSplitBoundary(child, pageBottom);
                                        if (tableBoundary) return tableBoundary;
                                        return boundaryBefore(child);
                                    }
                                    continue;
                                }
                                if (child.matches(atomicSelector)) {
                                    if (rect.bottom > pageBottom + overflowTolerance) return boundaryBefore(child);
                                    continue;
                                }

                                const nestedBoundary = visit(child);
                                if (nestedBoundary) return nestedBoundary;

                                // 兼容没有文字节点、但自身有高度的站点组件。
                                if (rect.bottom > pageBottom + overflowTolerance && !child.textContent.trim()) {
                                    return boundaryBefore(child);
                                }
                            }
                            return null;
                        }

                        return visit(root);
                    }

                    // 8.4：连续答案源图流。
                    // 不再先把所有切块塞进一个答案容器再交给通用分页，而是把答案图片抽成独立流节点，
                    // 每一段直接按当前 A4 的实际剩余高度生成并放入页面。最后一段结束后，下一张完整答案图立即续排。
                    function normalizeManualCutOffsets(imageKey) {
                        const current = previewSettings.answerImageCutOffsets[imageKey];
                        if (!current || typeof current !== 'object') previewSettings.answerImageCutOffsets[imageKey] = {};
                        return previewSettings.answerImageCutOffsets[imageKey];
                    }

                    function normalizeManualCutPositions(imageKey) {
                        const current = previewSettings.answerImageCutPositions[imageKey];
                        if (!current || typeof current !== 'object') previewSettings.answerImageCutPositions[imageKey] = {};
                        return previewSettings.answerImageCutPositions[imageKey];
                    }

                    function getAnswerImageKey(image) {
                        const existing = image.dataset.answerImageKey;
                        if (existing) return existing;
                        const source = String(image.getAttribute('src') || '');
                        const hash = Math.abs(source.split('').reduce((a, c) => ((a << 5) - a + c.charCodeAt(0)) | 0, 0));
                        const key = 'answer-image-' + hash;
                        image.dataset.answerImageKey = key;
                        return key;
                    }


                    // 8.2.1 精准行间切割：不再把整张超长图纵向压缩后寻找白线。
                    // 每次只分析“理论分页点附近”的原图局部条带，并尽量保持原始 Y 像素分辨率。
                    // 这样即使答案图非常长，也不会因为纵向缩小而把一行文字压成 1~2 个扫描像素。
                    function loadImageElementFromDataUrl(dataUrl) {
                        return new Promise((resolve, reject) => {
                            const image = new Image();
                            image.onload = () => resolve(image);
                            image.onerror = () => reject(new Error('答案图片分析资源加载失败'));
                            image.src = dataUrl;
                        });
                    }

                    function waitForAnswerImageReady(image, timeout = 16000) {
                        if (image?.complete && Number(image.naturalWidth) > 0 && Number(image.naturalHeight) > 0) {
                            return Promise.resolve(image);
                        }
                        return new Promise((resolve, reject) => {
                            if (!image) return reject(new Error('答案图片不存在'));
                            let settled = false;
                            const done = (fn, value) => {
                                if (settled) return;
                                settled = true;
                                clearTimeout(timer);
                                image.removeEventListener('load', onLoad);
                                image.removeEventListener('error', onError);
                                fn(value);
                            };
                            const onLoad = () => done(resolve, image);
                            const onError = () => done(reject, new Error('答案图片加载失败'));
                            const timer = setTimeout(() => done(reject, new Error('等待答案图片加载超时')), timeout);
                            image.addEventListener('load', onLoad, { once: true });
                            image.addEventListener('error', onError, { once: true });
                        });
                    }

                    function probeAnswerImagePixelAccess(image) {
                        const canvas = document.createElement('canvas');
                        canvas.width = 2;
                        canvas.height = 2;
                        const ctx = canvas.getContext('2d', { willReadFrequently: true });
                        if (!ctx) throw new Error('浏览器不支持 Canvas 像素分析');
                        ctx.drawImage(image, 0, 0, 2, 2);
                        ctx.getImageData(0, 0, 1, 1);
                        return true;
                    }

                    function buildAnswerImageRowModel(pixelSource, naturalWidth, naturalHeight) {
                        const width = Math.max(1, Math.round(Number(naturalWidth) || 0));
                        const height = Math.max(1, Math.round(Number(naturalHeight) || 0));
                        if (!pixelSource || width < 2 || height < 2) return null;

                        // 8.4.3.2 根治策略：对整张答案原图建立纵向“文字行包围盒”。
                        // 横向适度缩小以降低内存，但纵向始终保持 1:1 原图像素，避免长图压缩后丢失行间空白。
                        const scanWidth = Math.max(240, Math.min(720, width));
                        const tileHeight = 1200;
                        const rowInk = new Uint16Array(height);
                        const canvas = document.createElement('canvas');
                        canvas.width = scanWidth;
                        const ctx = canvas.getContext('2d', { willReadFrequently: true });
                        if (!ctx) return null;

                        for (let sourceTop = 0; sourceTop < height; sourceTop += tileHeight) {
                            const sourceHeight = Math.min(tileHeight, height - sourceTop);
                            canvas.height = sourceHeight;
                            ctx.fillStyle = '#fff';
                            ctx.fillRect(0, 0, scanWidth, sourceHeight);
                            ctx.drawImage(pixelSource, 0, sourceTop, width, sourceHeight, 0, 0, scanWidth, sourceHeight);
                            const pixels = ctx.getImageData(0, 0, scanWidth, sourceHeight).data;

                            const samples = [];
                            const sx = Math.max(1, Math.floor(scanWidth / 90));
                            const sy = Math.max(1, Math.floor(sourceHeight / 90));
                            for (let y = 0; y < sourceHeight; y += sy) {
                                for (let x = 0; x < scanWidth; x += sx) {
                                    const i = (y * scanWidth + x) * 4;
                                    if (pixels[i + 3] < 32) continue;
                                    samples.push(0.2126 * pixels[i] + 0.7152 * pixels[i + 1] + 0.0722 * pixels[i + 2]);
                                }
                            }
                            const backgroundLuma = percentileLocal(samples, 0.92);
                            // 强墨迹阈值刻意避开组卷网浅灰水印；蓝色“答案/解析”等标签用色差单独捕获。
                            const hardThreshold = Math.max(145, Math.min(205, backgroundLuma - 46));
                            const colorThreshold = Math.max(165, Math.min(224, backgroundLuma - 18));
                            const left = Math.max(0, Math.floor(scanWidth * 0.006));
                            const right = Math.max(left + 1, scanWidth - left);

                            for (let y = 0; y < sourceHeight; y++) {
                                let ink = 0;
                                for (let x = left; x < right; x++) {
                                    const i = (y * scanWidth + x) * 4;
                                    if (pixels[i + 3] < 32) continue;
                                    const r = pixels[i], g = pixels[i + 1], b = pixels[i + 2];
                                    const luma = 0.2126 * r + 0.7152 * g + 0.0722 * b;
                                    const maxC = Math.max(r, g, b), minC = Math.min(r, g, b);
                                    const chroma = maxC - minC;
                                    if (luma < hardThreshold || (chroma > 38 && minC < 215 && luma < colorThreshold)) ink++;
                                }
                                rowInk[sourceTop + y] = Math.min(65535, ink);
                            }
                        }

                        // 一条真正的文字行横向会包含大量墨迹；纵向表格边线/水印只有少量像素，不能把整块误判成文字行。
                        const inkLimit = Math.max(6, Math.floor(scanWidth * 0.009));
                        const active = new Uint8Array(height);
                        for (let y = 0; y < height; y++) active[y] = rowInk[y] >= inkLimit ? 1 : 0;

                        // 填掉同一文字行内部最多 3px 的抗锯齿/标点裂缝。
                        let zeroStart = -1;
                        for (let y = 0; y <= height; y++) {
                            const isZero = y < height && active[y] === 0;
                            if (isZero && zeroStart < 0) zeroStart = y;
                            if ((!isZero || y === height) && zeroStart >= 0) {
                                const zeroEnd = y - 1;
                                const zeroLength = zeroEnd - zeroStart + 1;
                                if (zeroStart > 0 && y < height && active[zeroStart - 1] && active[y] && zeroLength <= 3) {
                                    for (let yy = zeroStart; yy <= zeroEnd; yy++) active[yy] = 1;
                                }
                                zeroStart = -1;
                            }
                        }

                        const rawBands = [];
                        let bandStart = -1;
                        for (let y = 0; y <= height; y++) {
                            const on = y < height && active[y] === 1;
                            if (on && bandStart < 0) bandStart = y;
                            if ((!on || y === height) && bandStart >= 0) {
                                rawBands.push({ start: bandStart, end: y - 1 });
                                bandStart = -1;
                            }
                        }

                        // 给每一行上下各留 2px 禁切保护，并合并非常接近的碎片。
                        const bands = [];
                        rawBands.forEach(item => {
                            const expanded = {
                                start: Math.max(0, item.start - 2),
                                end: Math.min(height - 1, item.end + 2)
                            };
                            const last = bands[bands.length - 1];
                            if (last && expanded.start <= last.end + 4) last.end = Math.max(last.end, expanded.end);
                            else bands.push(expanded);
                        });

                        const median = (values, fallback) => {
                            const clean = values.filter(Number.isFinite).sort((a, b) => a - b);
                            if (!clean.length) return fallback;
                            const mid = Math.floor(clean.length / 2);
                            return clean.length % 2 ? clean[mid] : (clean[mid - 1] + clean[mid]) / 2;
                        };
                        const bandHeights = bands.map(band => band.end - band.start + 1).filter(value => value >= 3 && value <= 120);
                        const gaps = [];
                        const pitches = [];
                        for (let i = 0; i < bands.length - 1; i++) {
                            const startY = bands[i].end + 1;
                            const endY = bands[i + 1].start - 1;
                            const gapHeight = endY - startY + 1;
                            if (gapHeight >= 2) {
                                gaps.push({ start: startY, end: endY, height: gapHeight, center: (startY + endY) / 2 });
                            }
                            const pitch = bands[i + 1].start - bands[i].start;
                            if (pitch >= 12 && pitch <= 180) pitches.push(pitch);
                        }
                        const gapHeights = gaps.map(gap => gap.height).filter(value => value <= 180);
                        const typicalBandHeight = median(bandHeights, 26);
                        const typicalGap = median(gapHeights, 24);
                        const typicalPitch = median(pitches, typicalBandHeight + typicalGap);

                        return {
                            version: 1,
                            method: 'global-row-box',
                            bands,
                            gaps,
                            contentTop: bands.length ? bands[0].start : 0,
                            contentBottom: bands.length ? bands[bands.length - 1].end : height - 1,
                            typicalBandHeight,
                            typicalGap,
                            typicalPitch,
                            scanWidth,
                            inkLimit
                        };
                    }

                    async function computeAnswerImageAnalysis(image, imageKey) {
                        await waitForAnswerImageReady(image);
                        const directWidth = Number(image.naturalWidth) || 0;
                        const directHeight = Number(image.naturalHeight) || 0;
                        if (!directWidth || !directHeight) throw new Error('答案图片没有有效原始尺寸');

                        let pixelSource = image;
                        let naturalWidth = directWidth;
                        let naturalHeight = directHeight;
                        let via = 'direct';

                        // 第一选择：直接使用页面中已经加载的原图。跨域污染 Canvas 时，再通过父页面油猴请求转成 Data URL。
                        try {
                            probeAnswerImagePixelAccess(image);
                        } catch (directError) {
                            const source = String(image.currentSrc || image.getAttribute('src') || '');
                            if (!source) throw directError;
                            const absoluteUrl = new URL(source, document.baseURI).href;
                            const dataUrl = source.startsWith('data:') ? source : await requestResourceFromParent(absoluteUrl, 28000);
                            const cleanImage = await loadImageElementFromDataUrl(dataUrl);
                            probeAnswerImagePixelAccess(cleanImage);
                            pixelSource = cleanImage;
                            naturalWidth = Number(cleanImage.naturalWidth) || directWidth;
                            naturalHeight = Number(cleanImage.naturalHeight) || directHeight;
                            via = 'gm-data-url';
                        }

                        const rowModel = buildAnswerImageRowModel(pixelSource, naturalWidth, naturalHeight);
                        return { naturalWidth, naturalHeight, pixelSource, via, rowModel };
                    }

                    function ensureAnswerImageAnalysis(image, imageKey) {
                        if (!previewSettings.answerCutSmartSnap || !image || !imageKey) return Promise.resolve(null);
                        const cached = answerImageAnalysisCache.get(imageKey);
                        if (cached?.status === 'ready') return Promise.resolve(cached.result);
                        if (cached?.status === 'failed') return Promise.resolve(null);
                        if (cached?.promise) return cached.promise;
                        const entry = { status: 'loading', result: null, promise: null };
                        entry.promise = enqueueAnswerImageAnalysis(() => computeAnswerImageAnalysis(image, imageKey))
                            .then(result => {
                                entry.status = 'ready';
                                entry.result = result;
                                entry.promise = null;
                                scheduleRender();
                                return result;
                            })
                            .catch(error => {
                                entry.status = 'failed';
                                entry.promise = null;
                                console.warn('答案图片精准切线分析失败，已回退到重叠保护：', imageKey, error);
                                return null;
                            });
                        answerImageAnalysisCache.set(imageKey, entry);
                        return entry.promise;
                    }

                    function primeAnswerImageAnalyses() {
                        if (!previewSettings.answerCutSmartSnap) return Promise.resolve([]);
                        const jobs = Array.from(sourceContent.querySelectorAll('img.zujuanjs-answer-render-image')).map(image => {
                            const imageKey = getAnswerImageKey(image);
                            return ensureAnswerImageAnalysis(image, imageKey);
                        });
                        return Promise.all(jobs);
                    }

                    function getReadyAnswerImageAnalysis(imageKey) {
                        const cached = answerImageAnalysisCache.get(imageKey);
                        return cached?.status === 'ready' ? cached.result : null;
                    }

                    function percentileLocal(values, ratio) {
                        if (!values.length) return 255;
                        const copy = values.slice().sort((a, b) => a - b);
                        const index = Math.max(0, Math.min(copy.length - 1, Math.round((copy.length - 1) * ratio)));
                        return copy[index];
                    }

                    function scanRowModelSafeCuts(rowModel, targetSourceY, minSourceY, maxSourceY) {
                        if (!rowModel?.gaps?.length) return null;
                        const minY = Math.max(0, Number(minSourceY) || 0);
                        const maxY = Math.max(minY + 1, Number(maxSourceY) || minY + 1);
                        const targetY = Math.max(minY, Math.min(maxY, Number(targetSourceY) || minY));
                        const results = [];

                        rowModel.gaps.forEach(gap => {
                            if (gap.end < minY || gap.start > maxY) return;
                            const startY = Math.max(minY, gap.start);
                            const endY = Math.min(maxY, gap.end);
                            const gapHeight = endY - startY + 1;
                            if (gapHeight < 2) return;
                            const margin = Math.min(Math.max(1, gapHeight * 0.18), Math.max(1, gapHeight / 2 - 0.5));
                            const innerStart = startY + margin;
                            const innerEnd = endY - margin;
                            if (innerEnd <= innerStart) return;
                            const candidateY = Math.max(innerStart, Math.min(innerEnd, targetY));
                            results.push({
                                y: Math.round(candidateY),
                                gap: Math.max(1, Math.round(gapHeight)),
                                density: 0,
                                energy: 0,
                                distance: Math.abs(candidateY - targetY),
                                sourceStart: Math.round(startY),
                                sourceEnd: Math.round(endY),
                                containsTarget: targetY >= startY && targetY <= endY
                            });
                        });

                        let targetBand = null;
                        for (const band of rowModel.bands || []) {
                            if (targetY >= band.start && targetY <= band.end) {
                                targetBand = band;
                                break;
                            }
                        }
                        results.profile = {
                            targetProtected: Boolean(targetBand),
                            targetBandStart: targetBand ? targetBand.start : null,
                            targetBandEnd: targetBand ? targetBand.end : null,
                            sourcePerScanY: 1,
                            targetSourceY: targetY,
                            typicalPitch: Number(rowModel.typicalPitch) || 0,
                            method: rowModel.method || 'global-row-box'
                        };
                        return results;
                    }

                    function scanLocalSafeSourceCuts(analysis, targetSourceY, minSourceY, maxSourceY, options = {}) {
                        if (!analysis?.pixelSource || !analysis.naturalWidth || !analysis.naturalHeight) return [];
                        const rowModelResult = scanRowModelSafeCuts(analysis.rowModel, targetSourceY, minSourceY, maxSourceY);
                        if (rowModelResult) return rowModelResult;
                        const naturalWidth = Math.max(1, Math.round(analysis.naturalWidth));
                        const naturalHeight = Math.max(1, Math.round(analysis.naturalHeight));
                        const minY = Math.max(1, Math.min(naturalHeight - 2, Number(minSourceY) || 0));
                        const maxY = Math.max(minY + 1, Math.min(naturalHeight - 1, Number(maxSourceY) || naturalHeight - 1));
                        const targetY = Math.max(minY, Math.min(maxY, Number(targetSourceY) || minY));

                        // 8.4：只分析目标分页点附近的原图条带。纵向优先保持原始像素，避免超长图压缩后丢失行间空白。
                        const context = 12;
                        const sourceTop = Math.max(0, Math.floor(minY - context));
                        const sourceBottom = Math.min(naturalHeight, Math.ceil(maxY + context));
                        const sourceHeight = Math.max(2, sourceBottom - sourceTop);
                        const scanWidth = Math.max(120, Math.min(1200, naturalWidth));
                        const yScale = Math.min(1, 1600 / sourceHeight);
                        const scanHeight = Math.max(2, Math.round(sourceHeight * yScale));
                        const sourcePerScanY = sourceHeight / scanHeight;

                        const canvas = document.createElement('canvas');
                        canvas.width = scanWidth;
                        canvas.height = scanHeight;
                        const ctx = canvas.getContext('2d', { willReadFrequently: true });
                        if (!ctx) return [];
                        ctx.fillStyle = '#fff';
                        ctx.fillRect(0, 0, scanWidth, scanHeight);
                        ctx.drawImage(analysis.pixelSource, 0, sourceTop, naturalWidth, sourceHeight, 0, 0, scanWidth, scanHeight);
                        const pixels = ctx.getImageData(0, 0, scanWidth, scanHeight).data;

                        const samples = [];
                        const sx = Math.max(1, Math.floor(scanWidth / 96));
                        const sy = Math.max(1, Math.floor(scanHeight / 110));
                        for (let y = 0; y < scanHeight; y += sy) {
                            for (let x = 0; x < scanWidth; x += sx) {
                                const i = (y * scanWidth + x) * 4;
                                if (pixels[i + 3] < 32) continue;
                                samples.push(0.2126 * pixels[i] + 0.7152 * pixels[i + 1] + 0.0722 * pixels[i + 2]);
                            }
                        }
                        const backgroundLuma = percentileLocal(samples, 0.92);
                        // 双阈值：硬墨迹负责抓正文骨架，软能量负责抓细字体的抗锯齿边缘。
                        // 即使因此把浅灰水印误认为内容，也只会少排一点，不会把真实文字误判为空白。
                        const hardThreshold = Math.max(150, Math.min(242, backgroundLuma - 22));
                        const softThreshold = Math.max(170, Math.min(250, backgroundLuma - 8));
                        const left = Math.max(0, Math.floor(scanWidth * 0.008));
                        const right = Math.max(left + 1, scanWidth - left);
                        const usableWidth = Math.max(1, right - left);
                        const rowInk = new Uint16Array(scanHeight);
                        const rowEnergy = new Float32Array(scanHeight);

                        for (let y = 0; y < scanHeight; y++) {
                            let ink = 0;
                            let energy = 0;
                            for (let x = left; x < right; x++) {
                                const i = (y * scanWidth + x) * 4;
                                if (pixels[i + 3] < 32) continue;
                                const r = pixels[i], g = pixels[i + 1], b = pixels[i + 2];
                                const luma = 0.2126 * r + 0.7152 * g + 0.0722 * b;
                                const chroma = Math.max(r, g, b) - Math.min(r, g, b);
                                if (luma < hardThreshold || (chroma > 30 && luma < backgroundLuma - 6)) ink++;
                                if (luma < softThreshold) energy += Math.min(72, softThreshold - luma);
                            }
                            rowInk[y] = Math.min(65535, ink);
                            rowEnergy[y] = energy;
                        }

                        const hardLimit = Math.max(1, Math.floor(usableWidth * 0.001));
                        const energyLimit = Math.max(35, usableWidth * 0.28);
                        const protectedRow = new Uint8Array(scanHeight);
                        const protectPx = Math.max(2, Number(options.protectPx) || 4);
                        const protectRadius = Math.max(1, Math.ceil(protectPx / sourcePerScanY));
                        for (let y = 0; y < scanHeight; y++) {
                            if (rowInk[y] <= hardLimit && rowEnergy[y] <= energyLimit) continue;
                            const from = Math.max(0, y - protectRadius);
                            const to = Math.min(scanHeight - 1, y + protectRadius);
                            for (let yy = from; yy <= to; yy++) protectedRow[yy] = 1;
                        }

                        const lowerScan = Math.max(0, Math.ceil((minY - sourceTop) / sourcePerScanY));
                        const upperScan = Math.min(scanHeight - 1, Math.floor((maxY - sourceTop) / sourcePerScanY));
                        const targetScan = Math.max(lowerScan, Math.min(upperScan, Math.round((targetY - sourceTop) / sourcePerScanY)));
                        const minimumGapPx = Math.max(1, Number(options.minimumGapPx) || 4);
                        const minimumGapScan = Math.max(1, Math.ceil(minimumGapPx / sourcePerScanY));
                        const results = [];
                        let runStart = -1;

                        for (let y = lowerScan; y <= upperScan + 1; y++) {
                            const safe = y <= upperScan && protectedRow[y] === 0;
                            if (safe && runStart < 0) runStart = y;
                            if ((!safe || y > upperScan) && runStart >= 0) {
                                const runEnd = y - 1;
                                const runLength = runEnd - runStart + 1;
                                if (runLength >= minimumGapScan) {
                                    const sourceStart = sourceTop + runStart * sourcePerScanY;
                                    const sourceEnd = sourceTop + (runEnd + 1) * sourcePerScanY;
                                    const gapHeight = sourceEnd - sourceStart;
                                    const safeMargin = Math.min(Math.max(0.5, gapHeight / 2 - 0.35), Math.max(1.25, gapHeight * 0.18));
                                    const innerStart = sourceStart + safeMargin;
                                    const innerEnd = sourceEnd - safeMargin;
                                    if (innerEnd > innerStart) {
                                        const candidateY = Math.max(innerStart, Math.min(innerEnd, targetY));
                                        let totalInk = 0;
                                        let totalEnergy = 0;
                                        for (let yy = runStart; yy <= runEnd; yy++) {
                                            totalInk += rowInk[yy];
                                            totalEnergy += rowEnergy[yy];
                                        }
                                        results.push({
                                            y: Math.round(candidateY),
                                            gap: Math.max(1, Math.round(gapHeight)),
                                            density: totalInk / Math.max(1, runLength * usableWidth),
                                            energy: totalEnergy / Math.max(1, runLength * usableWidth),
                                            distance: Math.abs(candidateY - targetY),
                                            sourceStart: Math.round(sourceStart),
                                            sourceEnd: Math.round(sourceEnd),
                                            containsTarget: targetY >= sourceStart && targetY <= sourceEnd
                                        });
                                    }
                                }
                                runStart = -1;
                            }
                        }

                        // 如果理论分页点正落在文字/公式/横线等内容带内，记录这整条内容带的边界。
                        // chooseSmartAnswerCut 会优先在它的上边界之前切，让整行进入下一页，而不是把一行劈成两半。
                        let bandStart = -1, bandEnd = -1;
                        if (protectedRow[targetScan]) {
                            let a = targetScan, b = targetScan;
                            while (a > lowerScan && protectedRow[a - 1]) a--;
                            while (b < upperScan && protectedRow[b + 1]) b++;
                            bandStart = sourceTop + a * sourcePerScanY;
                            bandEnd = sourceTop + (b + 1) * sourcePerScanY;
                        }
                        results.profile = {
                            targetProtected: Boolean(protectedRow[targetScan]),
                            targetBandStart: bandStart >= 0 ? bandStart : null,
                            targetBandEnd: bandEnd >= 0 ? bandEnd : null,
                            sourcePerScanY,
                            targetSourceY: targetY
                        };
                        return results;
                    }

                    function chooseSmartAnswerCut(imageKey, logicalStart, maxEnd, automaticBase, sourceScale) {
                        if (!previewSettings.answerCutSmartSnap) return null;
                        const analysis = getReadyAnswerImageAnalysis(imageKey);
                        if (!analysis) return null;
                        const searchWindow = clamp(previewSettings.answerCutSearchWindow, 24, 180, 72);
                        const targetSource = automaticBase * sourceScale;
                        const minSource = (logicalStart + 2) * sourceScale;
                        const maxSource = Math.max(minSource + 1, (maxEnd - 0.5) * sourceScale);
                        const rowModel = analysis.rowModel;

                        if (rowModel?.gaps?.length) {
                            // 根治策略：不再“向上搜索最多112个显示像素”。
                            // 直接选择本页容量内最后一个完整文字行之后的真实行间空白。
                            // 对标准答案图，最坏只退让约一个文字行距，不会空掉四五行。
                            const desiredSource = Math.max(minSource, Math.min(maxSource, targetSource));
                            const typicalPitch = Math.max(18, Number(rowModel.typicalPitch) || 54);
                            const configuredBacktrack = Math.max(12, searchWindow * sourceScale);
                            const maxSafeBacktrack = Math.max(20, Math.min(configuredBacktrack, typicalPitch * 1.55));
                            let chosen = null;

                            for (const gap of rowModel.gaps) {
                                if (gap.end < minSource) continue;
                                if (gap.start > desiredSource) break;
                                const gapHeight = Math.max(1, gap.end - gap.start + 1);
                                const margin = Math.min(Math.max(1.5, gapHeight * 0.18), Math.max(1.5, gapHeight / 2 - 0.5));
                                const innerStart = Math.max(minSource, gap.start + margin);
                                const innerEnd = Math.min(desiredSource, gap.end - margin);
                                if (innerEnd <= innerStart) continue;
                                // 取该空白带尽量靠下的位置，以最大化本页利用率，同时保留离下一行的安全边距。
                                const sourceY = innerEnd;
                                chosen = { sourceY, gapHeight, gap };
                            }

                            if (chosen) {
                                const backtrack = desiredSource - chosen.sourceY;
                                if (backtrack <= maxSafeBacktrack + 1) {
                                    return {
                                        cut: Math.max(logicalStart + 2, Math.min(maxEnd, chosen.sourceY / sourceScale)),
                                        sourceY: chosen.sourceY,
                                        gap: chosen.gapHeight,
                                        density: 0,
                                        precision: 'row-box-gap',
                                        requiresOverlap: false,
                                        backtrackSource: backtrack,
                                        typicalPitch
                                    };
                                }
                            }

                            // 理论切点若落在普通文字行内部，只允许退到这一行之前；
                            // 若遇到超高连续图形/表格区域，不再远距离退让，而是维持满页切点并交给重叠保护兜底。
                            let targetBand = null;
                            for (const band of rowModel.bands || []) {
                                if (desiredSource >= band.start && desiredSource <= band.end) {
                                    targetBand = band;
                                    break;
                                }
                            }
                            if (targetBand) {
                                const bandHeight = targetBand.end - targetBand.start + 1;
                                const sourceCut = Math.max(minSource, targetBand.start - 2);
                                const backtrack = desiredSource - sourceCut;
                                if (bandHeight <= typicalPitch * 1.8 && backtrack >= 0 && backtrack <= maxSafeBacktrack + 1) {
                                    return {
                                        cut: Math.max(logicalStart + 2, Math.min(maxEnd, sourceCut / sourceScale)),
                                        sourceY: sourceCut,
                                        gap: 0,
                                        density: 1,
                                        precision: 'row-box-before-line',
                                        requiresOverlap: false,
                                        backtrackSource: backtrack,
                                        typicalPitch
                                    };
                                }
                            }

                            return {
                                cut: automaticBase,
                                sourceY: targetSource,
                                gap: 0,
                                density: 1,
                                precision: 'row-box-overlap-fallback',
                                requiresOverlap: true,
                                typicalPitch
                            };
                        }

                        // 极少数无法建立全图行模型的图片，保留8.4.3的局部扫描作为兼容兜底，
                        // 但最大回退收紧到约两行，禁止再次制造大块页尾空白。
                        const minDisplay = logicalStart + 3;
                        const maxDisplay = Math.max(minDisplay + 1, maxEnd - 0.5);
                        const maxBacktrackDisplay = Math.max(36, Math.min(64, searchWindow * 0.9));
                        const forwardDisplay = Math.max(4, Math.min(12, searchWindow * 0.14));
                        const lowerDisplay = Math.max(minDisplay, automaticBase - maxBacktrackDisplay);
                        const upperDisplay = Math.min(maxDisplay, automaticBase + forwardDisplay);
                        if (upperDisplay <= lowerDisplay + 1) {
                            return { cut: automaticBase, sourceY: targetSource, gap: 0, density: 1, precision: 'overlap-fallback', requiresOverlap: true };
                        }

                        let candidates = [];
                        try {
                            candidates = scanLocalSafeSourceCuts(
                                analysis,
                                targetSource,
                                lowerDisplay * sourceScale,
                                upperDisplay * sourceScale,
                                { minimumGapPx: 4, protectPx: 4 }
                            );
                        } catch (error) {
                            console.warn('局部行带扫描失败，改用重叠保护：', imageKey, error);
                            return { cut: automaticBase, sourceY: targetSource, gap: 0, density: 1, precision: 'overlap-fallback', requiresOverlap: true };
                        }

                        if (candidates.length) {
                            candidates.sort((a, b) => {
                                if (a.containsTarget !== b.containsTarget) return a.containsTarget ? -1 : 1;
                                return a.distance - b.distance;
                            });
                            const chosen = candidates[0];
                            const displayCut = chosen.y / sourceScale;
                            if (automaticBase - displayCut <= maxBacktrackDisplay + 1) {
                                return {
                                    cut: Math.max(logicalStart + 2, Math.min(maxEnd, displayCut)),
                                    sourceY: chosen.y,
                                    gap: chosen.gap,
                                    density: chosen.density,
                                    precision: chosen.containsTarget ? 'target-gap' : 'local-gap',
                                    requiresOverlap: chosen.gap < 3
                                };
                            }
                        }

                        return { cut: automaticBase, sourceY: targetSource, gap: 0, density: 1, precision: 'overlap-fallback', requiresOverlap: true };
                    }

                    function findNearestSafeSourceCut(imageKey, sourceY, minSourceY = 0, maxSourceY = Infinity) {
                        const analysis = getReadyAnswerImageAnalysis(imageKey);
                        if (!analysis) return null;
                        try {
                            const candidates = scanLocalSafeSourceCuts(analysis, sourceY, minSourceY, maxSourceY);
                            if (!candidates.length) return null;
                            candidates.sort((a, b) => Math.abs(a.y - sourceY) - Math.abs(b.y - sourceY));
                            return candidates[0];
                        } catch (error) {
                            console.warn('手动切线吸附扫描失败：', imageKey, error);
                            return null;
                        }
                    }

                    function sanitizeAnswerFlowFragment(fragment) {
                        if (!fragment) return fragment;
                        fragment.querySelectorAll('.knowledge-box, .q-toolbar, script, style, noscript').forEach(element => element.remove());
                        fragment.style.setProperty('margin', '0', 'important');
                        fragment.style.setProperty('padding-top', '0', 'important');
                        fragment.style.setProperty('padding-bottom', '0', 'important');
                        fragment.style.setProperty('min-height', '0', 'important');
                        fragment.style.setProperty('height', 'auto', 'important');
                        fragment.querySelectorAll('[style]').forEach(element => {
                            if (element.matches('img, svg, table, canvas, mjx-container')) return;
                            const inline = element.style;
                            if (inline.minHeight) inline.minHeight = '0px';
                            if (inline.height && /^(?:auto|\d+(?:\.\d+)?px)$/.test(inline.height)) inline.height = 'auto';
                        });
                        return fragment;
                    }

                    function makeAnswerHtmlFlowPart(original, contents, partIndex) {
                        const part = original.cloneNode(false);
                        part.dataset.answerFlowExpanded = '1';
                        part.classList.add('zujuanjs-answer-flow-html');
                        part.appendChild(contents);
                        sanitizeAnswerFlowFragment(part);
                        if (partIndex > 0 && part.classList.contains('q-wrapper')) {
                            part.classList.add('continued-from-previous');
                            preserveQuestionIndent(part, original);
                        }
                        return hasPrintableContent(part) ? part : null;
                    }

                    function expandNodeAroundAnswerImages(node) {
                        if (!node || node.dataset.answerFlowExpanded === '1') return null;
                        if (!node.matches('.q-wrapper, .zujuanjs-answer-content, .zujuanjs-answer-item')) return null;
                        const images = Array.from(node.querySelectorAll('img')).filter(isRenderedAnswerImage);
                        if (!images.length) return null;

                        const parts = [];
                        let startContainer = node;
                        let startOffset = 0;
                        let htmlPartIndex = 0;

                        images.forEach(image => {
                            const before = document.createRange();
                            before.setStart(startContainer, startOffset);
                            before.setEndBefore(image);
                            const htmlPart = makeAnswerHtmlFlowPart(node, before.cloneContents(), htmlPartIndex++);
                            if (htmlPart) parts.push(htmlPart);

                            const flow = document.createElement('div');
                            flow.className = 'zujuanjs-answer-image-flow';
                            flow.dataset.answerFlowExpanded = '1';
                            flow.dataset.answerImageKey = getAnswerImageKey(image);
                            if (node.dataset.blockId) flow.dataset.blockId = node.dataset.blockId;
                            const clone = image.cloneNode(true);
                            clone.dataset.answerImageKey = flow.dataset.answerImageKey;
                            clone.classList.add('zujuanjs-answer-render-image');
                            flow.appendChild(clone);
                            parts.push(flow);

                            startContainer = image.parentNode;
                            startOffset = Array.prototype.indexOf.call(image.parentNode.childNodes, image) + 1;
                        });

                        const after = document.createRange();
                        after.setStart(startContainer, startOffset);
                        after.setEnd(node, node.childNodes.length);
                        const tailPart = makeAnswerHtmlFlowPart(node, after.cloneContents(), htmlPartIndex);
                        if (tailPart) parts.push(tailPart);
                        return parts;
                    }

                    function getPageFlowCursorOffset(pageContent) {
                        const probe = document.createElement('span');
                        probe.setAttribute('aria-hidden', 'true');
                        probe.style.cssText = 'display:block;width:0;height:0;margin:0;padding:0;border:0;line-height:0;';
                        pageContent.appendChild(probe);
                        const pageRect = pageContent.getBoundingClientRect();
                        const probeRect = probe.getBoundingClientRect();
                        // 预览纸张会通过 transform: scale(...) 缩放。getBoundingClientRect() 返回缩放后的视觉像素，
                        // clientHeight 返回未缩放的排版像素；按实际比例还原，避免 70% 预览时只填满 70% 页面。
                        const scaleY = pageContent.clientHeight > 0 && pageRect.height > 0
                            ? pageRect.height / pageContent.clientHeight
                            : 1;
                        const offset = (probeRect.top - pageRect.top) / (scaleY || 1);
                        probe.remove();
                        return Math.max(0, offset);
                    }

                    function getRemainingPageHeight(pageContent) {
                        return Math.max(0, pageContent.clientHeight - getPageFlowCursorOffset(pageContent));
                    }

                    function normalizeFlowImage(image, displayWidth, displayHeight) {
                        image.removeAttribute('id');
                        image.classList.add('zujuanjs-answer-render-image');
                        image.style.setProperty('display', 'block', 'important');
                        image.style.setProperty('width', displayWidth + 'px', 'important');
                        image.style.setProperty('max-width', 'none', 'important');
                        image.style.setProperty('height', displayHeight + 'px', 'important');
                        image.style.setProperty('max-height', 'none', 'important');
                        image.style.setProperty('margin', '0', 'important');
                        image.style.setProperty('padding', '0', 'important');
                        image.style.setProperty('border', '0', 'important');
                        image.style.setProperty('transform', 'none', 'important');
                        return image;
                    }

                    function createAnswerImageSlice(image, displayWidth, displayHeight, visualStart, logicalEnd, imageKey, previousCutIndex, logicalStart, previousCutMeta = null) {
                        const slice = document.createElement('div');
                        slice.className = 'zujuanjs-answer-image-slice';
                        slice.dataset.answerImageKey = imageKey;
                        slice.dataset.sliceStart = String(visualStart);
                        slice.dataset.sliceEnd = String(logicalEnd);
                        slice.style.width = displayWidth + 'px';
                        slice.style.maxWidth = '100%';
                        slice.style.height = Math.max(1, logicalEnd - visualStart) + 'px';

                        const clone = normalizeFlowImage(image.cloneNode(true), displayWidth, displayHeight);
                        clone.style.setProperty('position', 'absolute', 'important');
                        clone.style.setProperty('left', '0', 'important');
                        clone.style.setProperty('top', (-visualStart) + 'px', 'important');
                        slice.appendChild(clone);

                        if (previousCutIndex >= 0) {
                            const positions = normalizeManualCutPositions(imageKey);
                            const legacyOffsets = normalizeManualCutOffsets(imageKey);
                            const hasPosition = Object.prototype.hasOwnProperty.call(positions, previousCutIndex);
                            const hasLegacy = !hasPosition && Object.prototype.hasOwnProperty.call(legacyOffsets, previousCutIndex);
                            const isManual = hasPosition || hasLegacy || Boolean(previousCutMeta?.manual);
                            const isSmart = !isManual && Boolean(previousCutMeta?.smartSafe);
                            const cutMode = isManual ? 'manual' : (isSmart ? 'smart' : 'protect');
                            const sourceScale = Number(previousCutMeta?.sourceScale) || 1;
                            const sourceY = hasPosition
                                ? Math.round(Number(positions[previousCutIndex]) || 0)
                                : Math.round((Number(logicalStart) || 0) * sourceScale);
                            const autoSourceY = Math.round((Number(previousCutMeta?.automaticBase) || Number(logicalStart) || 0) * sourceScale);
                            const line = document.createElement('div');
                            line.className = 'zujuanjs-answer-cut-line';
                            line.dataset.answerCutDrag = '1';
                            slice.appendChild(line);

                            const control = document.createElement('div');
                            control.className = 'zujuanjs-answer-cut-control';
                            control.dataset.answerImageKey = imageKey;
                            control.dataset.cutIndex = String(previousCutIndex);
                            control.dataset.cutMode = cutMode;
                            control.dataset.cutPosition = String(sourceY);
                            control.dataset.cutAutoPosition = String(autoSourceY);
                            control.dataset.sourceScale = String(sourceScale);
                            control.dataset.cutMinPosition = String(Math.max(0, Math.round((Number(previousCutMeta?.minimumCut) || 0) * sourceScale)));
                            control.dataset.cutMaxPosition = String(Math.max(sourceY + 1, Math.round((Number(previousCutMeta?.maximumCut) || Number(previousCutMeta?.displayHeight) || 99999) * sourceScale)));
                            const modeLabel = isManual ? '手动' : (isSmart ? '行间' : '保护');
                            control.innerHTML = '<span>' + modeLabel + ' · Y=' + sourceY + '</span>'
                                + '<button type="button" data-answer-cut-shift="-10" title="原图切线向上 10px">−10</button>'
                                + '<button type="button" data-answer-cut-shift="-1" title="原图切线向上 1px">−1</button>'
                                + '<label class="zujuanjs-answer-cut-numeric" title="原始答案图片的 Y 坐标；人工确认后下一页不再重复">Y'
                                + '<input type="number" data-answer-cut-value min="0" max="99999" step="1" value="' + sourceY + '" aria-label="原图切线 Y 坐标">px</label>'
                                + '<button type="button" data-answer-cut-shift="1" title="原图切线向下 1px">+1</button>'
                                + '<button type="button" data-answer-cut-shift="10" title="原图切线向下 10px">+10</button>'
                                + '<button type="button" data-answer-cut-snap="1" title="吸附到最近检测到的行间空白">吸附</button>'
                                + '<button type="button" data-answer-cut-drag-button="1" title="按住并上下拖动切线">↕</button>'
                                + '<button type="button" data-answer-cut-reset="1" title="恢复自动智能切线">复位</button>';
                            slice.appendChild(control);
                        }
                        return slice;
                    }

                    function measureFlowImage(image, pageContent) {
                        const ratio = clamp(Number(previewSettings.answerImageScale || 100) / 100, 0.5, 1, 1);
                        const naturalWidth = Number(image.naturalWidth) || 0;
                        const naturalHeight = Number(image.naturalHeight) || 0;
                        const displayWidth = Math.max(1, pageContent.clientWidth * ratio);
                        if (naturalWidth > 0 && naturalHeight > 0) {
                            return { displayWidth, displayHeight: displayWidth * naturalHeight / naturalWidth };
                        }
                        const measure = image.cloneNode(true);
                        measure.style.cssText = 'display:block;width:' + (ratio * 100) + '%;max-width:' + (ratio * 100) + '%;height:auto;margin:0;padding:0;border:0;';
                        pageContent.appendChild(measure);
                        const result = { displayWidth: measure.offsetWidth, displayHeight: measure.offsetHeight };
                        measure.remove();
                        return result;
                    }

                    function renderAnswerImageFlow(flowNode, page) {
                        const image = flowNode.querySelector('img');
                        if (!image) return page;
                        const imageKey = flowNode.dataset.answerImageKey || getAnswerImageKey(image);
                        const dimensions = measureFlowImage(image, page.content);
                        const displayWidth = dimensions.displayWidth;
                        const displayHeight = dimensions.displayHeight;
                        if (!displayWidth || !displayHeight) {
                            page.content.appendChild(flowNode);
                            return page;
                        }
                        const naturalHeight = Number(image.naturalHeight) || displayHeight;
                        const sourceScale = naturalHeight / displayHeight;

                        if (previewSettings.answerLongImageMode === 'fit') {
                            let remaining = getRemainingPageHeight(page.content);
                            const minimum = clamp(previewSettings.answerCutMinFill, 4, 160, 8);
                            if (remaining < minimum && pageHasContent(page.content)) {
                                page = createPaper();
                                remaining = getRemainingPageHeight(page.content);
                            }
                            const scale = Math.min(1, remaining / displayHeight);
                            const fitWidth = displayWidth * scale;
                            const fitHeight = displayHeight * scale;
                            page.content.appendChild(createAnswerImageSlice(image, fitWidth, fitHeight, 0, fitHeight, imageKey, -1, 0));
                            return page;
                        }

                        ensureAnswerImageAnalysis(image, imageKey);
                        const positions = normalizeManualCutPositions(imageKey);
                        const legacyOffsets = normalizeManualCutOffsets(imageKey);
                        const autoLead = clamp(previewSettings.answerCutAutoLead, 0, 80, 8);
                        const overlap = clamp(previewSettings.answerCutOverlap, 0, 80, 22);
                        const minimumFill = clamp(previewSettings.answerCutMinFill, 4, 160, 8);
                        let logicalStart = 0;
                        let cutIndex = 0;
                        let previousCutManual = true;
                        let previousCutSafe = true;
                        let previousCutMeta = null;
                        let guard = 0;

                        while (logicalStart < displayHeight - 1 && guard++ < 500) {
                            let remaining = getRemainingPageHeight(page.content);
                            // 8.1 填页策略：只在剩余空间真的小到连一行都不值得利用时才新建页面。
                            if (remaining < minimumFill && pageHasContent(page.content)) {
                                page = createPaper();
                                remaining = getRemainingPageHeight(page.content);
                            }
                            if (remaining < 3) {
                                page = createPaper();
                                remaining = getRemainingPageHeight(page.content);
                            }

                            const overlapProtectionEnabled = previewSettings.answerCutOverlapEnabled !== false;
                            const visualOverlap = !overlapProtectionEnabled || cutIndex === 0 || previousCutManual || previousCutSafe
                                ? 0
                                : Math.max(overlap, 52); // 开启时保持8.4.3的低置信度重叠保护；关闭时绝不重复切线附近内容。
                            const visualStart = Math.max(0, logicalStart - visualOverlap);
                            const maxEnd = Math.min(displayHeight, visualStart + Math.max(1, remaining - 0.5));
                            const previousCutIndex = cutIndex - 1;

                            if (maxEnd >= displayHeight - 0.5) {
                                page.content.appendChild(createAnswerImageSlice(image, displayWidth, displayHeight, visualStart, displayHeight, imageKey, previousCutIndex, logicalStart, previousCutMeta));
                                logicalStart = displayHeight;
                                break;
                            }

                            const minSegment = Math.min(48, Math.max(6, remaining * 0.05));
                            const automaticBase = Math.max(logicalStart + minSegment, maxEnd - autoLead);
                            const hasPosition = Object.prototype.hasOwnProperty.call(positions, cutIndex);
                            const hasLegacy = !hasPosition && Object.prototype.hasOwnProperty.call(legacyOffsets, cutIndex);
                            let cut = automaticBase;
                            let smartSafe = false;
                            let smartMeta = null;
                            if (hasPosition) {
                                cut = clamp(Number(positions[cutIndex]) / sourceScale, logicalStart + 2, maxEnd, automaticBase);
                            } else if (hasLegacy) {
                                cut = automaticBase + clamp(legacyOffsets[cutIndex], -500, 500, 0);
                            } else {
                                smartMeta = chooseSmartAnswerCut(imageKey, logicalStart, maxEnd, automaticBase, sourceScale);
                                if (smartMeta) {
                                    cut = smartMeta.cut;
                                    smartSafe = smartMeta.requiresOverlap !== true;
                                }
                            }
                            const minimumCut = logicalStart + 2;
                            const maximumCut = maxEnd;
                            cut = Math.max(minimumCut, Math.min(maximumCut, cut));
                            if (cut <= logicalStart + 1) cut = Math.min(maximumCut, logicalStart + Math.max(4, remaining * 0.5));

                            page.content.appendChild(createAnswerImageSlice(image, displayWidth, displayHeight, visualStart, cut, imageKey, previousCutIndex, logicalStart, previousCutMeta));
                            logicalStart = cut;
                            previousCutManual = hasPosition || hasLegacy;
                            previousCutSafe = smartSafe;
                            previousCutMeta = {
                                automaticBase,
                                manual: previousCutManual,
                                smartSafe,
                                smartMeta,
                                sourceScale,
                                displayHeight,
                                minimumCut,
                                maximumCut
                            };
                            cutIndex++;
                            page = createPaper();
                        }
                        return page;
                    }

                    function preserveQuestionIndent(fragment, original) {
                        const layout = fragment.querySelector('.zujuanjs-question-layout');
                        if (!layout || layout.querySelector('.zujuanjs-question-number')) return;
                        const originalNumber = original.querySelector('.zujuanjs-question-number');
                        if (!originalNumber) return;
                        const placeholder = originalNumber.cloneNode(true);
                        placeholder.classList.add('continuation-placeholder');
                        placeholder.setAttribute('aria-hidden', 'true');
                        layout.insertBefore(placeholder, layout.firstChild);
                    }

                    function childNodeIndexPath(root, target) {
                        const path = [];
                        let current = target;
                        while (current && current !== root) {
                            const parent = current.parentNode;
                            if (!parent) return null;
                            path.unshift(Array.prototype.indexOf.call(parent.childNodes, current));
                            current = parent;
                        }
                        return current === root ? path : null;
                    }

                    function nodeAtChildIndexPath(root, path) {
                        let current = root;
                        for (const index of path || []) {
                            if (!current || !current.childNodes || !current.childNodes[index]) return null;
                            current = current.childNodes[index];
                        }
                        return current;
                    }

                    function removeEmptyTableRowGroups(table) {
                        table.querySelectorAll('thead, tbody, tfoot').forEach(group => {
                            if (!group.querySelector('tr')) group.remove();
                        });
                    }

                    function trimTableRows(table, startRow, endRow) {
                        const rows = Array.from(table.rows || []);
                        for (let i = rows.length - 1; i >= 0; i--) {
                            if (i < startRow || i >= endRow) rows[i].remove();
                        }
                        removeEmptyTableRowGroups(table);
                    }

                    function pruneContentAfter(root, target) {
                        let current = target;
                        while (current && current !== root) {
                            let sibling = current.nextSibling;
                            while (sibling) {
                                const next = sibling.nextSibling;
                                sibling.remove();
                                sibling = next;
                            }
                            current = current.parentNode;
                        }
                    }

                    function pruneContentBefore(root, target) {
                        let current = target;
                        while (current && current !== root) {
                            let sibling = current.previousSibling;
                            while (sibling) {
                                const previous = sibling.previousSibling;
                                sibling.remove();
                                sibling = previous;
                            }
                            current = current.parentNode;
                        }
                    }

                    function splitQuestionAtTableRow(node, boundary) {
                        if (!boundary || boundary.type !== 'table-row' || !boundary.table) return null;
                        const path = childNodeIndexPath(node, boundary.table);
                        if (!path) return null;
                        const totalRows = Number(boundary.totalRows) || Array.from(boundary.table.rows || []).length;
                        const splitRowIndex = Number(boundary.rowIndex);
                        if (!Number.isInteger(splitRowIndex) || splitRowIndex <= 0 || splitRowIndex >= totalRows) return null;
                        if (tableBoundaryCrossesRowspan(boundary.table, splitRowIndex)) return null;

                        const first = node.cloneNode(true);
                        const second = node.cloneNode(true);
                        const firstTable = nodeAtChildIndexPath(first, path);
                        const secondTable = nodeAtChildIndexPath(second, path);
                        if (!(firstTable instanceof HTMLTableElement) || !(secondTable instanceof HTMLTableElement)) return null;

                        trimTableRows(firstTable, 0, splitRowIndex);
                        trimTableRows(secondTable, splitRowIndex, totalRows);
                        // caption 属于整张表的标题；续页不重复创造原卷没有的第二份标题。
                        secondTable.querySelectorAll(':scope > caption').forEach(caption => caption.remove());

                        pruneContentAfter(first, firstTable);
                        pruneContentBefore(second, secondTable);
                        preserveQuestionIndent(first, node);
                        preserveQuestionIndent(second, node);

                        first.classList.add('continues-on-next', 'zujuanjs-table-split-part');
                        second.classList.remove('continues-on-next');
                        second.classList.add('continued-from-previous', 'zujuanjs-table-split-part');
                        first.dataset.tableSplit = 'before-row-' + splitRowIndex;
                        second.dataset.tableSplit = 'from-row-' + splitRowIndex;

                        if (!hasPrintableContent(first) || !hasPrintableContent(second)) return null;
                        return { first, second, tableSplit: true };
                    }

                    function splitQuestionAt(node, boundary) {
                        if (!boundary) return null;
                        if (boundary.type === 'table-row') return splitQuestionAtTableRow(node, boundary);
                        const firstRange = document.createRange();
                        firstRange.selectNodeContents(node);
                        firstRange.setEnd(boundary.container, boundary.offset);

                        const secondRange = document.createRange();
                        secondRange.selectNodeContents(node);
                        secondRange.setStart(boundary.container, boundary.offset);

                        const first = node.cloneNode(false);
                        const second = node.cloneNode(false);
                        first.appendChild(firstRange.cloneContents());
                        second.appendChild(secondRange.cloneContents());
                        preserveQuestionIndent(first, node);
                        preserveQuestionIndent(second, node);

                        first.classList.add('continues-on-next');
                        second.classList.remove('continues-on-next');
                        second.classList.add('continued-from-previous');

                        if (!hasPrintableContent(first) || !hasPrintableContent(second)) return null;
                        return { first: first, second: second };
                    }

                    function addManualBreakMarker(pageContent) {
                        const marker = document.createElement('div');
                        marker.className = 'manual-break-indicator';
                        marker.textContent = '手动分页';
                        pageContent.appendChild(marker);
                    }

                    function updatePageNumbers() {
                        const papers = Array.from(paperContainer.querySelectorAll('.paper'));
                        const segmented = previewSettings.pageNumberScope === 'sectioned' && previewSettings.mode === 'qe';
                        const answerStartIndex = segmented
                            ? papers.findIndex(paper => Boolean(paper.querySelector('.zujuanjs-answer-section-title')))
                            : -1;

                        if (segmented && answerStartIndex >= 0) {
                            const questionTotal = answerStartIndex;
                            const answerTotal = papers.length - answerStartIndex;
                            papers.forEach((paper, index) => {
                                const isAnswerPage = index >= answerStartIndex;
                                const current = isAnswerPage ? index - answerStartIndex + 1 : index + 1;
                                const total = isAnswerPage ? answerTotal : questionTotal;
                                const footer = paper.querySelector('.page-footer');
                                // 分段计数只显示数字，不附加“试题/答案”文字。
                                if (previewSettings.pageNumberFormat === 'current') footer.textContent = String(current);
                                else if (previewSettings.pageNumberFormat === 'chinese') footer.textContent = '第 ' + current + ' 页，共 ' + total + ' 页';
                                else footer.textContent = current + ' / ' + total;
                                footer.style.display = previewSettings.showPageNumber === false ? 'none' : '';
                                paper.dataset.pageSection = isAnswerPage ? 'answer' : 'question';
                            });
                            const summary = [];
                            if (questionTotal > 0) summary.push(String(questionTotal));
                            if (answerTotal > 0) summary.push(String(answerTotal));
                            document.getElementById('preview-page-count').textContent = summary.length > 1
                                ? (summary.join(' + ') + ' = ' + papers.length + ' 页')
                                : ('共 ' + papers.length + ' 页');
                            return;
                        }

                        papers.forEach((paper, index) => {
                            const current = index + 1;
                            const total = papers.length;
                            const footer = paper.querySelector('.page-footer');
                            if (previewSettings.pageNumberFormat === 'current') footer.textContent = String(current);
                            else if (previewSettings.pageNumberFormat === 'chinese') footer.textContent = '第 ' + current + ' 页，共 ' + total + ' 页';
                            else footer.textContent = current + ' / ' + total;
                            footer.style.display = previewSettings.showPageNumber === false ? 'none' : '';
                            delete paper.dataset.pageSection;
                        });
                        document.getElementById('preview-page-count').textContent = '共 ' + papers.length + ' 页';
                    }

                    function renderPages() {
                        if (isRendering) {
                            renderAgain = true;
                            return;
                        }
                        isRendering = true;
                        const readingAnchor = pendingReadingAnchor || captureReadingAnchor();
                        pendingReadingAnchor = null;
                        paperContainer.textContent = '';

                        const queue = Array.from(sourceContent.children, child => child.cloneNode(true));
                        let page = createPaper();
                        let guard = 0;

                        while (queue.length && guard < 10000) {
                            guard++;
                            const node = queue.shift();

                            if (node.classList.contains('zujuanjs-answer-page-break')) {
                                if (previewSettings.answerStartNewPage && pageHasContent(page.content)) page = createPaper();
                                continue;
                            }

                            if (node.classList.contains('page-break')) {
                                if (pageHasContent(page.content)) {
                                    addManualBreakMarker(page.content);
                                    page = createPaper();
                                }
                                continue;
                            }

                            if (node.classList.contains('zujuanjs-section-title')) {
                                const pageAlreadyHasContent = pageHasContent(page.content);
                                page.content.appendChild(node);
                                const pageBottom = page.content.getBoundingClientRect().bottom;
                                const nodeBottom = node.getBoundingClientRect().bottom;
                                const followSpace = Math.max(72, clamp(parseFloat(previewSettings.size), 12, 30, 16) * 4.2);
                                if (nodeOverflowsPage(node, page.content) || (pageAlreadyHasContent && pageBottom - nodeBottom < followSpace)) {
                                    node.remove();
                                    if (pageAlreadyHasContent) page = createPaper();
                                    page.content.appendChild(node);
                                }
                                continue;
                            }

                            const expandedFlow = expandNodeAroundAnswerImages(node);
                            if (expandedFlow && expandedFlow.length) {
                                for (let index = expandedFlow.length - 1; index >= 0; index--) queue.unshift(expandedFlow[index]);
                                continue;
                            }
                            if (node.classList.contains('zujuanjs-answer-image-flow')) {
                                page = renderAnswerImageFlow(node, page);
                                continue;
                            }

                            const pageAlreadyHasContent = pageHasContent(page.content);
                            page.content.appendChild(node);
                            if (!nodeOverflowsPage(node, page.content)) continue;

                            let split = null;
                            if (node.classList.contains('q-wrapper') || node.classList.contains('zujuanjs-answer-content')) {
                                const pageBottom = page.content.getBoundingClientRect().bottom;
                                split = splitQuestionAt(node, findOverflowBoundary(node, pageBottom));
                            }

                            if (split) {
                                node.remove();
                                page.content.appendChild(split.first);
                                if (!nodeOverflowsPage(split.first, page.content)) {
                                    page = createPaper();
                                    queue.unshift(split.second);
                                    continue;
                                }
                                split.first.remove();
                            } else {
                                node.remove();
                            }

                            if (pageAlreadyHasContent) {
                                page = createPaper();
                                queue.unshift(node);
                            } else {
                                // 超过一整页且无法拆分的单个图片/表格，保留在本页以避免死循环。
                                page.content.appendChild(node);
                            }
                        }

                        const papers = Array.from(paperContainer.querySelectorAll('.paper'));
                        for (let i = papers.length - 1; i > 0; i--) {
                            const content = papers[i].querySelector('.paper-content');
                            if (pageHasContent(content)) break;
                            papers[i].closest('.paper-shell')?.remove();
                        }

                        updatePageNumbers();
                        renderVersion++;
                        paperContainer.dataset.renderVersion = String(renderVersion);
                        if (activeBlockId) {
                            paperContainer.querySelectorAll('.q-wrapper').forEach(fragment => {
                                fragment.classList.toggle('is-selected', fragment.dataset.blockId === activeBlockId);
                            });
                        }
                        bindRenderedQuestionInteractions();
                        restoreReadingAnchor(readingAnchor);
                        updateQuestionToolbarState();
                        scheduleQuestionToolbarPosition();
                        isRendering = false;
                        document.dispatchEvent(new CustomEvent('zujuan-preview-rendered', {
                            detail: { version: renderVersion, pages: paperContainer.querySelectorAll('.paper').length }
                        }));
                        if (renderAgain) {
                            renderAgain = false;
                            scheduleRender();
                        }
                    }

                    function scheduleRender() {
                        if (renderFrame) cancelAnimationFrame(renderFrame);
                        renderFrame = requestAnimationFrame(function() {
                            renderFrame = 0;
                            renderPages();
                        });
                    }


                    function waitForRenderAfter(version, timeout = 4500) {
                        return new Promise(resolve => {
                            let settled = false;
                            const finish = () => {
                                if (settled) return;
                                settled = true;
                                clearTimeout(timer);
                                document.removeEventListener('zujuan-preview-rendered', onRendered);
                                resolve();
                            };
                            const onRendered = event => {
                                if (Number(event.detail?.version || 0) > version) finish();
                            };
                            const timer = setTimeout(finish, timeout);
                            document.addEventListener('zujuan-preview-rendered', onRendered);
                        });
                    }

                    async function printWithPrecisionAnalysis() {
                        const button = document.getElementById('print-now');
                        if (button?.dataset.busy === '1') return;
                        if (button) {
                            button.dataset.busy = '1';
                            button.disabled = true;
                        }
                        try {
                            if (previewSettings.answerLongImageMode === 'split' && previewSettings.answerCutSmartSnap) {
                                setSaveStatus('正在完成全部答案的打印安全行间分析…');
                                await primeAnswerImageAnalyses();
                                const before = renderVersion;
                                scheduleRender();
                                await waitForRenderAfter(before);
                                // 再让浏览器完成一次布局/图片合成，避免刚重排就进入打印。
                                await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
                            }
                            setSaveStatus('打印排版已就绪');
                            window.print();
                        } catch (error) {
                            console.warn('打印前精准分析失败，使用当前排版继续打印：', error);
                            setSaveStatus('部分答案分析失败，已启用保护切线');
                            window.print();
                        } finally {
                            if (button) {
                                button.dataset.busy = '0';
                                button.disabled = false;
                            }
                        }
                    }

                    document.querySelectorAll('[data-layout]').forEach(button => {
                        button.addEventListener('click', () => setPreviewLayout(button.dataset.layout));
                    });
                    document.querySelector('[data-zoom="out"]').addEventListener('click', () => stepPreviewZoom(-1));
                    document.querySelector('[data-zoom="in"]').addEventListener('click', () => stepPreviewZoom(1));
                    document.querySelector('[data-zoom="auto"]').addEventListener('click', () => setPreviewZoom('auto'));
                    editorToggle.addEventListener('click', () => toggleEditor());
                    document.getElementById('editor-close').addEventListener('click', () => toggleEditor(false));
                    document.getElementById('export-html').addEventListener('click', exportStandaloneHtml);
                    document.getElementById('export-word').addEventListener('click', exportWordDocument);
                    document.getElementById('export-json').addEventListener('click', exportStructuredJson);
                    document.getElementById('print-now')?.addEventListener('click', printWithPrecisionAnalysis);
                    document.getElementById('quick-answer-end').addEventListener('click', () => {
                        const nextSettings = readSettingsFromEditor();
                        nextSettings.mode = previewSettings.mode === 'qe' ? 'q' : 'qe';
                        window.parent.postMessage({ type: 'rebuildZujuanPreview', settings: nextSettings }, '*');
                    });
                    window.addEventListener('message', event => {
                        if (!event.data || event.data.type !== 'zujuanWordExportResult') return;
                        const button = document.getElementById('export-word');
                        setExportButtonBusy(button, false, '处理中');
                        if (!event.data.ok) alert('Word 导出失败：' + (event.data.message || '未知错误'));
                        else if (event.data.unresolvedImages) console.warn('Word 导出中有 ' + event.data.unresolvedImages + ' 张图片未能内嵌，可能需要联网显示。');
                    });
                    updateQuickAnswerEndButton();

                    document.querySelectorAll('[data-editor-tab]').forEach(button => {
                        button.addEventListener('click', () => setEditorTab(button.dataset.editorTab));
                        button.addEventListener('keydown', event => {
                            if (!['ArrowLeft', 'ArrowRight'].includes(event.key)) return;
                            event.preventDefault();
                            const nextTab = button.dataset.editorTab === 'document' ? 'page' : 'document';
                            setEditorTab(nextTab);
                            document.querySelector('[data-editor-tab="' + nextTab + '"]').focus();
                        });
                    });

                    document.querySelectorAll('[data-layout-preset]').forEach(button => {
                        button.addEventListener('click', () => applyLayoutPreset(button.dataset.layoutPreset));
                    });

                    const editorControls = Array.from(editorPanel.querySelectorAll('input, select'));
                    editorControls.forEach(control => {
                        const eventName = control.tagName === 'INPUT' && ['text', 'number'].includes(control.type) ? 'input' : 'change';
                        control.addEventListener(eventName, () => {
                            previewSettings.layoutPreset = 'custom';
                            updateLayoutPresetUI();
                            setSaveStatus('正在保存…');
                            if (control.id === 'setting-answer-cut-overlap-enabled') {
                                const overlapInput = document.getElementById('setting-answer-cut-overlap');
                                if (overlapInput) overlapInput.disabled = !control.checked;
                            }
                            if (control.id === 'setting-page-number-scope' && control.value === 'sectioned' && document.getElementById('setting-mode').value === 'qe') {
                                const answerStart = document.getElementById('setting-answer-start-new-page');
                                if (answerStart && !answerStart.checked) answerStart.checked = true;
                            }
                            if (control.id === 'setting-mode') {
                                const nextSettings = readSettingsFromEditor();
                                previewSettings.mode = nextSettings.mode;
                                updateQuickAnswerEndButton();
                                window.parent.postMessage({ type: 'rebuildZujuanPreview', settings: nextSettings }, '*');
                                return;
                            }
                            if (control.id === 'setting-title') {
                                window.parent.postMessage({ type: 'setZujuanTitleOverride', title: control.value }, '*');
                            }
                            queueEditorSettings();
                        });
                    });
                    document.getElementById('editor-reset').addEventListener('click', () => {
                        applyLayoutPreset('exam');
                    });
                    document.getElementById('title-redetect')?.addEventListener('click', event => {
                        event.preventDefault();
                        restoreDetectedTitle();
                    });
                    document.getElementById('title-copy')?.addEventListener('click', async event => {
                        event.preventDefault();
                        const value = document.getElementById('setting-title')?.value || '';
                        try { await navigator.clipboard.writeText(value); setSaveStatus('标题已复制'); }
                        catch (error) { console.warn('复制标题失败', error); }
                    });
                    document.getElementById('editor-export-settings').addEventListener('click', exportPortableSettings);
                    document.getElementById('editor-import-settings').addEventListener('click', () => {
                        document.getElementById('editor-import-settings-file').click();
                    });
                    document.getElementById('editor-import-settings-file').addEventListener('change', event => {
                        const file = event.target.files && event.target.files[0];
                        importPortableSettings(file);
                        event.target.value = '';
                    });

                    let resizePointerId = null;
                    let resizeStartX = 0;
                    let resizeStartWidth = currentPanelWidth;
                    editorResizeHandle.addEventListener('pointerdown', event => {
                        if (window.innerWidth <= 680) return;
                        resizePointerId = event.pointerId;
                        resizeStartX = event.clientX;
                        resizeStartWidth = currentPanelWidth;
                        editorResizeHandle.setPointerCapture(event.pointerId);
                        document.body.classList.add('editor-resizing');
                        event.preventDefault();
                    });
                    editorResizeHandle.addEventListener('pointermove', event => {
                        if (event.pointerId !== resizePointerId) return;
                        applyEditorPanelWidth(resizeStartWidth + resizeStartX - event.clientX);
                    });
                    const finishPanelResize = event => {
                        if (event.pointerId !== resizePointerId) return;
                        resizePointerId = null;
                        document.body.classList.remove('editor-resizing');
                        applyEditorPanelWidth(currentPanelWidth, true);
                        scheduleSettingsSave();
                    };
                    editorResizeHandle.addEventListener('pointerup', finishPanelResize);
                    editorResizeHandle.addEventListener('pointercancel', finishPanelResize);
                    editorResizeHandle.addEventListener('keydown', event => {
                        if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
                        event.preventDefault();
                        if (event.key === 'Home') applyEditorPanelWidth(280, true);
                        else if (event.key === 'End') applyEditorPanelWidth(520, true);
                        else applyEditorPanelWidth(currentPanelWidth + (event.key === 'ArrowLeft' ? 10 : -10), true);
                        scheduleSettingsSave();
                    });

                    questionToolbar.addEventListener('click', event => {
                        const button = event.target.closest('[data-question-action]');
                        if (!button || button.disabled) return;
                        event.preventDefault();
                        performQuestionAction(button.dataset.questionAction);
                    });

                    function commitAnswerCutPosition(control, value) {
                        if (!control) return;
                        const imageKey = control.dataset.answerImageKey;
                        const cutIndex = Number(control.dataset.cutIndex);
                        if (!imageKey || !Number.isInteger(cutIndex) || cutIndex < 0) return;
                        const positions = normalizeManualCutPositions(imageKey);
                        const min = Math.max(0, Number(control.dataset.cutMinPosition) || 0);
                        const max = Math.max(min + 1, Number(control.dataset.cutMaxPosition) || 99999);
                        // 原图 Y 坐标是稳定坐标：改变预览缩放或答案显示宽度后仍可复用。
                        positions[cutIndex] = Math.max(min, Math.min(max, Math.round(Number(value) || 0)));
                        const legacyOffsets = normalizeManualCutOffsets(imageKey);
                        delete legacyOffsets[cutIndex];
                        if (!Object.keys(legacyOffsets).length) delete previewSettings.answerImageCutOffsets[imageKey];
                        setSaveStatus('正在保存…');
                        scheduleSettingsSave();
                        scheduleRender();
                    }

                    paperContainer.addEventListener('click', async event => {
                        const adjustButton = event.target.closest('[data-answer-cut-shift], [data-answer-cut-reset], [data-answer-cut-snap]');
                        if (!adjustButton) return;
                        const control = adjustButton.closest('.zujuanjs-answer-cut-control');
                        if (!control) return;
                        event.preventDefault();
                        event.stopPropagation();
                        const imageKey = control.dataset.answerImageKey;
                        const cutIndex = Number(control.dataset.cutIndex);
                        if (!imageKey || !Number.isInteger(cutIndex) || cutIndex < 0) return;
                        const positions = normalizeManualCutPositions(imageKey);
                        if (adjustButton.dataset.answerCutSnap) {
                            const current = Object.prototype.hasOwnProperty.call(positions, cutIndex)
                                ? Number(positions[cutIndex]) || 0
                                : Number(control.dataset.cutPosition) || Number(control.dataset.cutAutoPosition) || 0;
                            const sourceImage = Array.from(sourceContent.querySelectorAll('img.zujuanjs-answer-render-image'))
                                .find(image => getAnswerImageKey(image) === imageKey);
                            if (sourceImage) await ensureAnswerImageAnalysis(sourceImage, imageKey);
                            const minimum = Math.max(0, Number(control.dataset.cutMinPosition) || 0);
                            const maximum = Math.max(minimum + 1, Number(control.dataset.cutMaxPosition) || 99999);
                            const sourceScale = Math.max(0.01, Number(control.dataset.sourceScale) || 1);
                            const snapRadius = clamp(previewSettings.answerCutSearchWindow, 24, 180, 72) * sourceScale;
                            const safe = findNearestSafeSourceCut(
                                imageKey,
                                current,
                                Math.max(minimum, current - snapRadius),
                                Math.min(maximum, current + snapRadius)
                            );
                            if (safe) commitAnswerCutPosition(control, safe.y);
                            else setSaveStatus('当前切线附近未找到可靠行间空白');
                            return;
                        }
                        if (adjustButton.dataset.answerCutReset) {
                            delete positions[cutIndex];
                            if (!Object.keys(positions).length) delete previewSettings.answerImageCutPositions[imageKey];
                            const legacyOffsets = normalizeManualCutOffsets(imageKey);
                            delete legacyOffsets[cutIndex];
                            if (!Object.keys(legacyOffsets).length) delete previewSettings.answerImageCutOffsets[imageKey];
                            setSaveStatus('正在保存…');
                            scheduleSettingsSave();
                            scheduleRender();
                            return;
                        }
                        const shift = Number(adjustButton.dataset.answerCutShift) || 0;
                        const current = Object.prototype.hasOwnProperty.call(positions, cutIndex)
                            ? Number(positions[cutIndex]) || 0
                            : Number(control.dataset.cutPosition) || Number(control.dataset.cutAutoPosition) || 0;
                        commitAnswerCutPosition(control, current + shift);
                    });

                    paperContainer.addEventListener('change', event => {
                        const input = event.target.closest('[data-answer-cut-value]');
                        if (!input) return;
                        const control = input.closest('.zujuanjs-answer-cut-control');
                        if (!control) return;
                        event.stopPropagation();
                        commitAnswerCutPosition(control, input.value);
                    });

                    paperContainer.addEventListener('keydown', event => {
                        const input = event.target.closest('[data-answer-cut-value]');
                        if (!input || event.key !== 'Enter') return;
                        event.preventDefault();
                        const control = input.closest('.zujuanjs-answer-cut-control');
                        if (!control) return;
                        commitAnswerCutPosition(control, input.value);
                    });

                    let answerCutDragState = null;
                    paperContainer.addEventListener('pointerdown', event => {
                        const dragTarget = event.target.closest('[data-answer-cut-drag], [data-answer-cut-drag-button]');
                        if (!dragTarget) return;
                        const slice = dragTarget.closest('.zujuanjs-answer-image-slice');
                        const control = slice?.querySelector('.zujuanjs-answer-cut-control');
                        if (!control) return;
                        const imageKey = control.dataset.answerImageKey;
                        const cutIndex = Number(control.dataset.cutIndex);
                        if (!imageKey || !Number.isInteger(cutIndex)) return;
                        const positions = normalizeManualCutPositions(imageKey);
                        const current = Object.prototype.hasOwnProperty.call(positions, cutIndex)
                            ? Number(positions[cutIndex]) || 0
                            : Number(control.dataset.cutPosition) || Number(control.dataset.cutAutoPosition) || 0;
                        const sourceScale = Math.max(0.01, Number(control.dataset.sourceScale) || 1);
                        answerCutDragState = { pointerId: event.pointerId, startY: event.clientY, current, sourceScale, control };
                        dragTarget.setPointerCapture?.(event.pointerId);
                        event.preventDefault();
                        event.stopPropagation();
                    });
                    paperContainer.addEventListener('pointerup', event => {
                        if (!answerCutDragState || event.pointerId !== answerCutDragState.pointerId) return;
                        const delta = Math.round((event.clientY - answerCutDragState.startY) * answerCutDragState.sourceScale);
                        commitAnswerCutPosition(answerCutDragState.control, answerCutDragState.current + delta);
                        answerCutDragState = null;
                        event.preventDefault();
                    });
                    paperContainer.addEventListener('pointercancel', () => { answerCutDragState = null; });

                    paperContainer.addEventListener('click', event => {
                        if (event.target.closest('.zujuanjs-answer-cut-control, .zujuanjs-answer-cut-line')) return;
                        if (!event.target.closest('.q-wrapper')) hideQuestionToolbar();
                    });
                    pageViewport.addEventListener('scroll', scheduleQuestionToolbarPosition, { passive: true });
                    pageViewport.addEventListener('pointerdown', () => pageViewport.focus({ preventScroll: true }));
                    pageViewport.addEventListener('wheel', event => {
                        if (!event.ctrlKey && !event.metaKey) return;
                        event.preventDefault();
                        pageViewport.focus({ preventScroll: true });
                        zoomAtPoint(event.deltaY, event.clientX, event.clientY);
                    }, { passive: false });
                    let gestureStartScale = currentScale;
                    pageViewport.addEventListener('gesturestart', event => {
                        gestureStartScale = currentScale;
                        event.preventDefault();
                    }, { passive: false });
                    pageViewport.addEventListener('gesturechange', event => {
                        event.preventDefault();
                        const viewportRect = pageViewport.getBoundingClientRect();
                        const clientX = Number.isFinite(event.clientX) ? event.clientX : (viewportRect.left + viewportRect.right) / 2;
                        const clientY = Number.isFinite(event.clientY) ? event.clientY : (viewportRect.top + viewportRect.bottom) / 2;
                        setPreviewZoom(String(Math.max(0.25, Math.min(2, gestureStartScale * (Number(event.scale) || 1)))), { clientX, clientY });
                    }, { passive: false });
                    document.addEventListener('keydown', event => {
                        const typingTarget = event.target instanceof Element && event.target.closest('input, select, textarea');
                        if (event.key === 'Escape') {
                            if (activeBlockId) hideQuestionToolbar();
                            else if (!typingTarget) window.parent.postMessage({ type: 'closeZujuanPreview' }, '*');
                            return;
                        }
                        if (event.altKey && !event.ctrlKey && !event.metaKey && String(event.key).toLowerCase() === 'a' && !typingTarget) {
                            event.preventDefault();
                            document.getElementById('quick-answer-end').click();
                            return;
                        }
                        if ((!event.ctrlKey && !event.metaKey) || event.altKey || typingTarget) return;
                        if (['+', '=', 'Add'].includes(event.key)) {
                            event.preventDefault();
                            stepPreviewZoom(1);
                        } else if (['-', '_', 'Subtract'].includes(event.key)) {
                            event.preventDefault();
                            stepPreviewZoom(-1);
                        } else if (event.key === '0') {
                            event.preventDefault();
                            setPreviewZoom('auto');
                        }
                    });

                    window.addEventListener('load', scheduleRender);
                    window.addEventListener('beforeprint', renderPages);
                    window.addEventListener('resize', () => {
                        if (window.innerWidth <= 680 && editorOpen && activeBlockId) hideQuestionToolbar();
                        if (currentZoom === 'auto') applyPreviewView();
                        else scheduleQuestionToolbarPosition();
                    });
                    if (document.fonts && document.fonts.ready) {
                        document.fonts.ready.then(scheduleRender);
                    }

                    sourceContent.querySelectorAll('img').forEach(img => {
                        if (!img.complete) {
                            img.addEventListener('load', () => {
                                applyFormulaScale();
                                applyAnswerImageScale();
                                primeAnswerImageAnalyses();
                                scheduleRender();
                            }, { once: true });
                            img.addEventListener('error', scheduleRender, { once: true });
                        }
                    });
                    editorPanel.classList.add('editor-initializing');
                    document.body.classList.toggle('question-tools-open', Boolean(activeBlockId));
                    applyBlockEditsToSource();
                    writeSettingsToEditor();
                    renderPaperDiagnostics();
                    applyDocumentStyles();
                    primeAnswerImageAnalyses();
                    applyEditorPanelWidth(currentPanelWidth);
                    setEditorTab(currentEditorTab, false);
                    toggleEditor(editorOpen, false);
                    requestAnimationFrame(() => editorPanel.classList.remove('editor-initializing'));
                    scheduleRender();
                </script>
            </body>
            </html>
            `;
        }

        exportWordDocument(payload = {}, targetWindow = null) {
            const sendResult = (ok, message = '') => {
                try {
                    targetWindow?.postMessage({
                        type: 'zujuanWordExportResult',
                        ok,
                        message,
                        unresolvedImages: Number(payload.unresolvedImages) || 0
                    }, '*');
                } catch (error) {
                    console.warn('无法回传 Word 导出状态：', error);
                }
            };

            try {
                if (typeof htmlDocx === 'undefined' || !htmlDocx || typeof htmlDocx.asBlob !== 'function') {
                    throw new Error('Word 导出组件未加载，请刷新页面后重试');
                }
                const html = String(payload.html || '').trim();
                if (!html) throw new Error('没有可导出的试卷内容');

                const values = Array.isArray(payload.margins) ? payload.margins.map(Number) : [18, 15, 22, 15];
                const valid = values.length === 4 && values.every(value => Number.isFinite(value) && value >= 8 && value <= 55);
                const [top, right, bottom, left] = valid ? values : [18, 15, 22, 15];
                const mmToTwips = value => Math.round(value * 1440 / 25.4);
                const blob = htmlDocx.asBlob(html, {
                    orientation: 'portrait',
                    margins: {
                        top: mmToTwips(top),
                        right: mmToTwips(right),
                        bottom: mmToTwips(bottom),
                        left: mmToTwips(left),
                        header: 360,
                        footer: 360,
                        gutter: 0
                    }
                });

                const fileName = String(payload.fileName || '试卷.docx').replace(/[\\/:*?"<>|]+/g, '_');
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = fileName.toLowerCase().endsWith('.docx') ? fileName : fileName + '.docx';
                link.style.display = 'none';
                document.body.appendChild(link);
                link.click();
                link.remove();
                window.setTimeout(() => URL.revokeObjectURL(url), 2000);
                sendResult(true);
            } catch (error) {
                console.error('Word 导出失败：', error);
                const message = error && error.message ? error.message : String(error);
                try {
                    if (typeof GM_notification === 'function') GM_notification({ title: 'Word 导出失败', text: message, timeout: 5000 });
                } catch (notifyError) {
                    console.warn('通知失败：', notifyError);
                }
                sendResult(false, message);
            }
        }

        openPreview(htmlContent) {
            let overlay = document.getElementById('zujuanjs-preview-overlay');
            if (!overlay) {
                overlay = document.createElement('div');
                overlay.id = 'zujuanjs-preview-overlay';
                document.body.appendChild(overlay);
            }
            const iframe = document.createElement('iframe');
            iframe.title = '试卷排版预览';
            iframe.srcdoc = htmlContent;
            overlay.innerHTML = '';
            overlay.appendChild(iframe);
        }
    }

    const zujuanPrintWorkbench = new PaperPrinter();
    // 仅暴露只读诊断接口，方便站点改版时快速定位选择器问题。
    window.__ZUJUAN_PRINT_WORKBENCH__ = Object.freeze({
        version: '8.4.3',
        getTitleInfo: () => zujuanPrintWorkbench.getPaperTitleInfo(),
        getDiagnostics: () => zujuanPrintWorkbench.getPaperDiagnostics(),
        getPaperModel: (includeHtml = false) => zujuanPrintWorkbench.extractPaperModel({ includeHtml: Boolean(includeHtml) })
    });
})();
