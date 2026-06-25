const { createApp } = Vue;

localforage.config({ driver: localforage.INDEXEDDB, name: 'ME_Portal_DB', version: 1.0, storeName: 'port_documents' });

// ==========================================
// 1. THUẬT TOÁN DEBOUNCE (TỐI ƯU HIỆU NĂNG I/O)
// ==========================================
const debouncedFunctions = {};

function debouncedSaveData(key, data, context) {
    if (debouncedFunctions[key]) {
        clearTimeout(debouncedFunctions[key]);
    }
    debouncedFunctions[key] = setTimeout(() => {
        localforage.setItem(key, JSON.parse(JSON.stringify(data))).then(() => { 
            context.showSavedToast = true; 
            setTimeout(() => context.showSavedToast = false, 2000); 
        });
    }, 2000); // 2 giây
}

// ==========================================
// 2. ABSTRACT CLASS: XÂY DỰNG EXCEL CHUẨN ĐẦU RA
// ==========================================
class ExcelReportBuilder {
    constructor(creator, isPortrait = false) {
        this.wb = new ExcelJS.Workbook();
        this.wb.creator = creator;
        this.ws = this.wb.addWorksheet('Bao_Cao', { views: [{ showGridLines: false, pageBreakPreview: true }] });
        
        // Setup in ấn chuẩn A4
        this.ws.pageSetup.paperSize = 9; 
        this.ws.pageSetup.orientation = isPortrait ? 'portrait' : 'landscape';
        this.ws.pageSetup.fitToPage = true;
        this.ws.pageSetup.fitToWidth = 1;
        this.ws.pageSetup.fitToHeight = 0; 
        this.ws.pageSetup.horizontalCentered = true;
        
        this.borderThin = { 
            top: {style:'thin', color: {argb: 'FF000000'}}, left: {style:'thin', color: {argb: 'FF000000'}}, 
            bottom: {style:'thin', color: {argb: 'FF000000'}}, right: {style:'thin', color: {argb: 'FF000000'}} 
        };
        this.borderBotOnly = { bottom: {style:'thin', color: {argb: 'FF000000'}} };
        this.alignCenter = { vertical: 'middle', horizontal: 'center', wrapText: true };
    }

    setMargins(baseMargin) {
        this.ws.pageSetup.margins = { left: baseMargin, right: baseMargin, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 };
    }

    addLogo(logoData, width, height, bounds) {
        if (!logoData) return;
        const id = this.wb.addImage({ base64: logoData.base64, extension: logoData.ext });
        if (bounds) {
            this.ws.addImage(id, bounds);
        } else {
            this.ws.addImage(id, { tl: { col: 0.1, row: 0.1 }, ext: { width: width, height: height } });
        }
    }

    applyTableStyle(startR, endR, colCnt, size = 10) {
        for (let i = startR; i <= endR; i++) { 
            const row = this.ws.getRow(i);
            let maxLines = 1;
            for(let j = 1; j <= colCnt; j++) {
                const c = row.getCell(j);
                c.border = this.borderThin; 
                if(!c.alignment) c.alignment = this.alignCenter; 
                if(!c.font) c.font = { name: 'Times New Roman', size: size, color: {argb: 'FF000000'} }; 
                
                // Thuật toán Wrap-text & Auto-height cho Excel
                if (c.value && typeof c.value === 'string') {
                    const colWidth = this.ws.getColumn(j).width || 15;
                    const lines = c.value.split('\n');
                    let cellLines = 0;
                    lines.forEach(l => { cellLines += Math.ceil(l.length / (colWidth * 1.1)); });
                    if (cellLines > maxLines) maxLines = cellLines;
                }
            }
            const calculatedHeight = Math.max(18, maxLines * (size === 9 ? 12 : 14));
            if (!row.height || row.height < calculatedHeight) row.height = calculatedHeight;
        }
    }

    async download(fileName) {
        const buffer = await this.wb.xlsx.writeBuffer();
        saveAs(new Blob([buffer], { type: "application/octet-stream" }), fileName);
    }
}

// ==========================================
// 3. VUE COMPONENT: HEADER IN ẤN
// ==========================================
const PrintHeader = {
    template: '#print-header-template',
    props: {
        title: String,
        logo: String,
        bordered: { type: Boolean, default: false },
        showMeta: { type: Boolean, default: false },
        meta: { type: Object, default: () => ({}) }
    },
    emits: ['logo-upload']
};

createApp({
    components: {
        'print-header': PrintHeader
    },
    data() {
        return {
            currentTab: 'database', 
            lang: 'vi', 
            isAdmin: false, 
            showSavedToast: false,
            roundDecimal: false, 
            vatRate: 8,
            
            // DỮ LIỆU NHÂN VIÊN (MỚI BỔ SUNG)
            employees: [
                { id: 'ME-001', name: 'Phạm Thanh Khang', position: 'Trưởng phòng Cơ Điện', team: '', group: '', phone: '' },
                { id: 'ME-002', name: 'Nguyễn Công Văn', position: 'Kỹ sư Cơ Điện', team: '6', group: '', phone: '' },
                { id: 'ME-003', name: 'Trần Minh Thiện', position: 'Kỹ sư Cơ Điện', team: '6', group: '14', phone: '' },
                { id: 'ME-004', name: 'Nguyễn Văn Hai', position: 'Công nhân Cơ Điện', team: '6', group: '14', phone: '' },
                { id: 'ME-005', name: 'Ngô Khúc Khải', position: 'Công nhân Cơ Điện', team: '6', group: '14', phone: '' },
                { id: 'ME-006', name: 'Huỳnh Ngọc Hoàng Em', position: 'Công nhân Cơ Điện', team: '6', group: '14', phone: '' },
                { id: 'ME-007', name: 'Nguyễn Ngọc Sáu', position: 'Công nhân Cơ Điện', team: '6', group: '14', phone: '' },
                { id: 'ME-008', name: 'Dương Hữu Nghĩa', position: 'Công nhân Cơ Điện', team: '6', group: '14', phone: '' },
                { id: 'ME-009', name: 'Nguyễn Thành Phát', position: 'Kỹ sư Cơ Điện', team: '6', group: '15', phone: '' },
                { id: 'ME-010', name: 'Trần Nguyễn Duy', position: 'Công nhân Cơ Điện', team: '6', group: '15', phone: '' },
                { id: 'ME-011', name: 'Trần Minh Tân', position: 'Công nhân Cơ Điện', team: '6', group: '15', phone: '' },
                { id: 'ME-012', name: 'Lê Văn Nhị', position: 'Công nhân Cơ Điện', team: '6', group: '15', phone: '' },
                { id: 'ME-013', name: 'Quách Văn Tăng', position: 'Kỹ sư Cơ Điện', team: '6', group: '16', phone: '' },
                { id: 'ME-014', name: 'Phạm Trần Hoàng Giang', position: 'Kỹ sư Cơ Điện', team: '6', group: '16', phone: '' },
                { id: 'ME-015', name: 'Huỳnh Minh Vương', position: 'Kỹ sư Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-016', name: 'Phạm Hữu Tùng', position: 'Công nhân Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-017', name: 'Trần Phương Bằng', position: 'Công nhân Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-018', name: 'Lê Trần Hoàng Lộc', position: 'Công nhân Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-019', name: 'Nguyễn Đức Toàn', position: 'Kỹ sư Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-020', name: 'Nguyễn Lê Duy Thịnh', position: 'Công nhân Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-021', name: 'Trần Hoài Hận', position: 'Công nhân Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-022', name: 'Tô Hoài Thanh', position: 'Nhân viên Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-023', name: 'Hồ Minh Nhịnh', position: 'Công nhân Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-024', name: 'Huỳnh Tấn Khương', position: 'Công nhân Cơ Điện', team: '6', group: '17', phone: '' },
                { id: 'ME-025', name: 'Nguyễn Hải Tiến', position: 'Công nhân Cơ Điện', team: '6', group: '15', phone: '' },
                { id: 'ME-026', name: 'Trịnh Quốc Việt', position: 'Kỹ sư Cơ Điện', team: '7', group: '', phone: '' },
                { id: 'ME-027', name: 'Lâm Võ Thanh', position: 'Kỹ sư Cơ Điện', team: '7', group: '18', phone: '' },
                { id: 'ME-028', name: 'Hồ Trọng Hiếu', position: 'Nhân viên Cơ Điện', team: '7', group: '18', phone: '' },
                { id: 'ME-029', name: 'Nguyễn Văn Phi', position: 'Công nhân Cơ Điện', team: '7', group: '18', phone: '' },
                { id: 'ME-030', name: 'Nguyễn Tuấn An', position: 'Công nhân Cơ Điện', team: '7', group: '18', phone: '' },
                { id: 'ME-031', name: 'Đào Ngọc Toàn', position: 'Công nhân Cơ Điện', team: '7', group: '19', phone: '' },
                { id: 'ME-032', name: 'Ngô Long Hội', position: 'Kỹ sư Cơ Điện', team: '7', group: '19', phone: '' },
                { id: 'ME-033', name: 'Trương Tuấn Vũ', position: 'Kỹ sư Cơ Điện', team: '7', group: '19', phone: '' },
                { id: 'ME-034', name: 'Phạm Khánh Duy', position: 'Công nhân Cơ Điện', team: '7', group: '19', phone: '' },
                { id: 'ME-035', name: 'Lê Thanh Phong', position: 'Công nhân Cơ Điện', team: '7', group: '19', phone: '' },
                { id: 'ME-036', name: 'Lê Thanh Hải', position: 'Nhân viên Cơ Điện', team: '7', group: '19', phone: '' },
                { id: 'ME-037', name: 'Phạm Ngọc Diện', position: 'Kỹ sư Cơ Điện', team: '8', group: '', phone: '' },
            ],
            newEmployee: { id: '', name: '', position: '', team: '', group: '', phone: '' },

            database: [], 
            newItem: { partNo: '', name: '', specs: '', unit: '', priceVT: 0, priceNC: 0 },
            searchQuery: '', 
            showSearchDropdown: false, 
            companyLogo: 'https://raw.githubusercontent.com/Uniekey/BM/main/ADG.png', 
            globalProjectName: '',
            
            libraryLinks: [
                { title: 'Link sharepoint Tài liệu MMTB', titleEn: 'Port Equipment Documents', desc: 'Hồ sơ MMTB', descEn: 'Equipment Profiles', icon: 'fas fa-anchor', url: 'https://dongtam365.sharepoint.com/sites/ctycpptca/05%20Phng%20Dch%20V%20C%20in%20Cng/Forms/AllItems.aspx?id=%2Fsites%2Fctycpptca%2F05%20Phng%20Dch%20V%20C%20in%20Cng%2F5%2E7%2ET%C3%80I%20LI%E1%BB%86U%20%2D%20H%E1%BB%92%20S%C6%A0%20MMTB&viewid=ca8054c9%2D7c9c%2D4c5b%2Da730%2D3722d554e029' },
                { title: 'Hiệu chỉnh cân xe tải - Cảng', titleEn: 'Truck Scale Calibration', desc: 'Chỉnh góc - calib cân xe tải', descEn: 'Corner adjustment', icon: 'fas fa-balance-scale', url: 'https://uniekey.github.io/truck-scale/' },
                { title: 'Quy định/Quy trình/HDCV', titleEn: 'SOPs & Regulations', desc: 'Tài liệu nội bộ công ty', descEn: 'Internal Documents', icon: 'fas fa-envelope-open-text', url: '#' },
                { title: 'Tính toán điện NLMT', titleEn: 'Solar PV Calculator', desc: 'Tính PV, INV, BESS', descEn: 'Calculate PV, INV, BESS', icon: 'fas fa-search-dollar', url: 'https://uniekey.github.io/solar/' },
                { title: 'Dự phòng', titleEn: 'Spare', desc: 'Chưa biết ghi gì', descEn: 'Nothing to do', icon: 'fas fa-anchor', url: '#' },
                { title: 'Dự phòng', titleEn: 'Spare', desc: 'Chưa biết ghi gì', descEn: 'Nothing to do', icon: 'fas fa-anchor', url: '#' }
            ],

            docDuToan: [], 
            tyleNCGT: 0, 
            tyleQLCVT: 0, 
            tyleQLCNC: 0, 
            dtVatRate: 8,
            
            bgHeader: { 
                isoDoc: 'ADG_QĐ_KT_02/BM01', isoDate: '28/06/2022', isoRev: '01',
                company: '                CÔNG TY CỔ PHẦN PHÁT TRIỂN CHÂU Á', 
                address: '                Số 7, khu phố 6, Xã Bến Lức, Tỉnh Tây Ninh, Việt Nam', 
                taxCode: '                1100837040',
                ref: '0000/2026/ADG-LAIP', 
                date: new Date().toLocaleDateString('vi-VN'),
                cusCompany: 'CÔNG TY CỔ PHẦN QUẢN LÝ & KHAI THÁC CẢNG QUỐC TẾ LONG AN',
                cusAddress: 'Số 68 Đường Tỉnh 830, ấp Vĩnh Hòa, Xã Tân Tập, Tây Ninh, Việt Nam',
                cusTax: '1101809254'
            },
            docBaoGia: [], 
            bgTerms: [ 
                { title: 'Thời gian giao hàng:', content: 'Trong vòng 15 ngày làm việc kể từ ngày nhận được tạm ứng hợp đồng.' }, 
                { title: 'Thời gian bảo hành:', content: '12 tháng theo tiêu chuẩn của nhà sản xuất.' }, 
                { title: 'Thanh toán:', content: 'Lần 1: Tạm ứng 30%. Lần 2: 70% sau khi bàn giao nghiệm thu.' },
                { title: 'Hiệu lực báo giá:', content: '30 ngày kể từ ngày ban hành.' }
            ],
            
            printPrWorkflow: true,
            prHeader: { isoDate: '…/…/...', isoRev: '...', isoCheck: '…', prNo: '', isPlan: true, company: 'Công ty Cổ Phần Phát Triển Châu Á', dept: 'Phòng Cơ điện', deptCode: '', warehouse: '', storekeeper: '', approvedDate: '' },
            docPR: [], 
            prWorkflow: [ 
                { time: '', stepVi: '01. Lập phiếu ĐNMH', stepEn: '01. Create Purchase Requisition', person: 'Người lập phiếu', note: '' }, 
                { time: '', stepVi: '02. Xác nhận tồn kho', stepEn: '02. Confirm Inventory / Stock', person: 'Nhân viên Kế hoạch', note: '' }, 
                { time: '', stepVi: '03. Tổng hợp lại các mặt hàng cần mua/cấp', stepEn: '03. Consolidate Items', person: 'Nhân viên Kế hoạch', note: '' }, 
                { time: '', stepVi: '04. Kiểm tra phiếu đề nghị', stepEn: '04. Check PR Details', person: 'Trưởng phòng Cơ Điện', note: '' }, 
                { time: '', stepVi: '05. Xác nhận đơn hàng', stepEn: '05. Order Confirmation', person: 'Trưởng phòng Cung ứng', note: '' }, 
                { time: '', stepVi: '06. Kiểm tra ngân sách ĐNMH', stepEn: '06. Budget Verification', person: 'Kế toán trưởng', note: '' }, 
                { time: '', stepVi: '07. Phê duyệt đề nghị mua hàng', stepEn: '07. Final Approval', person: 'Ban Giám đốc', note: '' }, 
                { time: '', stepVi: '08. Tiếp nhận phiếu đề nghị', stepEn: '08. PR Received', person: 'Nhân viên Kế hoạch', note: '' } 
            ],
            
            dailyType: 'thicong', 
            dailyHeader: { date: '', location: '', pic: '' }, 
            dailyManpower: [], 
            dailyEquipment: [], 
            dailyTasks: [],
            
            ntHeader: { contract: '', date: '' }, 
            docNghiemThu: [],

            ntbgHeader: {
                project: 'V/v: Sửa chữa, bảo trì máy móc thiết bị',
                monthYear: `Tháng ${(new Date().getMonth() + 1).toString().padStart(2, '0')} /${new Date().getFullYear()}`,
                refNo: 'Số: .....................',
                benA_company: 'CÔNG TY CỔ PHẦN PHÁT TRIỂN CHÂU Á',
                benA_rep1_name: '', benA_rep1_title: '',
                benA_rep2_name: '', benA_rep2_title: '',
                benB_company: '',
                benB_rep1_name: '', benB_rep1_title: '',
                benB_rep2_name: '', benB_rep2_title: '',
                baseOn: '',
                conclusion: 'Sau khi kiểm tra thực tế khối lượng công việc, Hai Bên thống nhất nghiệm thu khối lượng công việc như sau:'
            },
            docNghiemThuBG: []
        }
    },
    computed: {
        dtSumVT() { 
            return this.docDuToan.filter(i => !i.isCategory).reduce((sum, item) => sum + ((item.qty || 0) * (item.priceVT || 0)), 0); 
        },
        dtSumNC() { 
            return this.docDuToan.filter(i => !i.isCategory).reduce((sum, item) => sum + ((item.qty || 0) * (item.priceNC || 0)), 0); 
        },
        dtCpNCGT() { return this.dtSumNC * (this.tyleNCGT / 100); }, 
        dtCpQLCVT() { return this.dtSumVT * (this.tyleQLCVT / 100); }, 
        dtCpQLCNC() { return (this.dtSumNC + this.dtCpNCGT) * (this.tyleQLCNC / 100); },
        
        dtTotalBeforeVAT() { 
            let t = this.dtSumVT + this.dtSumNC + this.dtCpNCGT + this.dtCpQLCVT + this.dtCpQLCNC; 
            return this.roundDecimal ? Math.round(t) : t; 
        },
        dtVATAmount() { 
            let v = this.dtTotalBeforeVAT * (this.dtVatRate / 100); 
            return this.roundDecimal ? Math.round(v) : v; 
        },
        dtTotalAfterVAT() { return this.dtTotalBeforeVAT + this.dtVATAmount; },

        bgTotalBeforeVAT() { 
            let t = this.docBaoGia.filter(i => !i.isCategory).reduce((sum, item) => sum + ((item.qty || 0) * (item.priceGop || 0)), 0); 
            return this.roundDecimal ? Math.round(t) : t; 
        },
        bgVATAmount() { 
            let v = this.bgTotalBeforeVAT * (this.vatRate / 100); 
            return this.roundDecimal ? Math.round(v) : v; 
        },
        bgTotalAfterVAT() { return this.bgTotalBeforeVAT + this.bgVATAmount; },

        prTotalAmount() { return this.docPR.reduce((sum, item) => sum + ((item.qtyReq || 0) * (item.price || 0)), 0); },
        filteredDatabase() { 
            if (!this.searchQuery) return []; 
            const q = this.searchQuery.toLowerCase(); 
            return this.database.filter(i => (i.partNo && i.partNo.toLowerCase().includes(q)) || (i.name && i.name.toLowerCase().includes(q))).slice(0, 30); 
        }
    },
    watch: {
        currentTab() { this.searchQuery = ''; },
        // Sử dụng debouncedSave thay vì ghi trực tiếp để bảo vệ CPU
	employees: { deep: true, handler(val) { debouncedSaveData('employees_db', val, this); } },
	globalProjectName(val) { debouncedSaveData('global_project_name', val, this); },
        docDuToan: { deep: true, handler(val) { debouncedSaveData('dt_table', val, this); } },
        docBaoGia: { deep: true, handler(val) { debouncedSaveData('bg_table', val); } }, 
        bgHeader: { deep: true, handler(val) { debouncedSaveData('bg_header', val); } },
        bgTerms: { deep: true, handler(val) { debouncedSaveData('bg_terms', val); } }, 
        docPR: { deep: true, handler(val) { debouncedSaveData('pr_table', val); } }, 
        prWorkflow: { deep: true, handler(val) { debouncedSaveData('pr_flow', val); } }, 
        docNghiemThu: { deep: true, handler(val) { debouncedSaveData('nt_table', val); } }, 
        prHeader: { deep: true, handler(val) { debouncedSaveData('pr_header', val); } }, 
        printPrWorkflow(val) { debouncedSaveData('pr_flow_print_setting', val); },
        vatRate(val) { debouncedSaveData('global_vat_rate', val); },
        dtVatRate(val) { debouncedSaveData('dt_vat_rate', val); },
        ntbgHeader: { deep: true, handler(val) { debouncedSaveData('ntbg_header', val); } },
        docNghiemThuBG: { deep: true, handler(val) { debouncedSaveData('ntbg_table', val); } }
    },
    async created() {
        // Cấu hình lưu nháp tự động sau khi NGỪNG gõ phím 2.5 giây
        this.debouncedSave = debounce((key, data) => {
            localforage.setItem(key, JSON.parse(JSON.stringify(data))).then(() => { 
                this.showSavedToast = true; 
                setTimeout(() => this.showSavedToast = false, 2000); 
            });
        }, 2500);
    },
    async mounted() {
        await this.migrateDataToIndexedDB(); 
        await this.loadDrafts(); 
        const db = await localforage.getItem('masterDataProject'); 
        if (db) this.database = db;
    },
    methods: {
        t(vi, en) { return this.lang === 'vi' ? vi : en; },
        getBase64Data(dataUrl) {
            if(!dataUrl || !dataUrl.startsWith('data:image')) return null;
            const arr = dataUrl.split(','); if(arr.length < 2) return null;
            const mime = arr[0].match(/:(.*?);/)[1]; const ext = mime.split('/')[1] || 'png';
            return { base64: arr[1], ext: ext };
        },

        toRoman(num) {
            if (!num || num === 0) return '';
            const roman = { M: 1000, CM: 900, D: 500, CD: 400, C: 100, XC: 90, L: 50, XL: 40, X: 10, IX: 9, V: 5, IV: 4, I: 1 };
            let str = '';
            for (let i of Object.keys(roman)) {
                let q = Math.floor(num / roman[i]);
                num -= q * roman[i];
                str += i.repeat(q);
            }
            return str;
        },

        getRowIndex(idx) {
            let catCount = 0; let itemCount = 0;
            for (let i = 0; i <= idx; i++) {
                if (this.docDuToan[i].isCategory) { catCount++; itemCount = 0; } else { itemCount++; }
            }
            if (this.docDuToan[idx].isCategory) return this.toRoman(catCount);
            return itemCount;
        },

        getCategoryTotal(index) {
            let vt = 0, nc = 0;
            for (let i = index + 1; i < this.docDuToan.length; i++) {
                if (this.docDuToan[i].isCategory) break; 
                const r = this.docDuToan[i];
                vt += ((r.qty || 0) * (r.priceVT || 0)); nc += ((r.qty || 0) * (r.priceNC || 0));
            }
            return { vt, nc, total: vt + nc };
        },

        getRowIndexBG(idx) {
            let catCount = 0; let itemCount = 0;
            for (let i = 0; i <= idx; i++) {
                if (this.docBaoGia[i].isCategory) { catCount++; itemCount = 0; } else { itemCount++; }
            }
            if (this.docBaoGia[idx].isCategory) return this.toRoman(catCount);
            return itemCount;
        },

        getBGCategoryTotal(index) {
            let total = 0;
            for (let i = index + 1; i < this.docBaoGia.length; i++) {
                if (this.docBaoGia[i].isCategory) break; 
                const r = this.docBaoGia[i];
                total += ((r.qty || 0) * (r.priceGop || 0));
            }
            return total;
        },

        async clearAllCache() {
            if (confirm(this.t("CẢNH BÁO: Hành động này sẽ xóa TOÀN BỘ dữ liệu lưu nháp và Kho vật tư. Trang sẽ được tải lại. Bạn chắc chắn chứ?", "WARNING: This will clear ALL cached data, drafts, and Master Data. The page will reload. Are you sure?"))) {
                try { await localforage.clear(); localStorage.clear(); window.location.reload(); } 
                catch (err) { alert("Có lỗi xảy ra khi xóa cache!"); }
            }
        },

        // QUẢN LÝ NHÂN VIÊN (MỚI)
        addEmployee() {
            if (!this.newEmployee.name || !this.newEmployee.id) {
                return alert(this.t("Vui lòng nhập Họ tên và MSNV", "Please enter Full Name and Emp ID"));
            }
            this.employees.unshift({...this.newEmployee});
            this.newEmployee = { id: '', name: '', position: '', team: '', group: '', phone: '' };
        },

        async exportToExcel(type) {
            if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
            this.showSavedToast = true;

            const str = (v) => v ? v.toString() : '';
            const num = (v) => Number(v) || 0;
            const projectName = this.globalProjectName || "Du_An_ME";
            const dateStr = new Date().toLocaleDateString('vi-VN').replace(/\//g, '-');
            let fileName = `Bao_Cao_${projectName}_${dateStr}.xlsx`;

            // Khởi tạo Excel bằng Lớp (Abstract Class)
            const builder = new ExcelReportBuilder('M&E Portal');
            const ws = builder.ws;
            const FORMAT_MONEY = this.roundDecimal ? '#,##0' : '#,##0.000'; 
            const logoData = this.getBase64Data(this.companyLogo);

            /* ========================= TAB NHÂN VIÊN (MỚI) ========================= */
            if (type === 'employees') {
                fileName = `Danh_Sach_Nhan_Vien_${dateStr}.xlsx`;
                ws.columns = [ {width: 6}, {width: 15}, {width: 30}, {width: 25}, {width: 15}, {width: 15}, {width: 20} ];
                ws.pageSetup.printTitlesRow = '1:6';

                builder.addLogo(logoData, 140, 50);
                
                ws.mergeCells('A1:G1'); ws.getCell('A1').value = this.t('DANH SÁCH NHÂN VIÊN', 'EMPLOYEE LIST'); ws.getCell('A1').font = { bold: true, name: 'Times New Roman', size: 16 }; ws.getCell('A1').alignment = builder.alignCenter; ws.getRow(1).height = 30;
                ws.addRow([]); ws.addRow([]); ws.addRow([]); ws.addRow([]);
                
                const head = ws.addRow([this.t('STT', 'No.'), this.t('MSNV', 'Emp ID'), this.t('Họ và tên', 'Full Name'), this.t('Vị trí công việc', 'Position'), this.t('Tổ', 'Team'), this.t('Nhóm', 'Group'), this.t('SĐT', 'Phone')]);
                head.eachCell(c => { c.font={ bold: true, name: 'Times New Roman', size: 11 }; c.fill={ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }; c.border=builder.borderThin; c.alignment=builder.alignCenter; });
                
                let rIdx = 7;
                this.employees.forEach((emp, i) => {
                    const r = ws.addRow([i + 1, str(emp.id), str(emp.name), str(emp.position), str(emp.team), str(emp.group), str(emp.phone)]);
                    r.getCell(1).alignment = builder.alignCenter; r.getCell(2).alignment = builder.alignCenter;
                    r.getCell(3).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                    r.getCell(4).alignment = builder.alignCenter; r.getCell(5).alignment = builder.alignCenter;
                    r.getCell(6).alignment = builder.alignCenter; r.getCell(7).alignment = builder.alignCenter;
                    rIdx++;
                });
                builder.applyTableStyle(6, rIdx-1, 7, 11);
            }

            /* ========================= TAB DỰ TOÁN (BOM) ========================= */
            else if (type === 'dutoan') {
                fileName = `BOM_${projectName}_${dateStr}.xlsx`;
                ws.columns = [ {width: 5}, {width: 35}, {width: 15}, {width: 6}, {width: 6}, {width: 12}, {width: 12}, {width: 12}, {width: 12}, {width: 15} ];
                ws.pageSetup.printTitlesRow = '1:6';

                builder.addLogo(logoData, 140, 50);
                
                ws.mergeCells('A1:J1'); ws.getCell('A1').value = this.t('BẢNG DỰ TOÁN CHI PHÍ (BOM)', 'BILL OF MATERIALS (BOM)'); ws.getCell('A1').font = { bold: true, name: 'Times New Roman', size: 16 }; ws.getCell('A1').alignment = builder.alignCenter; ws.getRow(1).height = 30;
                ws.mergeCells('A2:J2'); ws.getCell('A2').value = `${this.t('Dự án:', 'Project:')} ${str(projectName)}`; ws.getCell('A2').font = { bold: true, name: 'Times New Roman', size: 11 }; ws.getCell('A2').alignment = builder.alignCenter; ws.getRow(2).height = 20;
                ws.mergeCells('A3:J3'); ws.getCell('A3').value = this.t('', ''); ws.getCell('A3').font = {italic:true, name:'Times New Roman', size:10}; ws.getCell('A3').alignment = builder.alignCenter;
                ws.addRow([]);
                
                ws.addRow([this.t('STT', 'No.'), this.t('Hạng mục', 'Description'), this.t('Quy cách', 'Specs'), this.t('ĐVT', 'Unit'), this.t('SL', 'Qty'), this.t('Đơn giá (VNĐ)', 'Unit Price (VND)'), '', this.t('Thành tiền (VNĐ)', 'Total Amount (VND)'), '', this.t('Tổng cộng (VNĐ)', 'Grand Total')]);
                ws.addRow(['', '', '', '', '', this.t('Vật tư', 'Material'), this.t('Nhân công', 'Labor'), this.t('Vật tư', 'Material'), this.t('Nhân công', 'Labor'), '']);
                ['A5:A6', 'B5:B6', 'C5:C6', 'D5:D6', 'E5:E6', 'F5:G5', 'H5:I5', 'J5:J6'].forEach(m => ws.mergeCells(m));
                [5, 6].forEach(r => ws.getRow(r).eachCell(c => { c.font = { bold: true, name: 'Times New Roman', size: 10 }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }; c.border = builder.borderThin; c.alignment = builder.alignCenter; }));
                
                let rIdx = 7;
                this.docDuToan.forEach((row, i) => {
                    const r = ws.addRow([]);
                    if (row.isCategory) {
                        r.getCell(1).value = this.getRowIndex(i); 
                        r.getCell(2).value = str(row.name).toUpperCase();
                        r.getCell(8).value = this.getCategoryTotal(i).vt;
                        r.getCell(9).value = this.getCategoryTotal(i).nc;
                        r.getCell(10).value = this.getCategoryTotal(i).total;
                        
                        ws.mergeCells(`B${rIdx}:G${rIdx}`);
                        r.getCell(1).alignment = builder.alignCenter; r.getCell(1).font = { bold: true, name: 'Times New Roman', size: 10, color: {argb: 'FF0284C7'} };
                        r.getCell(2).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; r.getCell(2).font = { bold: true, name: 'Times New Roman', size: 10, color: {argb: 'FF0284C7'} }; 
                        r.getCell(8).font = { bold: true, name: 'Times New Roman', size: 10 }; r.getCell(9).font = { bold: true, name: 'Times New Roman', size: 10 }; r.getCell(10).font = { bold: true, name: 'Times New Roman', size: 10, color: {argb: 'FFDC2626'} };
                        r.getCell(8).numFmt = FORMAT_MONEY; r.getCell(9).numFmt = FORMAT_MONEY; r.getCell(10).numFmt = FORMAT_MONEY;
                        r.getCell(8).alignment = { vertical: 'middle', horizontal: 'right' }; r.getCell(9).alignment = { vertical: 'middle', horizontal: 'right' }; r.getCell(10).alignment = { vertical: 'middle', horizontal: 'right' };
                        r.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
                    } else {
                        r.getCell(1).value = this.getRowIndex(i); 
                        r.getCell(2).value = str(row.name); r.getCell(3).value = str(row.specs); r.getCell(4).value = str(row.unit);
                        let q = num(row.qty); let vt = num(row.priceVT); let nc = num(row.priceNC);
                        if(this.roundDecimal) { q=Math.round(q); vt=Math.round(vt); nc=Math.round(nc); }
                        r.getCell(5).value = q; r.getCell(6).value = vt; r.getCell(7).value = nc; r.getCell(8).value = q * vt; r.getCell(9).value = q * nc; r.getCell(10).value = (q * vt) + (q * nc);
                        
                        r.getCell(2).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; r.getCell(3).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; 
                        for(let c=5; c<=10; c++) { r.getCell(c).numFmt=FORMAT_MONEY; r.getCell(c).alignment={ vertical: 'middle', horizontal: 'right' }; } 
                        r.getCell(10).font = { bold: true, name: 'Times New Roman', size: 10, color: {argb: 'FFDC2626'} };
                    }
                    rIdx++;
                });
                
                const rT = ws.addRow([]);
                rT.getCell(8).value = this.t('Tổng chi phí trực tiếp:', 'Subtotal Direct Cost:'); rT.getCell(9).value = this.roundDecimal ? Math.round(this.dtSumVT) : this.dtSumVT; rT.getCell(10).value = this.roundDecimal ? Math.round(this.dtSumNC) : this.dtSumNC;
                ws.mergeCells(`A${rIdx}:H${rIdx}`); rT.font={ bold: true, name: 'Times New Roman', size: 10 }; rT.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; rT.getCell(9).numFmt=FORMAT_MONEY; rT.getCell(10).numFmt=FORMAT_MONEY;
                builder.applyTableStyle(5, rIdx, 10);
                
                ws.addRow([]); rIdx++;
                const sumTitle = ws.addRow([this.t('TỔNG HỢP CHI PHÍ & LỢI NHUẬN', 'COST & PROFIT SUMMARY')]);
                ws.mergeCells(`A${rIdx}:J${rIdx}`); sumTitle.font = {bold:true, name:'Times New Roman', size:11}; sumTitle.alignment = builder.alignCenter; sumTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }; sumTitle.border = builder.borderThin;
                rIdx++;

                const startSumR = rIdx;
                const sumData = [
                    { t: this.t(`1. TỔNG CHI PHÍ TRỰC TIẾP (Vật tư + Nhân công):`, `1. TOTAL BASE COST (Mat. + Lab.):`), v: this.dtSumVT + this.dtSumNC },
                    { t: this.t(`2. Chi phí Quản lý (${this.tyleNCGT}% NC)`, `2. Management Cost (${this.tyleNCGT}% Lab)`), v: this.dtCpNCGT },
                    { t: this.t(`3. LN định mức Vật tư (${this.tyleQLCVT}% VT)`, `3. Mat. Profit Margin (${this.tyleQLCVT}% Mat)`), v: this.dtCpQLCVT },
                    { t: this.t(`4. LN định mức Nhân công (${this.tyleQLCNC}% NC+QL)`, `4. Lab. Profit Margin (${this.tyleQLCNC}% Lab)`), v: this.dtCpQLCNC },
                    { t: this.t(`TỔNG DỰ TOÁN (TRƯỚC THUẾ):`, `ESTIMATE BEFORE TAX:`), v: this.dtTotalBeforeVAT, bold: true },
                    { t: this.t(`Thuế GTGT (VAT ${this.dtVatRate}%):`, `VAT (${this.dtVatRate}%):`), v: this.dtVATAmount },
                    { t: this.t(`TỔNG GIÁ TRỊ (SAU THUẾ):`, `TOTAL CONTRACT VALUE:`), v: this.dtTotalAfterVAT, bold: true, color: 'FFDC2626', size: 12 },
                ];
                
                sumData.forEach(d => {
                    const r = ws.addRow([]); r.getCell(1).value = d.t; r.getCell(9).value = this.roundDecimal ? Math.round(d.v) : d.v;
                    ws.mergeCells(`A${rIdx}:H${rIdx}`); ws.mergeCells(`I${rIdx}:J${rIdx}`); 
                    r.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; r.getCell(9).numFmt=FORMAT_MONEY; r.getCell(9).alignment={ vertical: 'middle', horizontal: 'right' };
                    if(d.bold) { r.font = {bold:true, name:'Times New Roman', size: d.size||10, color: d.color ? {argb: d.color} : undefined}; } 
                    rIdx++;
                });
                
                const rWords = ws.addRow([`${this.t('Bằng chữ:', 'In words:')} ${this.readMoney(this.dtTotalAfterVAT)}`]); ws.mergeCells(`A${rIdx}:J${rIdx}`); 
                rWords.font = {italic:true, name:'Times New Roman', size:10}; rWords.getCell(1).alignment={ vertical: 'middle', horizontal: 'left' }; rIdx++;

                builder.applyTableStyle(startSumR, rIdx-1, 10);
                
                ws.addRow([]); rIdx++; 
                const signR = ws.addRow([]);
                signR.getCell(2).value = this.t('NGƯỜI LẬP BẢNG\n\n(Ký và ghi rõ họ tên)', 'PREPARED BY\n\n(Signature & Full Name)'); 
                signR.getCell(8).value = this.t('NGƯỜI PHÊ DUYỆT\n\n(Ký và ghi rõ họ tên)', 'APPROVED BY\n\n(Signature & Full Name)');
                ws.mergeCells(`B${rIdx}:D${rIdx}`); ws.mergeCells(`H${rIdx}:J${rIdx}`); 
                signR.font={ bold: true, name: 'Times New Roman', size: 10 }; signR.alignment={ vertical: 'top', horizontal: 'center', wrapText: true }; signR.height = 100;
            }
            
            /* ========================= TAB BÁO GIÁ ========================= */
            else if (type === 'baogia') {
                fileName = `Bao_Gia_${projectName}_${dateStr}.xlsx`;
                ws.columns = [ {width: 5}, {width: 40}, {width: 25}, {width: 8}, {width: 8}, {width: 20}, {width: 22} ];
                ws.pageSetup.printTitlesRow = '1:18';

                let r1 = ws.addRow([]); 
                r1.getCell(1).value = this.t('BẢNG BÁO GIÁ', 'QUOTATION'); r1.getCell(1).font={bold:true, name:'Times New Roman', size:18}; r1.getCell(1).alignment=builder.alignCenter; r1.height = 25;
                r1.getCell(6).value = `${this.t('Số hiệu:', 'Doc No:')} ${str(this.bgHeader.isoDoc)}`; r1.getCell(6).font={ bold: true, name: 'Times New Roman', size: 10 }; r1.getCell(6).alignment = { vertical: 'middle', horizontal: 'left' };
                let r2 = ws.addRow([]); r2.getCell(1).value = str(projectName).toUpperCase(); r2.getCell(1).font={bold:true, name:'Times New Roman', size:12}; r2.getCell(1).alignment=builder.alignCenter; r2.height = 20;
                r2.getCell(6).value = `${this.t('Ngày hiệu lực:', 'Effective:')} ${str(this.bgHeader.isoDate)}\n${this.t('Lần ban hành:', 'Issue No:')} ${str(this.bgHeader.isoRev)}`; r2.getCell(6).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                ws.mergeCells('A1:E1'); ws.mergeCells('F1:G1'); ws.mergeCells('A2:E3'); ws.mergeCells('F2:G3'); 
                builder.applyTableStyle(1, 3, 7); ws.addRow([]); ws.addRow([]);

                let r6 = ws.addRow([]); 
                r6.getCell(1).value = this.t('Công ty:', 'Company:'); r6.getCell(2).value = str(this.bgHeader.company); 
                r6.getCell(6).value = this.t('Số:', 'Ref No:'); r6.getCell(7).value = str(this.bgHeader.ref); 
                ws.mergeCells('B6:E6'); r6.getCell(1).font={ bold: true, name: 'Times New Roman', size: 10 }; r6.getCell(2).font={ bold: true, name: 'Times New Roman', size: 10 }; r6.getCell(6).font={ bold: true, name: 'Times New Roman', size: 10 }; r6.getCell(7).font={bold:true, color:{argb:'FFDC2626'}, name:'Times New Roman', size: 10};
                let r7 = ws.addRow([]); 
                r7.getCell(1).value = this.t('Địa chỉ:', 'Address:'); r7.getCell(2).value = str(this.bgHeader.address); 
                r7.getCell(6).value = this.t('Ngày:', 'Date:'); r7.getCell(7).value = str(this.bgHeader.date); 
                ws.mergeCells('B7:E7'); r7.getCell(1).font={ bold: true, name: 'Times New Roman', size: 10 }; r7.getCell(6).font={ bold: true, name: 'Times New Roman', size: 10 };
                let r8 = ws.addRow([]); 
                r8.getCell(1).value = this.t('MST:', 'Tax Code:'); r8.getCell(2).value = str(this.bgHeader.taxCode); ws.mergeCells('B8:E8'); r8.getCell(1).font={ bold: true, name: 'Times New Roman', size: 10 }; ws.addRow([]);

                ws.addRow([this.t('Kính gửi Quý Khách hàng:', 'To our valued Customer:')]).font={ bold: true, name: 'Times New Roman', size: 10 };
                const cusArr = [ [this.t('Công ty/ Cá nhân:', 'Company/ Individual:'), this.bgHeader.cusCompany], [this.t('Địa chỉ:', 'Address:'), this.bgHeader.cusAddress], [this.t('MST:', 'Tax Code:'), this.bgHeader.cusTax] ];
                let startCus = 11;
                cusArr.forEach(c => {
                    let r = ws.addRow([]); r.getCell(1).value = c[0]; r.getCell(2).value = str(c[1]); ws.mergeCells(`B${startCus}:G${startCus}`);
                    r.getCell(1).font={ bold: true, name: 'Times New Roman', size: 10 }; r.getCell(2).font = (c[0] === this.t('Công ty/ Cá nhân:', 'Company/ Individual:')) ? { bold: true, name: 'Times New Roman', size: 10 } : { name: 'Times New Roman', size: 10 };
                    startCus++;
                });

                ws.addRow([this.t('Công ty chúng tôi trân trọng báo giá đến quý khách như sau:', 'We are pleased to submit our quotation as follows:')]).font = {italic:true, name:'Times New Roman', size:10};
                
                const head = ws.addRow([this.t('STT', 'No.'), this.t('Hạng mục', 'Item'), this.t('Thông số kỹ thuật', 'Specifications'), this.t('Đvt', 'Unit'), this.t('Số lượng', 'Qty'), this.t('Đơn giá', 'Price'), this.t('Thành tiền', 'Total Amount')]);
                head.eachCell(c => { c.font={ bold: true, name: 'Times New Roman', size: 10 }; c.fill={ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }; c.border=builder.borderThin; c.alignment=builder.alignCenter; });
                head.height = 25; let rIdx = startCus + 1;
                
                this.docBaoGia.forEach((row, i) => {
                    const r = ws.addRow([]);
                    if (row.isCategory) {
                        r.getCell(1).value = this.getRowIndexBG(i); r.getCell(2).value = str(row.name).toUpperCase(); r.getCell(7).value = this.getBGCategoryTotal(i);
                        ws.mergeCells(`B${rIdx}:F${rIdx}`);
                        r.getCell(1).alignment = builder.alignCenter; r.getCell(1).font = { bold: true, name: 'Times New Roman', size: 10, color: {argb: 'FF0284C7'} };
                        r.getCell(2).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }; r.getCell(2).font = { bold: true, name: 'Times New Roman', size: 10, color: {argb: 'FF0284C7'} }; 
                        r.getCell(7).font = { bold: true, name: 'Times New Roman', size: 10, color: {argb: 'FFDC2626'} }; r.getCell(7).numFmt = FORMAT_MONEY; r.getCell(7).alignment = { vertical: 'middle', horizontal: 'right' };
                        r.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
                    } else {
                        r.getCell(1).value = this.getRowIndexBG(i); r.getCell(2).value = str(row.name); r.getCell(3).value = str(row.specs); r.getCell(4).value = str(row.unit);
                        let q = num(row.qty); let p = num(row.priceGop); if(this.roundDecimal){ q=Math.round(q); p=Math.round(p); }
                        r.getCell(5).value = q; r.getCell(6).value = p; r.getCell(7).value = q * p;
                        r.getCell(2).alignment={ vertical: 'middle', horizontal: 'left', wrapText: true }; r.getCell(3).alignment={ vertical: 'middle', horizontal: 'left', wrapText: true }; 
                        r.getCell(5).numFmt=FORMAT_MONEY; r.getCell(6).numFmt=FORMAT_MONEY; r.getCell(7).numFmt=FORMAT_MONEY; 
                    }
                    rIdx++;
                });
                
                const rT1 = ws.addRow([]); rT1.getCell(6).value = this.t('Tổng giá (trước thuế):', 'Subtotal:'); rT1.getCell(7).value = this.bgTotalBeforeVAT; ws.mergeCells(`A${rIdx}:E${rIdx}`); rT1.font={ bold: true, name: 'Times New Roman', size: 10 }; rT1.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; rT1.getCell(7).numFmt=FORMAT_MONEY; rIdx++;
                const rT2 = ws.addRow([]); rT2.getCell(6).value = this.t(`Thuế GTGT (VAT ${this.vatRate}%):`, `VAT (${this.vatRate}%):`); rT2.getCell(7).value = this.bgVATAmount; ws.mergeCells(`A${rIdx}:E${rIdx}`); rT2.font={ bold: true, name: 'Times New Roman', size: 10 }; rT2.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; rT2.getCell(7).numFmt=FORMAT_MONEY; rIdx++;
                const rT3 = ws.addRow([]); rT3.getCell(6).value = this.t('TỔNG GIÁ TRỊ (SAU THUẾ):', 'TOTAL AMOUNT:'); rT3.getCell(7).value = this.bgTotalAfterVAT; ws.mergeCells(`A${rIdx}:E${rIdx}`); rT3.font={bold:true, color:{argb:'FFDC2626'}, name:'Times New Roman', size:11}; rT3.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; rT3.getCell(7).numFmt=FORMAT_MONEY; rIdx++;
                const rT4 = ws.addRow([]); rT4.getCell(1).value = `${this.t('Bằng chữ:', 'In words:')} ${this.readMoney(this.bgTotalAfterVAT)}`; ws.mergeCells(`A${rIdx}:G${rIdx}`); rT4.font={italic:true, name:'Times New Roman', size:10}; rT4.getCell(1).alignment={ vertical: 'middle', horizontal: 'left' }; rIdx++;

                builder.applyTableStyle(startCus, rIdx-5, 7);
                
                ws.addRow([]); rIdx++; 
                ws.addRow([this.t('ĐIỀU KHOẢN THƯƠNG MẠI (TERMS & CONDITIONS):', 'TERMS & CONDITIONS:')]).font = {bold:true, underline:true, name:'Times New Roman', size:10}; rIdx++;
                this.bgTerms.forEach(t => { 
                    const tr = ws.addRow([]); tr.getCell(1).value = '-'; tr.getCell(2).value = str(t.title); tr.getCell(3).value = str(t.content);
                    ws.mergeCells(`C${rIdx}:G${rIdx}`); tr.getCell(2).font={ bold: true, name: 'Times New Roman', size: 10 }; tr.getCell(3).alignment={ vertical: 'middle', horizontal: 'left', wrapText: true }; 
                    tr.height = Math.max(20, Math.ceil(str(t.content).length / 60) * 15);
                    rIdx++; 
                });
                
                ws.addRow([]); rIdx++; 
                const signR = ws.addRow([]); signR.getCell(2).value = this.t('ĐẠI DIỆN KHÁCH HÀNG\n\n(Ký và ghi rõ họ tên)', 'CUSTOMER REPRESENTATIVE\n\n(Signature & Full Name)'); signR.getCell(6).value = this.t('ĐẠI DIỆN CÔNG TY\n\n(Ký và ghi rõ họ tên)', 'COMPANY REP.\n\n(Signature & Full Name)');
                ws.mergeCells(`B${rIdx}:C${rIdx}`); ws.mergeCells(`F${rIdx}:G${rIdx}`); signR.font={ bold: true, name: 'Times New Roman', size: 10 }; signR.alignment={ vertical: 'top', horizontal: 'center', wrapText: true }; signR.height = 100;
            }
            
            /* ========================= TAB PR (ĐỀ NGHỊ MUA HÀNG) ========================= */
            else if (type === 'pr') {
                fileName = `PR_${str(this.prHeader.prNo) || 'No'}_${dateStr}.xlsx`;
                ws.columns = [ 
                    {width: 4}, {width: 15}, {width: 10}, {width: 10}, {width: 6}, {width: 6}, {width: 6}, {width: 6}, 
                    {width: 10}, {width: 12}, {width: 10}, {width: 12}, {width: 10}, {width: 10}, {width: 10}, {width: 10}, {width: 8}, {width: 12} 
                ];
                ws.pageSetup.printTitlesRow = '1:11'; builder.setMargins(0.15);
                builder.addLogo(logoData, 100, 40);
                
                let r1 = ws.addRow([]); r1.getCell(1).value = this.t('PHIẾU ĐỀ NGHỊ MUA HÀNG', 'PURCHASE REQUISITION'); r1.getCell(1).font={bold:true, name:'Times New Roman', size:16}; r1.getCell(1).alignment=builder.alignCenter; r1.getCell(15).value = this.t('Số hiệu: ADG_QT_KH_01/BM01', 'Doc No: ADG_QT_KH_01/BM01'); r1.getCell(15).font={ bold: true, name: 'Times New Roman', size: 10 };
                let r2 = ws.addRow([]); r2.getCell(15).value = `${this.t('Ngày HL:', 'Effective:')} ${str(this.prHeader.isoDate)} | ${this.t('Lần BH:', 'Issue No:')} ${str(this.prHeader.isoRev)} | ${this.t('Lần SX:', 'Rev No:')} ${str(this.prHeader.isoCheck)}`;
                ws.mergeCells('A1:N2'); ws.mergeCells('O1:R1'); ws.mergeCells('O2:R2'); builder.applyTableStyle(1, 2, 18, 9);
                
                ws.addRow([]); 
                let r4 = ws.addRow([]); r4.getCell(1).value = this.t('Số phiếu:', 'PR No:'); r4.getCell(1).font = { bold: true, name: 'Times New Roman', size: 10 }; r4.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' };
                r4.getCell(2).value = `ADG_ĐNMH_${str(this.prHeader.prNo)}`; r4.getCell(2).font={bold:true, color:{argb:'FFDC2626'}, name:'Times New Roman', size:11}; ws.mergeCells('B4:H4'); r4.getCell(2).border = builder.borderBotOnly; r4.getCell(2).alignment={ vertical: 'middle', horizontal: 'left' };
                r4.getCell(9).value = this.prHeader.isPlan ? this.t('☑ Theo kế hoạch', '☑ Planned') : this.t('☐ Theo kế hoạch', '☐ Planned'); r4.getCell(9).font={ bold: true, name: 'Times New Roman', size: 10 }; ws.mergeCells('I4:M4');
                
                let r5 = ws.addRow([]); r5.getCell(1).value = this.t('Đơn vị lập:', 'Company:'); r5.getCell(2).value = str(this.prHeader.company); ws.mergeCells('B5:H5'); r5.getCell(1).font={ bold: true, name: 'Times New Roman', size: 10 }; r5.getCell(2).font={ bold: true, name: 'Times New Roman', size: 10 }; r5.getCell(2).border = builder.borderBotOnly; r5.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; r5.getCell(2).alignment={ vertical: 'middle', horizontal: 'left' };
                let r6 = ws.addRow([]); r6.getCell(1).value = this.t('Bộ phận:', 'Department:'); r6.getCell(2).value = str(this.prHeader.dept); ws.mergeCells('B6:F6'); r6.getCell(1).font={ bold: true, name: 'Times New Roman', size: 10 }; r6.getCell(2).font={ bold: true, name: 'Times New Roman', size: 10 }; r6.getCell(2).border = builder.borderBotOnly; r6.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; r6.getCell(2).alignment={ vertical: 'middle', horizontal: 'left' };
                r6.getCell(9).value = this.t('Mã code:', 'Dept Code:'); r6.getCell(10).value = str(this.prHeader.deptCode); ws.mergeCells('J6:M6'); r6.getCell(9).font={ bold: true, name: 'Times New Roman', size: 10 }; r6.getCell(10).font={ bold: true, name: 'Times New Roman', size: 10 }; r6.getCell(10).border = builder.borderBotOnly; r6.getCell(9).alignment={ vertical: 'middle', horizontal: 'right' }; r6.getCell(10).alignment={ vertical: 'middle', horizontal: 'left' };
                
                let r7 = ws.addRow([]); r7.getCell(1).value = this.t('Kho hàng:', 'Warehouse:'); r7.getCell(2).value = str(this.prHeader.warehouse); ws.mergeCells('B7:C7'); r7.getCell(1).font={ bold: true, name: 'Times New Roman', size: 10 }; r7.getCell(2).border = builder.borderBotOnly; r7.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; r7.getCell(2).alignment={ vertical: 'middle', horizontal: 'left' };
                r7.getCell(5).value = this.t('Thủ kho:', 'Storekeeper:'); r7.getCell(6).value = str(this.prHeader.storekeeper); ws.mergeCells('F7:H7'); r7.getCell(5).font={ bold: true, name: 'Times New Roman', size: 10 }; r7.getCell(6).border = builder.borderBotOnly; r7.getCell(5).alignment={ vertical: 'middle', horizontal: 'right' }; r7.getCell(6).alignment={ vertical: 'middle', horizontal: 'left' };
                r7.getCell(9).value = this.t('Ngày duyệt:', 'Approved:'); r7.getCell(10).value = str(this.prHeader.approvedDate); ws.mergeCells('J7:M7'); r7.getCell(9).font={ bold: true, name: 'Times New Roman', size: 10 }; r7.getCell(10).border = builder.borderBotOnly; r7.getCell(9).alignment={ vertical: 'middle', horizontal: 'right' }; r7.getCell(10).alignment={ vertical: 'middle', horizontal: 'left' };
                ws.addRow([]);

                ws.addRow([ this.t('STT', 'No.'), this.t('Mô tả hàng hóa', 'Item Description'), '', this.t('Mã vật tư\n(*)', 'Mat. Code\n(*)'), this.t('ĐVT\n(*)', 'Unit\n(*)'), this.t('Bình quân SD hàng tháng\n(*)', 'Avg Monthly Usage\n(*)'), this.t('Số lượng', 'Quantity'), '', this.t('Đơn giá\ntham khảo\n(có VAT)\n(*)', 'Est. Price\n(VAT incl.)\n(*)'), this.t('Thành tiền\n(Có VAT)\n(*)', 'Total Amount\n(VAT incl.)\n(*)'), this.t('Thời gian\ncần hàng\n(*)', 'Req. Date\n(*)'), this.t('Mục đích sử dụng\n(*)', 'Purpose\n(*)'), this.t('TK chi phí\n(Người lập)', 'Expense Acc\n(Creator)'), this.t('TK chi phí\n(KTT)', 'Expense Acc\n(Chief Acct)'), this.t('Ngân sách\ncòn lại', 'Remaining\nBudget'), this.t('Nhà cung cấp tham khảo', 'Suggested\nSupplier'), this.t('Công ty\nsử dụng\n(*)', 'Company Use\n(*)'), this.t('Điều kiện nghiệm thu/ Đánh giá chất lượng (nếu có)', 'QC Conditions') ]);
                ws.addRow([ '', this.t('Tên hàng hóa\n(*)', 'Item Name\n(*)'), this.t('Quy cách\n(*)', 'Specs\n(*)'), '', '', '', this.t('Tồn Kho\n(*)', 'Stock\n(*)'), this.t('Đề xuất\n(*)', 'Request\n(*)'), '', '', '', '', '', '', '', '', '', '' ]);
                ws.addRow(['(1)','(2)','(3)','(4)','(5)','(6)','(7)','(8)','(9)','(10)','(11)','(12)','(13)','(14)','(15)','(16)','(17)','(18)']);
                ['A9:A10','B9:C9','D9:D10','E9:E10','F9:F10','G9:H9','I9:I10','J9:J10','K9:K10','L9:L10','M9:M10','N9:N10','O9:O10','P9:P10','Q9:Q10','R9:R10'].forEach(m => ws.mergeCells(m));
                [9, 10].forEach(r => ws.getRow(r).eachCell(c => { c.font={name:'Times New Roman', size:9, bold:true, color:{argb:'FF000000'}}; c.fill=fillHeader; c.border=builder.borderThin; c.alignment=builder.alignCenter; }));
                ws.getRow(11).eachCell(c => { c.font={name:'Times New Roman', size:8, italic:true}; c.fill={type:'pattern', pattern:'solid', fgColor:{argb:'FFD9D9D9'}}; c.border=builder.borderThin; c.alignment=builder.alignCenter; });
                
                let rIdx = 12;
                this.docPR.forEach((row, i) => {
                    const r = ws.addRow([]);
                    let avg = num(row.avgUsage); let qS = num(row.qtyStock); let qR = num(row.qtyReq); let p = num(row.price); let bud = num(row.budget);
                    if(this.roundDecimal) { avg=Math.round(avg); qS=Math.round(qS); qR=Math.round(qR); p=Math.round(p); bud=Math.round(bud); }
                    r.getCell(1).value = i+1; r.getCell(2).value = str(row.name); r.getCell(3).value = str(row.specs); r.getCell(4).value = str(row.partNo); r.getCell(5).value = str(row.unit); 
                    r.getCell(6).value = avg; r.getCell(7).value = qS; r.getCell(8).value = qR; r.getCell(9).value = p; r.getCell(10).value = qR * p; 
                    r.getCell(11).value = str(row.timeNeeded); r.getCell(12).value = str(row.purpose); r.getCell(13).value = str(row.accCreator); r.getCell(14).value = str(row.accKTT); r.getCell(15).value = bud; r.getCell(16).value = str(row.supplier); r.getCell(17).value = str(row.companyUse); r.getCell(18).value = str(row.qcCond);
                    [2,3,12,16,18].forEach(c=>r.getCell(c).alignment={ vertical: 'middle', horizontal: 'left', wrapText: true });
                    r.getCell(4).font={bold:true, color:{argb:'FFDC2626'}, name:'Times New Roman', size:9};
                    [6,7,8,9,10,15].forEach(c => r.getCell(c).numFmt=FORMAT_MONEY); r.getCell(8).font={bold:true, name:'Times New Roman', size:9}; r.getCell(10).font={bold:true, name:'Times New Roman', size:9};
                    rIdx++;
                });
                
                let rTot = ws.addRow([]); rTot.getCell(9).value = this.t('Tổng cộng:', 'Total:'); rTot.getCell(10).value = this.roundDecimal ? Math.round(this.prTotalAmount) : this.prTotalAmount; ws.mergeCells(`A${rIdx}:I${rIdx}`); rTot.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; rTot.font={bold:true, name:'Times New Roman', size:9}; rTot.getCell(10).numFmt=FORMAT_MONEY; rTot.getCell(10).font={bold:true, color:{argb:'FFDC2626'}, name:'Times New Roman', size:10}; rIdx++;
                let rWords = ws.addRow([]); rWords.getCell(9).value = `${this.t('Bằng chữ:', 'In words:')} ${this.readMoney(this.prTotalAmount)}`; ws.mergeCells(`A${rIdx}:R${rIdx}`); rWords.getCell(1).alignment={ vertical: 'middle', horizontal: 'right' }; rWords.font={italic:true, name:'Times New Roman', size:9}; rIdx++;
                builder.applyTableStyle(12, rIdx-2, 18, 9);
                ws.addRow([]); rIdx++; 

                if (this.printPrWorkflow) {
                    ws.addRow([this.t('QUÁ TRÌNH XỬ LÝ:', 'WORKFLOW PROCESS:')]).font = {bold:true, name:'Times New Roman', size:10}; rIdx++;
                    let wH = ws.addRow([]); wH.getCell(1).value = this.t('Thời điểm xử lý', 'Timestamp'); wH.getCell(2).value = this.t('Bước xử lý', 'Process Step'); wH.getCell(5).value = this.t('Người xử lý', 'Person in Charge'); wH.getCell(9).value = this.t('Ý kiến người xử lý', 'Remarks/Comments'); wH.eachCell(c => { c.font={bold:true, name:'Times New Roman', size:9}; c.fill=fillHeader; c.alignment=builder.alignCenter; });
                    ws.mergeCells(`B${rIdx}:D${rIdx}`); ws.mergeCells(`E${rIdx}:H${rIdx}`); ws.mergeCells(`I${rIdx}:R${rIdx}`); builder.applyTableStyle(rIdx, rIdx, 18, 9); rIdx++;
                    this.prWorkflow.forEach(s => {
                        let r = ws.addRow([]); r.getCell(1).value = str(s.time); r.getCell(2).value = this.lang === 'vi' ? str(s.stepVi) : str(s.stepEn); r.getCell(5).value = str(s.person); r.getCell(9).value = str(s.note);
                        ws.mergeCells(`B${rIdx}:D${rIdx}`); ws.mergeCells(`E${rIdx}:H${rIdx}`); ws.mergeCells(`I${rIdx}:R${rIdx}`); builder.applyTableStyle(rIdx, rIdx, 18, 9);
                        r.getCell(1).alignment=builder.alignCenter; r.getCell(2).alignment={ vertical: 'middle', horizontal: 'left', wrapText: true }; r.getCell(5).alignment=builder.alignCenter; r.getCell(9).alignment={ vertical: 'middle', horizontal: 'left', wrapText: true }; rIdx++;
                    });
                    ws.addRow([]); rIdx++;
                }
                
                let s1 = ws.addRow([]); let s2 = ws.addRow([]); let dText = this.t('Ngày...tháng...năm...', 'Date...../...../.....');
                [1, 5, 8, 11, 14].forEach(c => s1.getCell(c).value = dText);
                s2.getCell(1).value = this.t('Người lập phiếu\n\n\n\n\n', 'Prepared by\n\n\n\n\n'); s2.getCell(5).value = this.t('Trưởng phòng\n\n\n\n\n', 'Dept. Manager\n\n\n\n\n'); s2.getCell(8).value = this.t('TP Cung ứng\n\n\n\n\n', 'Purchasing Mgr\n\n\n\n\n'); s2.getCell(11).value = this.t('Kế toán trưởng\n\n\n\n\n', 'Chief Accountant\n\n\n\n\n'); s2.getCell(14).value = this.t('Ban Giám đốc\n\n\n\n\n', 'Board of Directors\n\n\n\n\n');
                [`A${rIdx}:D${rIdx}`, `E${rIdx}:G${rIdx}`, `H${rIdx}:J${rIdx}`, `K${rIdx}:M${rIdx}`, `N${rIdx}:R${rIdx}`].forEach(m => ws.mergeCells(m)); rIdx++;
                [`A${rIdx}:D${rIdx}`, `E${rIdx}:G${rIdx}`, `H${rIdx}:J${rIdx}`, `K${rIdx}:M${rIdx}`, `N${rIdx}:R${rIdx}`].forEach(m => ws.mergeCells(m));
                s1.alignment = builder.alignCenter; s1.font = {name:'Times New Roman', size:9};
                s2.alignment = { vertical: 'top', horizontal: 'center', wrapText: true }; s2.font = {bold:true, name:'Times New Roman', size:9}; s2.height = 80;
            }

            else if (type === 'daily' || type === 'nghiemthu' || type === 'nghiemthubg') {
                alert(this.t('Tính năng Export Excel cho form này đang bảo trì để nâng cấp UI.', 'Excel export for this form is under maintenance.'));
                this.showSavedToast = false;
                return;
            }

            builder.download(fileName);
        },

        importFromExcel(event) {
            const file = event.target.files[0]; if (!file) return; const reader = new FileReader();
            reader.onload = async (e) => {
                const data = new Uint8Array(e.target.result); const workbook = XLSX.read(data, { type: 'array' }); const firstSheet = workbook.Sheets[workbook.SheetNames[0]]; const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                let newDb = []; let headerFound = false; let colIdx = {};
                const parseCleanNumber = (val) => { 
                    if (typeof val === 'number') return val; if (!val) return 0; 
                    let s = val.toString().trim().replace(/[^0-9\.,\-]/g, ''); if(!s) return 0; 
                    let lastDot = s.lastIndexOf('.'); let lastComma = s.lastIndexOf(','); 
                    if (lastComma > lastDot && lastDot !== -1) { s = s.replace(/\./g, '').replace(/,/g, '.'); } else if (lastDot > lastComma && lastComma !== -1) { s = s.replace(/,/g, ''); } 
                    else { if (s.indexOf('.') !== -1) { let parts = s.split('.'); if (parts.length > 2 || parts[1].length === 3) s = s.replace(/\./g, ''); } else if (s.indexOf(',') !== -1) { let parts = s.split(','); if (parts.length > 2 || parts[1].length === 3) s = s.replace(/,/g, ''); else s = s.replace(/,/g, '.'); } } 
                    let num = parseFloat(s); return isNaN(num) ? 0 : num; 
                };
                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i]; if (!row || row.length === 0) continue;
                    if (!headerFound) { 
                        let rowStr = row.join('').toLowerCase(); 
                        if (rowStr.includes('mã vật tư') && rowStr.includes('tên vật tư')) { 
                            headerFound = true; 
                            colIdx.partNo = row.findIndex(c => c && c.toString().toLowerCase().includes('mã vật tư')); colIdx.name = row.findIndex(c => c && c.toString().toLowerCase().includes('tên vật tư')); colIdx.unit = row.findIndex(c => c && c.toString().toLowerCase().includes('đvt')); colIdx.qty = row.findIndex(c => c && c.toString().toLowerCase().includes('số lượng')); colIdx.val = row.findIndex(c => c && c.toString().toLowerCase().includes('giá trị')); 
                        } continue; 
                    }
                    const name = row[colIdx.name] ? row[colIdx.name].toString().trim() : ''; if (!name || name.toLowerCase().includes('tổng cộng')) continue;
                    const partNo = row[colIdx.partNo] ? row[colIdx.partNo].toString().trim() : ''; const unit = row[colIdx.unit] ? row[colIdx.unit].toString().trim() : ''; 
                    const qty = colIdx.qty !== -1 ? parseCleanNumber(row[colIdx.qty]) : 0; const val = colIdx.val !== -1 ? parseCleanNumber(row[colIdx.val]) : 0; const priceVT = qty > 0 ? (val / qty) : 0;
                    newDb.push({ partNo: partNo, name: name, specs: '', unit: unit, priceVT: priceVT, priceNC: 0 });
                }
                if (newDb.length > 0) { this.database = newDb; await localforage.setItem('masterDataProject', JSON.parse(JSON.stringify(newDb))); alert(this.t(`Cập nhật thành công ${newDb.length} vật tư.`, `Updated ${newDb.length} items.`)); } else { alert('LỖI: Cấu trúc file sai.'); } 
                event.target.value = ''; 
            }; reader.readAsArrayBuffer(file);
        },
        
        importNTBGFromExcel(event) {
            const file = event.target.files[0]; if (!file) return;
            const reader = new FileReader();
            reader.onload = async (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

                let newDoc = [];
                let newHeader = { ...this.ntbgHeader };
                let isTable = false;

                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row || row.length === 0) continue;
                    const rowStr = row.join(' ').toLowerCase();

                    if (rowStr.includes('v/v:')) {
                        let projectStr = row.find(c => c && c.toString().toLowerCase().includes('v/v:'));
                        if (projectStr) newHeader.project = projectStr.toString().trim();
                    }
                    if (rowStr.includes('tháng') && (rowStr.includes('/20') || rowStr.includes('/ 20'))) {
                         let mStr = row.find(c => c && c.toString().toLowerCase().includes('tháng'));
                         if(mStr) newHeader.monthYear = mStr.toString().trim();
                    }
                    if (rowStr.includes('bên a') || rowStr.includes('bên thực hiện') || rowStr.includes('party a')) {
                         let comp = row.find(c => c && c.toString().toLowerCase() !== 'bên a' && c.toString().toLowerCase() !== 'party a' && (c.toString().toLowerCase().includes('công ty') || c.toString().toLowerCase().includes('company')));
                         if(comp) newHeader.benA_company = comp.toString().trim();
                    }
                    if (rowStr.includes('bên b') || rowStr.includes('bên thuê') || rowStr.includes('party b')) {
                         let comp = row.find(c => c && c.toString().toLowerCase() !== 'bên b' && c.toString().toLowerCase() !== 'party b' && (c.toString().toLowerCase().includes('công ty') || c.toString().toLowerCase().includes('company')));
                         if(comp) newHeader.benB_company = comp.toString().trim();
                    }
                    if (rowStr.includes('căn cứ vào') || rowStr.includes('based on')) {
                        let baseStr = row.find(c => c && (c.toString().toLowerCase().includes('căn cứ vào') || c.toString().toLowerCase().includes('based on')));
                        if(baseStr) newHeader.baseOn = baseStr.toString().trim();
                    }

                    if (!isTable && row.some(c => c && (c.toString().toLowerCase() === 'stt' || c.toString().toLowerCase() === 'no.'))) {
                        isTable = true;
                        if (rows[i+1] && (rows[i+1].join(' ').toLowerCase().includes('bắt đầu') || rows[i+1].join(' ').toLowerCase().includes('start'))) i++; 
                        continue;
                    }

                    if (isTable) {
                        const stt = row[0];
                        if (stt && (stt.toString().toLowerCase().includes('ngày') || stt.toString().toLowerCase().includes('date'))) break;

                        if (stt && !isNaN(parseInt(stt))) {
                            let rest = [];
                            for(let k = 1; k < row.length; k++) {
                                if(row[k] !== undefined && row[k] !== null && row[k].toString().trim() !== '') {
                                    rest.push(row[k].toString().trim());
                                }
                            }
                            
                            let date = rest[0] || ''; let eq = rest[1] || ''; let desc = rest[2] || ''; let timeStart = rest[3] || ''; let timeEnd = rest[4] || ''; let note = '';
                            
                            if(rest.length > 5) {
                                let lastItem = rest[rest.length - 1];
                                if (lastItem.toLowerCase() !== 'x' && lastItem.toLowerCase() !== 'đạt' && lastItem.toLowerCase() !== 'pass') note = lastItem;
                            }

                            newDoc.push({ requestDate: date, equipment: eq, description: desc, timeStart: timeStart, timeEnd: timeEnd, statusPass: 'x', statusFail: '', note: note });
                        }
                    }
                }
                
                if (newDoc.length > 0) {
                    this.ntbgHeader = newHeader;
                    this.docNghiemThuBG = newDoc;
                    this.saveDraft('ntbg_header', newHeader);
                    this.saveDraft('ntbg_table', newDoc);
                    alert(this.t(`Đã Import thành công ${newDoc.length} công việc từ nhà thầu!`, `Imported ${newDoc.length} items!`));
                } else {
                    alert(this.t('Không tìm thấy dữ liệu bảng hợp lệ trong file Excel.', 'No valid table data found.'));
                }
                event.target.value = '';
            };
            reader.readAsArrayBuffer(file);
        },

        addBlankRow(t) { 
            if(t==='dutoan') this.docDuToan.push({name: '', specs: '', unit: '', qty: 1, priceVT: 0, priceNC: 0}); 
            if(t==='baogia') this.docBaoGia.push({partNo: '', name: '', specs: '', unit: '', qty: 1, priceGop: 0}); 
            if(t==='pr') this.docPR.push({name: '', specs: '', partNo: '', unit: '', avgUsage: 0, qtyStock: 0, qtyReq: 1, price: 0, timeNeeded: '', accCreator: '', accKTT: '', budget: 0, supplier: '', companyUse: 'ADG', qcCond: 'Đúng thông số kỹ thuật yêu cầu'}); 
            if(t==='dailyTasks') this.dailyTasks.push({col1: '', col2: '', col3: '', col4: '', col5: ''}); 
            if(t==='dailyManpower') this.dailyManpower.push({role: '', qty: 1}); 
            if(t==='dailyEquipment') this.dailyEquipment.push({name: '', qty: 1});
            if(t==='nghiemthu') this.docNghiemThu.push({name: '', specContract: '', specActual: '', doc: '', status: 'ĐẠT YÊU CẦU', note: '', image: null});
            if(t==='nghiemthubg') this.docNghiemThuBG.push({requestDate: '', equipment: '', description: '', timeStart: '', timeEnd: '', statusPass: 'x', statusFail: '', note: ''});
        },

        addCategoryRow(t) {
            if(t==='dutoan') this.docDuToan.push({ isCategory: true, name: '', specs: '', unit: '', qty: 1, priceVT: 0, priceNC: 0 });
            if(t==='baogia') this.docBaoGia.push({ isCategory: true, name: '', specs: '', unit: '', qty: 1, priceGop: 0 });
        },
        
        removeRow(arr, i) { if(confirm(this.t("Xóa dòng này khỏi Bảng?", "Delete this row?"))) arr.splice(i, 1); },
        addTerm() { this.bgTerms.push({ title: 'Điều khoản mới:', content: '...' }); }, 
        removeTerm(i) { this.bgTerms.splice(i, 1); },
        
        copyFromDuToan() { 
            if(confirm(this.t("Gộp dữ liệu từ BOM sang Báo Giá?", "Import from BOM?"))) { 
                this.docBaoGia = this.docDuToan.map(i => {
                    if (i.isCategory) return { isCategory: true, name: i.name, specs: '', unit: '', qty: 1, priceGop: 0, partNo: '' };
                    return { ...i, partNo: i.partNo || '', name: i.name || '', specs: i.specs || '', unit: i.unit || '', qty: i.qty || 1, priceGop: (((i.priceVT * i.qty * (1 + this.tyleQLCVT/100)) + (i.priceNC * i.qty * (1 + this.tyleNCGT/100) * (1 + this.tyleQLCNC/100))) / i.qty) || 0 };
                }); 
            } 
        },
        
        copyFromDuToanToPR() {
            if(confirm(this.t("Lấy dữ liệu từ BOM sang Phiếu PR?", "Import from BOM to PR?"))) {
                const materialsOnly = this.docDuToan.filter(i => !i.isCategory && (i.priceVT > 0 || (i.priceVT === 0 && i.priceNC === 0)));
                const newPRItems = materialsOnly.map(i => ({ name: i.name || '', specs: i.specs || '', partNo: i.partNo || '', unit: i.unit || '', avgUsage: 0, qtyStock: 0, qtyReq: i.qty || 1, price: i.priceVT || 0, timeNeeded: '', purpose: 'Dự án / Bảo trì', accCreator: '', accKTT: '', budget: 0, supplier: '', companyUse: 'ADG', qcCond: 'Đúng thông số kỹ thuật yêu cầu' }));
                this.docPR = [...this.docPR, ...newPRItems]; alert(this.t(`Đã thêm ${newPRItems.length} mã vật tư.`, `Added ${newPRItems.length} items.`));
            }
        },

        async hashPassword(message) {
            const msgBuffer = new TextEncoder().encode(message); const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer); const hashArray = Array.from(new Uint8Array(hashBuffer)); return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
        },
        
        async loginAdmin() { 
            const pass = prompt(this.t("Nhập mã số NV (Password):", "Enter Staff ID (Password):")); if (!pass) return;
            const inputHash = await this.hashPassword(pass); const targetPass = atob("MjEwOTAwNA=="); const VALID_HASH = await this.hashPassword(targetPass); 
            if (inputHash === VALID_HASH) { this.isAdmin = true; alert(this.t("Truy cập thành công!", "Access Granted!")); } else { alert(this.t("Từ chối truy cập! Mật khẩu sai.", "Access Denied! Wrong password.")); }
        },
        
        async migrateDataToIndexedDB() {
            const keys = ['masterDataProject', 'employees_db', 'global_project_name', 'dt_table', 'bg_table', 'bg_header', 'bg_terms', 'pr_table', 'pr_flow', 'nt_table', 'company_logo_base64', 'pr_header', 'pr_flow_print_setting', 'global_vat_rate'];
            for (let k of keys) { 
                let oldData = localStorage.getItem(k); 
                if (oldData !== null) { 
                    try { await localforage.setItem(k, JSON.parse(oldData)); } 
                    catch(e) { await localforage.setItem(k, oldData); } 
                    localStorage.removeItem(k); 
                } 
            }
        },

        formatNum(val) { 
            if (val === null || val === undefined || isNaN(val) || val === '') return ''; 
            if (val === 0) return '0';
            let num = Number(val);
            if (this.roundDecimal) { return Math.round(num).toLocaleString('vi-VN', { maximumFractionDigits: 0 }); }
            return num.toLocaleString('vi-VN', { maximumFractionDigits: 3 }); 
        },
        
        focusNum(e, val) { e.target.value = val === 0 ? '' : val.toString().replace('.', ','); e.target.select(); },
        
        blurNum(e) { 
            let str = e.target.value.toString().trim();
            if (!str) return 0;
            str = str.replace(/[^\d\.,\-]/g, '');
            if(!str) return 0;

            let dotCount = (str.match(/\./g) || []).length;
            let commaCount = (str.match(/,/g) || []).length;
            let num = 0;
            
            if (commaCount > 0 && dotCount > 0) {
                let lastDot = str.lastIndexOf('.'); let lastComma = str.lastIndexOf(',');
                if (lastDot > lastComma) num = parseFloat(str.replace(/,/g, ''));
                else num = parseFloat(str.replace(/\./g, '').replace(/,/g, '.'));
            } else if (commaCount > 0) {
                if (commaCount > 1 || (str.length - str.lastIndexOf(',') === 4)) num = parseFloat(str.replace(/,/g, ''));
                else num = parseFloat(str.replace(/,/g, '.'));
            } else if (dotCount > 0) {
                if (dotCount > 1 || (str.length - str.lastIndexOf('.') === 4)) num = parseFloat(str.replace(/\./g, ''));
                else num = parseFloat(str); 
            } else {
                num = parseFloat(str);
            }
            if (this.roundDecimal) { num = Math.round(num); }
            return isNaN(num) ? 0 : num;
        },
        
        readMoney(num) {
            if (!num || isNaN(num) || num === 0) return this.t("Không đồng", "Zero VND"); 
            num = Math.round(num); 
            if (this.lang === 'en') return "VND " + num.toLocaleString('en-US') + " only.";
            
            const defaultNumbers = ' hai ba bốn năm sáu bảy tám chín'; 
            const chuHangDonVi = ('1 một' + defaultNumbers).split(' '); 
            const chuHangChuc = ('lẻ mười' + defaultNumbers).split(' '); 
            const chuHangTram = ('không một' + defaultNumbers).split(' '); 
            const mangGiaTri = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ', 'triệu tỷ'];
            let str = parseInt(Math.abs(num)) + ''; let arr = []; 
            while (str.length > 0) { arr.push(str.length >= 3 ? str.slice(-3) : str); str = str.length >= 3 ? str.slice(0, -3) : ''; }
            let rsString = '';
            for (let i = 0; i < arr.length; i++) {
                let chuoi = ''; let temp = arr[i].padStart(3, '0'); let tram = parseInt(temp[0]); let chuc = parseInt(temp[1]); let donVi = parseInt(temp[2]);
                if (tram === 0 && chuc === 0 && donVi === 0) continue;
                if (tram !== 0 || i !== arr.length - 1) { chuoi += chuHangTram[tram] + ' trăm '; if (chuc === 0 && donVi !== 0) chuoi += 'lẻ '; }
                if (chuc !== 0 && chuc !== 1) { chuoi += chuHangChuc[chuc] + ' mươi '; if (chuc === 4 && donVi !== 0) chuoi = chuoi.replace('bốn mươi', 'bốn mươi'); } else if (chuc === 1) { chuoi += 'mười '; }
                switch (donVi) { 
                    case 1: chuoi += (chuc !== 0 && chuc !== 1) ? 'mốt ' : chuHangDonVi[donVi] + ' '; break; 
                    case 5: chuoi += (chuc !== 0) ? 'lăm ' : chuHangDonVi[donVi] + ' '; break; 
                    case 4: chuoi += (chuc !== 0 && chuc !== 1) ? 'tư ' : chuHangDonVi[donVi] + ' '; break; 
                    default: if (donVi !== 0) chuoi += chuHangDonVi[donVi] + ' '; break; 
                }
                rsString = chuoi + mangGiaTri[i] + ' ' + rsString;
            }
            rsString = rsString.trim().replace(/\s+/g, ' '); 
            if (rsString.startsWith('không trăm lẻ ')) rsString = rsString.substring(14); else if (rsString.startsWith('không trăm ')) rsString = rsString.substring(11);
            let res = rsString.charAt(0).toUpperCase() + rsString.slice(1) + ' đồng chẵn.'; 
            return num < 0 ? "Âm " + res.toLowerCase() : res;
        },
        
        hideSearch() { setTimeout(() => { this.showSearchDropdown = false; }, 200); },
        
        uploadRowImage(event, row) {
            const file = event.target.files[0];
            if (file) { 
                const reader = new FileReader(); 
                reader.onload = (e) => { 
                    const img = new Image(); 
                    img.onload = () => { 
                        const canvas = document.createElement('canvas'); let width = img.width; let height = img.height; const MAX_SIZE = 600;
                        if (width > height) { if (width > MAX_SIZE) { height *= MAX_SIZE / width; width = MAX_SIZE; } } else { if (height > MAX_SIZE) { width *= MAX_SIZE / height; height = MAX_SIZE; } } 
                        canvas.width = width; canvas.height = height; const ctx = canvas.getContext('2d'); ctx.drawImage(img, 0, 0, width, height); 
                        row.image = canvas.toDataURL('image/jpeg', 0.5); 
                    }; img.src = e.target.result; 
                }; reader.readAsDataURL(file); 
            } event.target.value = ''; 
        },
        
        uploadLogo(event) { 
            const file = event.target.files[0]; 
            if (file) { 
                const reader = new FileReader(); reader.onload = (e) => { this.companyLogo = e.target.result; localforage.setItem('company_logo_base64', this.companyLogo); }; reader.readAsDataURL(file); 
            } 
        },
        
        selectItemFromSearch(item, targetArray) {
            let newItem = { ...item, qty: 1 }; newItem.priceGop = newItem.priceVT + newItem.priceNC; 
            if (this.currentTab === 'pr') { targetArray.push({ name: newItem.name, specs: newItem.specs, partNo: newItem.partNo, unit: newItem.unit, avgUsage: 0, qtyStock: 0, qtyReq: 1, price: newItem.priceVT, timeNeeded: '', purpose: 'Dự án / Bảo trì', accCreator: '', accKTT: '', budget: 0, supplier: '', companyUse: 'ADG', qcCond: 'Đúng thông số kỹ thuật yêu cầu' }); } 
            else { targetArray.push(newItem); }
            this.searchQuery = ''; this.showSearchDropdown = false;
        },

        saveDraft(key, data) { 
            if (this.draftTimeouts[key]) { clearTimeout(this.draftTimeouts[key]); }
            this.draftTimeouts[key] = setTimeout(() => {
                localforage.setItem(key, JSON.parse(JSON.stringify(data))).then(() => { 
                    this.showSavedToast = true; setTimeout(() => this.showSavedToast = false, 2000); 
                });
            }, 800); 
        },
        
        async loadDrafts() {
            const parseL = async (k, defaultData) => { const data = await localforage.getItem(k); return data !== null ? data : defaultData; };
            this.docDuToan = await parseL('dt_table', this.docDuToan); 
            this.docBaoGia = await parseL('bg_table', this.docBaoGia); 
            this.bgTerms = await parseL('bg_terms', this.bgTerms); 
            this.docPR = await parseL('pr_table', this.docPR); 
            this.prWorkflow = await parseL('pr_flow', this.prWorkflow); 
            this.docNghiemThu = await parseL('nt_table', this.docNghiemThu); 
            this.globalProjectName = await parseL('global_project_name', this.globalProjectName); 
            this.prHeader = await parseL('pr_header', this.prHeader);
            this.printPrWorkflow = await parseL('pr_flow_print_setting', this.printPrWorkflow);
            this.vatRate = await parseL('global_vat_rate', this.vatRate);
            this.dtVatRate = await parseL('dt_vat_rate', this.dtVatRate);
            this.bgHeader = await parseL('bg_header', this.bgHeader);
            this.employees = await parseL('employees_db', this.employees);
            const cachedLogo = await parseL('company_logo_base64', null); if(cachedLogo) this.companyLogo = cachedLogo;

            this.ntbgHeader = await parseL('ntbg_header', this.ntbgHeader);
            this.docNghiemThuBG = await parseL('ntbg_table', this.docNghiemThuBG);
        }
    }
}).mount('#app');
