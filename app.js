const { createApp } = Vue;

localforage.config({ driver: localforage.INDEXEDDB, name: 'ME_Portal_DB', version: 1.0, storeName: 'port_documents' });

createApp({
    data() {
        return {
            currentTab: 'database', 
            lang: 'vi', 
            isAdmin: false, 
            showSavedToast: false,
            roundDecimal: false, 
            vatRate: 8,
            
            draftTimeouts: {},
            
            database: [], 
            newItem: { partNo: '', name: '', specs: '', unit: '', priceVT: 0, priceNC: 0 },
            searchQuery: '', 
            showSearchDropdown: false, 
            companyLogo: 'https://raw.githubusercontent.com/Uniekey/BM/main/ADG.png', 
            globalProjectName: '',
            
            // Dữ liệu Thư Viện
            libraryLinks: [
                { title: 'Link sharepoint Tài liệu MMTB Cảng', titleEn: 'Port Equipment Documents (Sharepoint)', desc: 'Hồ sơ MMTB', descEn: 'Equipment Profiles', icon: 'fas fa-anchor', url: 'https://dongtam365.sharepoint.com/sites/ctycpptca/05%20Phng%20Dch%20V%20C%20in%20Cng/Forms/AllItems.aspx?id=%2Fsites%2Fctycpptca%2F05%20Phng%20Dch%20V%20C%20in%20Cng%2F5%2E7%2ET%C3%80I%20LI%E1%BB%86U%20%2D%20H%E1%BB%92%20S%C6%A0%20MMTB&viewid=ca8054c9%2D7c9c%2D4c5b%2Da730%2D3722d554e029' },
                { title: 'Hiệu chỉnh cân xe tải - Cảng', titleEn: 'Truck Scale Calibration', desc: 'Chỉnh góc - calib cân xe tải', descEn: 'Corner adjustment & Calibration', icon: 'fas fa-balance-scale', url: 'https://uniekey.github.io/truck-scale/' },
                { title: 'Quy định/Quy trình/HDCV', titleEn: 'SOPs & Regulations', desc: 'Tài liệu nội bộ công ty', descEn: 'Internal Documents', icon: 'fas fa-envelope-open-text', url: '#' },
                { title: 'Tính toán điện NLMT', titleEn: 'Solar PV Calculator', desc: 'Tính PV, INV, BESS', descEn: 'Calculate PV, INV, BESS', icon: 'fas fa-search-dollar', url: 'https://uniekey.github.io/solar/' }
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
                { time: '', stepVi: '02. Xác nhận tồn kho', stepEn: '02. Confirm Inventory / Stock', person: 'Phạm Thị Như Huỳnh', note: '' }, 
                { time: '', stepVi: '03. Tổng hợp lại các mặt hàng cần mua/cấp', stepEn: '03. Consolidate Items', person: 'Phạm Thị Như Huỳnh', note: '' }, 
                { time: '', stepVi: '04. Kiểm tra phiếu đề nghị', stepEn: '04. Check PR Details', person: 'Trưởng phòng Cơ Điện', note: '' }, 
                { time: '', stepVi: '05. Xác nhận đơn hàng', stepEn: '05. Order Confirmation', person: 'Lưu Thị Cẩm Hồng', note: '' }, 
                { time: '', stepVi: '06. Kiểm tra ngân sách ĐNMH', stepEn: '06. Budget Verification', person: 'Kế toán trưởng', note: '' }, 
                { time: '', stepVi: '07. Phê duyệt đề nghị mua hàng', stepEn: '07. Final Approval', person: 'Ban Giám đốc', note: '' }, 
                { time: '', stepVi: '08. Tiếp nhận phiếu đề nghị', stepEn: '08. PR Received', person: 'Phạm Thị Như Huỳnh', note: '' } 
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
        dtSumVT() { return this.docDuToan.reduce((sum, item) => sum + ((item.qty || 0) * (item.priceVT || 0)), 0); },
        dtSumNC() { return this.docDuToan.reduce((sum, item) => sum + ((item.qty || 0) * (item.priceNC || 0)), 0); },
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
            let t = this.docBaoGia.reduce((sum, item) => sum + ((item.qty || 0) * (item.priceGop || 0)), 0); 
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
        globalProjectName(val) { this.saveDraft('global_project_name', val); }, 
        docDuToan: { deep: true, handler(val) { this.saveDraft('dt_table', val); } }, 
        docBaoGia: { deep: true, handler(val) { this.saveDraft('bg_table', val); } }, 
        bgHeader: { deep: true, handler(val) { this.saveDraft('bg_header', val); } },
        bgTerms: { deep: true, handler(val) { this.saveDraft('bg_terms', val); } }, 
        docPR: { deep: true, handler(val) { this.saveDraft('pr_table', val); } }, 
        prWorkflow: { deep: true, handler(val) { this.saveDraft('pr_flow', val); } }, 
        docNghiemThu: { deep: true, handler(val) { this.saveDraft('nt_table', val); } }, 
        prHeader: { deep: true, handler(val) { this.saveDraft('pr_header', val); } }, 
        printPrWorkflow(val) { this.saveDraft('pr_flow_print_setting', val); },
        vatRate(val) { this.saveDraft('global_vat_rate', val); },
        dtVatRate(val) { this.saveDraft('dt_vat_rate', val); },
        ntbgHeader: { deep: true, handler(val) { this.saveDraft('ntbg_header', val); } },
        docNghiemThuBG: { deep: true, handler(val) { this.saveDraft('ntbg_table', val); } }
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

        async clearAllCache() {
            if (confirm(this.t("CẢNH BÁO: Hành động này sẽ xóa TOÀN BỘ dữ liệu lưu nháp và Kho vật tư. Trang sẽ được tải lại. Bạn chắc chắn chứ?", "WARNING: This will clear ALL cached data, drafts, and Master Data. The page will reload. Are you sure?"))) {
                try {
                    await localforage.clear();
                    localStorage.clear();
                    window.location.reload();
                } catch (err) {
                    console.error(err);
                    alert("Có lỗi xảy ra khi xóa cache!");
                }
            }
        },

        async exportToExcel(type) {
            if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
            this.showSavedToast = true;

            const str = (v) => v ? v.toString() : '';
            const num = (v) => Number(v) || 0;

            const projectName = this.globalProjectName || "Du_An_ME";
            const dateStr = new Date().toLocaleDateString('vi-VN').replace(/\//g, '-');
            let fileName = `Bao_Cao_${projectName}_${dateStr}.xlsx`;

            const wb = new ExcelJS.Workbook(); 
            wb.creator = 'M&E Portal';
            const ws = wb.addWorksheet('Bao_Cao', { 
                views: [{ showGridLines: false, pageBreakPreview: true }] 
            });

            ws.pageSetup.paperSize = 9; 
            ws.pageSetup.orientation = 'landscape';
            ws.pageSetup.fitToPage = true;
            ws.pageSetup.fitToWidth = 1;
            ws.pageSetup.fitToHeight = 0; 
            ws.pageSetup.margins = { left: 0.3, right: 0.3, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 };
            ws.pageSetup.horizontalCentered = true;

            let logoId = null;
            const logoData = this.getBase64Data(this.companyLogo);
            if (logoData) logoId = wb.addImage({ base64: logoData.base64, extension: logoData.ext });

            const borderThin = { 
                top: {style:'thin', color: {argb: 'FF000000'}}, 
                left: {style:'thin', color: {argb: 'FF000000'}}, 
                bottom: {style:'thin', color: {argb: 'FF000000'}}, 
                right: {style:'thin', color: {argb: 'FF000000'}} 
            };
            const borderBotOnly = { bottom: {style:'thin'} };
            const fontTitle = { bold: true, name: 'Times New Roman', size: 16, color: {argb: 'FF000000'} };
            const fontBold = { bold: true, name: 'Times New Roman', size: 11, color: {argb: 'FF000000'} };
            const fontNormal = { name: 'Times New Roman', size: 11, color: {argb: 'FF000000'} };
            const fillHeader = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            const alignCenter = { vertical: 'middle', horizontal: 'center', wrapText: true };
            const alignLeft = { vertical: 'middle', horizontal: 'left', wrapText: true };
            const alignRight = { vertical: 'middle', horizontal: 'right', wrapText: true };

            const applyStyle = (startR, endR, colCnt, size = 11) => {
                for (let i = startR; i <= endR; i++) { 
                    const row = ws.getRow(i);
                    for(let j = 1; j <= colCnt; j++) {
                        const c = row.getCell(j);
                        c.border = borderThin; 
                        if(!c.alignment) c.alignment = alignCenter; 
                        if(!c.font) c.font = { name: 'Times New Roman', size: size, color: {argb: 'FF000000'} }; 
                    }
                    row.height = Math.max(row.height || 18, 22);
                }
            };

            const FORMAT_MONEY = this.roundDecimal ? '#,##0' : '#,##0.000'; 

            /* ========================= TAB DỰ TOÁN (BOM) ========================= */
            if (type === 'dutoan') {
                fileName = `BOM_${projectName}_${dateStr}.xlsx`;
                ws.columns = [ 
                    {width: 5}, {width: 40}, {width: 20}, {width: 8}, {width: 8}, 
                    {width: 15}, {width: 15}, {width: 15}, {width: 15} 
                ];
                ws.pageSetup.printTitlesRow = '1:6';

                if(logoId !== null) ws.addImage(logoId, { tl: { col: 0.1, row: 0.1 }, ext: { width: 140, height: 50 } });
                
                ws.mergeCells('A1:I1'); ws.getCell('A1').value = this.t('BẢNG DỰ TOÁN CHI PHÍ (BOM)', 'BILL OF MATERIALS (BOM)'); ws.getCell('A1').font = fontTitle; ws.getCell('A1').alignment = alignCenter; ws.getRow(1).height = 30;
                ws.mergeCells('A2:I2'); ws.getCell('A2').value = `${this.t('Dự án:', 'Project:')} ${str(projectName)}`; ws.getCell('A2').font = fontBold; ws.getCell('A2').alignment = alignCenter; ws.getRow(2).height = 20;
                ws.mergeCells('A3:I3'); ws.getCell('A3').value = this.t('', ''); ws.getCell('A3').font = {italic:true, name:'Times New Roman', size:11}; ws.getCell('A3').alignment = alignCenter;
                ws.addRow([]);
                
                ws.addRow([this.t('STT', 'No.'), this.t('Hạng mục', 'Description'), this.t('Quy cách', 'Specs'), this.t('ĐVT', 'Unit'), this.t('SL', 'Qty'), this.t('Đơn giá (VNĐ)', 'Unit Price (VND)'), '', this.t('Thành tiền (VNĐ)', 'Total Amount (VND)'), '']);
                ws.addRow(['', '', '', '', '', this.t('Vật tư', 'Material'), this.t('Nhân công', 'Labor'), this.t('Vật tư', 'Material'), this.t('Nhân công', 'Labor')]);
                ['A5:A6', 'B5:B6', 'C5:C6', 'D5:D6', 'E5:E6', 'F5:G5', 'H5:I5'].forEach(m => ws.mergeCells(m));
                [5, 6].forEach(r => ws.getRow(r).eachCell(c => { c.font = fontBold; c.fill = fillHeader; c.border = borderThin; c.alignment = alignCenter; }));
                
                let rIdx = 7;
                this.docDuToan.forEach((row, i) => {
                    const r = ws.addRow([]);
                    r.getCell(1).value = i + 1;
                    r.getCell(2).value = str(row.name);
                    r.getCell(3).value = str(row.specs);
                    r.getCell(4).value = str(row.unit);
                    let q = num(row.qty); let vt = num(row.priceVT); let nc = num(row.priceNC);
                    if(this.roundDecimal) { q=Math.round(q); vt=Math.round(vt); nc=Math.round(nc); }
                    r.getCell(5).value = q;
                    r.getCell(6).value = vt;
                    r.getCell(7).value = nc;
                    r.getCell(8).value = q * vt;
                    r.getCell(9).value = q * nc;
                    
                    r.getCell(2).alignment=alignLeft; r.getCell(3).alignment=alignLeft; 
                    for(let c=5; c<=9; c++) { r.getCell(c).numFmt=FORMAT_MONEY; r.getCell(c).alignment=alignRight; } 
                    rIdx++;
                });
                
                const rT = ws.addRow([]);
                rT.getCell(7).value = this.t('Tổng chi phí trực tiếp:', 'Subtotal Direct Cost:'); 
                rT.getCell(8).value = this.roundDecimal ? Math.round(this.dtSumVT) : this.dtSumVT; 
                rT.getCell(9).value = this.roundDecimal ? Math.round(this.dtSumNC) : this.dtSumNC;
                ws.mergeCells(`A${rIdx}:G${rIdx}`); rT.font=fontBold; rT.getCell(1).alignment=alignRight; rT.getCell(8).numFmt=FORMAT_MONEY; rT.getCell(9).numFmt=FORMAT_MONEY;
                applyStyle(5, rIdx, 9);
                
                ws.addRow([]); rIdx++;
                
                const sumTitle = ws.addRow([this.t('TỔNG HỢP CHI PHÍ & LỢI NHUẬN', 'COST & PROFIT SUMMARY')]);
                ws.mergeCells(`A${rIdx}:I${rIdx}`);
                sumTitle.font = {bold:true, name:'Times New Roman', size:12}; 
                sumTitle.alignment = alignCenter;
                sumTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
                sumTitle.border = borderThin;
                rIdx++;

                const startSumR = rIdx;
                const sumData = [
                    { t: this.t(`1. TỔNG CHI PHÍ TRỰC TIẾP (Vật tư + Nhân công):`, `1. TOTAL BASE COST (Mat. + Lab.):`), v: this.dtSumVT + this.dtSumNC },
                    { t: this.t(`2. Chi phí Quản lý (${this.tyleNCGT}% NC)`, `2. Management Cost (${this.tyleNCGT}% Lab)`), v: this.dtCpNCGT },
                    { t: this.t(`3. LN định mức Vật tư (${this.tyleQLCVT}% VT)`, `3. Mat. Profit Margin (${this.tyleQLCVT}% Mat)`), v: this.dtCpQLCVT },
                    { t: this.t(`4. LN định mức Nhân công (${this.tyleQLCNC}% NC)`, `4. Lab. Profit Margin (${this.tyleQLCNC}% Lab)`), v: this.dtCpQLCNC },
                    { t: this.t(`TỔNG DỰ TOÁN (TRƯỚC THUẾ):`, `ESTIMATE BEFORE TAX:`), v: this.dtTotalBeforeVAT, bold: true },
                    { t: this.t(`Thuế GTGT (VAT ${this.dtVatRate}%):`, `VAT (${this.dtVatRate}%):`), v: this.dtVATAmount },
                    { t: this.t(`TỔNG GIÁ TRỊ (SAU THUẾ):`, `TOTAL CONTRACT VALUE:`), v: this.dtTotalAfterVAT, bold: true, color: 'FFDC2626', size: 12 },
                ];
                
                sumData.forEach(d => {
                    const r = ws.addRow([]);
                    r.getCell(1).value = d.t; 
                    r.getCell(8).value = this.roundDecimal ? Math.round(d.v) : d.v;
                    ws.mergeCells(`A${rIdx}:G${rIdx}`); 
                    ws.mergeCells(`H${rIdx}:I${rIdx}`); 
                    r.getCell(1).alignment=alignRight; 
                    r.getCell(8).numFmt=FORMAT_MONEY; 
                    r.getCell(8).alignment=alignRight;
                    if(d.bold) { r.font = {bold:true, name:'Times New Roman', size: d.size||11, color: d.color ? {argb: d.color} : undefined}; } 
                    rIdx++;
                });
                
                const rWords = ws.addRow([`${this.t('Bằng chữ:', 'In words:')} ${this.readMoney(this.dtTotalAfterVAT)}`]);
                ws.mergeCells(`A${rIdx}:I${rIdx}`); 
                rWords.font = {italic:true, name:'Times New Roman', size:11}; 
                rWords.getCell(1).alignment=alignLeft;
                rIdx++;

                applyStyle(startSumR, rIdx-1, 9);
                
                ws.addRow([]); rIdx++; 
                const signR = ws.addRow([]);
                signR.getCell(2).value = this.t('NGƯỜI LẬP BẢNG', 'PREPARED BY'); 
                signR.getCell(7).value = this.t('NGƯỜI PHÊ DUYỆT', 'APPROVED BY');
                ws.mergeCells(`B${rIdx}:D${rIdx}`); 
                ws.mergeCells(`G${rIdx}:I${rIdx}`); 
                signR.font=fontBold; signR.alignment={ vertical: 'top', horizontal: 'center' }; signR.height = 100;
            }
            
            /* ========================= TAB BÁO GIÁ ========================= */
            else if (type === 'baogia') {
                fileName = `Bao_Gia_${projectName}_${dateStr}.xlsx`;
                ws.columns = [ {width: 5}, {width: 40}, {width: 25}, {width: 8}, {width: 8}, {width: 20}, {width: 22} ];
                ws.pageSetup.printTitlesRow = '1:18';

                let r1 = ws.addRow([]); 
                r1.getCell(1).value = this.t('BẢNG BÁO GIÁ', 'QUOTATION'); r1.getCell(1).font={bold:true, name:'Times New Roman', size:18}; r1.getCell(1).alignment=alignCenter; r1.height = 25;
                r1.getCell(6).value = `${this.t('Số hiệu:', 'Doc No:')} ${str(this.bgHeader.isoDoc)}`; r1.getCell(6).font=fontBold; r1.getCell(6).alignment = alignLeft;
                
                let r2 = ws.addRow([]); 
                r2.getCell(1).value = str(projectName).toUpperCase(); r2.getCell(1).font={bold:true, name:'Times New Roman', size:12}; r2.getCell(1).alignment=alignCenter; r2.height = 20;
                r2.getCell(6).value = `${this.t('Ngày hiệu lực:', 'Effective:')} ${str(this.bgHeader.isoDate)}\n${this.t('Lần ban hành:', 'Issue No:')} ${str(this.bgHeader.isoRev)}`; r2.getCell(6).alignment = alignLeft;
                
                ws.mergeCells('A1:E1'); ws.mergeCells('F1:G1'); ws.mergeCells('A2:E3'); ws.mergeCells('F2:G3'); 
                applyStyle(1, 3, 7); ws.addRow([]); ws.addRow([]);

                let r6 = ws.addRow([]); 
                r6.getCell(1).value = this.t('Công ty:', 'Company:'); r6.getCell(2).value = str(this.bgHeader.company); 
                r6.getCell(6).value = this.t('Số:', 'Ref No:'); r6.getCell(7).value = str(this.bgHeader.ref); 
                ws.mergeCells('B6:E6'); r6.getCell(1).font=fontBold; r6.getCell(2).font=fontBold; r6.getCell(6).font=fontBold; r6.getCell(7).font={bold:true, color:{argb:'FFDC2626'}};
                
                let r7 = ws.addRow([]); 
                r7.getCell(1).value = this.t('Địa chỉ:', 'Address:'); r7.getCell(2).value = str(this.bgHeader.address); 
                r7.getCell(6).value = this.t('Ngày:', 'Date:'); r7.getCell(7).value = str(this.bgHeader.date); 
                ws.mergeCells('B7:E7'); r7.getCell(1).font=fontBold; r7.getCell(6).font=fontBold;
                
                let r8 = ws.addRow([]); 
                r8.getCell(1).value = this.t('MST:', 'Tax Code:'); r8.getCell(2).value = str(this.bgHeader.taxCode); 
                ws.mergeCells('B8:E8'); r8.getCell(1).font=fontBold; ws.addRow([]);

                ws.addRow([this.t('Kính gửi Quý Khách hàng:', 'To our valued Customer:')]).font=fontBold;
                const cusArr = [
                    [this.t('Công ty/ Cá nhân:', 'Company/ Individual:'), this.bgHeader.cusCompany], 
                    [this.t('Địa chỉ:', 'Address:'), this.bgHeader.cusAddress], 
                    [this.t('MST:', 'Tax Code:'), this.bgHeader.cusTax]
                ];
                let startCus = 11;
                cusArr.forEach(c => {
                    let r = ws.addRow([]); r.getCell(1).value = c[0]; r.getCell(2).value = str(c[1]);
                    ws.mergeCells(`B${startCus}:G${startCus}`);
                    r.getCell(1).font=fontBold; r.getCell(2).font = (c[0] === this.t('Công ty/ Cá nhân:', 'Company/ Individual:')) ? fontBold : fontNormal;
                    startCus++;
                });

                ws.addRow([this.t('Công ty chúng tôi trân trọng báo giá đến quý khách như sau:', 'We are pleased to submit our quotation as follows:')]).font = {italic:true, name:'Times New Roman', size:11};
                
                const head = ws.addRow([this.t('STT', 'No.'), this.t('Hạng mục', 'Item'), this.t('Thông số kỹ thuật', 'Specifications'), this.t('Đvt', 'Unit'), this.t('Số lượng', 'Qty'), this.t('Đơn giá', 'Price'), this.t('Thành tiền', 'Total Amount')]);
                head.eachCell(c => { c.font=fontBold; c.fill=fillHeader; c.border=borderThin; c.alignment=alignCenter; });
                head.height = 25;
                let rIdx = startCus + 1;
                
                this.docBaoGia.forEach((row, i) => {
                    const r = ws.addRow([]);
                    r.getCell(1).value = i + 1; r.getCell(2).value = str(row.name); r.getCell(3).value = str(row.specs); r.getCell(4).value = str(row.unit);
                    let q = num(row.qty); let p = num(row.priceGop);
                    if(this.roundDecimal){ q=Math.round(q); p=Math.round(p); }
                    r.getCell(5).value = q; r.getCell(6).value = p; r.getCell(7).value = q * p;
                    r.getCell(2).alignment=alignLeft; r.getCell(3).alignment=alignLeft; 
                    r.getCell(5).numFmt=FORMAT_MONEY; r.getCell(6).numFmt=FORMAT_MONEY; r.getCell(7).numFmt=FORMAT_MONEY; 
                    rIdx++;
                });
                
                const rT1 = ws.addRow([]); rT1.getCell(6).value = this.t('Tổng giá (trước thuế):', 'Subtotal:'); rT1.getCell(7).value = this.bgTotalBeforeVAT; ws.mergeCells(`A${rIdx}:E${rIdx}`); rT1.font=fontBold; rT1.getCell(1).alignment=alignRight; rT1.getCell(7).numFmt=FORMAT_MONEY; rIdx++;
                const rT2 = ws.addRow([]); rT2.getCell(6).value = this.t(`Thuế GTGT (VAT ${this.vatRate}%):`, `VAT (${this.vatRate}%):`); rT2.getCell(7).value = this.bgVATAmount; ws.mergeCells(`A${rIdx}:E${rIdx}`); rT2.font=fontBold; rT2.getCell(1).alignment=alignRight; rT2.getCell(7).numFmt=FORMAT_MONEY; rIdx++;
                const rT3 = ws.addRow([]); rT3.getCell(6).value = this.t('TỔNG GIÁ TRỊ (SAU THUẾ):', 'TOTAL AMOUNT:'); rT3.getCell(7).value = this.bgTotalAfterVAT; ws.mergeCells(`A${rIdx}:E${rIdx}`); rT3.font={bold:true, color:{argb:'FFDC2626'}, name:'Times New Roman', size:12}; rT3.getCell(1).alignment=alignRight; rT3.getCell(7).numFmt=FORMAT_MONEY; rIdx++;
                const rT4 = ws.addRow([]); rT4.getCell(1).value = `${this.t('Bằng chữ:', 'In words:')} ${this.readMoney(this.bgTotalAfterVAT)}`; ws.mergeCells(`A${rIdx}:G${rIdx}`); rT4.font={italic:true, name:'Times New Roman', size:11}; rT4.getCell(1).alignment=alignLeft; rIdx++;

                applyStyle(startCus, rIdx-5, 7);
                
                ws.addRow([]); rIdx++; 
                ws.addRow([this.t('ĐIỀU KHOẢN THƯƠNG MẠI (TERMS & CONDITIONS):', 'TERMS & CONDITIONS:')]).font = {bold:true, underline:true, name:'Times New Roman', size:11}; rIdx++;
                this.bgTerms.forEach(t => { 
                    const tr = ws.addRow([]); tr.getCell(1).value = '-'; tr.getCell(2).value = str(t.title); tr.getCell(3).value = str(t.content);
                    ws.mergeCells(`C${rIdx}:G${rIdx}`); tr.getCell(2).font=fontBold; tr.getCell(3).alignment=alignLeft; rIdx++; 
                });
                
                ws.addRow([]); rIdx++; 
                const signR = ws.addRow([]); signR.getCell(2).value = this.t('ĐẠI DIỆN KHÁCH HÀNG', 'CUSTOMER REPRESENTATIVE'); signR.getCell(6).value = this.t('ĐẠI DIỆN CÔNG TY', 'COMPANY REP.');
                ws.mergeCells(`B${rIdx}:C${rIdx}`); ws.mergeCells(`F${rIdx}:G${rIdx}`); signR.font=fontBold; signR.alignment={ vertical: 'top', horizontal: 'center' }; signR.height = 100;
            }
            
            /* ========================= TAB PR ========================= */
            else if (type === 'pr') {
                fileName = `PR_${str(this.prHeader.prNo) || 'No'}_${dateStr}.xlsx`;
                ws.columns = [ 
                    {width: 4}, {width: 15}, {width: 10}, {width: 10}, {width: 6}, {width: 6}, {width: 6}, {width: 6}, 
                    {width: 10}, {width: 12}, {width: 10}, {width: 12}, {width: 10}, {width: 10}, {width: 10}, {width: 10}, {width: 8}, {width: 12} 
                ];
                ws.pageSetup.printTitlesRow = '1:11';

                if(logoId !== null) ws.addImage(logoId, { tl: { col: 0.1, row: 0.1 }, ext: { width: 100, height: 40 } });
                
                let r1 = ws.addRow([]); r1.getCell(1).value = this.t('PHIẾU ĐỀ NGHỊ MUA HÀNG', 'PURCHASE REQUISITION'); r1.getCell(1).font={bold:true, name:'Times New Roman', size:18}; r1.getCell(1).alignment=alignCenter; r1.getCell(15).value = this.t('Số hiệu: ADG_QT_KH_01/BM01', 'Doc No: ADG_QT_KH_01/BM01'); r1.getCell(15).font=fontBold;
                let r2 = ws.addRow([]); r2.getCell(15).value = `${this.t('Ngày HL:', 'Effective:')} ${str(this.prHeader.isoDate)} | ${this.t('Lần BH:', 'Issue No:')} ${str(this.prHeader.isoRev)} | ${this.t('Lần SX:', 'Rev No:')} ${str(this.prHeader.isoCheck)}`;
                ws.mergeCells('A1:N2'); ws.mergeCells('O1:R1'); ws.mergeCells('O2:R2'); applyStyle(1, 2, 18);
                
                ws.addRow([]); 
                
                let r4 = ws.addRow([]); 
                r4.getCell(1).value = this.t('Số phiếu:', 'PR No:'); r4.getCell(1).font = fontBold; r4.getCell(1).alignment=alignRight;
                r4.getCell(2).value = `ADG_ĐNMH_${str(this.prHeader.prNo)}`; r4.getCell(2).font={bold:true, color:{argb:'FFDC2626'}, name:'Times New Roman', size:12}; 
                ws.mergeCells('B4:H4'); r4.getCell(2).border = borderBotOnly; r4.getCell(2).alignment=alignLeft;
                r4.getCell(9).value = this.prHeader.isPlan ? this.t('☑ Theo kế hoạch', '☑ Planned') : this.t('☐ Theo kế hoạch', '☐ Planned'); r4.getCell(9).font=fontBold; ws.mergeCells('I4:M4');
                
                let r5 = ws.addRow([]); r5.getCell(1).value = this.t('Đơn vị lập:', 'Company:'); r5.getCell(2).value = str(this.prHeader.company); ws.mergeCells('B5:H5'); r5.getCell(1).font=fontBold; r5.getCell(2).font=fontBold; r5.getCell(2).border = borderBotOnly; r5.getCell(1).alignment=alignRight; r5.getCell(2).alignment=alignLeft;
                let r6 = ws.addRow([]); r6.getCell(1).value = this.t('Bộ phận:', 'Department:'); r6.getCell(2).value = str(this.prHeader.dept); ws.mergeCells('B6:F6'); r6.getCell(1).font=fontBold; r6.getCell(2).font=fontBold; r6.getCell(2).border = borderBotOnly; r6.getCell(1).alignment=alignRight; r6.getCell(2).alignment=alignLeft;
                r6.getCell(9).value = this.t('Mã code:', 'Dept Code:'); r6.getCell(10).value = str(this.prHeader.deptCode); ws.mergeCells('J6:M6'); r6.getCell(9).font=fontBold; r6.getCell(10).font=fontBold; r6.getCell(10).border = borderBotOnly; r6.getCell(9).alignment=alignRight; r6.getCell(10).alignment=alignLeft;
                
                let r7 = ws.addRow([]); r7.getCell(1).value = this.t('Kho hàng:', 'Warehouse:'); r7.getCell(2).value = str(this.prHeader.warehouse); ws.mergeCells('B7:C7'); r7.getCell(1).font=fontBold; r7.getCell(2).border = borderBotOnly; r7.getCell(1).alignment=alignRight; r7.getCell(2).alignment=alignLeft;
                r7.getCell(5).value = this.t('Thủ kho:', 'Storekeeper:'); r7.getCell(6).value = str(this.prHeader.storekeeper); ws.mergeCells('F7:H7'); r7.getCell(5).font=fontBold; r7.getCell(6).border = borderBotOnly; r7.getCell(5).alignment=alignRight; r7.getCell(6).alignment=alignLeft;
                r7.getCell(9).value = this.t('Ngày duyệt:', 'Approved:'); r7.getCell(10).value = str(this.prHeader.approvedDate); ws.mergeCells('J7:M7'); r7.getCell(9).font=fontBold; r7.getCell(10).border = borderBotOnly; r7.getCell(9).alignment=alignRight; r7.getCell(10).alignment=alignLeft;
                ws.addRow([]);

                ws.addRow([
                    this.t('STT', 'No.'), this.t('Mô tả hàng hóa', 'Item Description'), '', this.t('Mã vật tư\n(*)', 'Mat. Code\n(*)'), this.t('ĐVT\n(*)', 'Unit\n(*)'), 
                    this.t('Bình quân SD hàng tháng\n(*)', 'Avg Monthly Usage\n(*)'), this.t('Số lượng', 'Quantity'), '', this.t('Đơn giá\ntham khảo\n(có VAT)\n(*)', 'Est. Price\n(VAT incl.)\n(*)'), 
                    this.t('Thành tiền\n(Có VAT)\n(*)', 'Total Amount\n(VAT incl.)\n(*)'), this.t('Thời gian\ncần hàng\n(*)', 'Req. Date\n(*)'), this.t('Mục đích sử dụng\n(*)', 'Purpose\n(*)'), 
                    this.t('TK chi phí\n(Người lập)', 'Expense Acc\n(Creator)'), this.t('TK chi phí\n(KTT)', 'Expense Acc\n(Chief Acct)'), this.t('Ngân sách\ncòn lại', 'Remaining\nBudget'), 
                    this.t('Nhà cung cấp tham khảo', 'Suggested\nSupplier'), this.t('Công ty\nsử dụng\n(*)', 'Company Use\n(*)'), this.t('Điều kiện nghiệm thu/ Đánh giá chất lượng (nếu có)', 'QC Conditions')
                ]);
                ws.addRow([
                    '', this.t('Tên hàng hóa\n(*)', 'Item Name\n(*)'), this.t('Quy cách\n(*)', 'Specs\n(*)'), '', '', '', this.t('Tồn Kho\n(*)', 'Stock\n(*)'), this.t('Đề xuất\n(*)', 'Request\n(*)'), 
                    '', '', '', '', '', '', '', '', '', ''
                ]);
                ws.addRow(['(1)','(2)','(3)','(4)','(5)','(6)','(7)','(8)','(9)','(10)','(11)','(12)','(13)','(14)','(15)','(16)','(17)','(18)']);
                
                ['A9:A10','B9:C9','D9:D10','E9:E10','F9:F10','G9:H9','I9:I10','J9:J10','K9:K10','L9:L10','M9:M10','N9:N10','O9:O10','P9:P10','Q9:Q10','R9:R10'].forEach(m => ws.mergeCells(m));
                [9, 10].forEach(r => ws.getRow(r).eachCell(c => { c.font=fontBold; c.fill=fillHeader; c.border=borderThin; c.alignment=alignCenter; }));
                ws.getRow(11).eachCell(c => { c.font={name:'Times New Roman', size:8, italic:true}; c.fill={type:'pattern', pattern:'solid', fgColor:{argb:'FFD9D9D9'}}; c.border=borderThin; c.alignment=alignCenter; });
                
                let rIdx = 12;
                this.docPR.forEach((row, i) => {
                    const r = ws.addRow([]);
                    let avg = num(row.avgUsage); let qS = num(row.qtyStock); let qR = num(row.qtyReq); let p = num(row.price); let bud = num(row.budget);
                    if(this.roundDecimal) { avg=Math.round(avg); qS=Math.round(qS); qR=Math.round(qR); p=Math.round(p); bud=Math.round(bud); }
                    
                    r.getCell(1).value = i+1; r.getCell(2).value = str(row.name); r.getCell(3).value = str(row.specs); r.getCell(4).value = str(row.partNo); r.getCell(5).value = str(row.unit); 
                    r.getCell(6).value = avg; r.getCell(7).value = qS; r.getCell(8).value = qR; r.getCell(9).value = p; r.getCell(10).value = qR * p; 
                    r.getCell(11).value = str(row.timeNeeded); r.getCell(12).value = str(row.purpose); r.getCell(13).value = str(row.accCreator); r.getCell(14).value = str(row.accKTT); r.getCell(15).value = bud; r.getCell(16).value = str(row.supplier); r.getCell(17).value = str(row.companyUse); r.getCell(18).value = str(row.qcCond);
                    
                    [2,3,12,16,18].forEach(c=>r.getCell(c).alignment=alignLeft);
                    r.getCell(4).font={bold:true, color:{argb:'FFDC2626'}, name:'Times New Roman', size:9};
                    [6,7,8,9,10,15].forEach(c => r.getCell(c).numFmt=FORMAT_MONEY); 
                    r.getCell(8).font=fontBold; r.getCell(10).font=fontBold;
                    rIdx++;
                });
                
                let rTot = ws.addRow([]); rTot.getCell(9).value = this.t('Tổng cộng:', 'Total:'); rTot.getCell(10).value = this.roundDecimal ? Math.round(this.prTotalAmount) : this.prTotalAmount; ws.mergeCells(`A${rIdx}:I${rIdx}`); rTot.getCell(1).alignment=alignRight; rTot.font=fontBold; rTot.getCell(10).numFmt=FORMAT_MONEY; rTot.getCell(10).font={bold:true, color:{argb:'FFDC2626'}, name:'Times New Roman', size:11}; rIdx++;
                let rWords = ws.addRow([]); rWords.getCell(9).value = `${this.t('Bằng chữ:', 'In words:')} ${this.readMoney(this.prTotalAmount)}`; ws.mergeCells(`A${rIdx}:R${rIdx}`); rWords.getCell(1).alignment=alignRight; rWords.font={italic:true, name:'Times New Roman', size:11}; rIdx++;
                applyStyle(12, rIdx-2, 18, 9);
                
                ws.addRow([]); rIdx++; 

                if (this.printPrWorkflow) {
                    ws.addRow([this.t('QUÁ TRÌNH XỬ LÝ:', 'WORKFLOW PROCESS:')]).font = {bold:true, name:'Times New Roman', size:11}; rIdx++;
                    let wH = ws.addRow([]); wH.getCell(1).value = this.t('Thời điểm xử lý', 'Timestamp'); wH.getCell(2).value = this.t('Bước xử lý', 'Process Step'); wH.getCell(5).value = this.t('Người xử lý', 'Person in Charge'); wH.getCell(9).value = this.t('Ý kiến người xử lý', 'Remarks/Comments'); wH.eachCell(c => { c.font=fontBold; c.fill=fillHeader; c.alignment=alignCenter; });
                    ws.mergeCells(`B${rIdx}:D${rIdx}`); ws.mergeCells(`E${rIdx}:H${rIdx}`); ws.mergeCells(`I${rIdx}:R${rIdx}`); applyStyle(rIdx, rIdx, 18); rIdx++;
                    
                    this.prWorkflow.forEach(s => {
                        let r = ws.addRow([]); r.getCell(1).value = str(s.time); r.getCell(2).value = this.lang === 'vi' ? str(s.stepVi) : str(s.stepEn); r.getCell(5).value = str(s.person); r.getCell(9).value = str(s.note);
                        ws.mergeCells(`B${rIdx}:D${rIdx}`); ws.mergeCells(`E${rIdx}:H${rIdx}`); ws.mergeCells(`I${rIdx}:R${rIdx}`); applyStyle(rIdx, rIdx, 18);
                        r.getCell(1).alignment=alignCenter; r.getCell(2).alignment=alignLeft; r.getCell(5).alignment=alignCenter; r.getCell(9).alignment=alignLeft;
                        rIdx++;
                    });
                    ws.addRow([]); rIdx++;
                }
                
                let s1 = ws.addRow([]); let s2 = ws.addRow([]);
                let dText = this.t('Ngày...tháng...năm...', 'Date...../...../.....');
                [1, 5, 8, 11, 14].forEach(c => s1.getCell(c).value = dText);
                s2.getCell(1).value = this.t('Người lập phiếu', 'Prepared by'); s2.getCell(5).value = this.t('Trưởng phòng', 'Dept. Manager'); s2.getCell(8).value = this.t('TP Cung ứng', 'Purchasing Mgr'); s2.getCell(11).value = this.t('Kế toán trưởng', 'Chief Accountant'); s2.getCell(14).value = this.t('Ban Giám đốc', 'Board of Directors');
                [`A${rIdx}:D${rIdx}`, `E${rIdx}:G${rIdx}`, `H${rIdx}:J${rIdx}`, `K${rIdx}:M${rIdx}`, `N${rIdx}:R${rIdx}`].forEach(m => ws.mergeCells(m)); rIdx++;
                [`A${rIdx}:D${rIdx}`, `E${rIdx}:G${rIdx}`, `H${rIdx}:J${rIdx}`, `K${rIdx}:M${rIdx}`, `N${rIdx}:R${rIdx}`].forEach(m => ws.mergeCells(m));
                s1.alignment = alignCenter; s2.alignment = { vertical: 'top', horizontal: 'center' }; s2.font = fontBold; s2.height = 80;
            }

            /* ========================= TAB DAILY (NHẬT KÝ) ========================= */
            else if (type === 'daily') {
                fileName = `Daily_Report_${dateStr}.xlsx`;
                ws.columns = [ {width: 6}, {width: 30}, {width: 25}, {width: 20}, {width: 30}, {width: 15} ];
                ws.pageSetup.printTitlesRow = '1:9';

                if(logoId !== null) ws.addImage(logoId, { tl: { col: 0.1, row: 0.1 }, ext: { width: 140, height: 50 } });

                ws.mergeCells('A1:F1'); ws.getCell('A1').value = this.t('BÁO CÁO NHẬT KÝ THI CÔNG', 'DAILY SITE REPORT').toUpperCase(); ws.getCell('A1').font=fontTitle; ws.getCell('A1').alignment=alignCenter;
                ws.mergeCells('A2:F2'); ws.getCell('A2').value = `${this.t('Dự án:', 'Project:')} ${str(projectName)}`; ws.getCell('A2').font=fontBold; ws.getCell('A2').alignment=alignCenter;
                ws.mergeCells('A3:F3'); ws.getCell('A3').value = this.dailyType === 'baotri' ? this.t('(Công tác Bảo trì & Khắc phục sự cố)', '(Maintenance & Troubleshooting)') : this.t('(Công tác Thi công & Lắp đặt)', '(Construction & Installation)'); ws.getCell('A3').alignment=alignCenter;
                ws.addRow([]);
                
                let h1 = ws.addRow([]); h1.getCell(1).value = this.t('Dự án/Hạng mục:', 'Project/Item:'); h1.getCell(2).value = str(projectName); h1.getCell(4).value = this.t('Ngày thực hiện:', 'Date:'); h1.getCell(5).value = str(this.dailyHeader.date);
                ws.mergeCells('B5:C5'); ws.mergeCells('E5:F5');
                let h2 = ws.addRow([]); h2.getCell(1).value = this.t('Địa chỉ:', 'Location:'); h2.getCell(2).value = str(this.dailyHeader.location); h2.getCell(4).value = this.t('Người phụ trách:', 'In Charge:'); h2.getCell(5).value = str(this.dailyHeader.pic);
                ws.mergeCells('B6:C6'); ws.mergeCells('E6:F6');
                [5, 6].forEach(r => { ws.getRow(r).getCell(1).font=fontBold; ws.getRow(r).getCell(4).font=fontBold; });
                ws.addRow([]);
                
                ws.addRow([this.t('1. NHÂN LỰC & THIẾT BỊ:', '1. MANPOWER & EQUIPMENT:')]).font=fontBold;
                const rHD1 = ws.addRow([this.t('STT', 'No.'), this.t('Vai trò / Chức danh', 'Role/Title'), this.t('Số lượng', 'Qty'), this.t('STT', 'No.'), this.t('Tên Thiết bị', 'Equipment'), this.t('Số lượng', 'Qty')]);
                rHD1.eachCell(c => { c.font=fontBold; c.fill=fillHeader; c.border=borderThin; c.alignment=alignCenter; });
                
                let rIdx = 9;
                const maxRows = Math.max(this.dailyManpower.length, this.dailyEquipment.length);
                for(let i=0; i<maxRows; i++) {
                    const mp = this.dailyManpower[i] || { role: '', qty: '' }; const eq = this.dailyEquipment[i] || { name: '', qty: '' };
                    let r = ws.addRow([]); r.getCell(1).value = mp.role ? i+1 : ''; r.getCell(2).value = str(mp.role); r.getCell(3).value = this.roundDecimal ? Math.round(num(mp.qty)) : num(mp.qty) || ''; r.getCell(4).value = eq.name ? i+1 : ''; r.getCell(5).value = str(eq.name); r.getCell(6).value = this.roundDecimal ? Math.round(num(eq.qty)) : num(eq.qty) || '';
                    r.getCell(3).numFmt=FORMAT_MONEY; r.getCell(6).numFmt=FORMAT_MONEY;
                    rIdx++;
                }
                applyStyle(9, rIdx-1, 6);
                
                ws.addRow([]); rIdx++; ws.addRow([this.t('2. CHI TIẾT CÔNG VIỆC:', '2. TASK DETAILS:')]).font=fontBold; rIdx++;
                const c2 = this.dailyType==='baotri' ? this.t('Khu vực / Tên Máy', 'Area / Machine') : this.t('Hạng mục', 'Construction Item');
                const c3 = this.dailyType==='baotri' ? this.t('Triệu chứng', 'Symptom') : this.t('Nội dung', 'Details');
                const c4 = this.dailyType==='baotri' ? this.t('Nguyên nhân (RCA)', 'Root Cause (RCA)') : this.t('Tiến độ (%)', 'Progress (%)');
                const c5 = this.dailyType==='baotri' ? this.t('Phương án khắc phục', 'Action Taken') : this.t('Vướng mắc', 'Site Issues');
                const rHD2 = ws.addRow([this.t('STT', 'No.'), c2, c3, c4, c5, this.t('Ghi chú', 'Remarks')]);
                rHD2.eachCell(c => { c.font=fontBold; c.fill=fillHeader; c.border=borderThin; c.alignment=alignCenter; }); rIdx++;
                
                const startT2 = rIdx;
                this.dailyTasks.forEach((tk, i) => {
                    let r = ws.addRow([]); r.getCell(1).value = i+1; r.getCell(2).value = str(tk.col1); r.getCell(3).value = str(tk.col2); r.getCell(4).value = str(tk.col3); r.getCell(5).value = str(tk.col4); r.getCell(6).value = str(tk.col5);
                    [2,3,4,5,6].forEach(c => r.getCell(c).alignment=alignLeft); rIdx++;
                });
                applyStyle(startT2-1, rIdx-1, 6);
                
                ws.addRow([]); rIdx++; 
                let sR = ws.addRow([]); sR.getCell(2).value = this.t('NGƯỜI LẬP BÁO CÁO', 'PREPARED BY'); sR.getCell(4).value = this.t('CHỈ HUY TRƯỞNG', 'SITE MANAGER'); sR.getCell(6).value = this.t('ĐẠI DIỆN KHÁCH HÀNG', 'CUSTOMER REP.');
                ws.mergeCells(`B${rIdx}:C${rIdx}`); ws.mergeCells(`F${rIdx}:H${rIdx}`); sR.font=fontBold; sR.alignment={ vertical: 'top', horizontal: 'center' }; sR.height = 80;
            }

            /* ========================= TAB NGHIỆM THU ========================= */
            else if (type === 'nghiemthu') {
                fileName = `Acceptance_${projectName}_${dateStr}.xlsx`;
                ws.columns = [ {width: 5}, {width: 20}, {width: 25}, {width: 20}, {width: 20}, {width: 15}, {width: 15}, {width: 15} ];
                ws.pageSetup.printTitlesRow = '1:8';

                if(logoId !== null) ws.addImage(logoId, { tl: { col: 0.1, row: 0.1 }, ext: { width: 140, height: 50 } });

                ws.mergeCells('A1:H1'); ws.getCell('A1').value = this.t('BIÊN BẢN NGHIỆM THU KỸ THUẬT', 'TECHNICAL ACCEPTANCE RECORD').toUpperCase(); ws.getCell('A1').font=fontTitle; ws.getCell('A1').alignment=alignCenter;
                ws.addRow([]);
                let h3 = ws.addRow([]); h3.getCell(1).value = this.t('Căn cứ Hợp đồng số:', 'Based on Contract No:'); h3.getCell(2).value = str(this.ntHeader.contract); h3.font=fontBold; ws.mergeCells('B3:H3');
                let h4 = ws.addRow([]); h4.getCell(1).value = this.t('Tên Dự án / Gói thầu:', 'Project / Package:'); h4.getCell(2).value = str(projectName); h4.font=fontBold; ws.mergeCells('B4:H4');
                let h5 = ws.addRow([]); h5.getCell(1).value = this.t('Thời gian nghiệm thu:', 'Acceptance Date:'); h5.getCell(2).value = str(this.ntHeader.date); h5.font=fontBold; ws.mergeCells('B5:H5');
                ws.addRow([]);
                
                let tbHead1 = ws.addRow([this.t('STT', 'No.'), this.t('Hạng mục', 'Description'), this.t('Hình ảnh thực tế', 'Actual Image'), this.t('Tiêu chí Đánh giá', 'Evaluation Criteria'), '', this.t('Hồ sơ', 'Doc.'), this.t('Kết luận', 'Status'), this.t('Ghi chú', 'Remarks')]);
                let tbHead2 = ws.addRow(['', '', '', this.t('Theo Hợp Đồng', 'By Contract'), this.t('Theo Thực Tế', 'Actual'), '', '', '']);
                ws.mergeCells('A7:A8'); ws.mergeCells('B7:B8'); ws.mergeCells('C7:C8'); ws.mergeCells('D7:E7'); ws.mergeCells('F7:F8'); ws.mergeCells('G7:G8'); ws.mergeCells('H7:H8');
                
                tbHead1.eachCell(c => { c.font=fontBold; c.fill=fillHeader; c.border=borderThin; c.alignment=alignCenter; });
                tbHead2.eachCell(c => { c.font=fontBold; c.fill=fillHeader; c.border=borderThin; c.alignment=alignCenter; });
                
                let rIdx = 9;
                this.docNghiemThu.forEach((row, i) => {
                    let r = ws.addRow([]);
                    r.getCell(1).value = i+1;
                    r.getCell(2).value = str(row.name);
                    r.getCell(3).value = ''; 
                    r.getCell(4).value = str(row.specContract);
                    r.getCell(5).value = str(row.specActual);
                    r.getCell(6).value = str(row.doc); 
                    r.getCell(7).value = str(row.status);
                    r.getCell(8).value = str(row.note);
                    
                    [2,4,5,6,8].forEach(c => r.getCell(c).alignment=alignLeft); 
                    
                    if (row.image) {
                        const imgData = this.getBase64Data(row.image);
                        if (imgData) {
                            const embedId = wb.addImage({ base64: imgData.base64, extension: imgData.ext });
                            ws.addImage(embedId, {
                                tl: { col: 2.1, row: rIdx - 1 + 0.1 },
                                br: { col: 2.9, row: rIdx - 0.1 },
                                editAs: 'oneCell'
                            });
                            r.height = 100;
                        }
                    }
                    rIdx++;
                });
                applyStyle(7, rIdx-1, 8);
                
                ws.addRow([]); rIdx++; 
                let sR = ws.addRow([]); sR.getCell(2).value = this.t('BÊN GIAO (NHÀ THẦU)', 'CONTRACTOR'); sR.getCell(4).value = this.t('TƯ VẤN GIÁM SÁT', 'CONSULTANT'); sR.getCell(6).value = this.t('BÊN NHẬN (CHỦ ĐẦU TƯ)', 'OWNER');
                ws.mergeCells(`B${rIdx}:C${rIdx}`); ws.mergeCells(`D${rIdx}:E${rIdx}`); ws.mergeCells(`F${rIdx}:H${rIdx}`); sR.font=fontBold; sR.alignment={ vertical: 'top', horizontal: 'center' }; sR.height = 80;
            }

            /* ========================= TAB NGHIỆM THU BÀN GIAO ========================= */
            else if (type === 'nghiemthubg') {
                fileName = `BBNT_BanGiao_${dateStr}.xlsx`;
                ws.columns = [ {width: 5}, {width: 15}, {width: 18}, {width: 35}, {width: 15}, {width: 15}, {width: 8}, {width: 8}, {width: 15} ];
                ws.pageSetup.printTitlesRow = '1:12';

                let r1 = ws.addRow([]); r1.getCell(1).value = this.t('BIÊN BẢN NGHIỆM THU', 'ACCEPTANCE RECORD'); r1.getCell(1).font={bold:true, name:'Times New Roman', size:18}; r1.getCell(1).alignment=alignCenter;
                let r2 = ws.addRow([]); r2.getCell(1).value = str(this.ntbgHeader.project).toUpperCase(); r2.getCell(1).font={bold:true, name:'Times New Roman', size:12}; r2.getCell(1).alignment=alignCenter;
                let r3 = ws.addRow([]); r3.getCell(1).value = str(this.ntbgHeader.monthYear); r3.getCell(1).font={italic:true, name:'Times New Roman', size:11}; r3.getCell(1).alignment=alignCenter;
                let r4 = ws.addRow([]); r4.getCell(1).value = str(this.ntbgHeader.refNo); r4.getCell(1).font=fontBold; r4.getCell(1).alignment=alignCenter;
                ws.mergeCells('A1:I1'); ws.mergeCells('A2:I2'); ws.mergeCells('A3:I3'); ws.mergeCells('A4:I4');
                ws.addRow([]);

                ws.addRow([this.t('I. THÀNH PHẦN:', 'I. PARTICIPANTS:')]).font = {bold:true, underline:true, name:'Times New Roman', size:11};
                
                let aTitle = ws.addRow([]); aTitle.getCell(2).value = this.t('BÊN A (BÊN THỰC HIỆN):', 'PARTY A (CONTRACTOR):'); aTitle.getCell(4).value = str(this.ntbgHeader.benA_company);
                aTitle.getCell(2).font=fontBold; aTitle.getCell(4).font=fontBold; ws.mergeCells('B7:C7'); ws.mergeCells('D7:I7');
                
                let aR1 = ws.addRow([]); aR1.getCell(2).value = this.t('- Ông/Bà:', '- Mr/Ms:'); aR1.getCell(3).value = str(this.ntbgHeader.benA_rep1_name); aR1.getCell(5).value = this.t('Chức vụ:', 'Position:'); aR1.getCell(6).value = str(this.ntbgHeader.benA_rep1_title); ws.mergeCells('C8:D8'); ws.mergeCells('F8:I8');
                let aR2 = ws.addRow([]); aR2.getCell(2).value = this.t('- Ông/Bà:', '- Mr/Ms:'); aR2.getCell(3).value = str(this.ntbgHeader.benA_rep2_name); aR2.getCell(5).value = this.t('Chức vụ:', 'Position:'); aR2.getCell(6).value = str(this.ntbgHeader.benA_rep2_title); ws.mergeCells('C9:D9'); ws.mergeCells('F9:I9');
                
                let bTitle = ws.addRow([]); bTitle.getCell(2).value = this.t('BÊN B (BÊN THUÊ DỊCH VỤ):', 'PARTY B (OWNER):'); bTitle.getCell(4).value = str(this.ntbgHeader.benB_company);
                bTitle.getCell(2).font=fontBold; bTitle.getCell(4).font=fontBold; ws.mergeCells('B10:C10'); ws.mergeCells('D10:I10');
                
                let bR1 = ws.addRow([]); bR1.getCell(2).value = this.t('- Ông/Bà:', '- Mr/Ms:'); bR1.getCell(3).value = str(this.ntbgHeader.benB_rep1_name); bR1.getCell(5).value = this.t('Chức vụ:', 'Position:'); bR1.getCell(6).value = str(this.ntbgHeader.benB_rep1_title); ws.mergeCells('C11:D11'); ws.mergeCells('F11:I11');
                let bR2 = ws.addRow([]); bR2.getCell(2).value = this.t('- Ông/Bà:', '- Mr/Ms:'); bR2.getCell(3).value = str(this.ntbgHeader.benB_rep2_name); bR2.getCell(5).value = this.t('Chức vụ:', 'Position:'); bR2.getCell(6).value = str(this.ntbgHeader.benB_rep2_title); ws.mergeCells('C12:D12'); ws.mergeCells('F12:I12');

                ws.addRow([this.t('II. NỘI DUNG:', 'II. CONTENT:')]).font = {bold:true, underline:true, name:'Times New Roman', size:11};
                let base = ws.addRow([]); base.getCell(2).value = this.t('Căn cứ vào:', 'Based on:'); base.getCell(3).value = str(this.ntbgHeader.baseOn); ws.mergeCells('C14:I14');
                let conc = ws.addRow([]); conc.getCell(2).value = str(this.ntbgHeader.conclusion); conc.font={italic:true, name:'Times New Roman', size:11}; ws.mergeCells('B15:I15');

                let tbHead1 = ws.addRow([this.t('STT', 'No.'), this.t('Ngày đề nghị', 'Request Date'), this.t('Thiết bị', 'Equipment'), this.t('Hạng mục thi công', 'Construction Item'), this.t('Thời gian', 'Duration'), '', this.t('Xác nhận nghiệm thu', 'Acceptance Confirm'), '', this.t('Ghi chú', 'Remarks')]);
                let tbHead2 = ws.addRow(['', '', '', '', this.t('Bắt đầu', 'Start'), this.t('Kết thúc', 'End'), this.t('Đạt', 'Pass'), this.t('Không đạt', 'Fail'), '']);
                ws.mergeCells('A16:A17'); ws.mergeCells('B16:B17'); ws.mergeCells('C16:C17'); ws.mergeCells('D16:D17'); ws.mergeCells('E16:F16'); ws.mergeCells('G16:H16'); ws.mergeCells('I16:I17');
                
                tbHead1.eachCell(c => { c.font=fontBold; c.fill=fillHeader; c.border=borderThin; c.alignment=alignCenter; });
                tbHead2.eachCell(c => { c.font=fontBold; c.fill=fillHeader; c.border=borderThin; c.alignment=alignCenter; });

                let rIdx = 18;
                this.docNghiemThuBG.forEach((row, i) => {
                    let r = ws.addRow([]);
                    r.getCell(1).value = i+1;
                    r.getCell(2).value = str(row.requestDate);
                    r.getCell(3).value = str(row.equipment);
                    r.getCell(4).value = str(row.description);
                    r.getCell(5).value = str(row.timeStart);
                    r.getCell(6).value = str(row.timeEnd);
                    r.getCell(7).value = str(row.statusPass);
                    r.getCell(8).value = str(row.statusFail);
                    r.getCell(9).value = str(row.note);
                    
                    r.getCell(1).alignment=alignCenter; r.getCell(2).alignment=alignCenter; r.getCell(5).alignment=alignCenter; r.getCell(6).alignment=alignCenter; r.getCell(7).alignment=alignCenter; r.getCell(8).alignment=alignCenter;
                    r.getCell(3).alignment=alignLeft; r.getCell(4).alignment=alignLeft; r.getCell(9).alignment=alignLeft;
                    rIdx++;
                });
                applyStyle(16, rIdx-1, 9);
                
                ws.addRow([]); rIdx++; 
                let sR = ws.addRow([]); sR.getCell(2).value = this.t('BÊN A', 'PARTY A'); sR.getCell(7).value = this.t('BÊN B', 'PARTY B');
                ws.mergeCells(`B${rIdx}:D${rIdx}`); ws.mergeCells(`G${rIdx}:I${rIdx}`); sR.font=fontBold; sR.alignment={ vertical: 'top', horizontal: 'center' }; sR.height = 80;
            }

            const buffer = await wb.xlsx.writeBuffer();
            saveAs(new Blob([buffer], { type: "application/octet-stream" }), fileName);
            setTimeout(() => this.showSavedToast = false, 2500);
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
                            
                            let date = rest[0] || '';
                            let eq = rest[1] || '';
                            let desc = rest[2] || '';
                            let timeStart = rest[3] || '';
                            let timeEnd = rest[4] || '';
                            let note = '';
                            
                            if(rest.length > 5) {
                                let lastItem = rest[rest.length - 1];
                                if (lastItem.toLowerCase() !== 'x' && lastItem.toLowerCase() !== 'đạt' && lastItem.toLowerCase() !== 'pass') note = lastItem;
                            }

                            newDoc.push({
                                requestDate: date,
                                equipment: eq,
                                description: desc,
                                timeStart: timeStart,
                                timeEnd: timeEnd,
                                statusPass: 'x', 
                                statusFail: '',
                                note: note
                            });
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

        async hashPassword(message) {
            const msgBuffer = new TextEncoder().encode(message);
            const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
            const hashArray = Array.from(new Uint8Array(hashBuffer)); return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
        },
        
        async loginAdmin() { 
            const pass = prompt(this.t("Nhập mã số NV (Password):", "Enter Staff ID (Password):")); if (!pass) return;
            const inputHash = await this.hashPassword(pass); 
            const targetPass = atob("MjEwOTAwNA==");
            const VALID_HASH = await this.hashPassword(targetPass); 
            
            if (inputHash === VALID_HASH) { 
                this.isAdmin = true; 
                alert(this.t("Truy cập thành công!", "Access Granted!"));
            } else { 
                alert(this.t("Từ chối truy cập! Mật khẩu sai.", "Access Denied! Wrong password.")); 
            }
        },
        
        async migrateDataToIndexedDB() {
            const keys = ['masterDataProject', 'global_project_name', 'dt_table', 'bg_table', 'bg_header', 'bg_terms', 'pr_table', 'pr_flow', 'nt_table', 'company_logo_base64', 'pr_header', 'pr_flow_print_setting', 'global_vat_rate'];
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
            if (this.roundDecimal) {
                num = Math.round(num);
                return num.toLocaleString('vi-VN', { maximumFractionDigits: 0 });
            }
            return num.toLocaleString('vi-VN', { maximumFractionDigits: 3 }); 
        },
        
        focusNum(e, val) { 
            e.target.value = val === 0 ? '' : val.toString().replace('.', ','); 
            e.target.select(); 
        },
        
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
            let str = parseInt(Math.abs(num)) + ''; 
            let arr = []; 
            while (str.length > 0) { 
                arr.push(str.length >= 3 ? str.slice(-3) : str); 
                str = str.length >= 3 ? str.slice(0, -3) : ''; 
            }
            let rsString = '';
            for (let i = 0; i < arr.length; i++) {
                let chuoi = ''; 
                let temp = arr[i].padStart(3, '0'); 
                let tram = parseInt(temp[0]); 
                let chuc = parseInt(temp[1]); 
                let donVi = parseInt(temp[2]);
                if (tram === 0 && chuc === 0 && donVi === 0) continue;
                if (tram !== 0 || i !== arr.length - 1) { 
                    chuoi += chuHangTram[tram] + ' trăm '; 
                    if (chuc === 0 && donVi !== 0) chuoi += 'lẻ '; 
                }
                if (chuc !== 0 && chuc !== 1) { 
                    chuoi += chuHangChuc[chuc] + ' mươi '; 
                    if (chuc === 4 && donVi !== 0) chuoi = chuoi.replace('bốn mươi', 'bốn mươi'); 
                } else if (chuc === 1) { chuoi += 'mười '; }
                switch (donVi) { 
                    case 1: chuoi += (chuc !== 0 && chuc !== 1) ? 'mốt ' : chuHangDonVi[donVi] + ' '; break; 
                    case 5: chuoi += (chuc !== 0) ? 'lăm ' : chuHangDonVi[donVi] + ' '; break; 
                    case 4: chuoi += (chuc !== 0 && chuc !== 1) ? 'tư ' : chuHangDonVi[donVi] + ' '; break; 
                    default: if (donVi !== 0) chuoi += chuHangDonVi[donVi] + ' '; break; 
                }
                rsString = chuoi + mangGiaTri[i] + ' ' + rsString;
            }
            rsString = rsString.trim().replace(/\s+/g, ' '); 
            if (rsString.startsWith('không trăm lẻ ')) rsString = rsString.substring(14); 
            else if (rsString.startsWith('không trăm ')) rsString = rsString.substring(11);
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
                        const canvas = document.createElement('canvas'); 
                        let width = img.width; 
                        let height = img.height; 
                        const MAX_SIZE = 600;
                        if (width > height) { 
                            if (width > MAX_SIZE) { height *= MAX_SIZE / width; width = MAX_SIZE; } 
                        } else { 
                            if (height > MAX_SIZE) { width *= MAX_SIZE / height; height = MAX_SIZE; } 
                        } 
                        canvas.width = width; 
                        canvas.height = height; 
                        const ctx = canvas.getContext('2d'); 
                        ctx.drawImage(img, 0, 0, width, height); 
                        row.image = canvas.toDataURL('image/jpeg', 0.5); 
                    }; 
                    img.src = e.target.result; 
                }; 
                reader.readAsDataURL(file); 
            } 
            event.target.value = ''; 
        },
        
        uploadLogo(event) { 
            const file = event.target.files[0]; 
            if (file) { 
                const reader = new FileReader(); 
                reader.onload = (e) => { 
                    this.companyLogo = e.target.result; 
                    localforage.setItem('company_logo_base64', this.companyLogo); 
                }; 
                reader.readAsDataURL(file); 
            } 
        },
        
        selectItemFromSearch(item, targetArray) {
            let newItem = { ...item, qty: 1 }; newItem.priceGop = newItem.priceVT + newItem.priceNC; 
            if (this.currentTab === 'pr') { 
                targetArray.push({ name: newItem.name, specs: newItem.specs, partNo: newItem.partNo, unit: newItem.unit, avgUsage: 0, qtyStock: 0, qtyReq: 1, price: newItem.priceVT, timeNeeded: '', purpose: 'Dự án / Bảo trì', accCreator: '', accKTT: '', budget: 0, supplier: '', companyUse: 'ADG', qcCond: 'Đúng thông số kỹ thuật yêu cầu' }); 
            } 
            else { targetArray.push(newItem); }
            this.searchQuery = ''; this.showSearchDropdown = false;
        },

        saveDraft(key, data) { 
            if (this.draftTimeouts[key]) { clearTimeout(this.draftTimeouts[key]); }
            this.draftTimeouts[key] = setTimeout(() => {
                localforage.setItem(key, JSON.parse(JSON.stringify(data))).then(() => { 
                    this.showSavedToast = true; 
                    setTimeout(() => this.showSavedToast = false, 2000); 
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
            const cachedLogo = await parseL('company_logo_base64', null); if(cachedLogo) this.companyLogo = cachedLogo;

            this.ntbgHeader = await parseL('ntbg_header', this.ntbgHeader);
            this.docNghiemThuBG = await parseL('ntbg_table', this.docNghiemThuBG);
        },
        
        clearDraft(t) { 
            if(!confirm(this.t("CẢNH BÁO: Xóa toàn bộ dữ liệu form hiện tại?", "WARNING: Clear all data in this form?"))) return; 
            if(t==='dutoan') this.docDuToan=[]; 
            if(t==='baogia') { 
                this.docBaoGia=[]; 
                this.bgTerms = [ 
                    { title: 'Thời gian giao hàng:', content: 'Trong vòng 15 ngày làm việc kể từ ngày nhận được tạm ứng hợp đồng.' }, 
                    { title: 'Thời gian bảo hành:', content: '12 tháng theo tiêu chuẩn của nhà sản xuất.' }, 
                    { title: 'Thanh toán:', content: 'Lần 1: Tạm ứng 30%. Lần 2: 70% sau khi bàn giao nghiệm thu.' },
                    { title: 'Hiệu lực báo giá:', content: '30 ngày kể từ ngày ban hành.' } 
                ]; 
            }
            if(t==='pr') this.docPR=[]; 
            if(t==='daily') { this.dailyTasks=[]; this.dailyManpower=[]; this.dailyEquipment=[]; } 
            if(t==='nghiemthu') this.docNghiemThu=[]; 
            if(t==='nghiemthubg') this.docNghiemThuBG=[];
        },
        
        async addToDatabase() { 
            if(!this.newItem.name) return alert(this.t("Vui lòng nhập Tên Hạng Mục", "Please enter Description"));
            this.database.unshift({...this.newItem}); 
            await localforage.setItem('masterDataProject', JSON.parse(JSON.stringify(this.database)));
            this.newItem = { partNo: '', name: '', specs: '', unit: '', priceVT: 0, priceNC: 0 }; 
            alert(this.t("Thêm dữ liệu thành công!", "Added successfully!"));
        },
        
        importFromExcel(event) {
            const file = event.target.files[0]; if (!file) return; const reader = new FileReader();
            reader.onload = async (e) => {
                const data = new Uint8Array(e.target.result); 
                const workbook = XLSX.read(data, { type: 'array' }); 
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]]; 
                const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                let newDb = []; let headerFound = false; let colIdx = {};
                
                const parseCleanNumber = (val) => { 
                    if (typeof val === 'number') return val; if (!val) return 0; 
                    let s = val.toString().trim().replace(/[^0-9\.,\-]/g, ''); if(!s) return 0; 
                    let lastDot = s.lastIndexOf('.'); let lastComma = s.lastIndexOf(','); 
                    if (lastComma > lastDot && lastDot !== -1) { s = s.replace(/\./g, '').replace(/,/g, '.'); } 
                    else if (lastDot > lastComma && lastComma !== -1) { s = s.replace(/,/g, ''); } 
                    else { 
                        if (s.indexOf('.') !== -1) { 
                            let parts = s.split('.'); if (parts.length > 2 || parts[1].length === 3) s = s.replace(/\./g, ''); 
                        } else if (s.indexOf(',') !== -1) { 
                            let parts = s.split(','); if (parts.length > 2 || parts[1].length === 3) s = s.replace(/,/g, ''); else s = s.replace(/,/g, '.'); 
                        } 
                    } 
                    let num = parseFloat(s); return isNaN(num) ? 0 : num; 
                };

                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i]; if (!row || row.length === 0) continue;
                    if (!headerFound) { 
                        let rowStr = row.join('').toLowerCase(); 
                        if (rowStr.includes('mã vật tư') && rowStr.includes('tên vật tư')) { 
                            headerFound = true; 
                            colIdx.partNo = row.findIndex(c => c && c.toString().toLowerCase().includes('mã vật tư')); 
                            colIdx.name = row.findIndex(c => c && c.toString().toLowerCase().includes('tên vật tư')); 
                            colIdx.unit = row.findIndex(c => c && c.toString().toLowerCase().includes('đvt')); 
                            colIdx.qty = row.findIndex(c => c && c.toString().toLowerCase().includes('số lượng')); 
                            colIdx.val = row.findIndex(c => c && c.toString().toLowerCase().includes('giá trị')); 
                        } 
                        continue; 
                    }
                    const name = row[colIdx.name] ? row[colIdx.name].toString().trim() : ''; 
                    if (!name || name.toLowerCase().includes('tổng cộng')) continue;
                    const partNo = row[colIdx.partNo] ? row[colIdx.partNo].toString().trim() : ''; 
                    const unit = row[colIdx.unit] ? row[colIdx.unit].toString().trim() : ''; 
                    const qty = colIdx.qty !== -1 ? parseCleanNumber(row[colIdx.qty]) : 0; 
                    const val = colIdx.val !== -1 ? parseCleanNumber(row[colIdx.val]) : 0; 
                    const priceVT = qty > 0 ? (val / qty) : 0;
                    newDb.push({ partNo: partNo, name: name, specs: '', unit: unit, priceVT: priceVT, priceNC: 0 });
                }
                if (newDb.length > 0) { 
                    this.database = newDb; 
                    await localforage.setItem('masterDataProject', JSON.parse(JSON.stringify(newDb))); 
                    alert(this.t(`Cập nhật thành công ${newDb.length} vật tư.`, `Updated ${newDb.length} items.`)); 
                } else { 
                    alert('LỖI: Cấu trúc file sai.'); 
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
        
        removeRow(arr, i) { 
            if(confirm(this.t("Xóa dòng này khỏi Bảng?", "Delete this row?"))) arr.splice(i, 1); 
        },
        
        addTerm() { 
            this.bgTerms.push({ title: 'Điều khoản mới:', content: '...' }); 
        }, 
        
        removeTerm(i) { 
            this.bgTerms.splice(i, 1); 
        },
        
        copyFromDuToan() { 
            if(confirm(this.t("Gộp dữ liệu từ BOM sang Báo Giá?", "Import from BOM?"))) { 
                this.docBaoGia = this.docDuToan.map(i => ({
                    ...i, 
                    partNo: i.partNo || '', 
                    name: i.name || '', 
                    specs: i.specs || '', 
                    unit: i.unit || '', 
                    qty: i.qty || 1, 
                    priceGop: (((i.priceVT * i.qty * (1 + this.tyleQLCVT/100)) + (i.priceNC * i.qty * (1 + this.tyleNCGT/100) * (1 + this.tyleQLCNC/100))) / i.qty) || 0 
                })); 
            } 
        },
        
        copyFromDuToanToPR() {
            if(confirm(this.t("Lấy dữ liệu từ BOM sang Phiếu PR?", "Import from BOM to PR?"))) {
                const materialsOnly = this.docDuToan.filter(i => i.priceVT > 0 || (i.priceVT === 0 && i.priceNC === 0));
                const newPRItems = materialsOnly.map(i => ({ 
                    name: i.name || '', specs: i.specs || '', partNo: i.partNo || '', unit: i.unit || '', 
                    avgUsage: 0, qtyStock: 0, qtyReq: i.qty || 1, price: i.priceVT || 0, timeNeeded: '', 
                    purpose: 'Dự án / Bảo trì', accCreator: '', accKTT: '', budget: 0, supplier: '', 
                    companyUse: 'ADG', qcCond: 'Đúng thông số kỹ thuật yêu cầu' 
                }));
                this.docPR = [...this.docPR, ...newPRItems]; 
                alert(this.t(`Đã thêm ${newPRItems.length} mã vật tư.`, `Added ${newPRItems.length} items.`));
            }
        }
    }
}).mount('#app');
