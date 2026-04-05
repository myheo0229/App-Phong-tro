/**
 * PHIẾU THU TIỀN NHÀ - Template Generator
 * Tái hiện chính xác từ XML gốc của file mẫu.
 *
 * Page: A5 ngang (landscape), margins 2.54cm tất cả 4 phía
 * Tab chính: 16cm (9072 DXA), dot leader, căn phải
 * Tab phụ:   ~6.5cm (3686 DXA), dot leader, căn phải (dùng cho dòng công thức điện/nước)
 *
 * Để dùng trong project: import hàm `generatePhieu(data)` và gọi với object data.
 */

const {
    Document, Packer, Paragraph, TextRun,
    AlignmentType, TabStopType,
} = require('docx');
const fs   = require('fs');
const path = require('path');

// ─── Dữ liệu mẫu ─────────────────────────────────────────────────────────────
// LƯU Ý: Các biến này PHẢI KHỚP với dữ liệu được truyền trong main.js hàm export-invoice
const sampleData = {
    phone:              '0982 141 407',      // Số điện thoại chủ nhà
    room_name:          '1A',                // Tên phòng
    day:                '01',                // Ngày hiện tại
    month:              '04',                // Tháng
    year:               '2026',              // Năm
    room_fee:           '2.200.000',         // Tiền phòng (đã format)
    electricity_old:    '4924',              // Chỉ số điện cũ
    electricity_new:    '4986',              // Chỉ số điện mới
    electricity_used:   '62',                // Số kWh tiêu thụ
    electricity_price:  '2.900',             // Đơn giá điện/kWh (đã format)
    electricity_amount: '180.000',           // Thành tiền điện (đã format)
    electricity_loss:   '13.000',           // Tiền điện hao tải (đã format)
    water_old:          '539',               // Chỉ số nước cũ
    water_new:          '545',               // Chỉ số nước mới
    water_used:         '6',                 // Số m³ tiêu thụ
    water_price:        '12.000',            // Đơn giá nước/m³ (đã format)
    water_amount:       '72.000',            // Thành tiền nước (đã format)
    garbage_fee:        '40.000',            // Tiền rác (đã format)
    internet_fee:       '24.000',            // Tiền internet (đã format)
    total:              '2.529.000',         // Tổng cộng (đã format)
};

// ─── Hằng số (lấy chính xác từ XML gốc) ──────────────────────────────────────
const FONT     = 'Times New Roman';
const TAB_MAIN = 9072; // 16cm — tab duy nhất hầu hết các dòng
const TAB_MID  = 3686; // ~6.5cm — tab đầu dòng công thức điện/nước

// ─── Helper functions ─────────────────────────────────────────────────────────

/** TextRun thông thường */
function r(text, bold = false, size = 22) {
    return new TextRun({ text, font: FONT, size, bold });
}

/** Ký tự "đ" superscript (đúng như file gốc) */
function superD(size = 22) {
    return new TextRun({ text: 'đ', font: FONT, size, bold: true, superScript: true });
}

/** Tab character */
const TAB = new TextRun({ text: '\t', font: FONT });

/** Tab stops cho dòng thông thường: chỉ 1 tab tại 9072 (dot) */
function mainTabs() {
    return [{ type: TabStopType.RIGHT, position: TAB_MAIN, leader: 'dot' }];
}

/** Tab stops cho dòng công thức điện/nước: 3686 (dot) + 9072 (dot) */
function formulaTabs() {
    return [
        { type: TabStopType.RIGHT, position: TAB_MID,  leader: 'dot' },
        { type: TabStopType.RIGHT, position: TAB_MAIN, leader: 'dot' },
    ];
}

/** Spacing mặc định cho mọi đoạn */
function sp(afterPt = 160) {
    return { after: afterPt, line: 276, lineRule: 'auto' };
}

// ─── Hàm tạo document ─────────────────────────────────────────────────────────
function generatePhieu(data) {
    return new Document({
        styles: {
            default: {
                document: {
                    run:       { font: FONT, size: 22 },
                    paragraph: { spacing: sp() }
                }
            }
        },
        sections: [{
            properties: {
                page: {
                    // A5 landscape: width=21cm=11906 DXA, height=14.8cm=8391 DXA
                    // Truyền portrait dimensions, docx-js tự swap khi có orientation
                    size: {
                        width:       11906,
                        height:      8391,
                        orientation: 'landscape',
                    },
                    // Margins 2.54cm = 1440 DXA (đúng theo yêu cầu)
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
                }
            },
            children: [

                // ── Dòng 1: tiêu đề + số điện thoại (center, bold, 12pt) ──────
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    tabStops:  mainTabs(),
                    spacing:   sp(),
                    children: [
                        new TextRun({
                            text: `PHIẾU THU TIỀN NHÀ   đt: ${data.phone}`,
                            font: FONT, size: 24, bold: true,
                        }),
                    ]
                }),

                // ── Dòng 2: phòng số (center, bold, 12pt) ────────────────────
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    tabStops:  mainTabs(),
                    spacing:   sp(),
                    children: [
                        new TextRun({
                            text: `Phòng số: ${data.room_name}`,
                            font: FONT, size: 24, bold: true,
                        }),
                    ]
                }),

                // ── Dòng 3: ngày tháng (right, italic, 12pt) ─────────────────
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    tabStops:  mainTabs(),
                    spacing:   sp(),
                    children: [
                        new TextRun({
                            text: `Ngày ${data.day} tháng ${data.month} năm ${data.year}`,
                            font: FONT, size: 24, italics: true,
                        }),
                    ]
                }),

                // ── Tiền phòng: ……\t=   Xđ ──────────────────────────────────
                // Pattern: label \t =   số superD
                new Paragraph({
                    tabStops: mainTabs(),
                    spacing:  sp(),
                    children: [
                        r('Tiền phòng: '),
                        TAB,
                        r('=     ', true),
                        r(data.room_fee, true),
                        superD(),
                    ]
                }),

                // ── Tiền điện dòng 1: label số_mới dots số_cũ \t ─────────────
                // KHÔNG có amount ở cuối, kết thúc bằng \t (dots chạy đến hết dòng)
                new Paragraph({
                    tabStops: mainTabs(),
                    spacing:  { after: 0, line: 276, lineRule: 'auto' },
                    children: [
                        r('Tiền điện: Số mới: ………'),
                        r(data.electricity_new, true),
                        r('……………..Số cũ: …………'),
                        r(data.electricity_old, true),
                        TAB,
                    ]
                }),

                // ── Tiền điện dòng 2: \t= X  x  Yđ  \t=  Zđ ─────────────────
                // \t → dots đến 3686, rồi "= X x Yđ", \t → dots đến 9072, rồi "= Zđ"
                new Paragraph({
                    tabStops: formulaTabs(),
                    spacing:  sp(),
                    children: [
                        TAB,
                        r(`=      ${data.electricity_used}`, true),
                        r(`       x       ${data.electricity_price}`, true),
                        superD(), r(' ', true),
                        TAB,
                        r(`=   ${data.electricity_amount}`, true),
                        superD(),
                    ]
                }),

                // ── Tiền nước dòng 1 ─────────────────────────────────────────
                new Paragraph({
                    tabStops: mainTabs(),
                    spacing:  { after: 0, line: 276, lineRule: 'auto' },
                    children: [
                        r('Tiền nước:  Số mới: …….'),
                        r(data.water_new, true),
                        r('……………..Số cũ: ………….'),
                        r(data.water_old, true),
                        TAB,
                    ]
                }),

                // ── Tiền nước dòng 2 ─────────────────────────────────────────
                new Paragraph({
                    tabStops: formulaTabs(),
                    spacing:  sp(),
                    children: [
                        TAB,
                        r(`=      ${data.water_used}`, true),
                        r(`       x       ${data.water_price}`, true),
                        superD(), r('   ', true),
                        TAB,
                        r(`=   ${data.water_amount}`, true),
                        superD(),
                    ]
                }),

                // ── Tiền rác ─────────────────────────────────────────────────
                new Paragraph({
                    tabStops: mainTabs(),
                    spacing:  sp(),
                    children: [
                        r('Tiền rác: '),
                        TAB,
                        r(`=   ${data.garbage_fee}`, true),
                        superD(),
                    ]
                }),

                // ── Internet ─────────────────────────────────────────────────
                new Paragraph({
                    tabStops: mainTabs(),
                    spacing:  sp(),
                    children: [
                        r('Internet: '),
                        TAB,
                        r(`=   ${data.internet_fee}`, true),
                        superD(),
                    ]
                }),

                // ── Tiền điện hao tải ────────────────────────────────────────
                new Paragraph({
                    tabStops: mainTabs(),
                    spacing:  sp(),
                    children: [
                        r('Tiền điện hao tải: '),
                        TAB,
                        r(`=   ${data.electricity_loss}`, true),
                        superD(),
                    ]
                }),

                // ── Tổng cộng (size 14pt = 28 half-pt, đúng file gốc) ────────
                new Paragraph({
                    tabStops: mainTabs(),
                    spacing:  sp(),
                    children: [
                        r('Tổng cộng:'),
                        TAB,
                        new TextRun({ text: data.total, font: FONT, size: 28, bold: true }),
                        new TextRun({ text: 'đ', font: FONT, size: 28, bold: true, superScript: true }),
                    ]
                }),
            ]
        }]
    });
}

// ─── Chạy trực tiếp ───────────────────────────────────────────────────────────
const doc     = generatePhieu(sampleData);
const outPath = path.join(__dirname, 'PHIẾU THU TIỀN NHÀ_output.docx');

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(outPath, buffer);
    console.log('✅ Tạo file thành công:', outPath);
}).catch(err => {
    console.error('❌ Lỗi:', err);
});

module.exports = { generatePhieu };