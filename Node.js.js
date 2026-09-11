const express = require('express');
const PptxGenJS = require('pptxgenjs');
const ExcelJS = require('exceljs');
const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// --- 1. التحكم في حالة التسجيل (فتح / قفل) ---
let registrationStatus = { isOpen: true }; // الحالة الافتراضية

app.post('/api/admin/toggle-registration', (-req, res) => {
    registrationStatus.isOpen = !registrationStatus.isOpen;
    res.json({ success: true, isOpen: registrationStatus.isOpen, message: registrationStatus.isOpen ? 'تم فتح التسجيل بنجاح' : 'تم إغلاق التسجيل بنجاح' });
});

// --- 2. لوحة التحكم: تعديل صورة الطالب أو منحه الإذن ---
app.post('/api/admin/update-student/:id', async (req, res) => {
    const studentId = req.params.id;
    const { image_url, can_edit_image } = req.body;
    
    // تحديث قاعدة البيانات (كمثال توضيحي)
    // await Database.query('UPDATE students SET image_url = ?, can_edit_image = ? WHERE id = ?', [image_url, can_edit_image, studentId]);
    
    res.json({ success: true, message: 'تم تحديث بيانات الطالب وصلاحياته بنجاح' });
});

// --- 3. إنشاء وتحميل ملف الـ PowerPoint بتصميمات مخصصة (فاتح مبهج / غامق احترافي) ---
app.post('/api/admin/generate-ppt/:id', async (req, res) => {
    const studentId = req.params.id;
    const { theme } = req.body; // 'light-cheerful' أو 'dark-elegant'
    
    // جلب بيانات الطالب من القاعدة (كمثال تجريبي)
    const student = {
        name: "أحمد محمد علي",
        major: "هندسة الحاسبات والذكاء الاصطناعي",
        gpa: "3.85",
        speech: "شكراً لكل من ساعدني في رحلتي التعليمية، أهلي وأساتذتي الأفاضل.",
        image_url: "https://via.placeholder.com/150" 
    };

    let pptx = new PptxGenJS();
    // ضبط اتجاه الشريحة ليكون ملائماً للغة العربية إذا لزم
    let slide = pptx.addSlide();

    // تخصيص الألوان بناءً على الاختيار
    if (theme === 'light-cheerful') {
        // تصميم فاتح ومبهج (ألوان حفلات تخرج مبهجة: أبيض، سماوي فاتح، لمسات ذهبية/برتقالية)
        slide.background = { color: "F4F9FF" };
        
        // بطاقة خلفية بيضاء للشريحة
        slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 0.8, w: 8.4, h: 5.4, fill: { color: "FFFFFF" }, line: { color: "E1E8F0", width: 1 } });
        
        // شريط تزييني مبهج بالأعلى
        slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 0.8, w: 8.4, h: 0.6, fill: { color: "FFA500" } });

        // النصوص (من اليمين لليسار RTL)
        slide.addText("حفل التخرج السنوي 🎓", { x: 1.0, y: 1.0, w: 8.0, h: 0.4, fontSize: 16, color: "FFFFFF", bold: true, align: "right" });
        slide.addText(student.name, { x: 1.5, y: 1.8, w: 5.5, h: 0.6, fontSize: 24, color: "2C3E50", bold: true, align: "right" });
        slide.addText(`التخصص: ${student.major}`, { x: 1.5, y: 2.5, w: 5.5, h: 0.4, fontSize: 14, color: "7F8C8D", align: "right" });
        slide.addText(`المعدل التراكمي: ${student.gpa}`, { x: 1.5, y: 3.0, w: 5.5, h: 0.4, fontSize: 14, color: "27AE60", bold: true, align: "right" });
        
        // كلمة الخريج
        slide.addText(`" ${student.speech} "`, { x: 1.5, y: 3.6, w: 7.0, h: 1.2, fontSize: 13, color: "34495E", italic: true, align: "right" });

    } else {
        // تصميم غامق وأنيق لحفلات التخرج (كحلي ملكي داكن مع ذهبي)
        slide.background = { color: "111827" }; // خلفية داكنة فاخرة

        // بطاقة داخلية
        slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 0.8, w: 8.4, h: 5.4, fill: { color: "1F2937" }, line: { color: "D4AF37", width: 1.5 } });

        slide.addText(student.name, { x: 1.5, y: 1.6, w: 5.5, h: 0.6, fontSize: 26, color: "F3F4F6", bold: true, align: "right" });
        slide.addText(`التخصص: ${student.major}`, { x: 1.5, y: 2.3, w: 5.5, h: 0.4, fontSize: 14, color: "D4AF37", align: "right" });
        slide.addText(`المعدل التراكمي: ${student.gpa}`, { x: 1.5, y: 2.8, w: 5.5, h: 0.4, fontSize: 14, color: "10B981", bold: true, align: "right" });
        
        // كلمة الخريج
        slide.addText(`كلمة الخريج:\n" ${student.speech} "`, { x: 1.5, y: 3.4, w: 7.0, h: 1.3, fontSize: 13, color: "E5E7EB", italic: true, align: "right" });
    }

    // إضافة صورة الطالب في الجانب الأيسر لتناسق التصميم الـ RTL
    try {
        slide.addImage({ path: student.image_url, x: 7.2, y: 1.8, w: 1.6, h: 1.6, rounding: true });
    } catch (e) {
        // صورة افتراضية في حال عدم توفر الرابط
    }

    // حفظ وإرسال الملف للمشرف
    let fileName = `Graduation_${studentId}.pptx`;
    await pptx.writeFile({ outputType: 'nodebuffer' }).then(buffer => {
        res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
        res.send(buffer);
    });
});

// --- 4. تقرير أكسيل (Excel) لأرقام المقاعد بناءً على ترتيب الدفعة ---
app.get('/api/admin/export-seats-report', async (req, res) => {
    // جلب الطلاب مرتبين تناسلياً أو تصاعدياً حسب معدلاتهم أو تاريخ تسجليهم (ترتيب الدفعة)
    // const students = await Database.query('SELECT name, major FROM students ORDER BY gpa DESC');
    
    // بيانات تجريبية لمحاكاة ترتيب الدفعة
    const students = [
        { name: 'محمد أحمد', major: 'حاسبات' },
        { name: 'سارة خالد', major: 'إدارة أعمال' },
        { name: 'محمود حسن', major: 'هندسة' }
    ];

    let workbook = new ExcelJS.Workbook();
    let sheet = workbook.addWorksheet('مقاعد الخريجين', { views: [{ rightToLeft: true }] }); // تفعيل اليمين لليسار في الإكسيل

    // تنسيق الأعمدة
    sheet.columns = [
        { header: 'رقم المقعد (الترتيب)', key: 'seat_number', width: 20 },
        { header: 'اسم الطالب', key: 'name', width: 30 },
        { header: 'التخصص', key: 'major', width: 30 }
    ];

    // إدخال البيانات تلقائياً مع رقم المقعد الذي يطابق ترتيب الطالب في الدفعة
    students.forEach((student, index) => {
        sheet.addRow({
            seat_number: index + 1, // رقم المقعد هو ترتيبه في الدفعة
            name: student.name,
            major: student.major
        });
    });

    // تنسيق صف العناوين
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F2937' } };

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Graduation_Seats_Report.xlsx');

    await workbook.xlsx.write(res);
    res.end();
});