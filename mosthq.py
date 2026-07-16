def extract_eligible_only_records(doc):
    records = {}
    for table in doc.tables:
        for row in table.rows:
            # استخراج النصوص من الخلايا
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            
            # التأكد من أن الجدول يحتوي على 6 أعمدة على الأقل كما طلبت
            if len(cells) >= 6:
                # تجاوز صف العناوين
                if "اسم" in cells[3] or "المركز" in cells[0]: continue
                
                seq = cells[0]
                # new_card = cells[1] # تم تجاهله حسب طلبك
                old_card = cells[2]
                name = cells[3]
                # cells[4] حقل فارغ أو بيانات غير مهمة
                eligible_str = cells[5]
                
                # التأكد من أن رقم البطاقة صالح
                if old_card.isdigit() and len(old_card) >= 4:
                    try:
                        # تصفية أي نصوص زائدة وسحب الرقم فقط
                        el_val = int(''.join(filter(str.isdigit, eligible_str)))
                        records[old_card] = {
                            "seq": seq,
                            "name": name,
                            "total": 0,       # تصفير الكلي لتجاهله
                            "eligible": el_val,
                            "withheld": 0     # تصفير المحجوب لتجاهله
                        }
                    except ValueError:
                        continue
    return records
