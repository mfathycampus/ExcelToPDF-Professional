"""Arabic / English strings. Issue codes from the core are translated here, so the core
never formats user-facing text (and never embeds cell contents)."""
from __future__ import annotations

STRINGS: dict[str, dict[str, str]] = {
    # app / UI
    "app_title": {"ar": "Excel to PDF Professional", "en": "Excel to PDF Professional"},
    "open_file": {"ar": "فتح ملف Excel…", "en": "Open Excel file…"},
    "drop_hint": {"ar": "اسحب ملف Excel وأفلته هنا\n(.xlsx، .xls)", "en": "Drop an Excel file here\n(.xlsx, .xls)"},
    "inspect": {"ar": "فحص المصنف", "en": "Inspect workbook"},
    "preview": {"ar": "معاينة PDF", "en": "Preview PDF"},
    "export": {"ar": "تصدير…", "en": "Export…"},
    "cancel": {"ar": "إلغاء", "en": "Cancel"},
    "sheets": {"ar": "الأوراق", "en": "Sheets"},
    "settings": {"ar": "إعدادات الطباعة", "en": "Print settings"},
    "report": {"ar": "تقرير التحقق", "en": "Verification report"},
    "paper": {"ar": "حجم الورق", "en": "Paper size"},
    "mode": {"ar": "وضع التخطيط", "en": "Layout mode"},
    "min_font": {"ar": "أصغر خط مسموح (pt)", "en": "Minimum font (pt)"},
    "measured_min_font": {"ar": "أصغر خط في PDF (مقاس): {pt} pt", "en": "Smallest font in PDF (measured): {pt} pt"},
    "est_pages": {"ar": "صفحات (تقدير)", "en": "Pages (est.)"},
    "long_text": {"ar": "النصوص الطويلة جدًا", "en": "Very long text"},
    "long_appendix": {"ar": "ملحق في آخر الملف (موصى به)", "en": "Appendix at the end (recommended)"},
    "long_wrap": {"ar": "لفّ داخل الخلية", "en": "Wrap inside the cell"},
    "orientation": {"ar": "الاتجاه", "en": "Orientation"},
    "portrait": {"ar": "عمودي", "en": "Portrait"},
    "landscape": {"ar": "أفقي", "en": "Landscape"},
    "scale": {"ar": "المقياس", "en": "Scale"},
    "margins": {"ar": "الهوامش", "en": "Margins"},
    "top": {"ar": "أعلى", "en": "Top"},
    "bottom": {"ar": "أسفل", "en": "Bottom"},
    "left": {"ar": "يسار", "en": "Left"},
    "right": {"ar": "يمين", "en": "Right"},
    "title_rows": {"ar": "صفوف العناوين المتكررة", "en": "Repeat rows"},
    "title_cols": {"ar": "أعمدة المفتاح المتكررة", "en": "Repeat columns"},
    "include_hidden": {"ar": "تضمين الأوراق المخفية", "en": "Include hidden sheets"},
    "allow_a3": {"ar": "السماح بورق A3", "en": "Allow A3 paper"},
    "engine": {"ar": "محرك التحويل", "en": "Engine"},
    "language": {"ar": "English", "en": "العربية"},
    "page_of": {"ar": "صفحة {n} من {total} — {sheet}", "en": "Page {n} of {total} — {sheet}"},
    "zoom_in": {"ar": "تكبير", "en": "Zoom in"},
    "zoom_out": {"ar": "تصغير", "en": "Zoom out"},
    "fit": {"ar": "ملاءمة العرض", "en": "Fit width"},
    "prev": {"ar": "السابق", "en": "Previous"},
    "next": {"ar": "التالي", "en": "Next"},
    "open_pdf": {"ar": "فتح PDF", "en": "Open PDF"},
    "open_folder": {"ar": "فتح المجلد", "en": "Open folder"},
    "overwrite_q": {"ar": "الملف موجود. هل تريد استبداله؟", "en": "The file exists. Replace it?"},
    "failed_save_q": {"ar": "لم يجتز الملف التحقق. هل تريد حفظه رغم ذلك؟ سيُوسم التقرير بأنه غير مكتمل.",
                      "en": "Verification failed. Save anyway? The report will mark it as not verified."},
    "engines": {"ar": "المحركات", "en": "Engines"},
    "engine_missing_long": {"ar": "لا يوجد محرك تحويل على هذا الجهاز.\n\nالبرنامج يحتاج LibreOffice (مجاني) أو Microsoft Excel.\nيمكنك تثبيت LibreOffice الآن، ثم الضغط على «المحركات» للتحقق.",
                            "en": "No conversion engine on this computer.\n\nThe app needs LibreOffice (free) or Microsoft Excel.\nInstall LibreOffice now, then click “Engines” to re-check."},
    "install_lo": {"ar": "تثبيت LibreOffice", "en": "Install LibreOffice"},
    "open_lo_site": {"ar": "صفحة التحميل", "en": "Download page"},
    "later": {"ar": "لاحقًا", "en": "Later"},
    "install_started": {"ar": "بدأ التثبيت في نافذة منفصلة. عند انتهائه اضغط «المحركات».", "en": "Installation started in a separate window. When it finishes, click “Engines”."},
    "engine_ready": {"ar": "المحرك: {name} — جاهز", "en": "Engine: {name} — ready"},
    "engine_missing": {"ar": "لا يوجد محرك تحويل. ثبّت LibreOffice من libreoffice.org",
                       "en": "No engine found. Install LibreOffice from libreoffice.org"},
    "visible": {"ar": "ظاهرة", "en": "visible"},
    "hidden": {"ar": "مخفية", "en": "hidden"},
    "veryHidden": {"ar": "مخفية بالكامل", "en": "very hidden"},
    # states
    "state_idle": {"ar": "لم يتم اختيار ملف", "en": "No file selected"},
    "state_inspecting": {"ar": "جارٍ فحص الملف…", "en": "Inspecting…"},
    "state_inspected": {"ar": "اكتمل الفحص", "en": "Inspection complete"},
    "state_exporting": {"ar": "جارٍ إنشاء PDF…", "en": "Creating PDF…"},
    "state_verifying": {"ar": "جارٍ التحقق…", "en": "Verifying…"},
    "state_retrying": {"ar": "تجربة خطة بديلة…", "en": "Trying alternative plan…"},
    "state_merging": {"ar": "تجميع الصفحات…", "en": "Assembling pages…"},
    "state_done": {"ar": "اكتمل", "en": "Done"},
    "state_cancelled": {"ar": "أُلغي العمل", "en": "Cancelled"},
    "state_failed": {"ar": "فشل", "en": "Failed"},
    # modes
    "auto": {"ar": "تلقائي موصى به", "en": "Automatic (recommended)"},
    "fit_width": {"ar": "ملاءمة العرض", "en": "Fit to width"},
    "one_page": {"ar": "صفحة واحدة", "en": "One page"},
    "keep_original": {"ar": "حفظ إعدادات Excel", "en": "Keep Excel settings"},
    "manual": {"ar": "يدوي", "en": "Manual"},
    # status
    "passed": {"ar": "اجتاز التحقق", "en": "Passed verification"},
    "passed_with_warnings": {"ar": "اجتاز مع تحذيرات", "en": "Passed with warnings"},
    "failed": {"ar": "لم يجتز التحقق", "en": "Did not pass verification"},
    # issue codes
    "file_not_found": {"ar": "الملف غير موجود.", "en": "File not found."},
    "format_unsupported": {"ar": "صيغة الملف غير مدعومة في هذا الإصدار.", "en": "This file format is not supported in this version."},
    "format_experimental": {"ar": "دعم هذه الصيغة تجريبي.", "en": "Support for this format is experimental."},
    "file_empty": {"ar": "الملف فارغ.", "en": "The file is empty."},
    "file_too_large": {"ar": "حجم الملف يتجاوز الحد المسموح.", "en": "The file exceeds the size limit."},
    "password_protected": {"ar": "الملف محمي بكلمة مرور ولا يمكن فتحه.", "en": "The file is password protected."},
    "signature_mismatch": {"ar": "محتوى الملف لا يطابق امتداده.", "en": "File content does not match its extension."},
    "corrupt_file": {"ar": "الملف تالف أو لا يمكن قراءته.", "en": "The file is damaged or unreadable."},
    "zip_bomb_suspected": {"ar": "بنية الملف غير طبيعية (حجم مفكوك مفرط).", "en": "Abnormal file structure (excessive uncompressed size)."},
    "zip_too_many_entries": {"ar": "بنية الملف غير طبيعية.", "en": "Abnormal file structure."},
    "zip_bad_path": {"ar": "بنية الملف غير آمنة.", "en": "Unsafe file structure."},
    "macros_present_not_executed": {"ar": "يحتوي المصنف على وحدات ماكرو — لن تُشغَّل.", "en": "Workbook contains macros — they will not run."},
    "external_links_not_updated": {"ar": "روابط خارجية ({count}) — لن تُحدَّث، تُستخدم القيم المحفوظة.", "en": "External links ({count}) — not updated; cached values are used."},
    "sheet_hidden": {"ar": "ورقة مخفية.", "en": "Hidden sheet."},
    "shapes_need_review": {"ar": "أشكال ({count}) — قد يختلف عرضها؛ راجع المعاينة.", "en": "Shapes ({count}) — rendering may differ; review the preview."},
    "xls_objects_not_inspected": {"ar": "ملف XLS: لم تُفحص الصور والمخططات مسبقًا.", "en": "XLS file: images/charts were not pre-inspected."},
    "no_sheets_to_export": {"ar": "لا توجد أوراق للتصدير.", "en": "No sheets to export."},
    "engine_not_available": {"ar": "محرك التحويل غير مثبت.", "en": "The conversion engine is not installed."},
    "no_engine_available": {"ar": "لا يوجد محرك تحويل. ثبّت LibreOffice.", "en": "No conversion engine. Install LibreOffice."},
    "engine_unknown": {"ar": "محرك تحويل غير معروف.", "en": "Unknown conversion engine."},
    "engine_timeout": {"ar": "تجاوزت العملية المهلة المحددة.", "en": "The operation timed out."},
    "engine_crashed": {"ar": "توقف محرك التحويل بشكل غير متوقع.", "en": "The engine stopped unexpectedly."},
    "engine_failed": {"ar": "فشل محرك التحويل في فتح الملف أو تصديره.", "en": "The engine failed to open or export the file."},
    "engine_start_failed": {"ar": "تعذر تشغيل محرك التحويل.", "en": "Could not start the engine."},
    "engine_open_failed": {"ar": "تعذر على المحرك فتح الملف.", "en": "The engine could not open the file."},
    "sheet_export_failed": {"ar": "فشل تصدير الورقة.", "en": "Sheet export failed."},
    "sheet_not_processed": {"ar": "لم تُعالَج الورقة.", "en": "Sheet was not processed."},
    "nothing_to_print": {"ar": "لا يوجد محتوى قابل للطباعة.", "en": "Nothing to print."},
    "no_pages_exported": {"ar": "لم تُنتج أي صفحة.", "en": "No pages were produced."},
    "text_missing": {"ar": "{missing} خلية لم يُعثر على نصها في PDF.", "en": "{missing} cell(s) not found in the PDF text."},
    "rtl_order_unverified": {"ar": "{count} خلية عربية: كل حروفها موجودة في الصفحة لكن ترتيبها في طبقة النص لم يمكن إثباته — راجعها بصريًا.", "en": "{count} RTL cell(s): all characters are on the page but their order in the text layer could not be proven — review visually."},
    "cell_overflow_hash": {"ar": "{count} خلية تظهر ### (العمود أضيق من القيمة).", "en": "{count} cell(s) show ### (column too narrow)."},
    "font_below_min": {"ar": "حجم النص ({min_font_pt} pt) أصغر من الحد ({threshold} pt).", "en": "Text size ({min_font_pt} pt) is below the minimum ({threshold} pt)."},
    "font_below_min_user_choice": {"ar": "اختيارك أدى إلى نص أصغر من الحد ({min_font_pt} pt).", "en": "Your choice produced text below the minimum ({min_font_pt} pt)."},
    "small_text_present": {"ar": "يوجد نص صغير جدًا بنسبة محدودة ({min_font_pt} pt).", "en": "A small amount of text is very small ({min_font_pt} pt)."},
    "blank_pages": {"ar": "صفحات فارغة: {pages}.", "en": "Blank pages: {pages}."},
    "content_outside_page": {"ar": "نص خارج حدود الصفحة ({count}).", "en": "Text outside the page ({count})."},
    "images_not_found_in_pdf": {"ar": "لم يُعثر على الصور المتوقعة في PDF.", "en": "Expected images were not found in the PDF."},
    "charts_not_found_in_pdf": {"ar": "لم يُعثر على المخططات المتوقعة في PDF.", "en": "Expected charts were not found in the PDF."},
    "shapes_not_verified": {"ar": "الأشكال ({count}) لا يمكن التحقق منها آليًا.", "en": "Shapes ({count}) cannot be verified automatically."},
    "coverage_sampled": {"ar": "فُحصت عينة من {checked} خلية (ورقة كبيرة).", "en": "A sample of {checked} cells was checked (large sheet)."},
    "chartsheet_visual_only": {"ar": "ورقة مخطط — تحقق بصري فقط.", "en": "Chart sheet — visual check only."},
    "fallback_plan_used": {"ar": "استُخدمت خطة بديلة بعد فشل الأولى.", "en": "An alternative plan was used after the first failed."},
    "overflow_text_wrapped": {"ar": "لُفَّ النص الطويل في {count} خلية (في نسخة العمل فقط) حتى لا يُقص.", "en": "Long text in {count} cell(s) was wrapped (working copy only) so it is not clipped."},
    "long_text_to_appendix": {"ar": "نُقل نص {count} خلية طويلة جدًا إلى «ملحق النصوص الطويلة»، والخلية تعرض رقم الإحالة.", "en": "Text of {count} very long cell(s) moved to the long-text appendix; the cell shows its reference number."},
    "appendix_title": {"ar": "ملحق النصوص الطويلة", "en": "Long text appendix"},
    "cards": {"ar": "بطاقة لكل سجل", "en": "Card per record"},
    "cards_suffix": {"ar": "بطاقات", "en": "cards"},
    "row": {"ar": "صف", "en": "row"},
    "cards_layout": {"ar": "طُبعت الورقة بطاقة لكل سجل (حقل: قيمة) بناءً على اختيارك؛ القيم كما يعرضها Excel.", "en": "Sheet printed as one card per record (field: value) as you chose; values as displayed."},
    "cards_visuals_not_included": {"ar": "وضع البطاقات لا يتضمن الصور والمخططات والأشكال الموجودة في الورقة.", "en": "Card mode does not include the sheet's images, charts and shapes."},
    "cards_suggested": {"ar": "الورقة عريضة ({parts} أجزاء أفقية لكل صفحة). جرّب وضع «بطاقة لكل سجل» لقراءة كل سجل كاملًا في مكان واحد (قارن عدد الصفحات المقدَّر).", "en": "Very wide sheet ({parts} horizontal parts per page). Try “card per record” to read each record in one place (compare the estimated pages)."},
    "formulas_computed_by_engine": {"ar": "{count} صيغة بلا نتيجة محفوظة في الملف — حسبها المحرك عند التحويل.", "en": "{count} formula(s) had no saved result — computed by the engine during conversion."},
    "error_values_present": {"ar": "{count} خلية قيمتها خطأ (مثل #N/A) — تُطبع كما هي.", "en": "{count} cell(s) contain error values (e.g. #N/A) — printed as they are."},
    "comments_not_printed": {"ar": "{count} تعليق/ملاحظة على الخلايا لا تُطبع (مثل سلوك Excel الافتراضي).", "en": "{count} cell comment(s)/note(s) are not printed (Excel's default behaviour)."},
    "horizontal_split": {"ar": "قُسمت الورقة أفقيًا للحفاظ على حجم الخط.", "en": "Sheet split horizontally to keep text readable."},
    "no_key_column_detected": {"ar": "لم يُكتشف عمود مفتاح لتكراره في الأجزاء الأفقية.", "en": "No key column detected to repeat across horizontal parts."},
    "original_setup_is_readable": {"ar": "أُبقيت إعدادات Excel الأصلية لأنها مقروءة.", "en": "Original Excel setup kept (readable)."},
    "no_cell_grid_keep_original": {"ar": "أُبقيت الإعدادات الأصلية.", "en": "Original setup kept."},
    "one_page_text_below_minimum": {"ar": "وضع الصفحة الواحدة سيجعل النص أصغر من الحد.", "en": "One-page mode makes text smaller than the minimum."},
    "fit_width_text_below_minimum": {"ar": "ملاءمة العرض ستجعل النص أصغر من الحد.", "en": "Fit-to-width makes text smaller than the minimum."},
    "fallback_explicit_min_scale": {"ar": "خطة بديلة: مقياس ثابت عند الحد الأدنى.", "en": "Fallback: fixed scale at the minimum."},
    "delivery_blocked_verification_failed": {"ar": "أُوقف التسليم لأن التحقق فشل.", "en": "Delivery blocked because verification failed."},
    "source_modified": {"ar": "تغيّر الملف الأصلي أثناء المعالجة!", "en": "The source file changed during processing!"},
    "cancelled": {"ar": "أُلغي العمل.", "en": "Cancelled."},
    "excluded_hidden": {"ar": "مخفية", "en": "hidden"},
    "excluded_empty": {"ar": "فارغة", "en": "empty"},
    "excluded_not_selected": {"ar": "غير محددة", "en": "not selected"},
}


class Translator:
    def __init__(self, lang: str = "ar"):
        self.lang = lang

    @property
    def rtl(self) -> bool:
        return self.lang == "ar"

    def __call__(self, key: str, **kw) -> str:
        entry = STRINGS.get(key)
        text = entry.get(self.lang, entry.get("en", key)) if entry else key
        if kw:
            try:
                return text.format(**kw)
            except (KeyError, IndexError, ValueError):
                return text
        return text

    def issue(self, issue) -> str:
        detail = {k: v for k, v in (issue.detail or {}).items() if not isinstance(v, (list, dict))}
        if "pages" in issue.detail:
            detail["pages"] = ", ".join(map(str, issue.detail["pages"]))
        return self(issue.code, **detail)
