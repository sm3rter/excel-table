/* English / Arabic strings. Add a language by adding one more object to STRINGS and a button label in app.js. */
(function (root) {
  'use strict';

  const STRINGS = {
    en: {
      'app.title': 'Excel Table',
      'welcome.title': 'Open a spreadsheet',
      'welcome.lead': 'View, edit and convert xlsx, xls, csv and json files. Nothing is uploaded: your file stays on this device.',
      'welcome.choose': 'Choose a file',
      'welcome.sample': 'Try the sample',
      'welcome.drop': 'You can also drop a file anywhere on this page.',
      'welcome.formats': 'Supported: xlsx, xls, csv, tsv and json.',
      'drop.overlay': 'Drop to open',
      'bar.open': 'Open file',
      'bar.language': 'Language',
      'formula.placeholder': 'Tap a cell to edit it',
      'formula.label': 'Cell value',
      'search.placeholder': 'Search rows',
      'header.toggle': 'First row is header',
      'row.add': 'Add row',
      'export': 'Export',
      'pager.info': '{from}–{to} of {total} rows',
      'pager.none': 'No rows to show',
      'pager.prev': 'Previous page',
      'pager.next': 'Next page',
      'menu.row.title': 'Row {n}',
      'menu.row.insert': 'Insert row below',
      'menu.row.delete': 'Delete row',
      'menu.col.title': 'Column {name}',
      'menu.col.asc': 'Sort A to Z',
      'menu.col.desc': 'Sort Z to A',
      'menu.col.clear': 'Clear sort',
      'menu.col.insert': 'Insert column after',
      'menu.col.delete': 'Delete column',
      'col.aria': 'Column {name} options',
      'row.aria': 'Row {n} options',
      'export.title': 'Export',
      'export.xlsx': 'Excel workbook (.xlsx)',
      'export.xlsx.hint': '{n} sheets',
      'export.xlsx.hint1': '1 sheet',
      'export.csv': 'CSV (.csv)',
      'export.csv.hint': 'Current sheet. Opens correctly in Excel, including Arabic.',
      'export.json': 'JSON (.json)',
      'export.json.hintHeader': 'Current sheet. The first row becomes the keys.',
      'export.json.hintPlain': 'Current sheet as a list of rows.',
      'dialog.close': 'Close',
      'toast.rowDeleted': 'Row deleted',
      'toast.colDeleted': 'Column deleted',
      'toast.undo': 'Undo',
      'toast.exported': 'Saved {name}',
      'busy': 'Opening…',
      'err.open': 'Could not open this file. {reason}',
      'err.xlsBad': 'This is not a valid Excel .xls file.',
      'lang.title': 'Language',
      'export.xls': 'Excel 97-2003 (.xls)',
      'export.xls.hint': 'Older Excel. Max 65,536 rows and 256 columns',
      'err.big': 'This file is larger than 100 MB.',
      'err.type': 'Unsupported file type. Use xlsx, xls, csv, tsv or json.',
      'err.export': 'Could not export: {reason}',
      'sheet.empty': 'Sheet',
    },
    ar: {
      'app.title': 'جدول إكسل',
      'welcome.title': 'افتح ملف جداول',
      'welcome.lead': 'اعرض وعدّل وحوّل ملفات xlsx وxls وcsv وjson. لا يُرفع أي شيء: ملفك يبقى على جهازك.',
      'welcome.choose': 'اختر ملفًا',
      'welcome.sample': 'جرّب العيّنة',
      'welcome.drop': 'ويمكنك أيضًا إسقاط الملف في أي مكان في هذه الصفحة.',
      'welcome.formats': 'المدعوم: xlsx وxls وcsv وtsv وjson.',
      'drop.overlay': 'أفلِت الملف لفتحه',
      'bar.open': 'فتح ملف',
      'bar.language': 'اللغة',
      'formula.placeholder': 'اضغط على خلية لتعديلها',
      'formula.label': 'قيمة الخلية',
      'search.placeholder': 'ابحث في الصفوف',
      'header.toggle': 'الصف الأول عناوين',
      'row.add': 'إضافة صف',
      'export': 'تصدير',
      'pager.info': '{from}–{to} من {total} صف',
      'pager.none': 'لا توجد صفوف للعرض',
      'pager.prev': 'الصفحة السابقة',
      'pager.next': 'الصفحة التالية',
      'menu.row.title': 'الصف {n}',
      'menu.row.insert': 'إدراج صف بالأسفل',
      'menu.row.delete': 'حذف الصف',
      'menu.col.title': 'العمود {name}',
      'menu.col.asc': 'ترتيب تصاعدي',
      'menu.col.desc': 'ترتيب تنازلي',
      'menu.col.clear': 'إلغاء الترتيب',
      'menu.col.insert': 'إدراج عمود بعده',
      'menu.col.delete': 'حذف العمود',
      'col.aria': 'خيارات العمود {name}',
      'row.aria': 'خيارات الصف {n}',
      'export.title': 'تصدير',
      'export.xlsx': 'مصنف إكسل (.xlsx)',
      'export.xlsx.hint': '{n} أوراق',
      'export.xlsx.hint1': 'ورقة واحدة',
      'export.csv': 'ملف CSV (.csv)',
      'export.csv.hint': 'الورقة الحالية. يُفتح في إكسل بشكل صحيح مع العربية.',
      'export.json': 'ملف JSON (.json)',
      'export.json.hintHeader': 'الورقة الحالية. الصف الأول يصبح أسماء الحقول.',
      'export.json.hintPlain': 'الورقة الحالية كقائمة صفوف.',
      'dialog.close': 'إغلاق',
      'toast.rowDeleted': 'حُذف الصف',
      'toast.colDeleted': 'حُذف العمود',
      'toast.undo': 'تراجع',
      'toast.exported': 'تم حفظ {name}',
      'busy': 'جارٍ الفتح…',
      'err.open': 'تعذّر فتح هذا الملف. {reason}',
      'err.xlsBad': 'هذا ليس ملف إكسل xls صالحًا.',
      'lang.title': 'اللغة',
      'export.xls': 'إكسل 97-2003 (.xls)',
      'export.xls.hint': 'إكسل القديم. حتى 65,536 صفًا و256 عمودًا',
      'err.big': 'حجم الملف أكبر من 100 ميجابايت.',
      'err.type': 'نوع الملف غير مدعوم. استخدم xlsx أو xls أو csv أو tsv أو json.',
      'err.export': 'تعذّر التصدير: {reason}',
      'sheet.empty': 'ورقة',
    },
  };

  let current = 'en';

  function t(key, vars) {
    let s = (STRINGS[current] && STRINGS[current][key]) || STRINGS.en[key] || key;
    if (vars) s = s.replace(/\{(\w+)\}/g, (m, k) => (vars[k] !== undefined ? vars[k] : m));
    return s;
  }

  function setLang(lang) {
    current = STRINGS[lang] ? lang : 'en';
    document.documentElement.lang = current;
    document.documentElement.dir = current === 'ar' ? 'rtl' : 'ltr';
  }

  /** Fill every element marked with data-i18n, data-i18n-placeholder or data-i18n-label. */
  function apply() {
    document.querySelectorAll('[data-i18n]').forEach((el) => { el.textContent = t(el.dataset.i18n); });
    document.querySelectorAll('[data-i18n-placeholder]').forEach((el) => { el.placeholder = t(el.dataset.i18nPlaceholder); });
    document.querySelectorAll('[data-i18n-label]').forEach((el) => { el.setAttribute('aria-label', t(el.dataset.i18nLabel)); });
  }

  root.I18n = { t, setLang, apply, get lang() { return current; } };
})(window);
