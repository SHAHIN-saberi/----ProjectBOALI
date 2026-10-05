# Prompt — شروع پیاده‌سازی XLSX Symbol Counter با Repo Agent

Repository را بررسی کن و `AGENT.md` را به‌عنوان Specification قطعی پروژه در نظر بگیر.

اکنون implementation واقعی را شروع کن.

## قواعد اجرایی

- اول `AGENT.md` را کامل بخوان.
- سپس کد موجود repository را بررسی کن و از ساختار فعلی استفاده کن؛ فایل سالم را بی‌دلیل بازنویسی نکن.
- Scope را قفل نگه دار و قابلیت جدید پیشنهاد نده یا اضافه نکن.
- هر تصمیمی که با Specification تعارض دارد باید قبل از تغییر، از متن `AGENT.md` تبعیت کند.
- کد production-quality، type-hinted، قابل‌تست و قابل‌بسته‌بندی با PyInstaller باشد.

## ترتیب پیاده‌سازی

### مرحله 1 — مدل‌ها و خطاها

مدل‌های صریح برای موارد زیر بساز:

- InputConfig
- SymbolConfig/List
- DetectedStructure
- ColumnMapping
- ColumnCountResult
- ProcessingSummary
- ProgressEvent

خطاهای مشخص برای:

- invalid input
- workbook open failure
- header not found
- data not found
- inconsistent mapping
- output failure
- cancellation

### مرحله 2 — Structure Detector

ساختاری بساز که:

1. XLSX را فقط read-only باز کند.
2. Worksheet مناسب را پیدا کند.
3. Header حاوی `#CHROM` و `FORMAT` را پیدا کند.
4. Sample Headerها را استخراج کند.
5. Data Start را پیدا کند.
6. FORMAT marker داده را در چند ردیف اولیه تشخیص دهد.
7. Sample Data Start را از روی FORMAT marker پیدا کند.
8. Mapping Header Sample → Physical Excel Data Column را بسازد.
9. offset را به‌صورت dynamic محاسبه کند.
10. trailing unnamed Data Columns را تشخیص دهد.
11. consistency mapping را در چند ردیف کنترل کند.

برای فایل مرجع انتظار داریم چیزی در این جنس به دست بیاید:

- Header first sample: `J253 = HG00096`
- Actual data column for that sample: `U`
- Last named sample: `CRQ → CSB`
- trailing unnamed columns: `CSC:CSS`

این مقادیر را در کد hard-code نکن.

### مرحله 3 — Streaming Scanner

اسکن اصلی را Row-by-Row انجام بده.

الزامات:

- کل فایل هرگز در RAM بارگذاری نشود.
- از `read_only=True` استفاده شود.
- فقط state موردنیاز برای Count نگهداری شود.
- در هر Row cancellation event بررسی شود.
- برای UI progress event ارسال شود.
- از `values_only=True` استفاده شود.

### مرحله 4 — Counter

برای هر Data Column واقعی:

- هر Symbol انتخابی → count دقیق
- empty → count
- unknown → count
- total → count

Exact Match را بدون normalization انجام بده.

مثال:

`0|0` باید فقط با `0|0` برابر باشد.

`0|0:35:99` نباید match شود.

همچنین مقدارهای unknown دقیق را با count نگه دار تا Sheet جزئیات قابل ساخت باشد.

Memory discipline:

- از ساختن لیست تمام مقدارها خودداری کن.
- Counter/Dictهای محدود و هدفمند استفاده کن.
- اگر لازم شد برای unknown cardinality بسیار بزرگ، spill-to-disk طراحی کن؛ ولی dependency خارجی اضافه نکن.

### مرحله 5 — Output Builder

Workbook خروجی را با:

`Workbook(write_only=True)`

بساز.

Sheetها:

1. `شمارش`
2. `خلاصه`
3. `مقادیر ناشناخته` فقط در صورت داشتن unknown غیرخالی

Sheet `شمارش` ستون‌های زیر را داشته باشد:

- ستون Excel مرجع
- نام نمونه / هدر
- وضعیت
- مقدار / نماد
- تعداد
- محدوده بررسی

برای هر Symbol حتی Count=0 نیز Row خروجی بساز.

برای هر ستون Data، حداقل Rowهای `خالی` و `ناشناخته` نیز ساخته شوند.

### مرحله 6 — Excel Formatting

خروجی:

- RTL
- Header خوانا
- فونت Vazir در صورت امکان
- Freeze header
- Widthهای کنترل‌شده
- متن‌های فارسی راست‌چین
- Countها عددی
- Range به‌شکل Excel A1 notation

ظاهر را حرفه‌ای ولی ساده نگه دار.

### مرحله 7 — GUI

با `tkinter/ttk` پیاده‌سازی کن.

صفحه اصلی باید شامل:

- فایل ورودی + انتخاب فایل
- لیست Symbolها + افزودن/حذف
- پوشه خروجی + انتخاب پوشه
- ساخت فایل خروجی
- نوار پیشرفت
- وضعیت
- لغو پردازش

تمام UI فارسی و RTL باشد.

پردازش سنگین در Main Thread اجرا نشود.

### مرحله 8 — Cancellation و Safe Output

- worker thread
- `threading.Event`
- temp output
- atomic finalization
- cleanup on error/cancel

فایل ناقص نباید با نام نتیجه نهایی باقی بماند.

### مرحله 9 — Tests

حداقل این تست‌ها را اجرا و سبز کن:

1. Header discovery
2. Data start discovery
3. FORMAT marker discovery
4. Header/Data offset mapping
5. trailing unnamed columns
6. exact symbol matching
7. empty counting
8. unknown counting
9. zero-count selected symbol
10. range generation
11. inconsistent mapping failure
12. output workbook structure
13. cancellation

یک integration test با fixture کوچک ساختاری انجام بده.

### مرحله 10 — Build

PyInstaller spec را کامل کن.

Target:

- Windows 64-bit
- Python 3.13.x
- GUI mode
- no console
- one EXE

Font asset Vazir را در build لحاظ کن.

بعد build را انجام بده و EXE را launch-test کن.

## Benchmark واقعی

فایل اصلی کاربر حدود 193MB است و Worksheet تا `CSS22165` دارد.

آن را از Git خارج از repository نگه دار، اما در محیط توسعه در صورت در دسترس بودن روی آن smoke test اجرا کن.

Smoke test باید حداقل این موارد را verify کند:

- Header row = 253
- Data start row = 254
- Sample header count = 2504
- First mapped sample = `HG00096`
- First mapped physical data column = `U`
- Last named mapped column = `CSB`
- trailing unnamed data columns = `CSC:CSS`
- exact count logic produces no obvious column-shift error

اگر اجرای فایل اصلی زمان‌بر بود، آن را بهانه‌ای برای تغییر الگوریتم به load-all یا pandas نکن.

## پایان کار

پس از implementation:

1. تست‌ها را اجرا کن.
2. Build را اجرا کن.
3. EXE را بررسی کن.
4. هر خطای واقعی را برطرف کن.
5. `README.md` را با دستورهای اجرای برنامه، تست و build به‌روزرسانی کن.
6. در گزارش نهایی دقیقاً بگو چه چیزهایی پیاده شده، چه تست‌هایی اجرا شده و نتیجه آنها چه بوده است.

هیچ بخش ناتمام را به‌عنوان complete گزارش نکن.
