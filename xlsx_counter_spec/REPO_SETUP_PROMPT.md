# Prompt — راه‌اندازی Repository برای XLSX Symbol Counter

تو در نقش یک Senior Python Desktop Engineer و Software Architect هستی.

من می‌خواهم یک برنامه Windows دسکتاپ با Python بسازم که فایل‌های Excel بزرگ `.xlsx` شبیه فایل مرجع من را دریافت کند، Symbolهای انتخابی کاربر را دقیقاً بشمارد و خروجی Excel تولید کند.

**فایل `AGENT.md` موجود در ریشه Repository قرارداد اصلی پروژه است. قبل از هر کاری آن را کامل بخوان و هیچ بخش آن را نادیده نگیر.**

## وظیفه تو در این مرحله

در این مرحله هنوز وارد پیاده‌سازی business logic نشو مگر برای ساخت اسکلت لازم. Repository را برای implementation آماده کن.

### کارهایی که باید انجام دهی

1. `AGENT.md` را مبنا قرار بده.
2. Repository را با معماری تعریف‌شده در `AGENT.md` بساز.
3. `pyproject.toml` ایجاد کن.
4. dependencyها را حداقلی نگه دار.
5. Python target را روی `3.13.16` تنظیم کن.
6. dependency اصلی Excel را روی `openpyxl` نگه دار.
7. برای UI فقط `tkinter/ttk` را از Standard Library استفاده کن.
8. تست‌ها را با `pytest` تنظیم کن.
9. فایل‌های build موردنیاز PyInstaller را ایجاد کن.
10. ساختار assets برای Vazir را ایجاد کن و README کوتاه برای محل font asset بنویس.
11. فایل `README.md` ایجاد کن که هدف، ساختار پروژه و دستورهای توسعه/تست/build را توضیح دهد.
12. یک fixture کوچک Excel بساز یا generator تستی بنویس که دقیقاً مشکل فایل مرجع را شبیه‌سازی کند: Header و Data با offset، FORMAT marker، trailing unnamed columns و values مختلف.
13. تست‌های اولیه برای detector/mapping بنویس، اما الگوریتم نهایی counter را هنوز به‌صورت کامل توسعه نده.
14. `.gitignore` مناسب Python/Windows/build artifacts اضافه کن.
15. هیچ فایل شخصی، Excel بزرگ اصلی، cache، venv یا font binary را بی‌دلیل commit نکن.

## تصمیم‌های فنی ممنوع برای تغییر بدون دلیل

- Pandas را اضافه نکن.
- PyQt/PySide/CustomTkinter را اضافه نکن.
- Database خارجی اضافه نکن.
- Web service یا اینترنت runtime اضافه نکن.
- Docker لازم نیست.
- Feature جدید خارج از `AGENT.md` اضافه نکن.

## شرط مهم

اگر بین «سادگی» و «پیچیدگی معماری» مردد شدی، ساده‌ترین معماری‌ای را انتخاب کن که نیازمندی‌های `AGENT.md` را کامل و قابل‌تست پوشش دهد.

## خروجی مورد انتظار از این مرحله

در پایان:

- ساختار repository آماده باشد.
- dependencyها مشخص باشند.
- تست‌ها بتوانند اجرا شوند.
- fixture ساختاری وجود داشته باشد.
- PyInstaller configuration وجود داشته باشد.
- README نحوه اجرای dev/test/build را توضیح دهد.
- هیچ business logic ناقص به‌عنوان «پیاده‌سازی‌شده» معرفی نشود.

قبل از پایان، `pytest` را اجرا کن و هر خطا را تا حد امکان رفع کن. سپس وضعیت فایل‌های ایجادشده و تست‌ها را خلاصه کن.
