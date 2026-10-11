# جدول مسح الطاولة

صفحة بسيطة تسجّل فيها غنون وملوك مسح الطاولة بعد الفطور والغداء، ويوافق عليها ولي الأمر بكلمة سر.
التسجيلات تُحفظ في Google Sheet وتُحذف تلقائياً بعد 45 يوماً.

## الخطوة 1: إعداد Google Sheet (مرة واحدة)

1. افتح [sheets.new](https://sheets.new) لإنشاء جدول جديد، وسمّه مثلاً "جدول مسح الطاولة".
2. من القائمة: **Extensions ← Apps Script** (الإضافات ← برمجة التطبيقات).
3. احذف الكود الموجود، والصق محتوى ملف `apps-script/Code.gs`، ثم احفظ.
4. اضغط **Deploy ← New deployment**، واختر النوع **Web app**:
   - **Execute as:** Me
   - **Who has access:** Anyone
5. اضغط Deploy ووافق على الصلاحيات (قد تظهر رسالة "Google hasn't verified this app": اضغط Advanced ← Go to ... ).
6. انسخ رابط **Web app URL** (ينتهي بـ `/exec`).

## الخطوة 2: ربط الصفحة

افتح `index.html` وضع الرابط في السطر:

```js
var API_URL = "https://script.google.com/macros/s/..../exec";
```

## الخطوة 3: النشر على GitHub Pages

1. ارفع الملفات إلى مستودع على GitHub.
2. من المستودع: **Settings ← Pages ← Branch: main / root ← Save**.
3. بعد دقيقة يظهر رابط الموقع، افتحه من التابلت أو اللابتوب.

## تغيير كلمة السر

كلمة السر في `Code.gs` داخل Apps Script (السطر `PARENT_PASSWORD`) وليست في الصفحة.
بعد تغييرها: **Deploy ← Manage deployments ← Edit ← Version: New version ← Deploy**.
