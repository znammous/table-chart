# جدول الأبطال

صفحة يسجّل فيها الأطفال مهام البيت، ويوافق عليها ولي الأمر، مع لوحة إنجاز وشارات وبطل الشهر.
أسماء الأطفال والمهام ونقاطها واسم التطبيق ورمز ولي الأمر تُضبط من زر ⚙️ الإعدادات داخل الصفحة.

**الموقع:** https://tablehero-d582b.web.app

## كيف يعمل

- **Firebase Authentication:** ولي الأمر يسجّل دخوله بحساب Google مرة واحدة على كل جهاز.
- **Cloud Firestore:** كل عائلة لها مستند `families/{uid}` فيه الإعدادات والجوائز والرمز، وتحته `entries` للتسجيلات.
- **قواعد الحماية** (`firestore.rules`): لا يقرأ ولا يكتب في بيانات العائلة إلا ولي أمرها.
- **Firebase Hosting:** يستضيف `public/index.html`.

## النشر

```bash
firebase deploy --only hosting          # الصفحة
firebase deploy --only firestore:rules  # قواعد الحماية
```

## ملفات قديمة

- `index.html` في الجذر يحوّل رابط GitHub Pages القديم إلى الموقع الجديد.
- `apps-script/Code.gs` نسخة Google Sheet السابقة، محفوظة كمرجع ونسخة احتياطية.
