# Taswiyah Hub

منصة عربية لمطابقة كشوف الكاشير والبنك، تصنيف الفروقات، وحفظ جلسات التسوية بين المنصات.

## Run & Operate

- `pnpm --filter @workspace/api-server run dev` — run the API server (port 5000)
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from the OpenAPI spec
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- Required env: `DATABASE_URL` — Postgres connection string

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- API: Express 5
- DB: PostgreSQL + Drizzle ORM
- Validation: Zod (`zod/v4`), `drizzle-zod`
- API codegen: Orval (from OpenAPI spec)
- Build: esbuild (CJS bundle)

## Where things live

- `artifacts/taswiyah-hub/` — تطبيق الويب الرئيسي ومساحة التسوية.
- `artifacts/taswiyah-hub/src/components/reconciliation-app.tsx` — استيراد Excel، المطابقة، الحفظ، التصدير، وتقرير المقارنة.
- `artifacts/taswiyah-hub/src/components/Reconciliation2.tsx` — سير عمل المنصات A → B → C.
- `artifacts/taswiyah-hub/src/index.css` — ألوان وهوية Taswiyah Hub.
- `artifacts/taswiyah-hub/public/taswiyah-logo.png` — الشعار المستخدم في الدخول والقائمة.

## Architecture decisions

- التطبيق يعمل محلياً في المتصفح ويحفظ جلسات الكشوف والمطابقات في storage المتاح، مع دعم الاستكمال من ملف Excel المصدّر.
- ملف التصدير يحتوي على ورقة مستقلة باسم `تقرير المقارنة` قبل أوراق التفاصيل.
- منطق المطابقة الأصلي محفوظ ومفصول عن طبقة الهوية والتنقل حتى يمكن تطوير واجهة التقرير دون تغيير نتائج المطابقة.

## Product

يتيح Taswiyah Hub رفع كشوف الكاشير والحوالات، تشغيل المطابقة التلقائية أو اليدوية بالاسم ونطاق المبلغ، تصنيف الفيزا ومحفظة محمود والمعلقات وجوال بي، حفظ الجلسات للترحيل من منصة A إلى B ثم C، ومراجعة تقرير فروقات مرئي وتصديره إلى Excel.

## User preferences

- اللغة الأساسية عربية وواجهة التطبيق RTL.
- الاسم التجاري هو Taswiyah Hub والعبارة: «محطتك الأولى لتسوية الحسابات بدقة.»

## Gotchas

- الحزمة `xlsx` يجب أن تبقى ضمن حزمة `@workspace/taswiyah-hub`، وليس في جذر مساحة العمل.
- تشغيل البناء اليدوي يحتاج متغيرات الخدمة التي يضيفها workflow؛ استخدم `typecheck` أو workflow المُدار للتحقق من التطبيق.

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
