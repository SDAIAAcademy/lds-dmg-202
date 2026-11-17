# البيانات الوصفية وكتالوجات البيانات وتتبع أصل البيانات: المختبرات

ستة مختبرات لمقرر أكاديمية سدايا **LDS-DMG-202**، مدة كل منها 70 دقيقة في جلسة من 120 دقيقة. المختبرات صفحات
إرشادية بالعربية: الهدف، والملفات، وخطوات بدقائقها، وأسئلة تحقق، وتمرين إضافي ✱، وملاحظات للمدرب.

كل المختبرات في جهة واحدة: **هيئة صِنوان للترخيص المهني**، جهة حكومية افتراضية، وكلها على مجال الترخيص. كل الأسماء
والأرقام وأرقام الهوية والجوالات اصطناعية: الهوية تبدأ بالرقم 9، والجوال بـ 0599، والبريد بنطاق example.com.

## ما يحتاجه المشارك

- حاسوب محمول عليه [uv](https://docs.astral.sh/uv/)، أو حساب Google Colab.
- جدول بيانات: LibreOffice Calc أو Microsoft Excel.
- لا خادم، ولا حساب في أداة كتالوج، ولا يُنزَّل أي نموذج. أداة الكتالوج هي سطر أوامر الاستخلاص في DataHub المفتوح
  المصدر (`acryl-datahub`)، ويكتب ما يستخلصه في ملف محلي في `out/`.
- لا يحتاج أي مختبر نموذجًا ليكتمل. المسودات الأربعون في المختبر 1.2 مكتوبة سلفًا بديلًا عن ناتج نموذج.

### على الحاسوب

```bash
cd labs
uv sync                      # مرة واحدة: ينشئ البيئة من uv.lock
uv run kit/harvest_check.py  # كل أمر في الصفحات يُشغَّل من مجلد labs
```

### في Google Colab

```python
!git clone <رابط المستودع> course
%cd course/labs
!pip install -q -r requirements-colab.txt
!datahub ingest -c files/recipes/licensing_db.yml
!python kit/harvest_check.py
```

في Colab يُكتب `python` بدل `uv run`، ويُكتب `datahub` بدل `uv run datahub`. الملفات التي تُعدّل (جداول البيانات
والوصفات) تُنزَّل وتُرفع من لوحة الملفات، أو تُعدّل في المحرر.

## نقطة البداية والحل

كل مختبر يبدأ من ملفات جاهزة في `files/` ولا يعتمد على ناتج مختبر آخر. مختبرات اليوم الثاني تبدأ من ملفات مرجعية
(أحداث المسار، وحالة الكتالوج، والسجلات)، لا من ناتج اليوم الأول.

| المختبر | الصفحة | يبدأ من | الحل المرجعي |
|---|---|---|---|
| 1.1 | `D1_1_harvest.html` | `files/recipes/`، `files/sources/` | `solutions/D1_1_harvest.md`، `solutions/recipes/` |
| 1.2 | `D1_2_drafts.html` | `files/drafts/` | `solutions/D1_2_drafts.md`، `solutions/reference/drafts_reviewed.xlsx` |
| 1.3 | `D1_3_glossary.html` | `files/glossary/` | `solutions/D1_3_glossary.md`، `solutions/reference/glossary_reference.xlsx` |
| 2.1 | `D2_1_impact.html` | `files/lineage/` | `solutions/D2_1_impact.md` |
| 2.2 | `D2_2_health.html` | `files/reference/` | `solutions/D2_2_health.md`، `solutions/reference/health_config_reference.yml` |
| 2.3 | `D2_3_records.html` | `files/records/` | `solutions/D2_3_records.md`، `solutions/reference/dataset_records_reference.xlsx` |
| المتابعة | `files/capstone/brief.txt` | `files/capstone/` (الموجز، وحالة الكتالوج، وورقة السجلات، وأحداث المسار ومالكو الأصول)، `files/sources/accreditation.db` | `solutions/capstone.md`، `solutions/recipes/accreditation_*.yml`، `solutions/reference/accreditation_records_reference.xlsx` |

## تشغيل الحلول

```bash
cd labs
uv run solutions/check.py            # كل المختبرات ومهمة المتابعة
uv run solutions/check.py 1.1 2.3    # بعضها
```

يشغّل الاستخلاص الحقيقي في المختبر 1.1 ومهمة المتابعة، والأدوات في `kit/` على الملفات المرجعية في بقية المختبرات،
ويطبع الأرقام التي تقتبسها شرائح ما بعد المختبر و`PASS` لكل منها. ويتأكد أيضًا أن ملفات البداية ما زالت تحمل مشكلاتها
المزروعة. لا يستدعي أي نموذج. يكتب في `out/`، ويمكن حذفه بعد ذلك.

## الملفات

`_tools/prepare_files.py` يولّد كل ما في `files/` و`solutions/reference/` ببذرة ثابتة، و`_tools/content.py` فيه النصوص
العربية. أعد التوليد بعد أي تعديل فيهما:

```bash
uv run --no-project _tools/prepare_files.py
```

بوابة البيانات المفتوحة في المختبر 2.3 وقواعد التحقق فيها (`files/records/target_profile.ttl`) كُتبت لهذه الحالة بمفردات
DCAT 3، ولا تُنسب إلى أي بوابة حقيقية.
