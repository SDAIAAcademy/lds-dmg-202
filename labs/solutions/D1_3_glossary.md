# الحل المرجعي: المختبر 1.3 · 15 مصطلحًا لمجال الترخيص

نقطة البداية: `files/glossary/glossary_workbook.xlsx`. الحل: `solutions/reference/glossary_reference.xlsx`. التحقق: `uv run solutions/check.py 1.3`.

| المعرّف | المصطلح | التصنيف | المالك | الأعمدة |
|---|---|---|---|---|
| T01 | ممارس مرخَّص | مقيد | راكان التريكي | practitioners.practitioner_id، practitioners.full_name، licences.practitioner_id |
| T02 | رخصة مهنية | عام | راكان التريكي | licences.licence_no، licences.issued_on، public_register.licence_no |
| T03 | رقم الهوية الوطنية | مقيد | راكان التريكي | practitioners.national_id |
| T04 | جوال التواصل | مقيد | راكان التريكي | practitioners.mobile، renewal_reminders.mobile |
| T05 | مهنة منظَّمة | عام | راكان التريكي | licences.profession_code، professions.profession_code |
| T06 | فئة الرخصة | عام | راكان التريكي | licences.licence_class |
| T07 | حالة الرخصة | عام | راكان التريكي | licences.status، public_register.status |
| T08 | رخصة سارية | عام | راكان التريكي | published_licence_stats.active_licences |
| T09 | تاريخ انتهاء الرخصة | عام | راكان التريكي | licences.expiry_date، renewal_reminders.expiry_date |
| T10 | طلب ترخيص | مقيد | راكان التريكي | applications.application_id، applications.kind، applications.submitted_on |
| T11 | قرار الطلب | مقيد | راكان التريكي | applications.status، applications.decided_on |
| T12 | تجديد الرخصة | مقيد | راكان التريكي | renewals.renewal_id، renewals.decision، licences.last_renewed_on |
| T13 | ساعات التطوير المهني | مقيد | راكان التريكي | renewals.cpd_hours، professions.cpd_hours_required |
| T14 | رسم الترخيص | مقيد | مدير الشؤون المالية | fee_payments.amount_sar، fee_payments.paid_on |
| T15 | منطقة الممارسة | عام | راكان التريكي | practitioners.region_code |

التعريفات المرجعية في ملف الحل. ما يُقبل غيرها: أي تعريف يبدأ بالنوع ثم يذكر ما يميز، ولا يعيد المصطلح، ويتفق مع الملاحظة.

«رخصة سارية» بتعريف المالك في 31 أكتوبر 2026: 1710 رخصة. بحالة active وحدها: 1804. الفرق 94 رخصة انتهت ولم تُحدّث حالتها؛ والإحصاءات المنشورة حتى سبتمبر تعدّ بالحالة وحدها، وهذا ما يكشفه المختبر 2.1.
