# Russian fixed Word templates, owner-approved revision 4

These are independently typeset, adapted working templates, not government-approved universal forms. No source logos, advertisements, hyperlinks or watermarks are embedded. References supplied/selected by the owner: services agreement DOCX, services-act PDF (Dogovor.expert), Rospotrebnadzor replacement claim (3330.doc), and Yandex apartment tenancy sample. Layout and branding were not copied.

Kinds: services_ru (ordinary services between organizations/IP), services_act_ru (completed services), rent_flat_ru (residential tenancy between individuals, owner signs personally), claim_replace_ru (consumer replacement with the same model). Scope is displayed before confirmation. These forms stay Russian regardless of interface language. Generic documents remain separate.

Models only extract exact supplied substrings and cannot rewrite clauses. Labeled input bypasses models. Explicit corrections replace values; blank values clear them. Unknown fields remain visible in brackets. Users review the filled fields, omissions and scope before confirming DOCX/PDF delivery. This does not establish identity, ownership, tax treatment or legal applicability to an individual case.

Legal references checked 2026-10-19:
- Services, Civil Code 782: https://zpp.rospotrebnadzor.ru/news/federal/104479
- Act, 402-FZ article 9: https://www.consultant.ru/document/cons_doc_LAW_122855/b24121ba7152f33673608559f2ca844ef5b6a74c/
- Consumer replacement periods: https://zpp.rospotrebnadzor.ru/news/regional/576055
- February 2026 consumer-law changes: https://zpp.rospotrebnadzor.ru/news/federal/572341
- Tenancy, Civil Code 674, 683, 687: https://www.consultant.ru/document/cons_doc_LAW_9027/

The claim does not promise unconditional penalties/fines. Technical goods may need additional grounds depending on timing, as shown in the preview. Payment and absence of complaints are not invented. Tenancy separates signing from actual handover.

The checked-in DOCX assets are the approved source of truth; older builder scripts reproduce earlier layouts and must not overwrite them. JSON sample fields are fictional test fixtures and are never automatically inserted into user documents. tools/check_fixed_word_runtime.py checks real DOCX/PDF export during deployment without external API calls.

Pricing: existing 5-credit Word page rate, capped at the two-page quote. Free blank DOCX download. Ready files and sent-file flags are retained for retries in the current bot session, with charging after both formats are sent. Retry state is not durable across process restarts. Failure/debit tests use mocked Telegram and balances, not real user accounts.

## Remaining Word menu forms

The 24 older non-study menu entries now use fixed forms for Russian requests. Other content languages retain their prior workflow. The specialised car-sale route retains priority when the generic sale entry explicitly concerns a car. The separate Mini App real-estate sale flow is not changed by this batch.

Editorial source: `tools/extend_fixed_catalog.py`; per-form scope, fields, fictional test data and reference URLs are in the catalog. These are independently drafted working forms based on the relevant rules, not reproductions of officially approved forms. They do not cover every transaction that could fall under a broad menu label. The precise scope is shown at input and confirmation, and on free download.

| Menu kind | Fixed form and scope |
| --- | --- |
| dkp | Ordinary movable property sale between adults; car route remains separate |
| rent | Ordinary movable property lease; separate apartment and commercial-premises forms remain available |
| proxy | Simple written authority to collect listed documents, no disposition of property or litigation powers |
| statement | General application to an organisation, not a prescribed administrative form |
| act | Actual handover and acceptance of listed property/documents |
| loan | Monetary loan between adult individuals; signing is not a receipt for money |
| claim | General civil pre-action demand; grounds, amounts and deadline supplied by the user |
| consent | Child travel consent draft for notarisation |
| marriage_contract | Marital property agreement draft for notarisation |
| gift | Ordinary movable property gift between adults, not real estate |
| lawsuit | Civil claim draft under the Civil Procedure Code, not an arbitration/administrative claim |
| alimony | Child-support agreement draft for notarisation |
| offer | Commercial proposal for discussion, not a deliberately binding offer |
| services | Reuses the existing reviewed B2B/IP service agreement |
| employment | Indefinite main-job adult employment, no probation or remote regime |
| work_act | Ordinary completed-work acceptance, not KS-2 or a tax invoice |
| supply | Domestic ordinary-goods B2B/IP supply |
| agency | Agent acts in the principal's name and at their expense |
| joint_activity | Entrepreneurial simple partnership between IP/commercial organisations, draft for review |
| nonresidential_rent | Direct commercial-premises lease from the owner |
| cession | Paid assignment of an existing monetary business claim |
| nda | Mutual confidentiality agreement, not a substitute for a trade-secret regime or personal-data processing basis |
| self_employed | Independent ordinary services under NPD, not employment or agency |
| warranty_letter | Debtor's letter about its own obligation, not an independent guarantee |

Legal reference set consulted 2026-10-19 (the applicable rule must still be checked for each real transaction):
- Civil Code Part I, representation, obligations and assignment: https://www.consultant.ru/document/cons_doc_LAW_5142/
- Civil Code Part II, sale, gift, lease, services, loan, agency and partnership: https://www.consultant.ru/document/cons_doc_LAW_9027/
- Family Code, articles 40–44, 46, 99–105: https://www.consultant.ru/document/cons_doc_LAW_8982/
- Labour Code, articles 57, 67, 68, 115, 136: https://www.consultant.ru/document/cons_doc_LAW_34683/
- Consular Department, child travel consent: https://www.kdmid.ru/cons/notary/certification-of-consent-to-the-departure-of-a-minor-citizen-of-the-Russian-Federation/
- Civil Procedure Code 131: https://www.consultant.ru/document/cons_doc_LAW_39570/69470b15e4eb033f5bfa691094e1d083f32238e9/
- Civil Procedure Code 132: https://www.consultant.ru/document/cons_doc_LAW_39570/c04bddd90fe8dd05299705ec675c65ecd0f668dc/
- NPD income exclusions for current/recent employers: https://www.nalog.gov.ru/rn27/news/smi/13907247/
- Trade secrets, 98-FZ article 10: https://www.consultant.ru/document/cons_doc_LAW_48699/0ca1b144b64eaa68cd80ca51ad37ac4047c47775/

Notarial drafts explicitly state that they are not certified. No certification inscription, registry number, signature image, logo or seal is fabricated. Unknown identity, ownership, payment, tax and case details remain placeholders. The review screen is a user confirmation step, not an independent legal opinion. The scoped regex checks catch only explicit incompatible cases; they are not legal classification or validation of every user condition.

## Approved layout integration — 2026-10-19

The owner reviewed all 28 distinct Word layouts and authorized release after filling checks. Existing menu routes remain unchanged (services aliases services_ru). The service agreement preserves its eleven sections. Role-specific labels disambiguate repeated fields; the service act separates scope/quantity from measurement units. No model rewrites form clauses. The first editable item row accepts the supplied description; other rows remain available for manual additions. Multi-row structured entry is not introduced in this release.

`docx_form_fill.py` replaces markers across split Word runs, including nested tables and footers, preserving formatting and unrelated breaks. Long table rows can expand and flow between pages; short signature blocks stay together. The employment receipt stays with the signatures. Missing fields remain visible. Versions invalidate previously prepared sessions so old layouts cannot be reused for a new confirmation. Prices are unchanged.

Offline QA: `tools/qa_approved_forms.py` uses the actual word_build handler and production DOCX/PDF conversion with fictional labeled data. Telegram transport and balances are mocked, and paid extraction is never called. Runtime deployment checks render every catalog route. Local page inspection complements XML/field tests; these checks do not certify legal applicability or prove the live Telegram round trip.
