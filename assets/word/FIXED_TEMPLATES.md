# Russian fixed Word templates, version 1

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

Build with tools/build_fixed_word_templates.py. JSON sample fields are fictional test fixtures and are never automatically inserted into user documents. tools/check_fixed_word_runtime.py checks real DOCX/PDF export during deployment without external API calls.

Pricing: existing 5-credit Word page rate, capped at the two-page quote. Free blank DOCX download. Ready files and sent-file flags are retained for retries in the current bot session, with charging after both formats are sent. Retry state is not durable across process restarts. Failure/debit tests use mocked Telegram and balances, not real user accounts.
