# DOCX + XLSX: bounded review and refactor report

Дата: 2026-08-13

## Решение

DOCX и XLSX остаются отдельными доменами.

```text
Общие policy patterns
≠ общий Office domain model
```

Не создавались `OfficeArtifactBase`, общий SceneSpec, универсальный OOXML runtime или общий mutation backend.

## DOCX

Проект в монорепозитории: `packages/docx`

Архитектура:

```text
closed create/edit contracts
→ transaction-bound snapshot and plan
→ preservation-first DOCX mutation
→ private candidate
→ OPC + semantic + collateral validation
→ atomic publication
```

### Исправленные findings

1. Public `create`/`apply` принимали output без `.docx` — исправлено на публичной границе.
2. Create-table contract принимал `NaN`, `Infinity`, `-Infinity` — теперь допускаются только finite numeric scalars.
3. `table_totals` принимал non-finite inputs и overflow конечных operands — теперь входы, line totals и grand total должны быть finite.
4. ZIP admission отклонял symlink, но не все явно nonregular Unix member types — теперь разрешены directory, regular file и type bits `0`; FIFO/device/symlink отклоняются.
5. `set_cell_text` semantic postcondition мог принять совпадающую координату из другой таблицы — теперь target связывается с `story_part + table ordinal + row/cell index`.
6. `apply` мог выбросить `IsADirectoryError` для directory с суффиксом `.docx` — public boundary теперь возвращает typed `validation_failure`.
7. Между fingerprint и mutation source мог измениться — `apply` теперь работает с private immutable source snapshot, hash-bound к плану.
8. `plan` доверял самозаявленным `source_sha256`/`snapshot_sha256` — теперь snapshot fingerprint пересчитывается на public boundary.
9. `replace_text` postcondition мог принять ожидаемый текст в другом paragraph того же story part — теперь postcondition связан с ordinal целевого paragraph/heading.
10. Несериализуемое значение forged operation могло выбросить `TypeError` до cleanup private snapshot — operation values теперь полностью типизированы, а fingerprint/processing находятся внутри cleanup boundary.

### Выполненная проверка

```text
Authoritative suite: 36/36 PASS
Schema JSON parsing: PASS
Fresh public lifecycle: create → inspect → plan → apply → validate
Source unchanged: PASS
Changed package member: word/document.xml
Unexpected changed members: []
Package/XML/relationships/content types: PASS
```

Fresh dogfood artifact был сгенерирован исходным review harness; в публичный source distribution бинарный результат не включён.

Machine-readable результаты воспроизводятся тестами и CI.

Application compatibility: `NOT_EXECUTED`
Visual fidelity: `NOT_EXECUTED`

## XLSX

Проект в монорепозитории: `packages/xlsx`

Архитектура:

```text
closed workbook/cell/range/formula contracts
→ transaction-bound inspection and plan
→ feature-risk admission
→ bounded XLSX mutation
→ private candidate
→ reopen + semantic + collateral validation
→ atomic publication
```

### Исправленные findings

1. Public `create`/`apply` принимали output без `.xlsx` — исправлено на публичной границе.
2. Scalar `value: "=..."` обходил explicit formula contract и превращался `openpyxl` в формулу — scalar mode теперь отклоняет formula-like strings; формулы разрешены только через explicit `formula` mode.
3. Тот же formula-mode bypass был возможен через correctly rehashed forged plans — apply повторно выполняет recursive scalar/formula preflight, включая append rows.
4. Create candidate проверялся как безопасный ZIP, но не доказывал deterministic-library reopen до публикации — добавлен private-candidate reopen gate.
5. Два structural edits одного region могли пройти вместе — `plan` и forged `apply` теперь отклоняют их как `conflict`.
6. Runtime не обеспечивал schema limits 250,000 create cells и 1,000 plan operations до тяжёлой работы — budgets добавлены до renderer/operation traversal.
7. Create/plan schemas синхронизированы с explicit scalar/formula runtime policy.
8. Candidate reopen errors нормализованы в typed `validation_failure` вместо утечки текста library exception.
9. Admission требует обязательные OPC members (`[Content_Types].xml`, root relationships, workbook и workbook relationships); malformed XML/package errors не покидают public boundary.
10. Column widths и row heights проверяются как finite positive numbers до запуска renderer.
11. Между fingerprint и mutation source мог измениться — `apply` теперь работает с private immutable source snapshot.
12. Forged structural plan мог менять `range`, подменять `row_ids` и дублировать `ordered_rows` — `region_id` теперь детерминированно связан с `source + sheet + range`, а reorder требует точную permutation row identities.
13. Create candidate мог быть открываемым, но семантически не соответствовать модели — до публикации проверяются sheet order/names, values/formulas, styles, state, panes, filters, merges и заявленные dimensions.
14. Formula mode пропускал DDE/command-like формулы — добавлен отдельный deny-policy для `cmd`, `powershell`, `DDEAUTO` наряду с external workbook references.
15. Digital-signature graph мог быть потерян или инвалидирован mutation — signed workbooks теперь read-only в bounded path и отклоняются до candidate.
16. Style postcondition проверял только number format — теперь проверяются релевантные font/fill/number-format свойства; официальный validator допускает ожидаемое изменение `xl/styles.xml`.
17. Malformed snapshot (`rows: null`) и повреждённый `workbook.xml` могли выбрасывать `TypeError`/`XMLSyntaxError` наружу — public `plan`/`apply` теперь всегда нормализуют непредвиденные parser/input failures в typed `validation_failure`.
18. Forged reorder мог сохранить разрешённые row IDs, но заменить `ordered_rows` payload внешней строкой — row ID теперь пересчитывается из source hash, sheet, row number и полного cell/style payload, а ordered rows обязаны быть exact source-row objects.
19. DDE deny-list была слишком узкой — теперь блокируется общий application-token pipe pattern (`=<app>|...`), включая `MSEXCEL|...`.
20. Managed style `normal` был no-op patch — теперь это полный bounded reset font/fill/number-format, проверяемый candidate oracle.

### Выполненная проверка

```text
Authoritative suite: 37/37 PASS
Schema JSON parsing: PASS
Fresh public lifecycle: create → inspect → plan → apply → validate
Source unchanged: PASS
Changed package member: xl/worksheets/sheet1.xml
Unexpected changed members: []
Package admission + deterministic reopen: PASS
```

Fresh dogfood artifact был сгенерирован исходным review harness; в публичный source distribution бинарный результат не включён.

Machine-readable результаты воспроизводятся тестами и CI.

Application compatibility: `NOT_EXECUTED`
Formula recalculation: `NOT_EXECUTED`
Visual fidelity: `NOT_EXECUTED`

## Independent review

Final focused read-only re-review of the DOCX and XLSX bounded implementations: `APPROVE`.

```text
Critical/High/Medium blockers: 0
Architecture: APPROVE
Bounded implementation: APPROVE
Production: HOLD
```

## Production HOLD gates

DOCX requires a real Word/LibreOffice open-save-reopen and visual/pagination check.

XLSX requires a real Excel/LibreOffice open-save-reopen, formula recalculation where applicable, and visual/print-layout check.

Structural validation and deterministic-library reopen do not prove application fidelity.
