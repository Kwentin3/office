# Rich Inspect and strict templates

## Rich Inspect

Each preservation-first domain exposes `inspect(..., view="inventory")`. The inventory is format-specific; there is no shared Office feature model.

The result contains:

```json
{
  "status": "ok",
  "artifact_type": "docx",
  "view": "inventory",
  "features": {"tracked_changes": 0},
  "mutation_policy": {
    "decision": "safe",
    "blockers": [],
    "warnings": []
  }
}
```

Decisions:

- `safe`: none of the inspected risk features are present;
- `safe_with_warnings`: preserved but unsupported features are present;
- `refuse_mutation`: a blocker is present and `apply` refuses before mutation.

Current blockers are deliberately closed and format-specific:

- DOCX: tracked changes, content controls, fields, altChunk, OLE/ActiveX, document protection, macros, signatures;
- XLSX: external links, connections, OLE/ActiveX, workbook/sheet protection, macros, signatures;
- PPTX: OLE/ActiveX, modification protection, macros, signatures.

Headers, footers, footnotes and endnotes are counted in DOCX but are not warnings because the current inspector already addresses those stories. An embedded package is not classified as OLE unless an OLE node or relationship is present.

The inventory is evidence about the listed features only. It is not a claim that every possible OOXML extension is understood.

## Strict Template Mode

The three preservation-first packages expose:

```python
result = tool.fill_template(source, values, output, strict=True)
```

Tokens use the closed syntax `{{name}}`, where names match `[A-Za-z][A-Za-z0-9_.-]{0,79}`. Values must be strings and cannot contain token markers. Both the value and the final rendered target are bounded to the format text limit; malformed nesting such as `{{{name}}}` refuses.

Strict mode requires an exact key set:

- a missing value refuses;
- an unknown value refuses;
- malformed or unresolved markers refuse;
- the source and output must differ;
- every package `.xml` and `.rels` member must parse under the hardened XML policy, which rejects DTD/entity declarations from parsed document metadata independently of XML encoding;
- the candidate is independently inspected package-wide and must contain no residual marker before atomic publication.

Each format compiles to its own existing plan operations:

- DOCX: paragraphs/headings, single-paragraph table cells, headers, footers, footnotes and endnotes already represented by the DOCX inspector; split runs are resolved through logical text while non-token paragraph and table-cell run formatting is preserved;
- XLSX: non-formula string cells across bounded worksheets; the package-preserving cell patch keeps unknown non-target members byte-identical;
- PPTX: only explicitly managed top-level `slot:*` text shapes and their table cells; non-token run and paragraph formatting is preserved.

DOCX tokens in text boxes, comments, unsupported package parts, or multi-paragraph table cells refuse rather than flattening structure. XLSX markers in formulas, headers/footers or non-cell package scope refuse. PPTX tokens in unmanaged/grouped shapes, speaker notes or other package scope refuse; marker-free speaker notes remain preserved. The packages do not silently broaden their editable surfaces.
