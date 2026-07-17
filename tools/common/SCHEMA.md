# Snapshot JSON schema (`opengex-snapshot/1`)

One snapshot file describes the full object model state of a single `GridEX` or
`GEXPreview` control instance. Produced by `mdSnapshot.bas` (runtime or IDE
add-in), consumed by the twin test harness (state diffing), the sample porting
code generator and later the object model round-trip tests.

All JSON reading/writing goes through `mdJson.bas` (`JsonValue`/`JsonDump`/
`JsonParse`) -- no other JSON code is allowed in the project, using XPath-like
slash paths (`"props/items/-1"`) as the preferred notation. Note: `mdJson`
treats path strings starting with `$` as JSONPath expressions, so keys that
literally start with `$` need a leading-slash escape: `"/$schema"`,
`"props/$errors"` (JSONPath `"$.$schema"` works too). Never use bare
`"$schema"` paths -- they misparse. Idioms: create empty object with
`JsonValue(oJson, vbNullString) = Empty`, empty array / append with index
`-1` (reading `-1` returns the count), remove keys by assigning `Empty`
(which is why `Empty` values are stored as `Null`), test existence with
`IsEmpty(JsonValue(...))`.

```jsonc
{
  "$schema": "opengex-snapshot/1",
  "class": "GridEX",              // ProgID-less class name (GridEX | GEXPreview)
  "mode": "persist",              // "persist" = design-time settable state only
                                  // "runtime" = every readable property
  "props": {
    // scalar properties: JSON number/string/bool; enums as numbers
    "AllowAddNew": false,
    "BackColor": -2147483643,     // OLE_COLOR as signed 32-bit number
    "View": 0,                    // jgexViewConstants

    // Font-valued properties
    "Font": { "Name": "MS Sans Serif", "Size": 8.25, "Bold": false,
              "Italic": false, "Underline": false, "Strikethrough": false,
              "Charset": 0 },     // null when unset

    // Picture-valued properties: null when empty
    "Picture": { "type": 1, "data": "<base64>" },

    // object-valued properties recurse into their own props
    "PrinterProperties": {
      "Copies": 1,
      // parameterized props expand positions as keys (jgexHFCenter/Left/Right)
      "HeaderString": { "1": "", "2": "", "3": "" },
      "FooterString": { "1": "", "2": "", "3": "" }
    },

    // collection-valued properties are plain arrays of item snapshots (in
    // collection order); the key is omitted entirely when the collection is
    // empty -- mdJson consumers see no difference (JsonKeys returns an empty
    // array, the "-1" count reads 0)
    "Columns": [
      { "Caption": "ProductID", "Width": 1000,
        "ValueList": [ { "Value": 1, "Text": "Yes" } ] },
      null                        // slot present but Item() returned Nothing
    ],

    // exception: collection classes with own persistable properties
    // (currently only JSFmtConditions) keep the object form with "items"
    "FmtConditions": {
      "ApplyGroupCondition": false,
      "GroupCondition": { "ColIndex": 0 },
      "items": [ { "ColIndex": 1, "Operator": 0 } ]
    },

    // reads that raised an error are collected per object, value = Err.Number;
    // failed Item() reads use bracketed 1-based indexes ("[3]" inside a
    // wrapped collection, "Columns[3]" when the array is inlined); a failed
    // Count read on an inlined collection reports under the property name
    "$errors": { "SomeProp": 438, "Columns[3]": 91, "HeaderString[2]": 438 }
  },

  // reserved: raw .frm propbag keys captured by the porting tooling
  // (MethodHoldFields, IntProp3/5/7/8, ColumnsCount, ...) -- text values
  // verbatim, binary blob keys as { "frx": "<base64>" }
  "raw": { "MethodHoldFields": "-1  'True" }
}
```

Rules:

- Key order follows profile order (`mdProfiles.bas`), which follows stub
  declaration order -- stable across runs so snapshots diff cleanly as text.
- `persist` mode omits properties flagged `bRuntimeOnly` in the profiles
  (window handles, cursor position, selection, live recordsets, ...).
- Formatting is whatever `JsonDump` emits (4-space indent, short objects
  inlined); always re-dump through `mdJson.bas` before diffing snapshots
  from other sources.
- Numbers always use `.` as decimal separator regardless of locale.
- `null` for: unset object values, Variant `Empty`, `Nothing` items, empty
  pictures.
- Unknown/extra keys must be ignored by consumers (forward compatibility);
  the `$schema` value changes only on breaking shape changes.
