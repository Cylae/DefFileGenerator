# Input Formats & Column Mapping Mechanics

The extractor pipeline accepts manufacturer register maps in **PDF, XLSX/XLSM/XLTX/XLTM, CSV, or XML** formats. Legacy `.xls` files are not supported by `openpyxl`. Column headers vary across manufacturers; the extractor automatically maps diverse input headers using staged heuristic matching and pattern dictionaries.

---

## 🔍 Heuristic Mapping Logic

When a document table is read, the extractor samples rows, collects column headers, and evaluates them in four sequential tiers:

1. **Exact Case-Insensitive Match**: Direct comparison against target internal names (`Address`, `Name`, `Type`, `Unit`, `Action`, `Factor`, `Offset`, `ScaleFactor`, `Length`, `StartBit`).
2. **Pattern Match**: Comparison against pre-compiled synonym dictionaries in `Extractor.COLUMN_MAPPING`.
3. **Partial Match**: Substring matching for descriptive vendor headers (e.g., `"Register Start Address"` -> `Address`). Short semantic fields such as `Tag` are excluded to prevent collisions such as `voltage` -> `tag`.
4. **Fragmented/Rotated PDF Fallback**: Conservative similarity matching recovers headings whose glyph order is reversed by PDF extraction (for example `s s e r d d A` -> `Address`). It runs only after exact and partial matching so normal documents retain priority.

---

## 📋 Recommended Column Headers

| Internal Field | Mandatory | Example Values | Recognized Vendor Headers |
|---|---|---|---|
| **Address** | Yes | `40001`, `0x9C41`, `40001_0_8` | `address`, `addr`, `register`, `reg`, `index` |
| **Name** | Yes | `Active Power`, `Grid Voltage` | `name`, `description`, `parameter`, `variable`, `signal` |
| **Type** | Yes | `U16`, `I32`, `F32_WB`, `STR20`, `BITS` | `data type`, `datatype`, `type`, `format` |
| **RegisterType** | Recommended | `Holding Register`, `Input Register` | `register type`, `reg type`, `modbus type` |
| **Unit** | No | `W`, `V`, `A`, `°C`, `Hz` | `unit`, `units` |
| **Factor** | No | `0.1`, `1/10`, `0.01` | `scale`, `factor`, `multiplier`, `ratio` |
| **Gain** | No | `10`, `100` | `gain` (converted to `Factor = 1 / Gain`) |
| **Offset** | No | `0`, `-10` | `offset`, `bias`, `coefficient b` |
| **ScaleFactor** | No | `-1`, `0`, `2` | `scalefactor`, `scale factor` |
| **Action** | No | `RO`, `RW`, `Read`, `Read/Write` | `action`, `access`, `read/write` |
| **Length** | No | `1`, `2`, `10` | `length`, `len`, `size`, `count`, `quantity` |
| **StartBit** | No | `0`, `8` | `startbit`, `bit offset`, `bit`, `start` |

---

## ⚙️ Custom Mapping JSON Override

To bypass auto-detection heuristics, provide an explicit JSON mapping file via `--mapping`:

```json
{
  "Address": "Modbus_Addr",
  "Name": "Signal_Description",
  "Type": "Format_Code",
  "Unit": "Engineering_Unit",
  "Action": "Access_Right"
}
```

Usage:

```bash
deffilegen extract datasheet.xlsx --mapping custom_mapping.json -o registers.csv
```

---

## 🔢 Address Format & Compound Notation

1. **Decimal**: `40001` -> parsed directly.
2. **Hexadecimal**: `0x9C41` or `9C41h` -> converted to decimal `40001`.
3. **Register Ranges**: `31657~31658`, `40001-40002`, `0x8232 ~ 0x82FE`, `30001..30002` -> automatically extracts base start register (`31657`, `40001`, `33330`, `30001`).
4. **Compound Bitfields (`BITS`)**: `address_startbit_length`, for example `30001_0_1`.
5. **Compound Strings (`STR<n>`)**: `address_length`, for example `30030_20`.
6. **Raw Register Sequences (`RAW`)**: `address_byte_length`, for example `30040_10` for five Modbus registers. The byte length must be positive and divisible by two.

---

## 🔠 Supported Data Types & Aliases

The generator normalizes a broad range of vendor type representations:

| Normalized Type | Vendor Input Examples | Notes |
|---|---|---|
| **U16** | `uint16`, `u16`, `int16u`, `unsigned int 16`, `unsigned integer`, `uint`, `hex`, `16-bit hex`, `bitfield16` | 1 Modbus word (16 bits unsigned) |
| **I16** | `int16`, `i16`, `int16s`, `sint16`, `s16`, `signed int 16`, `signed integer`, `int` | 1 Modbus word (16 bits signed) |
| **U32** | `uint32`, `u32`, `int32u`, `unsigned int 32`, `hex32`, `32-bit hex`, `bitfield32`, `datetime` | 2 Modbus words (32 bits unsigned) |
| **I32** | `int32`, `i32`, `int32s`, `sint32`, `s32`, `signed int 32` | 2 Modbus words (32 bits signed) |
| **F32** | `float`, `float32`, `f32`, `32-bit IEEE 754`, `IEEE-754` | 2 Modbus words (IEEE 754 single precision) |
| **U64** | `uint64`, `u64`, `int64u`, `unsigned int 64`, `bitfield64` | 4 Modbus words (64 bits unsigned) |
| **I64** | `int64`, `i64`, `int64s`, `sint64`, `s64`, `signed int 64` | 4 Modbus words (64 bits signed) |
| **F64** | `double`, `float64`, `f64`, `64-bit IEEE 754` | 4 Modbus words (IEEE 754 double precision) |
| **BITS** | `bit16`, `bitmap16`, `bits16`, `bit32`, `bitmap32` | Bitfield slice (`addr_startbit_length`) |
| **IP / IPV6** | `ip`, `ip4`, `ipv4`, `ipv6` | Network IP addresses |
| **STRING** | `string 20`, `str*30`, `string_16`, `STR` | Modbus ASCII character string |
| **RAW** | `raw` | Uninterpreted register sequence; address uses `address_byte_length` and byte length must be a positive multiple of 2 |

### Action codes

WebdynSunPM action codes `0`, `1`, `2`, `4`, `6`, `7`, `8`, `9`, and `10` are preserved. Action `10` represents a constant and must not be replaced by a default action.

### BITS constraints

A `BITS` slice represents bits inside a single 16-bit Modbus register. Therefore:

- `startbit` is in the range `0..15`;
- `length` must be at least `1`;
- `startbit + length` must not exceed `16`.

Examples:

| Address | Result | Reason |
|---|---|---|
| `30001_0_1` | valid | bit 0 only |
| `30001_8_8` | valid | bits 8 through 15 |
| `30001_15_1` | valid | final bit only |
| `30001_15_2` | invalid | would extend beyond bit 15 |
| `30001_0_0` | invalid | empty bit slice |

Strict validation also checks bit-level overlap. Disjoint slices on the same register are valid (`30001_0_4` and `30001_4_4`); overlapping slices (`30001_0_4` and `30001_2_4`) are rejected.

## Output definition structure

Definitions use semicolons and UTF-8 with BOM. The first row has five metadata fields followed by six empty fields so every row has eleven columns:

```text
Protocol;Category;Manufacturer;Model;Forced writing code;;;;;;
```

Every following row has eleven fields:

```text
Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action
```

For Modbus, `Info1` is `1` (coil), `2` (discrete input), `3` (holding register), or `4` (input register); `Info2` contains the normalized address notation and `Info3` the normalized data type. `CoefA` is computed from the source factor and scale factor, while `CoefB` is the offset.

## Extraction boundaries

The PDF reader consumes tables exposed by `pdfplumber`; it does not perform OCR and does not infer addresses that are absent from extracted cells. For difficult manuals, restrict processing with `--pages`, OCR image-only pages externally, or use `--mapping` after inspecting an intermediate `extract` CSV. This conservative behavior prevents fabricated registers.
