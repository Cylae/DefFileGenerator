# Input Formats & Column Mapping Mechanics

<p align="center">
  <a href="../README.md"><b>README</b></a> •
  <a href="quickstart.md"><b>Quick Start</b></a> •
  <a href="architecture.md"><b>Architecture</b></a> •
  <a href="security.md"><b>Security</b></a> •
  <a href="../DEVELOPER_GUIDE.md"><b>Developer Guide</b></a>
</p>

---

The extractor pipeline accepts manufacturer Modbus register documentation in **PDF, Excel (XLSX/XLSM/XLTX/XLTM), CSV, or XML** formats. Because column headers and data type names vary across manufacturers, the engine applies staged heuristic matching and pattern dictionaries to normalize heterogeneous inputs into a canonical schema.

---

## 📑 Table of Contents

- [🔍 4-Tier Heuristic Mapping Logic](#-4-tier-heuristic-mapping-logic)
- [📋 Recognized Vendor Column Headers](#-recognized-vendor-column-headers)
- [⚙️ Explicit JSON Column Override (`--mapping`)](#️-explicit-json-column-override---mapping)
- [🔢 Address Notations & Compound Syntax](#-address-notations--compound-syntax)
- [🔠 Normalized Data Types & Aliases](#-normalized-data-types--aliases)
- [🧩 Strict Rules for Bitfield Slices (`BITS`)](#-strict-rules-for-bitfield-slices-bits)
- [📐 Mathematical Scaling Coefficients Calculation](#-mathematical-scaling-coefficients-calculation)
- [📄 Official WebdynSunPM CSV Definition Structure](#-official-webdynsunpm-csv-definition-structure)

---

## 🔍 4-Tier Heuristic Mapping Logic

When a table is ingested, the extractor evaluates header rows and sample entries across 4 sequential matching tiers:

```text
┌────────────────────────────────────────────────────────┐
│  Tier 1: Exact Case-Insensitive Match                  │
│  (e.g., "Address" == "address", "Type" == "type")      │
└───────────────────────────┬────────────────────────────┘
                            │ Not matched
                            ▼
┌────────────────────────────────────────────────────────┐
│  Tier 2: Synonym Dictionary Lookup (COLUMN_MAPPING)    │
│  (e.g., "register", "addr", "offset", "adresse")       │
└───────────────────────────┬────────────────────────────┘
                            │ Not matched
                            ▼
┌────────────────────────────────────────────────────────┐
│  Tier 3: Guarded Partial Match (Substring search)      │
│  (e.g., "Modbus Start Address" -> "Address")           │
│  (Note: "Tag" is excluded to avoid collisions)         │
└───────────────────────────┬────────────────────────────┘
                            │ Not matched (PDF only)
                            ▼
┌────────────────────────────────────────────────────────┐
│  Tier 4: Fragmented or Reversed Glyph Fallback         │
│  (e.g., "s s e r d d A" -> "Address")                  │
└────────────────────────────────────────────────────────┘
```

---

## 📋 Recognized Vendor Column Headers

| Canonical Target | Mandatory | Sample Values | Automatically Recognized Vendor Headers |
|---|---|---|---|
| **Address** | **Yes** | `40001`, `0x9C41`, `30001_0_1` | `address`, `addr`, `register`, `reg`, `registre`, `adresse`, `offset`, `index` |
| **Name** | **Yes** | `Active Power`, `Grid Voltage` | `name`, `description`, `parameter`, `variable`, `signal`, `grandeur`, `nom` |
| **Type** | **Yes** | `U16`, `I32`, `F32_WB`, `STR20`, `BITS` | `data type`, `datatype`, `type`, `format`, `taille`, `longueur` |
| **RegisterType**| Recommended | `Holding Register`, `Input Register` | `register type`, `reg type`, `modbus type`, `type registre` |
| **Unit** | No | `W`, `V`, `A`, `°C`, `Hz`, `kWh` | `unit`, `units`, `unite`, `unité`, `symbole` |
| **Factor** | No | `0.1`, `1/10`, `0.01` | `scale`, `factor`, `multiplier`, `ratio`, `pas`, `coeff`, `multiplicateur` |
| **Gain** | No | `10`, `100` *(inverted to 1/Gain)* | `gain`, `diviseur` |
| **Offset** | No | `0`, `-10`, `40` | `offset`, `bias`, `decalage`, `b` |
| **ScaleFactor** | No | `-1`, `0`, `2` | `scalefactor`, `scale factor`, `puissance` |
| **Action** | No | `RO`, `RW`, `Read`, `4`, `10` | `action`, `access`, `read/write`, `droits` |
| **Length** | No | `1`, `2`, `10` | `length`, `len`, `size`, `count`, `quantity` |
| **StartBit** | No | `0`, `8` | `startbit`, `bit offset`, `bit`, `start` |

---

## ⚙️ Explicit JSON Column Override (`--mapping`)

If a manufacturer documentation uses obscure or conflicting column titles, supply an explicit JSON mapping file to bypass automatic heuristics:

```json
{
  "Address": "Vendor_Reg_Num",
  "Name": "Signal_Description",
  "Type": "Format_Code",
  "Unit": "Engineering_Unit",
  "Action": "Access_Right"
}
```

Usage:
```bash
deffilegen extract "datasheet.xlsx" --mapping "custom_mapping.json" -o "registers.csv"
```

---

## 🔢 Address Notations & Compound Syntax

The generator parses and normalizes diverse vendor notations:

1. **Standard Decimal**: `40001` $\rightarrow$ parsed directly.
2. **Hexadecimal**: `0x9C41` or `9C41h` $\rightarrow$ converted to decimal `40001`.
3. **Register Ranges**: `31657~31658`, `40001-40002`, `0x8232 ~ 0x82FE`, `30001..30002` $\rightarrow$ automatically resolves the base starting register (`31657`, `40001`, `33330`, `30001`).
4. **Bitfield Slices (`BITS`)**: `address_startbit_length` notation (e.g. `30001_0_1`).
5. **Character Strings (`STR<n>`)**: `address_length` notation (e.g. `30030_20` for 20 characters).
6. **Raw Bytes (`RAW`)**: `address_byte_length` notation (e.g. `30040_10` for 5 Modbus 16-bit words). Byte length must be an even positive integer.

Every register's complete occupied span must fit in the address space `0..65535`, including multiword numeric types, strings, raw bytes, and network-address types. For example, `U16` at `65535` is valid, while `U32` there extends past the final Modbus word and fails strict validation.

Numeric Excel addresses equal to zero are preserved. A supplied `Factor`, including zero, takes precedence over `Gain`; a reciprocal gain retains its precision until final six-decimal coefficient formatting. When `RAW` uses a separate `Length` column, that length is interpreted in bytes; a `Quantity` or register-count column is converted from words to bytes.

---

## 🔠 Normalized Data Types & Aliases

| Canonical Code | Accepted Vendor Synonyms | Size & Details |
|---|---|---|
| **U16** | `uint16`, `u16`, `int16u`, `unsigned short`, `hex`, `bitfield16` | 1 Modbus word (16-bit unsigned) |
| **I16** | `int16`, `i16`, `int16s`, `signed short`, `sint16` | 1 Modbus word (16-bit signed) |
| **U32** | `uint32`, `u32`, `int32u`, `unsigned int 32`, `datetime` | 2 Modbus words (32-bit unsigned) |
| **I32** | `int32`, `i32`, `int32s`, `signed int 32` | 2 Modbus words (32-bit signed) |
| **F32** | `float`, `float32`, `f32`, `32-bit IEEE 754` | 2 Modbus words (single-precision float) |
| **U64** | `uint64`, `u64`, `int64u`, `unsigned int 64` | 4 Modbus words (64-bit unsigned) |
| **I64** | `int64`, `i64`, `int64s`, `signed int 64` | 4 Modbus words (64-bit signed) |
| **F64** | `double`, `float64`, `f64`, `64-bit IEEE 754` | 4 Modbus words (double-precision float) |
| **BITS** | `bit16`, `bitmap16`, `bits16`, `bitfield` | Bitfield slice within a 16-bit register |
| **IP / IPV6** | `ip`, `ip4`, `ipv4`, `ipv6` | Network IP address representation |
| **STRING** | `string 20`, `str*30`, `string_16`, `STR` | ASCII character string |
| **RAW** | `raw` | Sequence of raw uninterpreted bytes (even length) |

---

## 🧩 Strict Rules for Bitfield Slices (`BITS`)

A `BITS` slice represents a discrete sequence of bits **within a single 16-bit Modbus register**.

The validator strictly enforces the following domain invariants:
- `startbit` must be between `0` and `15` inclusive.
- `length` must be $\ge 1$.
- `startbit + length` must **never exceed 16**.

| Address Notation | Validation Status | Explanation |
|---|---|---|
| `30001_0_1` | ✅ Valid | Bit 0 isolated |
| `30001_8_8` | ✅ Valid | Upper byte (bits 8 through 15) |
| `30001_15_1` | ✅ Valid | Highest register bit (bit 15) |
| `30001_15_2` | ❌ Invalid | Extends past register boundary (bit 16 does not exist) |
| `30001_0_0` | ❌ Invalid | Zero-length bit slice |

> [!IMPORTANT]
> **Bit-Level Overlap Resolution:**  
> Disjoint bit slices on the same base register are fully valid (e.g. `30001_0_4` and `30001_4_4`). Conversely, overlapping slices (e.g. `30001_0_4` and `30001_2_4`) are rejected by strict validation.

---

## 📐 Mathematical Scaling Coefficients Calculation

WebdynSunPM data loggers compute the physical reading using the linear equation:

$$\text{Physical Value} = \text{CoefA} \times \text{Raw Register Value} + \text{CoefB}$$

The generator computes these coefficients automatically:
- $\text{CoefA} = \text{Factor} \times 10^{\text{ScaleFactor}}$ (default: `1.000000`)
- $\text{CoefB} = \text{Offset}$ (default: `0.000000`)

*Concrete examples:*
- Voltage with vendor scale factor $0.1$: $\text{CoefA} = 0.100000$
- Active Power with vendor factor $10$ and $\text{ScaleFactor} = -2$: $\text{CoefA} = 10 \times 10^{-2} = 0.100000$
- Ambient temperature with $-40^\circ\text{C}$ offset: $\text{CoefB} = -40.000000$

---

## 📄 Official WebdynSunPM CSV Definition Structure

The generated definition is a semicolon-delimited (`;`) CSV file encoded in **UTF-8 with BOM**.

### Line 1: Equipment Metadata Header
Contains 5 metadata fields followed by 6 empty separators to match the required 11-column width:
```text
Protocol;Category;Manufacturer;Model;Forced writing code;;;;;;
```

### Lines 2+: Register Rows (11 mandatory columns)
```text
Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action
```

- **Index**: 1-based sequential line counter (`1`, `2`, `3`...).
- **Info1**: Modbus function code (`1`=Coil, `2`=Discrete Input, `3`=Holding Register, `4`=Input Register).
- **Info2**: Normalized Modbus address (e.g. `40001`, `30001_0_1`).
- **Info3**: Normalized Webdyn type code (e.g. `U16`, `F32`, `BITS`).
- **Info4**: Reserved field (always empty `""`).
- **Name**: Clear variable description.
- **Tag**: Unique variable identifier in lowercase ASCII.
- **CoefA**: Multiplier factor formatted to 6 decimal places.
- **CoefB**: Offset addition bias formatted to 6 decimal places.
- **Unit**: Engineering measurement unit (`V`, `A`, `W`, `°C`...).
- **Action**: Webdyn action code (`4` for read-only, `10` for constant).

---

<p align="center">
  <a href="../README.md"><b>⬅ Back to README</b></a> •
  <a href="architecture.md"><b>Architecture & Design ➡</b></a>
</p>
