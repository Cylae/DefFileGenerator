# Input Formats & Column Mapping Mechanics

The extractor pipeline accepts manufacturer register maps in **PDF, XLSX, CSV, or XML** formats. Column headers vary across manufacturers (e.g., Huawei, SMA, SolarEdge, Schneider). The extractor automatically maps diverse input headers using two-pass heuristic matching and pattern dictionaries.

---

## 🔍 Two-Pass Heuristic Mapping Logic

When a document table is read, the extractor collects all column headers and evaluates them in three sequential tiers:

1. **Exact Case-Insensitive Match**: Direct comparison against target internal names (`Address`, `Name`, `Type`, `Unit`, `Action`, `Factor`, `Offset`, `ScaleFactor`, `Length`, `StartBit`).
2. **Pattern Match**: Comparison against pre-compiled synonym dictionaries in `Extractor.COLUMN_MAPPING`.
3. **Partial Match**: Substring matching for descriptive vendor headers (e.g., `"Register Start Address"` -> `Address`).

---

## 📋 Recommended Column Headers

| Internal Field | Mandatory | Example Values | Recognized Vendor Headers |
|---|---|---|---|
| **Address** | Yes | `40001`, `0x9C41`, `40001_0_8` | `address`, `addr`, `offset`, `register`, `reg`, `index` |
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
| **StartBit** | No | `0`, `8` | `startbit`, `bit offset`, `bit` |

---

## ⚙️ Custom Mapping JSON Override

To bypass auto-detection heuristics, provide a explicit JSON mapping file via `--mapping`:

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
3. **Compound Bitfields (`BITS`)**: Formatted as `address_startbit_length` (e.g. `30001_0_1`).
4. **Compound Strings (`STR<n>`)**: Formatted as `address_length` (e.g. `30030_20`).
