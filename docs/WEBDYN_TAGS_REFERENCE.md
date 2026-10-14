# 🏷️ Webdyn SunPM Definition Tags Reference

> **Official Reference Specification** based on `SunPM_DEfFile_Tags_v1.xlsx`.  
> Defines standard telemetry, command, status, and metadata variable tags used in WebdynSunPM definition files (`Field 7` of CSV definitions) and their integration with supervisory systems (SynaptiQ / 3E).

---

## 📌 Introduction & Architectural Role

In WebdynSunPM definition files, **Field 7** specifies the **Tag** associated with a Modbus variable:

```csv
Index;Info1;Info2;Info3;Info4;Info5;Info6;Info7;Info8;Info9;Info10
1;3;40071;U16;;Active Power;P_AC;1.0;0.0;W;4
```

Tags allow supervisory platforms (Webdyn gateways, Plant Manager, SynaptiQ, SCADA) to:
1. **Identify standardized measures** without relying on manufacturer-specific variable labels in various languages.
2. **Execute automated control functions** (e.g. active power limitation via `CMD_P_AC_SET_PERCENT`, remote start/stop via `CMD_ON`/`CMD_OFF`).
3. **Calculate key performance indicators (KPIs)** such as Performance Ratio (PR), availability, yield, and avoided $CO_2$.

---

## 📊 Category Taxonomies & Tag Dictionary

### 1. Inverters (`Inverter` / `INV_`)

#### 1.1 Metadata & Identification (Basic Info)
| Tag | Previous Tag | Description | MVP | Webdyn Action |
|:---|:---|:---|:---:|:---:|
| `SERIAL` | `DisplaySerial`, `DisplaySerialNumber` | Inverter serial number | ⭐ | `4` |
| `MODEL` | `DisplayModel` | Inverter model string | ⭐ | `4` |
| `MANUFACTURER` | `ManufacturerStr` | Inverter manufacturer name | | `4` |
| `UNIT_ID` | `SlaveAddress` | Modbus slave unit address | | `4` |
| `P_RATED` | `NominalPower` | Nominal / rated active power ($P_n$) | ⭐ | `4` |
| `Q_RATED` | | Nominal reactive power ($Q_n$) | | `4` |

#### 1.2 Basic Status & Diagnostics
| Tag | Description | Webdyn Action |
|:---|:---|:---:|
| `STATE[1..x]` | Inverter operating state word / enum | `9` |
| `ALARM[1..x]` | Fault codes, alarm bitmasks, trip status | `9` |
| `TOT` | Total Operating Time (hours/seconds) | `4` |
| `TEMP_INT` | Internal cabinet / heatsink temperature | `4` |
| `R_ISO` | Insulation resistance to earth | `4` |

#### 1.3 Remote Commands & Control
| Tag | Previous Tag | Description | MVP | Webdyn Action |
|:---|:---|:---|:---:|:---:|
| `CMD_ON` | `cmdOn` | Start inverter switch (Active High) | ⭐ | `1` |
| `CMD_OFF` | `cmdOff` | Stop inverter switch (Active Low) | ⭐ | `1` |
| `CMD_P_AC_LIM_PERCENT` | | Active Power limitation in percent (%) | ⭐ | `1` |
| `CMD_P_AC_LIM_ABS` | | Active Power limitation in absolute value (W/kW) | | `1` |
| `CMD_P_AC_SET_PERCENT` | `cmdPwrPercent` | Active Power setpoint in percent (%) | ⭐ | `1` |
| `CMD_P_AC_SET_ABS` | `cmdPwr` | Active Power setpoint in absolute value (W/kW) | ⭐ | `1` |
| `CMD_P_ENA` | `WMaxLim_Ena` | Power limitation enable latch | ⭐ | `1` |
| `CMD_P_MODE` | | Active power control mode selector | | `1` |
| `CMD_RAMP_TIME_S` | `WMaxLimPct_RmpTms` | Ramp time from current to target setpoint (seconds) | ⭐ | `1` |
| `CMD_REVERT_TIME_S` | `WMaxLimPct_RvrtTms` | Timeout fallback period for power limit (seconds) | ⭐ | `1` |
| `CMD_Q_AC_LIM_PERCENT` | | Reactive Power limitation in percent (%) | | `1` |
| `CMD_Q_AC_LIM_ABS` | | Reactive Power limitation in absolute value (var/kvar) | | `1` |
| `CMD_Q_AC_SET_PERCENT` | `cmdQPct` | Reactive Power setpoint in percent (%) | | `1` |
| `CMD_Q_AC_SET_ABS` | | Reactive Power setpoint in absolute value (var/kvar) | | `1` |
| `CMD_Q_ENA` | `VArPct_Ena` | Reactive power setpoint enable latch | | `1` |
| `CMD_Q_MODE` | `VArPct_Mod` | Reactive power control mode selector | | `1` |
| `CMD_Q_RAMP_TIME_S` | `VArPct_RmpTms` | Reactive power ramp time (seconds) | ⭐ | `1` |
| `CMD_Q_REVERT_TIME_S` | `VArPct_RvrtTms` | Reactive power revert timeout (seconds) | ⭐ | `1` |
| `CMD_Q_ANGLE` | `cmdAngle` | Reactive power phase angle setpoint | | `1` |
| `CMD_PF` | `cmdPF`, `OutPFSet` | Power Factor ($cos\,\phi$) setpoint | ⭐ | `1` |
| `CMD_PF_INT` | `cmdPFint` | Power Factor integer setpoint | | `1` |
| `CMD_PF_ENA` | `OutPFSet_Ena` | Power Factor setpoint enable | | `1` |
| `CMD_PF_RAMP_TIME_S` | `OutPFSet_RmpTms` | Power Factor ramp time | | `1` |
| `CMD_PF_REVERT_TIME_S` | `OutPFSet_RvrtTms` | Power Factor revert timeout | | `1` |
| `CMD_COMMIT` | `Commit` | Commit / persist configuration changes | | `1` |

#### 1.4 AC Measurements (AC Monitoring)
| Tag | Previous Tag | Description | MVP |
|:---|:---|:---|:---:|
| `P_AC` | `RealPower` | Total AC active power | ⭐ |
| `P_AC_[1..3]` | | AC active power per phase ($P_{L1}, P_{L2}, P_{L3}$) | ⭐ |
| `Q_AC` | `ReactivePower` | Total AC reactive power | ⭐ |
| `Q_AC_[1..3]` | | AC reactive power per phase ($Q_{L1}, Q_{L2}, Q_{L3}$) | ⭐ |
| `S_AC` | | Total AC apparent power | |
| `S_AC_[1..3]` | | AC apparent power per phase ($S_{L1}, S_{L2}, S_{L3}$) | |
| `PF` | `CosPhi`, `cosPHI` | Total Power factor ($cos\,\phi$) | ⭐ |
| `PF_[1..3]` | | Power factor per phase | ⭐ |
| `FREQ` | `GridFrequency` | Grid frequency (Hz) | ⭐ |
| `FREQ_[1..3]` | `GridFrequency1-2-3` | Grid frequency per phase (Hz) | ⭐ |
| `I_AC` | | Total AC current | ⭐ |
| `I_AC_[1..3]` | | AC current per phase ($I_{L1}, I_{L2}, I_{L3}$) | ⭐ |
| `U_AC_[1..3]` | | AC phase-to-neutral voltage ($V_{L1}, V_{L2}, V_{L3}$) | ⭐ |
| `U_AC_L1L2`, `L2L3`, `L3L1` | | AC phase-to-phase voltage | ⭐ |
| `E_DAY` | `EnergyDay` | Daily energy yield (kWh) | ⭐ |
| `E_TOT` | `EnergyTotal` | Cumulative total energy yield (kWh) | ⭐ |

#### 1.5 DC Measurements (DC Monitoring)
| Tag | Description | MVP |
|:---|:---|:---:|
| `U_DC` | Total DC bus voltage | ⭐ |
| `I_DC` | Total DC bus current | ⭐ |
| `P_DC` | Total DC power | ⭐ |
| `U_DC_MPPT_[1..x]` | MPPT input voltage 1..x ($U_{mppt1}, U_{mppt2}\dots$) | ⭐ |
| `I_DC_MPPT_[1..x]` | MPPT input current 1..x ($I_{mppt1}, I_{mppt2}\dots$) | ⭐ |
| `P_DC_MPPT_[1..x]` | MPPT input power 1..x ($P_{mppt1}, P_{mppt2}\dots$) | ⭐ |
| `U_DC_STRING_[1..x]` | PV string voltage 1..x | |
| `I_DC_STRING_[1..x]` | PV string current 1..x | |
| `U_DC_NE` | Negative pole voltage to earth | |
| `U_DC_PE` | Positive pole voltage to earth | |

---

### 2. Meters & Network Analyzers (`Meter` / `MET_`)

| Tag | Previous Tag | Description | MVP |
|:---|:---|:---|:---:|
| `P_AC` | `ActivePowSumkW` | Total active power | ⭐ |
| `P_AC_[1..3]` | `ActivePow1..2..3kW` | Active power per phase ($P_{L1}, P_{L2}, P_{L3}$) | ⭐ |
| `Q_AC` | `ReactivePowSumkW` | Total reactive power | ⭐ |
| `Q_AC_[1..3]` | `ReactivePow1..2..3kW` | Reactive power per phase | ⭐ |
| `S_AC` | | Total apparent power | ⭐ |
| `S_AC_[1..3]` | | Apparent power per phase | ⭐ |
| `I_AC_[1..3]` | | Current per phase | ⭐ |
| `U_AC_[1..3]` | | Phase-to-neutral voltage | |
| `U_AC_L1L2`, `L2L3`, `L3L1` | | Phase-to-phase voltage | ⭐ |
| `U_LE_AC_[1..3]` | | Phase to earth voltage | |
| `U_NE` | | Neutral to earth voltage | |
| `PF` | | Total power factor | ⭐ |
| `PF_[1..3]` | | Power factor per phase | ⭐ |
| `FREQ` | | Grid frequency | ⭐ |
| `E_TOT` | | Total cumulative energy | ⭐ |
| `E_TOT_[1..3]` | | Total cumulative energy per phase | ⭐ |
| `E_EXP`, `E_EXP_[1..3]` | | Exported active energy (total & per phase) | |
| `E_IMP`, `E_IMP_[1..3]` | | Imported active energy (total & per phase) | |
| `ES`, `ES_[1..3]` | | Total apparent energy (total & per phase) | |
| `ES_EXP`, `ES_IMP` | | Export / import apparent energy | |
| `EQ`, `EQ_[1..3]` | | Total reactive energy (total & per phase) | |
| `EQ_EXP`, `EQ_IMP` | | Export / import reactive energy | |
| `SERIAL`, `MODEL` | `DisplaySerial`, `DisplayModel` | Meter identification | |

---

### 3. Sensors & Weather Stations (`Sensor` / `SEN_`)

| Tag | Previous Tag | Description | MVP | Unit |
|:---|:---|:---|:---:|:---:|
| `POA` | | Plane-of-Array Solar Irradiance | ⭐ | $W/m^2$ |
| `GHI` | | Global Horizontal Irradiance | ⭐ | $W/m^2$ |
| `TEMP_AMB` | | Ambient outdoor temperature | ⭐ | $^\circ C$ |
| `TEMP_CELL` | | PV module / cell temperature | ⭐ | $^\circ C$ |
| `WIND_SPEED` | | Anemometer wind speed | | $m/s$ |
| `WIND_BEARING` | | Wind direction angle | | $^\circ$ |
| `HUMIDITY` | | Relative air humidity | | $\%$ |
| `PRESSURE` | | Barometric air pressure | | $hPa$ |
| `RAIN` | | Precipitation rainfall | | $mm$ |
| `SERIAL`, `MODEL`, `MANUFACTURER` | | Sensor metadata | | |

---

### 4. Battery Energy Storage Systems (`Battery` / `BESS` / `BAT_`)

| Tag | SynaptiQ Indicator | Description | MVP |
|:---|:---|:---|:---:|
| `SOC` / `B_SoC` | `charge.state` | Battery State of Charge ($0-100\%$) | ⭐ |
| `SOH` / `B_SoH` | `charge.health` | Battery State of Health ($0-100\%$) | ⭐ |
| `TEMP` / `B_TEMP` | `temperature.battery` | Battery pack / cell temperature ($^\circ C$) | ⭐ |
| `U_DC` / `B_U_DC` | `voltage.battery` | Battery rack DC voltage ($V$) | ⭐ |

---

### 5. Communication Loggers & Displays (`COM_` / `OTH_`)

#### 5.1 Communication Gateways (`COM_` / `Logger`)
- `CMD_ON`, `CMD_OFF`: Remote control commands.
- `CMD_P_AC_SET_PERCENT`: Plant-level active power curtailment.
- `P_AC`, `E_DAY`, `E_TOT`, `PF`: Aggregated supervisory measures.

#### 5.2 Displays & Public Dashboards (`Display`)
- `P_AC` (prev: `Pac-Tag`): Current instantaneous power.
- `E_DAY` (prev: `EacDay-Tag`): Day yield.
- `E_TOT` (prev: `Eac-Tag`): Total generation.
- `AVOIDED_CO2` (prev: `Carb-Tag`): Total avoided greenhouse gases ($kg/t$).
- `AVOIDED_CO2_DAY` (prev: `CarbDay-Tag`): Today's avoided $CO_2$.
- `SELF_CONSUMPTION_RATIO` (prev: `AutoConso-Tag`): Self-consumption rate ($\%$).
- `LOAD_P_AC` (prev: `Conso-Tag`): Site active consumption power.

---

## 🔄 Legacy Tag Migration Matrix

When parsing or upgrading older definition files, `DefFileGenerator` automatically normalizes legacy tags to the official SunPM tags:

| Legacy Tag | Canonical SunPM Tag | Device Type |
|:---|:---|:---|
| `RealPower` | `P_AC` | Inverter / Meter / Logger |
| `ActivePowSumkW` | `P_AC` | Meter |
| `ActivePow1kW` / `2kW` / `3kW` | `P_AC_1` / `P_AC_2` / `P_AC_3` | Meter |
| `EnergyTotal` | `E_TOT` | Inverter / Meter / Logger |
| `EnergyDay` | `E_DAY` | Inverter / Meter / Logger |
| `NominalPower` | `P_RATED` | Inverter |
| `GridFrequency` | `FREQ` | Inverter / Meter |
| `GridFrequency1` / `2` / `3` | `FREQ_1` / `FREQ_2` / `FREQ_3` | Inverter / Meter |
| `CosPhi` / `cosPHI` | `PF` | Inverter / Meter / Logger |
| `ReactivePower` / `ReactivePowSumkW` | `Q_AC` | Inverter / Meter |
| `ReactivePow1kW` / `2kW` / `3kW` | `Q_AC_1` / `Q_AC_2` / `Q_AC_3` | Meter |
| `DisplayModel` | `MODEL` | All |
| `DisplaySerialNumber` / `DisplaySerial` | `SERIAL` | All |
| `ManufacturerStr` | `MANUFACTURER` | All |
| `SlaveAddress` | `UNIT_ID` | Inverter |
| `cmdOn` | `CMD_ON` | Inverter / Logger |
| `cmdOff` | `CMD_OFF` | Inverter / Logger |
| `cmdPwrPercent` | `CMD_P_AC_SET_PERCENT` | Inverter / Logger |
| `cmdPwr` | `CMD_P_AC_SET_ABS` | Inverter |
| `WMaxLim_Ena` | `CMD_P_ENA` | Inverter |
| `WMaxLimPct_RmpTms` | `CMD_RAMP_TIME_S` | Inverter |
| `WMaxLimPct_RvrtTms` | `CMD_REVERT_TIME_S` | Inverter |
| `cmdPF` / `OutPFSet` | `CMD_PF` | Inverter |
| `OutPFSet_Ena` | `CMD_PF_ENA` | Inverter |
| `VArPct_Ena` | `CMD_Q_ENA` | Inverter |
| `Commit` | `CMD_COMMIT` | Inverter |
| `Pac-Tag` | `P_AC` | Display |
| `Eac-Tag` | `E_TOT` | Display |
| `EacDay-Tag` | `E_DAY` | Display |
| `Carb-Tag` | `AVOIDED_CO2` | Display |
| `CarbDay-Tag` | `AVOIDED_CO2_DAY` | Display |
| `AutoConso-Tag` | `SELF_CONSUMPTION_RATIO` | Display |
| `Conso-Tag` | `LOAD_P_AC` | Display |

---

## 🛠️ Tooling & API Usage

### Command Line Interface (CLI)

```bash
# Enrich an existing definition CSV with inferred SunPM tags
deffilegen tag Definitions/INV_GoodWe_SMT_Series.csv -o enriched.csv

# Normalize legacy tags to standard SunPM in-place
deffilegen tag Definitions/INV_Growatt_Generic.csv --normalize --in-place
```

### Python API

```python
from DefFileGenerator.knowledge_miner import (
    infer_sunpm_tag,
    normalize_tag_to_sunpm,
    enrich_definition_rows,
)

# 1. Infer SunPM tag from register label
tag = infer_sunpm_tag("Phase A Voltage", equipment_type="Inverter")
# Returns: "U_AC_1"

# 2. Normalize legacy tags
canonical = normalize_tag_to_sunpm("RealPower")
# Returns: "P_AC"
```
