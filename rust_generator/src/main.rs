use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::{self, Write};
use regex::Regex;
use once_cell::sync::Lazy;
use serde::{Deserialize, Serialize};
use clap::Parser;

static RE_TYPE_INT: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^[UI](8|16|32|64)(_(W|B|WB))?$").unwrap());
static RE_TYPE_STR_CONV: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^STR(\d+)$").unwrap());
static RE_ADDR_STRING: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$").unwrap());
static RE_ADDR_BITS: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$").unwrap());
static RE_ADDR_INT: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$").unwrap());
static RE_COUNT_16_8: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([UI](16|8)(_(W|B|WB))?|BITS)$").unwrap());
static RE_COUNT_32: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([UI]32(_(W|B|WB))?|F32|IP)$").unwrap());
static RE_COUNT_64: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([UI]64(_(W|B|WB))?|F64)$").unwrap());

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    #[arg(help = "Path to the simplified CSV input file.")]
    input_file: String,

    #[arg(short, long, help = "Path to the output CSV file.")]
    output: Option<String>,

    #[arg(long, default_value = "modbusRTU", help = "Protocol name (default: modbusRTU).")]
    protocol: String,

    #[arg(long, default_value = "Inverter", help = "Device category (default: Inverter).")]
    category: String,

    #[arg(long, help = "Manufacturer name.")]
    manufacturer: String,

    #[arg(long, help = "Model name.")]
    model: String,

    #[arg(long, default_value = "", help = "Forced write code (default: empty).")]
    forced_write: String,

    #[arg(long, default_value_t = 0, help = "Value to subtract from register addresses.")]
    address_offset: i32,
}

#[derive(Debug, Deserialize)]
struct InputRow {
    #[serde(alias = "Name")]
    name: String,
    #[serde(alias = "Tag", default)]
    tag: String,
    #[serde(alias = "RegisterType")]
    register_type: String,
    #[serde(alias = "Address")]
    address: String,
    #[serde(alias = "Type")]
    dtype: String,
    #[serde(alias = "Factor", default)]
    factor: Option<String>,
    #[serde(alias = "Offset", default)]
    offset: Option<String>,
    #[serde(alias = "Unit", default)]
    unit: Option<String>,
    #[serde(alias = "Action", default)]
    action: Option<String>,
    #[serde(alias = "ScaleFactor", default)]
    scale_factor: Option<String>,
}

#[derive(Debug, Serialize)]
struct ProcessedRow {
    #[serde(rename = "Info1")]
    info1: String,
    #[serde(rename = "Info2")]
    info2: String,
    #[serde(rename = "Info3")]
    info3: String,
    #[serde(rename = "Info4")]
    info4: String,
    #[serde(rename = "Name")]
    name: String,
    #[serde(rename = "Tag")]
    tag: String,
    #[serde(rename = "CoefA")]
    coef_a: String,
    #[serde(rename = "CoefB")]
    coef_b: String,
    #[serde(rename = "Unit")]
    unit: String,
    #[serde(rename = "Action")]
    action: String,
}

struct Generator {
    address_offset: i32,
    register_type_map: HashMap<String, String>,
    allowed_actions: Vec<String>,
}

impl Generator {
    fn new(address_offset: i32) -> Self {
        let mut register_type_map = HashMap::new();
        register_type_map.insert("coil".to_string(), "1".to_string());
        register_type_map.insert("coils".to_string(), "1".to_string());
        register_type_map.insert("discrete input".to_string(), "2".to_string());
        register_type_map.insert("holding register".to_string(), "3".to_string());
        register_type_map.insert("holding".to_string(), "3".to_string());
        register_type_map.insert("input register".to_string(), "4".to_string());
        register_type_map.insert("input".to_string(), "4".to_string());

        let allowed_actions = vec!["0", "1", "2", "4", "6", "7", "8", "9"]
            .into_iter()
            .map(|s| s.to_string())
            .collect();

        Generator {
            address_offset,
            register_type_map,
            allowed_actions,
        }
    }

    fn normalize_type(&self, dtype: &str) -> String {
        if dtype.is_empty() {
            return "U16".to_string();
        }
        let mut t_str = dtype.to_lowercase().trim().to_string();
        t_str = t_str.replace("unsigned", "u").replace("signed", "i").replace(" ", "");

        let mut synonyms = HashMap::new();
        synonyms.insert("uint16", "U16");
        synonyms.insert("int16", "I16");
        synonyms.insert("uint32", "U32");
        synonyms.insert("int32", "I32");
        synonyms.insert("uint64", "U64");
        synonyms.insert("int64", "I64");
        synonyms.insert("float32", "F32");
        synonyms.insert("float", "F32");
        synonyms.insert("float64", "F64");
        synonyms.insert("double", "F64");

        if let Some(val) = synonyms.get(t_str.as_str()) {
            return val.to_string();
        }

        dtype.to_uppercase()
    }

    fn validate_type(&self, dtype: &str) -> bool {
        let dtype_upper = dtype.to_uppercase();
        let base_types = vec!["STRING", "BITS", "IP", "IPV6", "MAC", "F32", "F64"];
        if base_types.contains(&dtype_upper.as_str()) {
            return true;
        }
        if RE_TYPE_INT.is_match(&dtype_upper) {
            return true;
        }
        if RE_TYPE_STR_CONV.is_match(&dtype_upper) {
            return true;
        }
        false
    }

    fn normalize_address_val(&self, addr_part: &str) -> String {
        let addr_part = addr_part.trim().replace(",", "");
        if addr_part.is_empty() {
            return "".to_string();
        }
        let addr_part_lower = addr_part.to_lowercase();
        if addr_part_lower.starts_with("0x") {
            if let Ok(val) = i64::from_str_radix(&addr_part[2..], 16) {
                return val.to_string();
            }
        } else if addr_part_lower.ends_with('h') {
            if let Ok(val) = i64::from_str_radix(&addr_part[..addr_part.len() - 1], 16) {
                return val.to_string();
            }
        }
        if addr_part.chars().any(|c| c.is_ascii_alphabetic()) {
             if let Ok(val) = i64::from_str_radix(&addr_part, 16) {
                return val.to_string();
            }
        }
        addr_part
    }

    fn validate_address(&self, address: &str, dtype: &str) -> bool {
        let dtype_upper = dtype.to_uppercase();
        if dtype_upper == "STRING" {
            RE_ADDR_STRING.is_match(address)
        } else if dtype_upper == "BITS" {
            RE_ADDR_BITS.is_match(address)
        } else {
            RE_ADDR_INT.is_match(address)
        }
    }

    fn get_register_count(&self, dtype: &str, address: &str) -> i32 {
        let dtype_upper = dtype.to_uppercase();
        if RE_COUNT_16_8.is_match(&dtype_upper) {
            1
        } else if RE_COUNT_32.is_match(&dtype_upper) {
            2
        } else if RE_COUNT_64.is_match(&dtype_upper) {
            4
        } else if dtype_upper == "MAC" {
            3
        } else if dtype_upper == "IPV6" {
            8
        } else if dtype_upper == "STRING" {
            let parts: Vec<&str> = address.split('_').collect();
            if parts.len() >= 2 {
                if let Ok(len) = parts[1].parse::<f64>() {
                    return (len / 2.0).ceil() as i32;
                }
            }
            0
        } else {
            1
        }
    }

    fn process_rows(&self, rows: Vec<InputRow>) -> Vec<ProcessedRow> {
        let mut processed_rows = Vec::new();
        let mut seen_names: HashMap<String, usize> = HashMap::new();
        let mut seen_tags: HashMap<String, usize> = HashMap::new();
        let mut used_addresses_by_type: HashMap<String, Vec<(i32, i32, usize, String, String)>> = HashMap::new();

        for (idx, row) in rows.into_iter().enumerate() {
            let line_num = idx + 2;

            if row.name.is_empty() && row.address.is_empty() {
                eprintln!("Line {}: Skipping row with missing Name and Address.", line_num);
                continue;
            }

            let mut dtype = self.normalize_type(&row.dtype);
            if !self.validate_type(&dtype) {
                eprintln!("Line {}: Invalid Type '{}'. Skipping row.", line_num, dtype);
                continue;
            }

            let dtype_upper = dtype.to_uppercase();
            let mut address = row.address.clone();

            if let Some(caps) = RE_TYPE_STR_CONV.captures(&dtype_upper) {
                dtype = "STRING".to_string();
                if !address.contains('_') {
                    address = format!("{}_{}", address, &caps[1]);
                }
            }

            let parts: Vec<&str> = address.split('_').collect();
            let norm_parts: Vec<String> = parts.iter().map(|p| self.normalize_address_val(p)).collect();
            address = norm_parts.join("_");

            if !self.validate_address(&address, &dtype) {
                eprintln!("Line {}: Invalid Address '{}' for Type '{}'. Skipping row.", line_num, address, dtype);
                continue;
            }

            if !row.name.is_empty() {
                if let Some(prev_line) = seen_names.get(&row.name) {
                    eprintln!("Line {}: Duplicate Name '{}' detected. Previous occurrence at line {}.", line_num, row.name, prev_line);
                } else {
                    seen_names.insert(row.name.clone(), line_num);
                }
            }

            let mut tag = row.tag.clone();
            if tag.is_empty() && !row.name.is_empty() {
                let mut base_tag = row.name.to_lowercase().replace(' ', "_");
                base_tag = base_tag.chars().filter(|c| c.is_alphanumeric() || *c == '_').collect();
                if base_tag.is_empty() {
                    base_tag = "var".to_string();
                }
                tag = base_tag.clone();
                let mut counter = 1;
                while seen_tags.contains_key(&tag) {
                    tag = format!("{}_{}", base_tag, counter);
                    counter += 1;
                }
            }

            if !tag.is_empty() {
                if let Some(prev_line) = seen_tags.get(&tag) {
                    eprintln!("Line {}: Duplicate Tag '{}' detected. Previous occurrence at line {}.", line_num, tag, prev_line);
                } else {
                    seen_tags.insert(tag.clone(), line_num);
                }
            }

            let mut info1 = "3".to_string();
            let reg_type_str = row.register_type.clone();
            if !reg_type_str.is_empty() {
                let lower_type = reg_type_str.to_lowercase();
                if let Some(val) = self.register_type_map.get(&lower_type) {
                    info1 = val.clone();
                } else if vec!["1", "2", "3", "4"].contains(&reg_type_str.as_str()) {
                    info1 = reg_type_str;
                } else {
                    eprintln!("Line {}: Unknown RegisterType '{}'. Defaulting to Holding Register (3).", line_num, reg_type_str);
                }
            }

            let mut info2 = address.clone();
            let mut start_addr = 0;
            let mut end_addr = 0;
            let mut calc_ok = false;

            if let Ok(raw_start) = info2.split('_').next().unwrap_or("").parse::<i32>() {
                start_addr = raw_start - self.address_offset;
                if start_addr < 0 {
                    eprintln!("Line {}: Address {} with offset {} results in negative address {}", line_num, raw_start, self.address_offset, start_addr);
                }

                let parts: Vec<&str> = info2.split('_').collect();
                if parts.len() > 1 {
                    info2 = format!("{}_{}", start_addr, parts[1..].join("_"));
                } else {
                    info2 = start_addr.to_string();
                }

                let reg_count = self.get_register_count(&dtype, &info2);
                end_addr = start_addr + reg_count - 1;
                calc_ok = true;
            }

            if calc_ok {
                let is_bits = dtype.to_uppercase() == "BITS";
                if let Some(used) = used_addresses_by_type.get(&info1) {
                    for (u_start, u_end, u_line, u_name, u_type) in used {
                        if std::cmp::max(start_addr, *u_start) <= std::cmp::min(end_addr, *u_end) {
                            let allowed = is_bits && u_type == "BITS";
                            if !allowed {
                                eprintln!("Line {}: Address overlap detected for '{}' (Addr: {}-{}). Overlaps with '{}' (Line {}, Addr: {}-{}) in register type {}.",
                                    line_num, row.name, start_addr, end_addr, u_name, u_line, u_start, u_end, info1);
                            }
                        }
                    }
                }
                used_addresses_by_type.entry(info1.clone()).or_insert(Vec::new()).push((start_addr, end_addr, line_num, row.name.clone(), dtype.to_uppercase()));
            }

            let factor_val: f64 = row.factor.as_ref().and_then(|f| f.parse().ok()).unwrap_or(1.0);
            let scale_val: i32 = row.scale_factor.as_ref().and_then(|s| s.parse::<f64>().ok()).map(|s| s as i32).unwrap_or(0);
            let coef_a = format!("{:.6}", factor_val * 10f64.powi(scale_val));

            let offset_val: f64 = row.offset.as_ref().and_then(|o| o.parse().ok()).unwrap_or(0.0);
            let coef_b = format!("{:.6}", offset_val);

            let mut action = row.action.unwrap_or_else(|| "1".to_string());
            if action.trim().is_empty() {
                action = "1".to_string();
            } else {
                let act_upper = action.trim().to_uppercase();
                if act_upper == "R" || act_upper == "READ" || act_upper == "4" {
                    action = "4".to_string();
                } else if act_upper == "RW" || act_upper == "W" || act_upper == "WRITE" || act_upper == "1" {
                    action = "1".to_string();
                } else if !self.allowed_actions.contains(&act_upper) {
                    eprintln!("Line {}: Invalid Action '{}'. Defaulting to '1'.", line_num, action);
                    action = "1".to_string();
                } else {
                    action = act_upper;
                }
            }

            processed_rows.push(ProcessedRow {
                info1,
                info2,
                info3: dtype.to_uppercase(),
                info4: "".to_string(),
                name: row.name,
                tag,
                coef_a,
                coef_b,
                unit: row.unit.unwrap_or_default(),
                action,
            });
        }
        processed_rows
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    let generator = Generator::new(args.address_offset);

    let file = File::open(&args.input_file)?;
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(true)
        .trim(csv::Trim::All)
        .flexible(true)
        .from_reader(file);

    let mut input_rows = Vec::new();
    for result in rdr.deserialize() {
        let row: InputRow = result?;
        input_rows.push(row);
    }

    let processed_rows = generator.process_rows(input_rows);

    let mut wtr: Box<dyn Write> = if let Some(output_path) = args.output {
        Box::new(File::create(output_path)?)
    } else {
        Box::new(io::stdout())
    };

    // Write header
    let header = format!(
        "{};{};{};{};{};;;;;;\n",
        args.protocol, args.category, args.manufacturer, args.model, args.forced_write
    );
    wtr.write_all(header.as_bytes())?;

    let mut csv_wtr = csv::WriterBuilder::new()
        .delimiter(b';')
        .quote_style(csv::QuoteStyle::Never)
        .from_writer(wtr);

    for (idx, row) in processed_rows.into_iter().enumerate() {
        let index = (idx + 1).to_string();
        csv_wtr.serialize((
            index,
            row.info1,
            row.info2,
            row.info3,
            row.info4,
            row.name,
            row.tag,
            row.coef_a,
            row.coef_b,
            row.unit,
            row.action,
        ))?;
    }

    csv_wtr.flush()?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_normalize_address_val() {
        let gen = Generator::new(0);
        assert_eq!(gen.normalize_address_val("0x10"), "16");
        assert_eq!(gen.normalize_address_val("10h"), "16");
        assert_eq!(gen.normalize_address_val("1,000"), "1000");
        assert_eq!(gen.normalize_address_val("ABC"), "2748");
        assert_eq!(gen.normalize_address_val("123"), "123");
    }

    #[test]
    fn test_validate_type() {
        let gen = Generator::new(0);
        assert!(gen.validate_type("U16"));
        assert!(gen.validate_type(&gen.normalize_type("uint16")));
        assert!(gen.validate_type("STR20"));
        assert!(gen.validate_type("F32"));
        assert!(!gen.validate_type("INVALID"));
    }

    #[test]
    fn test_address_offset() {
        let gen = Generator::new(10);
        let rows = vec![InputRow {
            name: "Test".to_string(),
            tag: "".to_string(),
            register_type: "Holding Register".to_string(),
            address: "100".to_string(),
            dtype: "U16".to_string(),
            factor: None,
            offset: None,
            unit: None,
            action: None,
            scale_factor: None,
        }];
        let processed = gen.process_rows(rows);
        assert_eq!(processed[0].info2, "90");
    }

    #[test]
    fn test_address_offset_composite() {
        let gen = Generator::new(10);
        let rows = vec![InputRow {
            name: "Test".to_string(),
            tag: "".to_string(),
            register_type: "Holding Register".to_string(),
            address: "100_20".to_string(),
            dtype: "STRING".to_string(),
            factor: None,
            offset: None,
            unit: None,
            action: None,
            scale_factor: None,
        }];
        let processed = gen.process_rows(rows);
        assert_eq!(processed[0].info2, "90_20");
    }

    #[test]
    fn test_get_register_count() {
        let gen = Generator::new(0);
        assert_eq!(gen.get_register_count("U16", "100"), 1);
        assert_eq!(gen.get_register_count("U32", "100"), 2);
        assert_eq!(gen.get_register_count("F64", "100"), 4);
        assert_eq!(gen.get_register_count("STRING", "100_10"), 5);
        assert_eq!(gen.get_register_count("IPV6", "100"), 8);
    }
}
