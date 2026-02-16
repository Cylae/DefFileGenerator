use clap::Parser;
use csv::{ReaderBuilder, WriterBuilder};
use regex::Regex;
use serde::Deserialize;
use std::collections::{HashMap, HashSet};
use std::error::Error;
use std::fs::File;
use std::io::{self, Write};
use lazy_static::lazy_static;

lazy_static! {
    static ref RE_TYPE_INT: Regex = Regex::new(r"(?i)^[UI](8|16|32|64)(_(W|B|WB))?$").unwrap();
    static ref RE_TYPE_STR_CONV: Regex = Regex::new(r"(?i)^STR(\d+)$").unwrap();
    static ref RE_ADDR_STRING: Regex = Regex::new(r"(?i)^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$").unwrap();
    static ref RE_ADDR_BITS: Regex = Regex::new(r"(?i)^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$").unwrap();
    static ref RE_ADDR_INT: Regex = Regex::new(r"(?i)^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$").unwrap();
    static ref RE_COUNT_16_8: Regex = Regex::new(r"(?i)^([UI](16|8)(_(W|B|WB))?|BITS)$").unwrap();
    static ref RE_COUNT_32: Regex = Regex::new(r"(?i)^([UI]32(_(W|B|WB))?|F32|IP)$").unwrap();
    static ref RE_COUNT_64: Regex = Regex::new(r"(?i)^([UI]64(_(W|B|WB))?|F64)$").unwrap();
}

#[derive(Parser, Debug)]
#[clap(author, version, about, long_about = None)]
struct Args {
    #[clap(value_parser)]
    input_file: String,

    #[clap(short, long)]
    output: Option<String>,

    #[clap(long)]
    manufacturer: String,

    #[clap(long)]
    model: String,

    #[clap(long, default_value = "modbusRTU")]
    protocol: String,

    #[clap(long, default_value = "Inverter")]
    category: String,

    #[clap(long, default_value = "")]
    forced_write: String,
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
    factor: String,
    #[serde(alias = "Offset", default)]
    offset: String,
    #[serde(alias = "Unit", default)]
    unit: String,
    #[serde(alias = "Action", default)]
    action: String,
    #[serde(alias = "ScaleFactor", default)]
    scale_factor: String,
}

struct Generator {
    register_type_map: HashMap<String, String>,
    allowed_actions: HashSet<String>,
}

impl Generator {
    fn new() -> Self {
        let mut register_type_map = HashMap::new();
        register_type_map.insert("coil".to_string(), "1".to_string());
        register_type_map.insert("coils".to_string(), "1".to_string());
        register_type_map.insert("discrete input".to_string(), "2".to_string());
        register_type_map.insert("holding register".to_string(), "3".to_string());
        register_type_map.insert("holding".to_string(), "3".to_string());
        register_type_map.insert("input register".to_string(), "4".to_string());
        register_type_map.insert("input".to_string(), "4".to_string());

        let mut allowed_actions = HashSet::new();
        for code in &["0", "1", "2", "4", "6", "7", "8", "9"] {
            allowed_actions.insert(code.to_string());
        }

        Generator {
            register_type_map,
            allowed_actions,
        }
    }

    fn validate_type(&self, dtype: &str) -> bool {
        let dtype_upper = dtype.to_uppercase();
        let base_types = ["STRING", "BITS", "IP", "IPV6", "MAC", "F32", "F64"];
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
        if addr_part.to_lowercase().starts_with("0x") {
            if let Ok(val) = u64::from_str_radix(&addr_part[2..], 16) {
                return val.to_string();
            }
        } else if addr_part.to_lowercase().ends_with('h') {
            if let Ok(val) = u64::from_str_radix(&addr_part[..addr_part.len() - 1], 16) {
                return val.to_string();
            }
        } else if addr_part.chars().any(|c| c.is_ascii_alphabetic()) {
             if let Ok(val) = u64::from_str_radix(&addr_part, 16) {
                return val.to_string();
            }
        }
        addr_part.to_string()
    }

    fn validate_address(&self, address: &str, dtype: &str) -> bool {
        let dtype_upper = dtype.to_uppercase();
        match dtype_upper.as_str() {
            "STRING" => RE_ADDR_STRING.is_match(address),
            "BITS" => RE_ADDR_BITS.is_match(address),
            _ => RE_ADDR_INT.is_match(address),
        }
    }

    fn get_register_count(&self, dtype: &str, address: &str) -> u32 {
        let dtype_upper = dtype.to_uppercase();
        if RE_COUNT_16_8.is_match(&dtype_upper) {
            return 1;
        } else if RE_COUNT_32.is_match(&dtype_upper) {
            return 2;
        } else if RE_COUNT_64.is_match(&dtype_upper) {
            return 4;
        } else if dtype_upper == "MAC" {
            return 3;
        } else if dtype_upper == "IPV6" {
            return 8;
        } else if dtype_upper == "STRING" {
            let parts: Vec<&str> = address.split('_').collect();
            if parts.len() >= 2 {
                if let Ok(len) = parts[1].parse::<u32>() {
                    return (len + 1) / 2;
                }
            }
            return 0;
        }
        1
    }

    fn normalize_action(&self, action: &str) -> String {
        let act_upper = action.trim().to_uppercase();
        if act_upper.is_empty() {
            return "1".to_string();
        }
        match act_upper.as_str() {
            "R" | "READ" | "4" => "4".to_string(),
            "RW" | "W" | "WRITE" | "1" => "1".to_string(),
            _ => {
                if self.allowed_actions.contains(&act_upper) {
                    act_upper
                } else {
                    "1".to_string()
                }
            }
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();
    let gen = Generator::new();

    let file = File::open(&args.input_file)?;
    let mut rdr = ReaderBuilder::new()
        .flexible(true)
        .trim(csv::Trim::All)
        .from_reader(file);

    let mut processed_rows = Vec::new();
    let mut seen_names: HashMap<String, usize> = HashMap::new();
    let mut seen_tags: HashMap<String, usize> = HashMap::new();
    let mut used_addresses_by_type: HashMap<String, Vec<(u32, u32, usize, String, String)>> = HashMap::new();

    for (i, result) in rdr.deserialize().enumerate() {
        let line_num = i + 2;
        let mut row: InputRow = match result {
            Ok(r) => r,
            Err(e) => {
                eprintln!("Line {}: Error parsing row: {}", line_num, e);
                continue;
            }
        };

        if row.name.is_empty() && row.address.is_empty() {
            continue;
        }

        if !gen.validate_type(&row.dtype) {
            eprintln!("Line {}: Invalid Type '{}'. Skipping row.", line_num, row.dtype);
            continue;
        }

        let mut dtype = row.dtype.to_uppercase();
        let mut address = row.address.clone();

        let is_str_n = if let Some(caps) = RE_TYPE_STR_CONV.captures(&dtype) {
            Some(caps.get(1).unwrap().as_str().to_string())
        } else {
            None
        };

        if let Some(len) = is_str_n {
            dtype = "STRING".to_string();
            if !address.contains('_') {
                address = format!("{}_{}", address, len);
            }
        }

        let parts: Vec<&str> = address.split('_').collect();
        let norm_parts: Vec<String> = parts.iter().map(|p| gen.normalize_address_val(p)).collect();
        address = norm_parts.join("_");

        if !gen.validate_address(&address, &dtype) {
            eprintln!("Line {}: Invalid Address '{}' for Type '{}'. Skipping row.", line_num, address, dtype);
            continue;
        }

        if !row.name.is_empty() {
            if let Some(&prev_line) = seen_names.get(&row.name) {
                eprintln!("Line {}: Duplicate Name '{}' detected. Previous occurrence at line {}.", line_num, row.name, prev_line);
            } else {
                seen_names.insert(row.name.clone(), line_num);
            }
        }

        if row.tag.is_empty() && !row.name.is_empty() {
            let mut base_tag = row.name.to_lowercase()
                .replace(" ", "_")
                .chars()
                .filter(|c| c.is_alphanumeric() || *c == '_')
                .collect::<String>();
            if base_tag.is_empty() {
                base_tag = "var".to_string();
            }
            let mut tag = base_tag.clone();
            let mut counter = 1;
            while seen_tags.contains_key(&tag) {
                tag = format!("{}_{}", base_tag, counter);
                counter += 1;
            }
            row.tag = tag;
        }

        if !row.tag.is_empty() {
            if let Some(&prev_line) = seen_tags.get(&row.tag) {
                eprintln!("Line {}: Duplicate Tag '{}' detected. Previous occurrence at line {}.", line_num, row.tag, prev_line);
            } else {
                seen_tags.insert(row.tag.clone(), line_num);
            }
        }

        let mut info1 = "3".to_string();
        let reg_type_lower = row.register_type.to_lowercase();
        if let Some(val) = gen.register_type_map.get(&reg_type_lower) {
            info1 = val.clone();
        } else if ["1", "2", "3", "4"].contains(&row.register_type.as_str()) {
            info1 = row.register_type.clone();
        }

        // Overlap detection
        if let Ok(start_addr) = address.split('_').next().unwrap().parse::<u32>() {
            let reg_count = gen.get_register_count(&dtype, &address);
            let end_addr = start_addr + reg_count - 1;

            let used_list = used_addresses_by_type.entry(info1.clone()).or_insert_with(Vec::new);
            let is_bits = dtype == "BITS";

            for &(u_start, u_end, u_line, ref u_name, ref u_type) in used_list.iter() {
                if std::cmp::max(start_addr, u_start) <= std::cmp::min(end_addr, u_end) {
                    let is_overlap_allowed = is_bits && (u_type == "BITS");
                    if !is_overlap_allowed {
                         eprintln!("Line {}: Address overlap detected for '{}' (Addr: {}-{}). Overlaps with '{}' (Line {}, Addr: {}-{}) in register type {}.",
                            line_num, row.name, start_addr, end_addr, u_name, u_line, u_start, u_end, info1);
                    }
                }
            }
            used_list.push((start_addr, end_addr, line_num, row.name.clone(), dtype.clone()));
        }

        let mut factor = 1.0;
        if !row.factor.is_empty() {
            if row.factor.contains('/') {
                let parts: Vec<&str> = row.factor.split('/').collect();
                if parts.len() == 2 {
                    if let (Ok(num), Ok(den)) = (parts[0].parse::<f64>(), parts[1].parse::<f64>()) {
                        if den != 0.0 {
                            factor = num / den;
                        }
                    }
                }
            } else {
                factor = row.factor.parse::<f64>().unwrap_or(1.0);
            }
        }
        let scale_factor = row.scale_factor.parse::<f64>().unwrap_or(0.0);
        let coef_a = factor * 10f64.powf(scale_factor);
        let coef_b = row.offset.parse::<f64>().unwrap_or(0.0);

        processed_rows.push(vec![
            "".to_string(), // placeholder for index
            info1,
            address,
            dtype,
            "".to_string(), // info4
            row.name,
            row.tag,
            format!("{:.6}", coef_a),
            format!("{:.6}", coef_b),
            row.unit,
            gen.normalize_action(&row.action),
        ]);
    }

    let mut out: Box<dyn Write> = match args.output {
        Some(path) => Box::new(File::create(path)?),
        None => Box::new(io::stdout()),
    };

    let header = format!("{};{};{};{};{};;;;;;\n", args.protocol, args.category, args.manufacturer, args.model, args.forced_write);
    out.write_all(header.as_bytes())?;

    let mut wtr = WriterBuilder::new()
        .delimiter(b';')
        .terminator(csv::Terminator::Any(b'\n'))
        .from_writer(out);

    for (i, row) in processed_rows.iter_mut().enumerate() {
        row[0] = (i + 1).to_string();
        wtr.write_record(row)?;
    }
    wtr.flush()?;

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_validate_type() {
        let gen = Generator::new();
        assert!(gen.validate_type("U16"));
        assert!(gen.validate_type("I32_WB"));
        assert!(gen.validate_type("STR20"));
        assert!(gen.validate_type("BITS"));
        assert!(!gen.validate_type("UNKNOWN"));
    }

    #[test]
    fn test_normalize_address_val() {
        let gen = Generator::new();
        assert_eq!(gen.normalize_address_val("0x10"), "16");
        assert_eq!(gen.normalize_address_val("10h"), "16");
        assert_eq!(gen.normalize_address_val("10"), "10");
        assert_eq!(gen.normalize_address_val("A0"), "160");
        assert_eq!(gen.normalize_address_val("1,234"), "1234");
    }

    #[test]
    fn test_get_register_count() {
        let gen = Generator::new();
        assert_eq!(gen.get_register_count("U16", "30000"), 1);
        assert_eq!(gen.get_register_count("U32", "30000"), 2);
        assert_eq!(gen.get_register_count("STRING", "30000_10"), 5);
        assert_eq!(gen.get_register_count("STRING", "30000_11"), 6);
    }
}
