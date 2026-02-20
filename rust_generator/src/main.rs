use clap::Parser;
use csv::{ReaderBuilder, WriterBuilder};
use regex::Regex;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::{self, Write};
use std::path::Path;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Input simplified CSV file
    input_file: String,

    /// Output definition CSV file
    #[arg(short, long)]
    output: Option<String>,

    /// Manufacturer name
    #[arg(long)]
    manufacturer: String,

    /// Model name
    #[arg(long)]
    model: String,

    /// Protocol name
    #[arg(long, default_value = "modbusRTU")]
    protocol: String,

    /// Device category
    #[arg(long, default_value = "Inverter")]
    category: String,

    /// Forced write code
    #[arg(long, default_value = "")]
    forced_write: String,

    /// Address offset to subtract
    #[arg(long, default_value_t = 0)]
    address_offset: i32,
}

#[derive(Debug, Deserialize)]
struct InputRow {
    #[serde(alias = "Name")]
    name: Option<String>,
    #[serde(alias = "Tag")]
    tag: Option<String>,
    #[serde(alias = "RegisterType")]
    register_type: Option<String>,
    #[serde(alias = "Address")]
    address: Option<String>,
    #[serde(alias = "Type")]
    dtype: Option<String>,
    #[serde(alias = "Factor")]
    factor: Option<String>,
    #[serde(alias = "Offset")]
    offset: Option<String>,
    #[serde(alias = "Unit")]
    unit: Option<String>,
    #[serde(alias = "Action")]
    action: Option<String>,
    #[serde(alias = "ScaleFactor")]
    scale_factor: Option<String>,
}

#[derive(Debug, Serialize)]
struct ProcessedRow {
    info1: String,
    info2: String,
    info3: String,
    info4: String,
    name: String,
    tag: String,
    coef_a: String,
    coef_b: String,
    unit: String,
    action: String,
}

struct Generator {
    address_offset: i32,
    register_type_map: HashMap<String, String>,
}

impl Generator {
    fn new(address_offset: i32) -> Self {
        let mut m = HashMap::new();
        m.insert("coil".to_string(), "1".to_string());
        m.insert("coils".to_string(), "1".to_string());
        m.insert("discrete input".to_string(), "2".to_string());
        m.insert("holding register".to_string(), "3".to_string());
        m.insert("holding".to_string(), "3".to_string());
        m.insert("input register".to_string(), "4".to_string());
        m.insert("input".to_string(), "4".to_string());
        Self {
            address_offset,
            register_type_map: m,
        }
    }

    fn normalize_type(&self, dtype: &str) -> String {
        let dtype_lower = dtype.to_lowercase();
        let cleaned = dtype_lower
            .replace("unsigned", "u")
            .replace("signed", "i")
            .replace(" ", "");

        match cleaned.as_str() {
            "uint8" => "U8".to_string(),
            "int8" => "I8".to_string(),
            "uint16" => "U16".to_string(),
            "int16" => "I16".to_string(),
            "uint32" => "U32".to_string(),
            "int32" => "I32".to_string(),
            "uint64" => "U64".to_string(),
            "int64" => "I64".to_string(),
            "float" | "float32" => "F32".to_string(),
            "double" | "float64" => "F64".to_string(),
            _ => {
                let re = Regex::new(r"(?i)([UI]\d+|F\d+|STRING|BITS|IP|IPV6|MAC|STR\d+)").unwrap();
                if let Some(caps) = re.captures(&cleaned.to_uppercase()) {
                    caps[1].to_string()
                } else {
                    cleaned.to_uppercase()
                }
            }
        }
    }

    fn validate_type(&self, dtype: &str) -> bool {
        let dtype_upper = dtype.to_uppercase();
        let base_types = vec!["STRING", "BITS", "IP", "IPV6", "MAC", "F32", "F64"];
        if base_types.contains(&dtype_upper.as_str()) {
            return true;
        }

        let re_int = Regex::new(r"(?i)^[UI](8|16|32|64)(_(W|B|WB))?$").unwrap();
        if re_int.is_match(&dtype_upper) {
            return true;
        }

        let re_str_conv = Regex::new(r"(?i)^STR(\d+)$").unwrap();
        if re_str_conv.is_match(&dtype_upper) {
            return true;
        }

        false
    }

    fn normalize_address_val(&self, addr_part: &str) -> String {
        let cleaned = addr_part.trim().replace(",", "");
        if cleaned.is_empty() {
            return "".to_string();
        }

        let is_negative = cleaned.starts_with('-');
        let abs_part = if is_negative { &cleaned[1..] } else { &cleaned };

        let val = if abs_part.to_lowercase().starts_with("0x") {
            u64::from_str_radix(&abs_part[2..], 16).ok()
        } else if abs_part.to_lowercase().ends_with('h') {
            u64::from_str_radix(&abs_part[..abs_part.len() - 1], 16).ok()
        } else if abs_part.chars().any(|c| c.is_ascii_alphabetic())
            && abs_part.chars().all(|c| c.is_ascii_hexdigit())
        {
            u64::from_str_radix(abs_part, 16).ok()
        } else {
            abs_part.parse::<u64>().ok()
        };

        match val {
            Some(v) => {
                let final_v = if is_negative { -(v as i64) } else { v as i64 };
                final_v.to_string()
            }
            None => cleaned,
        }
    }

    fn validate_address(&self, address: &str, dtype: &str) -> bool {
        let dtype_upper = dtype.to_uppercase();
        let addr_hex_pattern = r"(-?[0-9A-F]+|-?0x[0-9A-F]+|-?[0-9A-F]+h)";

        if dtype_upper == "STRING" {
            let re = Regex::new(&format!(r"(?i)^{}_\d+$", addr_hex_pattern)).unwrap();
            re.is_match(address)
        } else if dtype_upper == "BITS" {
            let re = Regex::new(&format!(r"(?i)^{}_\d+_\d+$", addr_hex_pattern)).unwrap();
            re.is_match(address)
        } else {
            let re = Regex::new(&format!(r"(?i)^{}$", addr_hex_pattern)).unwrap();
            re.is_match(address)
        }
    }

    fn get_register_count(&self, dtype: &str, address: &str) -> i32 {
        let dtype_upper = dtype.to_uppercase();
        let re_16_8 = Regex::new(r"(?i)^([UI](16|8)(_(W|B|WB))?|BITS)$").unwrap();
        let re_32 = Regex::new(r"(?i)^([UI]32(_(W|B|WB))?|F32|IP)$").unwrap();
        let re_64 = Regex::new(r"(?i)^([UI]64(_(W|B|WB))?|F64)$").unwrap();

        if re_16_8.is_match(&dtype_upper) {
            1
        } else if re_32.is_match(&dtype_upper) {
            2
        } else if re_64.is_match(&dtype_upper) {
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

    fn normalize_action(&self, action: &str) -> String {
        let a = action.trim().to_uppercase();
        if a.is_empty() {
            return "1".to_string();
        }
        match a.as_str() {
            "R" | "READ" | "4" => "4".to_string(),
            "RW" | "W" | "WRITE" | "1" => "1".to_string(),
            "0" | "2" | "6" | "7" | "8" | "9" => a,
            _ => "1".to_string(),
        }
    }

    fn process_rows(&self, rows: Vec<InputRow>) -> Vec<ProcessedRow> {
        let mut processed = Vec::new();
        let mut seen_names: HashMap<String, usize> = HashMap::new();
        let mut seen_tags: HashMap<String, usize> = HashMap::new();
        let mut used_addresses: HashMap<String, Vec<(i32, i32, usize, String, String)>> =
            HashMap::new();

        let re_str_conv = Regex::new(r"(?i)^STR(\d+)$").unwrap();

        for (i, row) in rows.into_iter().enumerate() {
            let line_num = i + 2;
            let name = row.name.clone().unwrap_or_default().trim().to_string();
            let mut address = row.address.clone().unwrap_or_default().trim().to_string();
            let raw_dtype = row.dtype.clone().unwrap_or_default();

            if name.is_empty() && address.is_empty() {
                continue;
            }

            let mut dtype = self.normalize_type(&raw_dtype);
            if !self.validate_type(&dtype) {
                eprintln!("Line {}: Invalid Type '{}'. Skipping row.", line_num, dtype);
                continue;
            }

            if let Some(caps) = re_str_conv.captures(&dtype) {
                let length = caps[1].to_string();
                dtype = "STRING".to_string();
                if !address.contains('_') {
                    address = format!("{}_{}", address, length);
                }
            }

            if !address.is_empty() {
                let parts: Vec<&str> = address.split('_').collect();
                let mut norm_parts: Vec<String> = parts
                    .iter()
                    .map(|p| self.normalize_address_val(p))
                    .collect();

                if !norm_parts.is_empty() {
                    if let Ok(raw_start) = norm_parts[0].parse::<i32>() {
                        let start_addr = raw_start - self.address_offset;
                        if start_addr < 0 {
                            eprintln!(
                                "Line {}: Address {} with offset {} results in negative address {}",
                                line_num, raw_start, self.address_offset, start_addr
                            );
                        }
                        norm_parts[0] = start_addr.to_string();
                    }
                }
                address = norm_parts.join("_");
            }

            if !self.validate_address(&address, &dtype) {
                eprintln!(
                    "Line {}: Invalid Address '{}' for Type '{}'. Skipping row.",
                    line_num, address, dtype
                );
                continue;
            }

            if !name.is_empty() {
                if let Some(prev_line) = seen_names.get(&name) {
                    eprintln!(
                        "Line {}: Duplicate Name '{}' detected. Previous occurrence at line {}.",
                        line_num, name, prev_line
                    );
                } else {
                    seen_names.insert(name.clone(), line_num);
                }
            }

            let mut tag = row.tag.clone().unwrap_or_default().trim().to_string();
            if tag.is_empty() && !name.is_empty() {
                let re_tag = Regex::new(r"[^a-z0-9_]").unwrap();
                let mut base_tag = re_tag
                    .replace_all(&name.to_lowercase().replace(" ", "_"), "")
                    .to_string();
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
                    eprintln!(
                        "Line {}: Duplicate Tag '{}' detected. Previous occurrence at line {}.",
                        line_num, tag, prev_line
                    );
                } else {
                    seen_tags.insert(tag.clone(), line_num);
                }
            }

            let reg_type_str = row.register_type.clone().unwrap_or_default();
            let info1 = if let Some(code) = self.register_type_map.get(&reg_type_str.to_lowercase())
            {
                code.clone()
            } else if vec!["1", "2", "3", "4"].contains(&reg_type_str.as_str()) {
                reg_type_str
            } else {
                "3".to_string()
            };

            // Overlap check
            if let Ok(start_addr) = address.split('_').next().unwrap_or("0").parse::<i32>() {
                let reg_count = self.get_register_count(&dtype, &address);
                let end_addr = start_addr + reg_count - 1;
                let current_range = (start_addr, end_addr, line_num, name.clone(), dtype.clone());

                let used = used_addresses.entry(info1.clone()).or_insert_with(Vec::new);
                let is_bits = dtype.to_uppercase() == "BITS";

                for (u_start, u_end, u_line, u_name, u_type) in used.iter() {
                    if std::cmp::max(start_addr, *u_start) <= std::cmp::min(end_addr, *u_end) {
                        let allowed = is_bits && u_type.to_uppercase() == "BITS";
                        if !allowed {
                            eprintln!("Line {}: Address overlap detected for '{}' (Addr: {}-{}). Overlaps with '{}' (Line {}, Addr: {}-{}) in register type {}.",
                                line_num, name, start_addr, end_addr, u_name, u_line, u_start, u_end, info1);
                        }
                    }
                }
                used.push(current_range);
            }

            let factor_str = row.factor.clone().unwrap_or_default();
            let val_factor = if factor_str.contains('/') {
                let parts: Vec<&str> = factor_str.split('/').collect();
                if parts.len() == 2 {
                    let p1 = parts[0].parse::<f64>().unwrap_or(1.0);
                    let p2 = parts[1].parse::<f64>().unwrap_or(1.0);
                    p1 / p2
                } else {
                    1.0
                }
            } else {
                factor_str.parse::<f64>().unwrap_or(1.0)
            };

            let scale_val = row
                .scale_factor
                .clone()
                .unwrap_or_default()
                .parse::<f64>()
                .unwrap_or(0.0);
            let final_coef_a = val_factor * 10.0f64.powf(scale_val);

            let coef_b_val = row
                .offset
                .clone()
                .unwrap_or_default()
                .parse::<f64>()
                .unwrap_or(0.0);

            processed.push(ProcessedRow {
                info1,
                info2: address,
                info3: dtype.to_uppercase(),
                info4: "".to_string(),
                name,
                tag,
                coef_a: format!("{:.6}", final_coef_a),
                coef_b: format!("{:.6}", coef_b_val),
                unit: row.unit.clone().unwrap_or_default().trim().to_string(),
                action: self.normalize_action(&row.action.unwrap_or_default()),
            });
        }
        processed
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    let path = Path::new(&args.input_file);
    let file = File::open(path)?;
    let mut reader = ReaderBuilder::new()
        .flexible(true)
        .from_reader(file);

    let mut rows = Vec::new();
    for result in reader.deserialize() {
        let row: InputRow = result?;
        rows.push(row);
    }

    let generator = Generator::new(args.address_offset);
    let processed = generator.process_rows(rows);

    let mut writer_buf: Box<dyn Write> = match args.output {
        Some(path) => Box::new(File::create(path)?),
        None => Box::new(io::stdout()),
    };

    // Header row
    let header = format!(
        "{};{};{};{};{};;;;;;\n",
        args.protocol, args.category, args.manufacturer, args.model, args.forced_write
    );
    writer_buf.write_all(header.as_bytes())?;

    let mut writer = WriterBuilder::new()
        .delimiter(b';')
        .has_headers(false)
        .from_writer(writer_buf);

    for (i, row) in processed.into_iter().enumerate() {
        writer.write_record(&[
            (i + 1).to_string(),
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
        ])?;
    }

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
        assert_eq!(gen.normalize_address_val("10"), "10");
        assert_eq!(gen.normalize_address_val("A0"), "160");
        assert_eq!(gen.normalize_address_val("1,234"), "1234");
        assert_eq!(gen.normalize_address_val("-10"), "-10");
    }

    #[test]
    fn test_normalize_type() {
        let gen = Generator::new(0);
        assert_eq!(gen.normalize_type("uint16"), "U16");
        assert_eq!(gen.normalize_type("float32"), "F32");
        assert_eq!(gen.normalize_type("STR20"), "STR20");
    }

    #[test]
    fn test_address_offset() {
        let gen = Generator::new(1);
        let row = InputRow {
            name: Some("Test".to_string()),
            tag: None,
            register_type: Some("Holding Register".to_string()),
            address: Some("100".to_string()),
            dtype: Some("U16".to_string()),
            factor: None,
            offset: None,
            unit: None,
            action: None,
            scale_factor: None,
        };
        let processed = gen.process_rows(vec![row]);
        assert_eq!(processed[0].info2, "99");
    }
}
