use clap::Parser;
use csv::{ReaderBuilder, WriterBuilder};
use serde::{Deserialize, Serialize};
use regex::Regex;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::{self, Read, Seek, SeekFrom};
use std::path::PathBuf;

static RE_TYPE_INT: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^[UI](8|16|32|64)(_(W|B|WB))?$").unwrap());
static RE_TYPE_STR_CONV: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^STR(\d+)$").unwrap());
static RE_ADDR_STRING: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$").unwrap());
static RE_ADDR_BITS: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$").unwrap());
static RE_ADDR_INT: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$").unwrap());
static RE_COUNT_16_8: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([UI](16|8)(_(W|B|WB))?|BITS)$").unwrap());
static RE_COUNT_32: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([UI]32(_(W|B|WB))?|F32|IP)$").unwrap());
static RE_COUNT_64: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([UI]64(_(W|B|WB))?|F64)$").unwrap());
static RE_TAG: Lazy<Regex> = Lazy::new(|| Regex::new(r"[^a-z0-9_]").unwrap());

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    input_file: Option<PathBuf>,

    #[arg(short, long)]
    output: Option<PathBuf>,

    #[arg(long)]
    protocol: Option<String>,

    #[arg(long)]
    category: Option<String>,

    #[arg(long)]
    manufacturer: String,

    #[arg(long)]
    model: String,

    #[arg(long, default_value = "")]
    forced_write: String,

    #[arg(long, default_value_t = 0)]
    address_offset: i32,

    #[arg(long)]
    template: bool,
}

#[derive(Debug, Deserialize)]
struct InputRow {
    #[serde(rename = "Name")]
    name: String,
    #[serde(rename = "Tag", default)]
    tag: String,
    #[serde(rename = "RegisterType")]
    register_type: String,
    #[serde(rename = "Address")]
    address: String,
    #[serde(rename = "Type")]
    dtype: String,
    #[serde(rename = "Factor", default)]
    factor: String,
    #[serde(rename = "Offset", default)]
    offset: String,
    #[serde(rename = "Unit", default)]
    unit: String,
    #[serde(rename = "Action", default)]
    action: String,
    #[serde(rename = "ScaleFactor", default)]
    scale_factor: String,
}

#[derive(Debug, Serialize)]
struct OutputRow {
    index: String,
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
        let mut register_type_map = HashMap::new();
        register_type_map.insert("coil".to_string(), "1".to_string());
        register_type_map.insert("coils".to_string(), "1".to_string());
        register_type_map.insert("discrete input".to_string(), "2".to_string());
        register_type_map.insert("holding register".to_string(), "3".to_string());
        register_type_map.insert("holding".to_string(), "3".to_string());
        register_type_map.insert("input register".to_string(), "4".to_string());
        register_type_map.insert("input".to_string(), "4".to_string());

        Generator {
            address_offset,
            register_type_map,
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
        let addr_part = addr_part.trim().replace(',', "");
        if addr_part.is_empty() {
            return "".to_string();
        }
        if addr_part.to_lowercase().starts_with("0x") {
            if let Ok(val) = i64::from_str_radix(&addr_part[2..], 16) {
                return val.to_string();
            }
        } else if addr_part.to_lowercase().ends_with('h') {
             if let Ok(val) = i64::from_str_radix(&addr_part[..addr_part.len()-1], 16) {
                return val.to_string();
            }
        }
        // If it contains A-F, try hex
        if addr_part.chars().any(|c| c.is_ascii_hexdigit() && !c.is_ascii_digit()) {
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
            if parts.len() > 1 {
                if let Ok(len) = parts[1].parse::<f64>() {
                    return (len / 2.0).ceil() as i32;
                }
            }
            0
        } else {
            1
        }
    }

    fn process_rows(&self, rows: Vec<InputRow>) -> Vec<OutputRow> {
        let mut processed_rows = Vec::new();
        let mut seen_names = HashMap::new();
        let mut seen_tags = HashMap::new();
        let mut used_addresses_by_type: HashMap<String, Vec<(i32, i32, usize, String, String)>> = HashMap::new();

        for (i, row) in rows.into_iter().enumerate() {
            let line_num = i + 2;
            if row.name.is_empty() && row.address.is_empty() {
                eprintln!("Line {}: Skipping row with missing Name and Address.", line_num);
                continue;
            }

            let mut dtype = row.dtype.clone();
            if !self.validate_type(&dtype) {
                 eprintln!("Line {}: Invalid Type '{}'. Skipping row.", line_num, dtype);
                 continue;
            }

            let dtype_upper = dtype.to_uppercase();
            let mut address = row.address.clone();
            if let Some(caps) = RE_TYPE_STR_CONV.captures(&dtype_upper) {
                dtype = "STRING".to_string();
                let length = caps.get(1).unwrap().as_str();
                if !address.contains('_') {
                    address = format!("{}_{}", address, length);
                }
            }

            if !address.is_empty() {
                let parts: Vec<&str> = address.split('_').collect();
                let norm_parts: Vec<String> = parts.into_iter().map(|p| self.normalize_address_val(p)).collect();
                address = norm_parts.join("_");
            }

            if !self.validate_address(&address, &dtype) {
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

            let mut tag = row.tag.clone();
            if tag.is_empty() && !row.name.is_empty() {
                let mut base_tag = row.name.to_lowercase().replace(' ', "_");
                base_tag = RE_TAG.replace_all(&base_tag, "").into_owned();
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
                if let Some(&prev_line) = seen_tags.get(&tag) {
                    eprintln!("Line {}: Duplicate Tag '{}' detected. Previous occurrence at line {}.", line_num, tag, prev_line);
                } else {
                    seen_tags.insert(tag.clone(), line_num);
                }
            }

            let mut info1 = "3".to_string();
            let lower_reg_type = row.register_type.to_lowercase();
            if let Some(val) = self.register_type_map.get(&lower_reg_type) {
                info1 = val.clone();
            } else if ["1", "2", "3", "4"].contains(&row.register_type.as_str()) {
                info1 = row.register_type.clone();
            } else {
                eprintln!("Line {}: Unknown RegisterType '{}'. Defaulting to Holding Register (3).", line_num, row.register_type);
            }

            let mut parts: Vec<String> = address.split('_').map(|s| s.to_string()).collect();
            if let Ok(raw_start_addr) = parts[0].parse::<i32>() {
                let start_addr = raw_start_addr - self.address_offset;
                if start_addr < 0 {
                    eprintln!("Line {}: Address {} with offset {} results in negative address {}", line_num, raw_start_addr, self.address_offset, start_addr);
                }
                parts[0] = start_addr.to_string();
                address = parts.join("_");

                let reg_count = self.get_register_count(&dtype, &address);
                let end_addr = start_addr + reg_count - 1;

                let current_range = (start_addr, end_addr, line_num, row.name.clone(), dtype.to_uppercase());
                let entries = used_addresses_by_type.entry(info1.clone()).or_insert_with(Vec::new);

                let is_bits = dtype.to_uppercase() == "BITS";
                for (used_start, used_end, used_line, used_name, used_type) in entries.iter() {
                    if std::cmp::max(start_addr, *used_start) <= std::cmp::min(end_addr, *used_end) {
                        let is_overlap_allowed = is_bits && used_type == "BITS";
                        if !is_overlap_allowed {
                             eprintln!("Line {}: Address overlap detected for '{}' (Addr: {}-{}). Overlaps with '{}' (Line {}, Addr: {}-{}) in register type {}.", line_num, row.name, start_addr, end_addr, used_name, used_line, used_start, used_end, info1);
                        }
                    }
                }
                entries.push(current_range);
            }

            let factor = row.factor.parse::<f64>().unwrap_or(1.0);
            let scale_factor = row.scale_factor.parse::<f64>().unwrap_or(0.0);
            let coef_a_val = factor * 10.0f64.powf(scale_factor);
            let coef_a = format!("{:.6}", coef_a_val);

            let offset = row.offset.parse::<f64>().unwrap_or(0.0);
            let coef_b = format!("{:.6}", offset);

            let mut action = row.action.to_uppercase();
            if action.is_empty() {
                action = "1".to_string();
            } else if action == "R" || action == "READ" || action == "4" {
                action = "4".to_string();
            } else if action == "RW" || action == "W" || action == "WRITE" || action == "1" {
                action = "1".to_string();
            } else if !["0", "1", "2", "4", "6", "7", "8", "9"].contains(&action.as_str()) {
                eprintln!("Line {}: Invalid Action '{}'. Defaulting to '1'.", line_num, action);
                action = "1".to_string();
            }

            processed_rows.push(OutputRow {
                index: (processed_rows.len() + 1).to_string(),
                info1,
                info2: address,
                info3: dtype.to_uppercase(),
                info4: "".to_string(),
                name: row.name,
                tag,
                coef_a,
                coef_b,
                unit: row.unit,
                action,
            });
        }
        processed_rows
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    if args.template {
        let stdout = io::stdout();
        let mut writer = csv::Writer::from_writer(stdout);
        writer.write_record(&["Name", "Tag", "RegisterType", "Address", "Type", "Factor", "Offset", "Unit", "Action", "ScaleFactor"])?;
        writer.write_record(&["Example Variable", "example_tag", "Holding Register", "30001", "U16", "1", "0", "V", "4", "0"])?;
        return Ok(());
    }

    let generator = Generator::new(args.address_offset);

    let input_path = args.input_file.ok_or("Input file is required")?;
    let mut file = File::open(input_path)?;

    // Delimiter detection
    let mut buffer = [0; 1024];
    let n = file.read(&mut buffer)?;
    let content = String::from_utf8_lossy(&buffer[..n]);
    let comma_count = content.matches(',').count();
    let semi_count = content.matches(';').count();
    let tab_count = content.matches('\t').count();

    let delimiter = if semi_count > comma_count && semi_count > tab_count {
        b';'
    } else if tab_count > comma_count {
        b'\t'
    } else {
        b','
    };

    file.seek(SeekFrom::Start(0))?;

    let mut reader = ReaderBuilder::new()
        .flexible(true)
        .delimiter(delimiter)
        .trim(csv::Trim::All)
        .from_reader(file);

    let mut rows = Vec::new();
    for result in reader.deserialize() {
        let row: InputRow = result?;
        rows.push(row);
    }

    let processed_rows = generator.process_rows(rows);

    let inner_writer: Box<dyn io::Write> = if let Some(output_path) = args.output {
        Box::new(File::create(output_path)?)
    } else {
        Box::new(io::stdout())
    };

    let mut writer = WriterBuilder::new()
        .delimiter(b';')
        .terminator(csv::Terminator::Any(b'\n'))
        .from_writer(inner_writer);

    // WebdynSunPM header
    writer.write_record(&[
        args.protocol.unwrap_or_else(|| "modbusRTU".to_string()),
        args.category.unwrap_or_else(|| "Inverter".to_string()),
        args.manufacturer,
        args.model,
        args.forced_write,
        "".to_string(), "".to_string(), "".to_string(), "".to_string(), "".to_string(), "".to_string()
    ])?;

    for row in processed_rows {
        writer.serialize(row)?;
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_normalize_address_val() {
        let gen = Generator::new(0);
        assert_eq!(gen.normalize_address_val("100"), "100");
        assert_eq!(gen.normalize_address_val("0x64"), "100");
        assert_eq!(gen.normalize_address_val("64h"), "100");
        assert_eq!(gen.normalize_address_val("1,000"), "1000");
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
            factor: "1".to_string(),
            offset: "0".to_string(),
            unit: "V".to_string(),
            action: "4".to_string(),
            scale_factor: "0".to_string(),
        }];
        let processed = gen.process_rows(rows);
        assert_eq!(processed[0].info2, "90");
    }

    #[test]
    fn test_coef_a_calculation() {
        let gen = Generator::new(0);
        let rows = vec![InputRow {
            name: "Test".to_string(),
            tag: "".to_string(),
            register_type: "Holding Register".to_string(),
            address: "100".to_string(),
            dtype: "U16".to_string(),
            factor: "0.1".to_string(),
            offset: "0".to_string(),
            unit: "V".to_string(),
            action: "4".to_string(),
            scale_factor: "-1".to_string(),
        }];
        let processed = gen.process_rows(rows);
        // 0.1 * 10^-1 = 0.01
        assert_eq!(processed[0].coef_a, "0.010000");
    }
}
