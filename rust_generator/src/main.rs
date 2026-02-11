use clap::Parser;
use serde::Deserialize;
use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::{self, Read, Write};
use regex::Regex;
use csv::ReaderBuilder;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    input_file: Option<String>,

    #[arg(short, long)]
    output: Option<String>,

    #[arg(long, default_value = "modbusRTU")]
    protocol: String,

    #[arg(long, default_value = "Inverter")]
    category: String,

    #[arg(long)]
    manufacturer: Option<String>,

    #[arg(long)]
    model: Option<String>,

    #[arg(long, default_value = "")]
    forced_write: String,

    #[arg(long)]
    template: bool,
}

#[derive(Debug, Deserialize, Clone)]
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
    register_type_map: HashMap<String, String>,
    allowed_actions: Vec<String>,
    re_type_int: Regex,
    re_type_str_conv: Regex,
    re_addr_string: Regex,
    re_addr_bits: Regex,
    re_addr_int: Regex,
    re_count_16_8: Regex,
    re_count_32: Regex,
    re_count_64: Regex,
}

impl Generator {
    fn new() -> Self {
        let mut m = HashMap::new();
        m.insert("coil".to_string(), "1".to_string());
        m.insert("coils".to_string(), "1".to_string());
        m.insert("discrete input".to_string(), "2".to_string());
        m.insert("holding register".to_string(), "3".to_string());
        m.insert("holding".to_string(), "3".to_string());
        m.insert("input register".to_string(), "4".to_string());
        m.insert("input".to_string(), "4".to_string());

        Self {
            register_type_map: m,
            allowed_actions: vec!["0", "1", "2", "4", "6", "7", "8", "9"].into_iter().map(|s| s.to_string()).collect(),
            re_type_int: Regex::new(r"(?i)^[UI](8|16|32|64)(_(W|B|WB))?$").unwrap(),
            re_type_str_conv: Regex::new(r"(?i)^STR(\d+)$").unwrap(),
            re_addr_string: Regex::new(r"(?i)^(\d+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$").unwrap(),
            re_addr_bits: Regex::new(r"(?i)^(\d+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$").unwrap(),
            re_addr_int: Regex::new(r"(?i)^(\d+|0x[0-9A-F]+|[0-9A-F]+h)$").unwrap(),
            re_count_16_8: Regex::new(r"(?i)^([UI](16|8)(_(W|B|WB))?|BITS)$").unwrap(),
            re_count_32: Regex::new(r"(?i)^([UI]32(_(W|B|WB))?|F32|IP)$").unwrap(),
            re_count_64: Regex::new(r"(?i)^([UI]64(_(W|B|WB))?|F64)$").unwrap(),
        }
    }

    fn validate_type(&self, dtype: &str) -> bool {
        let dtype_upper = dtype.to_uppercase();
        let base_types = vec!["STRING", "BITS", "IP", "IPV6", "MAC", "F32", "F64"];
        if base_types.contains(&dtype_upper.as_str()) {
            return true;
        }
        if self.re_type_int.is_match(&dtype_upper) {
            return true;
        }
        if self.re_type_str_conv.is_match(&dtype_upper) {
            return true;
        }
        false
    }

    fn normalize_address_val(&self, addr_part: &str) -> String {
        let part = addr_part.trim();
        if part.is_empty() { return String::new(); }

        if part.to_lowercase().starts_with("0x") {
            if let Ok(val) = u64::from_str_radix(&part[2..], 16) {
                return val.to_string();
            }
        } else if part.to_lowercase().ends_with('h') {
             if let Ok(val) = u64::from_str_radix(&part[..part.len()-1], 16) {
                return val.to_string();
            }
        }
        part.to_string()
    }

    fn validate_address(&self, address: &str, dtype: &str) -> bool {
        let dtype_upper = dtype.to_uppercase();
        if dtype_upper == "STRING" {
            self.re_addr_string.is_match(address)
        } else if dtype_upper == "BITS" {
            self.re_addr_bits.is_match(address)
        } else {
            self.re_addr_int.is_match(address)
        }
    }

    fn get_register_count(&self, dtype: &str, address: &str) -> u32 {
        let dtype_upper = dtype.to_uppercase();
        if self.re_count_16_8.is_match(&dtype_upper) {
            return 1;
        } else if self.re_count_32.is_match(&dtype_upper) {
            return 2;
        } else if self.re_count_64.is_match(&dtype_upper) {
            return 4;
        } else if dtype_upper == "MAC" {
            return 3;
        } else if dtype_upper == "IPV6" {
            return 8;
        } else if dtype_upper == "STRING" {
            let parts: Vec<&str> = address.split('_').collect();
            if parts.len() >= 2 {
                if let Ok(len) = parts[1].parse::<f64>() {
                    return (len / 2.0).ceil() as u32;
                }
            }
        }
        1
    }

    fn process_rows(&self, rows: Vec<InputRow>) -> Vec<ProcessedRow> {
        let mut processed = Vec::new();
        let mut seen_names: HashMap<String, usize> = HashMap::new();
        let mut seen_tags: HashMap<String, usize> = HashMap::new();
        let mut used_addresses_by_type: HashMap<String, Vec<(u32, u32, usize, String, String)>> = HashMap::new();

        for (idx, row) in rows.into_iter().enumerate() {
            let line_num = idx + 2;
            let name = row.name.trim().to_string();
            let mut tag = row.tag.trim().to_string();
            let reg_type_str = row.register_type.trim();
            let mut address = row.address.trim().to_string();
            let mut dtype = row.dtype.trim().to_string();
            let factor = row.factor.trim();
            let offset = row.offset.trim();
            let unit = row.unit.trim().to_string();
            let mut action = row.action.trim().to_string();
            let scale_factor_str = row.scale_factor.trim();

            if name.is_empty() && address.is_empty() {
                eprintln!("Line {}: Skipping row with missing Name and Address.", line_num);
                continue;
            }

            if !self.validate_type(&dtype) {
                 eprintln!("Line {}: Invalid Type '{}'. Skipping row.", line_num, dtype);
                 continue;
            }

            // STR<n> conversion
            let mut caps_info = None;
            if let Some(caps) = self.re_type_str_conv.captures(&dtype) {
                caps_info = Some(caps[1].to_string());
            }

            if let Some(length) = caps_info {
                dtype = "STRING".to_string();
                if !address.contains('_') {
                    address = format!("{}_{}", address, length);
                }
            }

            if !self.validate_address(&address, &dtype) {
                eprintln!("Line {}: Invalid Address '{}' for Type '{}'. Skipping row.", line_num, address, dtype);
                 continue;
            }

            // Normalize address
            let parts: Vec<&str> = address.split('_').collect();
            let norm_parts: Vec<String> = parts.into_iter().map(|p| self.normalize_address_val(p)).collect();
            address = norm_parts.join("_");

            if !name.is_empty() {
                if let Some(prev) = seen_names.get(&name) {
                    eprintln!("Line {}: Duplicate Name '{}' detected. Previous occurrence at line {}.", line_num, name, prev);
                } else {
                    seen_names.insert(name.clone(), line_num);
                }
            }

            if tag.is_empty() && !name.is_empty() {
                let re_tag = Regex::new(r"[^a-z0-9_]").unwrap();
                let mut base_tag = re_tag.replace_all(&name.to_lowercase().replace(' ', "_"), "").to_string();
                if base_tag.is_empty() { base_tag = "var".to_string(); }
                tag = base_tag.clone();
                let mut counter = 1;
                while seen_tags.contains_key(&tag) {
                    tag = format!("{}_{}", base_tag, counter);
                    counter += 1;
                }
            }

            if !tag.is_empty() {
                if let Some(prev) = seen_tags.get(&tag) {
                    eprintln!("Line {}: Duplicate Tag '{}' detected. Previous occurrence at line {}.", line_num, tag, prev);
                } else {
                    seen_tags.insert(tag.clone(), line_num);
                }
            }

            let mut info1 = "3".to_string();
            if !reg_type_str.is_empty() {
                if let Some(val) = self.register_type_map.get(&reg_type_str.to_lowercase()) {
                    info1 = val.clone();
                } else if vec!["1", "2", "3", "4"].contains(&reg_type_str) {
                    info1 = reg_type_str.to_string();
                } else {
                    eprintln!("Line {}: Unknown RegisterType '{}'. Defaulting to Holding Register (3).", line_num, reg_type_str);
                }
            }

            // Overlap detection
            let addr_parts: Vec<&str> = address.split('_').collect();
            if let Ok(start_addr) = addr_parts[0].parse::<u32>() {
                let reg_count = self.get_register_count(&dtype, &address);
                let end_addr = start_addr + reg_count - 1;

                let current_range = (start_addr, end_addr, line_num, name.clone(), dtype.to_uppercase());

                let type_ranges = used_addresses_by_type.entry(info1.clone()).or_insert(Vec::new());
                let is_bits = dtype.to_uppercase() == "BITS";

                for (used_start, used_end, used_line, used_name, used_type) in type_ranges.iter() {
                    if std::cmp::max(start_addr, *used_start) <= std::cmp::min(end_addr, *used_end) {
                        let is_overlap_allowed = is_bits && (used_type == "BITS");
                        if !is_overlap_allowed {
                            eprintln!("Line {}: Address overlap detected for '{}' (Addr: {}-{}). Overlaps with '{}' (Line {}, Addr: {}-{}) in register type {}.",
                                line_num, name, start_addr, end_addr, used_name, used_line, used_start, used_end, info1);
                        }
                    }
                }
                type_ranges.push(current_range);
            }

            let val_factor = factor.parse::<f64>().unwrap_or(1.0);
            let val_scale = scale_factor_str.parse::<f64>().unwrap_or(0.0);
            let final_coef_a = val_factor * 10.0f64.powf(val_scale);
            let coef_a = format!("{:.6}", final_coef_a);

            let coef_b = format!("{:.6}", offset.parse::<f64>().unwrap_or(0.0));

            if action.is_empty() {
                action = "1".to_string();
            } else {
                let act_upper = action.to_uppercase();
                if act_upper == "R" || act_upper == "READ" {
                    action = "4".to_string();
                } else if act_upper == "RW" || act_upper == "W" || act_upper == "WRITE" {
                    action = "1".to_string();
                } else if !self.allowed_actions.contains(&action) {
                    eprintln!("Line {}: Invalid Action '{}'. Defaulting to '1'.", line_num, action);
                    action = "1".to_string();
                }
            }

            processed.push(ProcessedRow {
                info1,
                info2: address,
                info3: dtype.to_uppercase(),
                info4: "".to_string(),
                name,
                tag,
                coef_a,
                coef_b,
                unit,
                action,
            });
        }
        processed
    }
}

fn detect_delimiter(path: &str) -> u8 {
    let mut f = File::open(path).expect("Could not open file for delimiter detection");
    let mut buffer = [0; 1024];
    let n = f.read(&mut buffer).unwrap_or(0);
    let s = String::from_utf8_lossy(&buffer[..n]);

    let semicolon_count = s.matches(';').count();
    let comma_count = s.matches(',').count();

    if semicolon_count > comma_count {
        b';'
    } else {
        b','
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    if args.template {
        generate_template(args.output.as_deref())?;
        return Ok(());
    }

    let input_file = match args.input_file {
        Some(f) => f,
        None => {
            eprintln!("Error: input_file is required");
            std::process::exit(1);
        }
    };

    let manufacturer = match args.manufacturer {
        Some(m) => m,
        None => {
            eprintln!("Error: --manufacturer is required");
            std::process::exit(1);
        }
    };

    let model = match args.model {
        Some(m) => m,
        None => {
            eprintln!("Error: --model is required");
            std::process::exit(1);
        }
    };

    let delimiter = detect_delimiter(&input_file);

    let mut rdr = ReaderBuilder::new()
        .delimiter(delimiter)
        .trim(csv::Trim::All)
        .from_path(&input_file)?;

    let mut rows = Vec::new();
    for result in rdr.deserialize() {
        let row: InputRow = result?;
        rows.push(row);
    }

    let generator = Generator::new();
    let processed = generator.process_rows(rows);

    let mut writer: Box<dyn Write> = match args.output {
        Some(path) => Box::new(File::create(path)?),
        None => Box::new(io::stdout()),
    };

    // Header row
    let header = format!("{};{};{};{};{};;;;;;\n", args.protocol, args.category, manufacturer, model, args.forced_write);
    writer.write_all(header.as_bytes())?;

    for (idx, row) in processed.iter().enumerate() {
        let line = format!("{};{};{};{};{};{};{};{};{};{};{}\n",
            idx + 1,
            row.info1,
            row.info2,
            row.info3,
            row.info4,
            row.name,
            row.tag,
            row.coef_a,
            row.coef_b,
            row.unit,
            row.action
        );
        writer.write_all(line.as_bytes())?;
    }

    Ok(())
}

fn generate_template(output: Option<&str>) -> Result<(), Box<dyn Error>> {
    let mut writer: Box<dyn Write> = match output {
        Some(path) => Box::new(File::create(path)?),
        None => Box::new(io::stdout()),
    };

    writeln!(writer, "Name,Tag,RegisterType,Address,Type,Factor,Offset,Unit,Action,ScaleFactor")?;
    writeln!(writer, "Example Variable,example_tag,Holding Register,30001,U16,1,0,V,4,0")?;
    writeln!(writer, "String Variable,string_tag,Holding Register,30010_10,String,,, ,4,")?;
    writeln!(writer, "Bit Variable,bit_tag,Holding Register,30020_0_1,Bits,,, ,4,")?;
    writeln!(writer, "Convenience String,str_tag,Holding Register,30030,STR20,,, ,4,")?;

    Ok(())
}
