use clap::Parser;
use csv::{ReaderBuilder, WriterBuilder, Terminator};
use regex::Regex;
use serde::Serialize;
use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::{self, Read};
use once_cell::sync::Lazy;

static RE_TYPE_INT: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^[UI](8|16|32|64)(_(W|B|WB))?$").unwrap());
static RE_TYPE_STR_CONV: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^STR(\d+)$").unwrap());
static RE_ADDR_STRING: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^(-?[0-9A-F]+|-?0x[0-9A-F]+|-?[0-9A-F]+h)_(\d+)$").unwrap());
static RE_ADDR_BITS: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^(-?[0-9A-F]+|-?0x[0-9A-F]+|-?[0-9A-F]+h)_(\d+)_(\d+)$").unwrap());
static RE_ADDR_INT: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^(-?[0-9A-F]+|-?0x[0-9A-F]+|-?[0-9A-F]+h)$").unwrap());
static RE_COUNT_16_8: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([UI](16|8)(_(W|B|WB))?|BITS)$").unwrap());
static RE_COUNT_32: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([UI]32(_(W|B|WB))?|F32|IP)$").unwrap());
static RE_COUNT_64: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)^([UI]64(_(W|B|WB))?|F64)$").unwrap());
static RE_TAG_CLEAN: Lazy<Regex> = Lazy::new(|| Regex::new(r"[^a-z0-9_]").unwrap());
static RE_HEX_EXPLICIT: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)(-?0x[0-9A-F]+|-?[0-9A-F]+h)").unwrap());
static RE_WORD_BOUNDARIES: Lazy<Regex> = Lazy::new(|| Regex::new(r"(?i)-?[0-9A-F]+").unwrap());

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Path to the simplified CSV input file
    input_file: Option<String>,

    /// Path to the output CSV file
    #[arg(short, long)]
    output: Option<String>,

    /// Protocol name
    #[arg(long, default_value = "modbusRTU")]
    protocol: String,

    /// Device category
    #[arg(long, default_value = "Inverter")]
    category: String,

    /// Manufacturer name
    #[arg(long)]
    manufacturer: Option<String>,

    /// Model name
    #[arg(long)]
    model: Option<String>,

    /// Forced write code
    #[arg(long, default_value = "")]
    forced_write: String,

    /// Value to subtract from all register addresses
    #[arg(long, default_value_t = 0)]
    address_offset: i32,

    /// Generate a template input CSV file
    #[arg(long)]
    template: bool,
}

#[derive(Debug)]
struct InputRow {
    name: String,
    tag: String,
    register_type: String,
    address: String,
    dtype: String,
    factor: String,
    offset: String,
    unit: String,
    action: String,
    scale_factor: String,
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
        let s = addr_part.trim().replace(',', "");
        if s.is_empty() {
            return "".to_string();
        }

        if let Some(caps) = RE_HEX_EXPLICIT.find(&s) {
            let hex_val = caps.as_str().to_lowercase();
            if hex_val.ends_with('h') {
                let is_neg = hex_val.starts_with('-');
                let clean_hex = if is_neg { &hex_val[1..hex_val.len() - 1] } else { &hex_val[..hex_val.len() - 1] };
                if let Ok(val) = i64::from_str_radix(clean_hex, 16) {
                    return if is_neg { (-val).to_string() } else { val.to_string() };
                }
            } else if hex_val.contains("0x") {
                let is_neg = hex_val.starts_with('-');
                let clean_hex = hex_val.replace("0x", "").replace('-', "");
                if let Ok(val) = i64::from_str_radix(&clean_hex, 16) {
                    return if is_neg { (-val).to_string() } else { val.to_string() };
                }
            }
        }

        if let Some(caps) = RE_WORD_BOUNDARIES.find(&s) {
            let word = caps.as_str();
            if word.chars().any(|c| c.is_ascii_alphabetic()) {
                if let Ok(val) = i64::from_str_radix(word.replace('-', "").as_str(), 16) {
                    return if word.starts_with('-') { (-val).to_string() } else { val.to_string() };
                }
            }
            return word.to_string();
        }

        s
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
                if let Ok(length) = parts[1].parse::<f64>() {
                    return (length / 2.0).ceil() as i32;
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

            let mut dtype = row.dtype.clone();
            if !self.validate_type(&dtype) {
                eprintln!("Line {}: Invalid Type '{}'. Skipping row.", line_num, dtype);
                continue;
            }

            let dtype_upper = dtype.to_uppercase();
            let mut address = row.address.clone();
            if let Some(caps) = RE_TYPE_STR_CONV.captures(&dtype_upper) {
                if let Ok(length) = caps[1].parse::<i32>() {
                    dtype = "STRING".to_string();
                    if !address.contains('_') {
                        address = format!("{}_{}", address, length);
                    }
                }
            }

            if !address.is_empty() {
                let parts: Vec<&str> = address.split('_').collect();
                let norm_parts: Vec<String> = parts.iter().map(|p| self.normalize_address_val(p)).collect();
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
                let base_tag = RE_TAG_CLEAN.replace_all(&row.name.to_lowercase().replace(' ', "_"), "").to_string();
                let base_tag = if base_tag.is_empty() { "var".to_string() } else { base_tag };
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
            if let Some(mapped) = self.register_type_map.get(&lower_reg_type) {
                info1 = mapped.clone();
            } else if vec!["1", "2", "3", "4"].contains(&row.register_type.as_str()) {
                info1 = row.register_type.clone();
            } else {
                 eprintln!("Line {}: Unknown RegisterType '{}'. Defaulting to Holding Register (3).", line_num, row.register_type);
            }

            let mut start_addr = 0;
            let parts: Vec<&str> = address.split('_').collect();
            if let Ok(raw_start_addr) = parts[0].parse::<i32>() {
                start_addr = raw_start_addr - self.address_offset;
                if start_addr < 0 {
                    eprintln!("Line {}: Address {} with offset {} results in negative address {}", line_num, raw_start_addr, self.address_offset, start_addr);
                }
                let mut new_parts = parts.clone();
                let start_addr_str = start_addr.to_string();
                new_parts[0] = &start_addr_str;
                address = new_parts.join("_");
            }

            let reg_count = self.get_register_count(&dtype, &address);
            let end_addr = start_addr + reg_count - 1;
            let dtype_up = dtype.to_uppercase();

            let entries = used_addresses_by_type.entry(info1.clone()).or_insert(Vec::new());
            let is_bits = dtype_up == "BITS";
            for (u_start, u_end, u_line, u_name, u_type) in entries.iter() {
                if std::cmp::max(start_addr, *u_start) <= std::cmp::min(end_addr, *u_end) {
                    let is_overlap_allowed = is_bits && (u_type == "BITS");
                    if !is_overlap_allowed {
                         eprintln!("Line {}: Address overlap detected for '{}' (Addr: {}-{}). Overlaps with '{}' (Line {}, Addr: {}-{}) in register type {}.", line_num, row.name, start_addr, end_addr, u_name, u_line, u_start, u_end, info1);
                    }
                }
            }
            entries.push((start_addr, end_addr, line_num, row.name.clone(), dtype_up.clone()));

            let val_factor = row.factor.parse::<f64>().unwrap_or(1.0);
            let val_scale = row.scale_factor.parse::<f64>().unwrap_or(0.0);
            let final_coef_a = val_factor * 10.0f64.powf(val_scale);
            let coef_a = format!("{:.6}", final_coef_a);

            let val_offset = row.offset.parse::<f64>().unwrap_or(0.0);
            let coef_b = format!("{:.6}", val_offset);

            let mut final_action = "1".to_string();
            let act_up = row.action.trim().to_uppercase();
            if act_up.is_empty() {
                final_action = "1".to_string();
            } else if act_up == "R" || act_up == "READ" || act_up == "4" {
                final_action = "4".to_string();
            } else if act_up == "RW" || act_up == "W" || act_up == "WRITE" || act_up == "1" {
                final_action = "1".to_string();
            } else if self.allowed_actions.contains(&act_up) {
                final_action = act_up;
            } else {
                eprintln!("Line {}: Invalid Action '{}'. Defaulting to '1'.", line_num, row.action);
            }

            processed_rows.push(ProcessedRow {
                info1,
                info2: address,
                info3: dtype_up,
                info4: "".to_string(),
                name: row.name,
                tag,
                coef_a,
                coef_b,
                unit: row.unit,
                action: final_action,
            });
        }
        processed_rows
    }
}

fn generate_template(output: &Option<String>) -> Result<(), Box<dyn Error>> {
    let headers = vec!["Name", "Tag", "RegisterType", "Address", "Type", "Factor", "Offset", "Unit", "Action", "ScaleFactor"];
    let rows = vec![
        vec!["Example Variable", "example_tag", "Holding Register", "30001", "U16", "1", "0", "V", "4", "0"],
        vec!["String Variable", "string_tag", "Holding Register", "30010_10", "String", "", "", "", "4", ""],
        vec!["Bit Variable", "bit_tag", "Holding Register", "30020_0_1", "Bits", "", "", "", "4", ""],
        vec!["Convenience String", "str_tag", "Holding Register", "30030", "STR20", "", "", "", "4", ""],
    ];

    let out_writer: Box<dyn io::Write> = match output {
        Some(path) => Box::new(File::create(path)?),
        None => Box::new(io::stdout()),
    };

    let mut writer = WriterBuilder::new().from_writer(out_writer);

    writer.write_record(&headers)?;
    for row in rows {
        writer.write_record(&row)?;
    }
    writer.flush()?;
    if let Some(path) = output {
        eprintln!("Template generated at {}", path);
    }
    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    if args.template {
        return generate_template(&args.output);
    }

    let input_path = args.input_file.ok_or("input_file is required")?;
    let manufacturer = args.manufacturer.ok_or("--manufacturer is required")?;
    let model = args.model.ok_or("--model is required")?;

    let mut delimiter = b',';
    {
        let mut f = File::open(&input_path)?;
        let mut buf = [0u8; 1024];
        let n = f.read(&mut buf)?;
        let s = String::from_utf8_lossy(&buf[..n]);
        if s.contains(';') && !s.contains(',') {
            delimiter = b';';
        } else if s.contains('\t') {
            delimiter = b'\t';
        }
    }

    let mut reader = ReaderBuilder::new()
        .delimiter(delimiter)
        .flexible(true)
        .from_path(&input_path)?;

    let headers = reader.headers()?.clone();
    let mut header_map = HashMap::new();
    for (i, h) in headers.iter().enumerate() {
        header_map.insert(h.to_lowercase().trim().to_string(), i);
    }

    let mut input_rows = Vec::new();
    for result in reader.records() {
        let record = result?;
        let get_val = |key: &str| -> String {
            if let Some(&i) = header_map.get(&key.to_lowercase()) {
                record.get(i).unwrap_or("").trim().to_string()
            } else {
                "".to_string()
            }
        };

        input_rows.push(InputRow {
            name: get_val("Name"),
            tag: get_val("Tag"),
            register_type: get_val("RegisterType"),
            address: get_val("Address"),
            dtype: get_val("Type"),
            factor: get_val("Factor"),
            offset: get_val("Offset"),
            unit: get_val("Unit"),
            action: get_val("Action"),
            scale_factor: get_val("ScaleFactor"),
        });
    }

    let generator = Generator::new(args.address_offset);
    let processed = generator.process_rows(input_rows);

    let mut writer_builder = WriterBuilder::new();
    writer_builder.delimiter(b';');

    writer_builder.terminator(Terminator::Any(b'\n'));

    let out_writer: Box<dyn io::Write> = match &args.output {
        Some(path) => Box::new(File::create(path)?),
        None => Box::new(io::stdout()),
    };

    let mut writer = writer_builder.from_writer(out_writer);

    // Header row
    let header_row = vec![
        args.protocol,
        args.category,
        manufacturer,
        model,
        args.forced_write,
        "".to_string(), "".to_string(), "".to_string(), "".to_string(), "".to_string(), "".to_string(),
    ];
    writer.write_record(&header_row)?;

    for (idx, row) in processed.iter().enumerate() {
        let data_row = vec![
            (idx + 1).to_string(),
            row.info1.clone(),
            row.info2.clone(),
            row.info3.clone(),
            row.info4.clone(),
            row.name.clone(),
            row.tag.clone(),
            row.coef_a.clone(),
            row.coef_b.clone(),
            row.unit.clone(),
            row.action.clone(),
        ];
        writer.write_record(&data_row)?;
    }
    writer.flush()?;

    if let Some(path) = &args.output {
        eprintln!("Definition file generated at {}", path);
    }

    Ok(())
}
