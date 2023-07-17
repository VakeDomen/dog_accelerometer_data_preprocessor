use std::collections::{HashMap, btree_map::Values};

use calamine::{open_workbook, Xlsx, Reader};
use chrono::{NaiveDate, NaiveTime, Duration};
use configparser::ini::Ini;
use rust_xlsxwriter::{XlsxError, Workbook, Format};

#[derive(Debug, PartialEq)]
enum Mode {
    Parsing,
    Waiting,
}

#[derive(Debug)]
struct SensorEntry {
    date: NaiveDate,
    time: NaiveTime,
    value: i32,
    vigorus: bool,
    moderate: bool,
    low: bool,
    sedentary: bool,
    con_vig: bool,
    con_mod: bool,
}

impl SensorEntry {
    fn from(data: &[calamine::DataType]) -> Option<Self> {
        let date = match extract_date(&data[0]) {
            Some(d) => d,
            None => return None,
        };
        let time = match extract_time(&data[1]) {
            Some(d) => d,
            None => return None,
        };

        let value = match extract_mag_value(&data[2]) {
            Some(d) => d,
            None => return None,
        };

        let vigorus = match extract_Y_N(&data[3]) {
            Some(d) => d,
            None => return None,
        };
        let moderate = match extract_Y_N(&data[4]) {
            Some(d) => d,
            None => return None,
        };
        let low = match extract_Y_N(&data[5]) {
            Some(d) => d,
            None => return None,
        };
        let sedentary = match extract_Y_N(&data[6]) {
            Some(d) => d,
            None => return None,
        };
        let con_vig = match extract_Y_N(&data[7]) {
            Some(d) => d,
            None => return None,
        };
        let con_mod = match extract_Y_N(&data[8]) {
            Some(d) => d,
            None => return None,
        };
        Some(Self {
            date,
            time,
            value,
            vigorus,
            moderate,
            low,
            sedentary,
            con_vig,
            con_mod,
        })
    }
}

fn main() {
    
    let config = match Ini::new().load("config.ini") {
        Ok(f) => f,
        Err(e) => {
            println!("Error: Can't find config.ini! Make sure it's in the same folder.\nSource: {:#?}", e.to_string());
            return;
        },
    };

    let general_config = match config.get("general") {
        Some(f) => f,
        None => {
            println!("Error: Can't find [general] section config.ini");
            return;
        },
    };

    let parsing_config = match config.get("parsing") {
        Some(f) => f,
        None => {
            println!("Error: Can't find [general] section config.ini");
            return;
        },
    };

    let format_config = match config.get("format") {
        Some(f) => f,
        None => {
            println!("Error: Can't find [general] section config.ini");
            return;
        },
    };

    // general - input_file
    let input_file_ref = match general_config.get("input_file") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"input_file\" attribute in the [general] section of config.ini");
            return;
        },
    };

    let input_file = match input_file_ref {
        Some(f) => f.clone(),
        None => {
            println!("Error: Can't find \"input_file\" attribute in the [general] section of config.ini");
            return;
        },
    };

    // general - input_file_sheet
    let input_file_sheet_ref = match general_config.get("input_file_sheet") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"input_file_sheet\" attribute in the [general] section of config.ini");
            return;
        },
    };

    let input_file_sheet = match input_file_sheet_ref {
        Some(f) => f.clone(),
        None => {
            println!("Error: Can't find \"input_file_sheet\" attribute in the [general] section of config.ini");
            return;
        },
    };

    // general - output_file
    let output_file_ref = match general_config.get("output_file") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"output_file\" attribute in the [general] section of config.ini");
            return;
        },
    };

    let output_file = match output_file_ref {
        Some(f) => f.clone(),
        None => {
            println!("Error: Can't find \"output_file\" attribute in the [general] section of config.ini");
            return;
        },
    };

    // format - output_number_decimals 
    let decimals_format_ref = match format_config.get("decimals") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"decimals\" attribute in the [format] section of config.ini");
            return;
        },
    };

    let decimals_format = match decimals_format_ref {
        Some(f) => f.clone(),
        None => {
            println!("Error: Can't find \"decimals\" attribute in the [format] section of config.ini");
            return;
        },
    };
    
    // format - date_format 
    let date_format_ref = match format_config.get("date") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"date\" attribute in the [format] section of config.ini");
            return;
        },
    };

    let date_format = match date_format_ref {
        Some(f) => f.clone(),
        None => {
            println!("Error: Can't find \"date\" attribute in the [format] section of config.ini");
            return;
        },
    };

    // parsing - skip_days_num
    let skip_days_num_ref = match parsing_config.get("skip_days_num") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"skip_days_num\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    let skip_days_num: i32 = match skip_days_num_ref {
        Some(f) => match f.clone().parse() {
            Ok(v) => v,
            Err(e) => {
                println!("Error: Can't parse \"skip_days_num\" attribute in the [parsing] section of config.ini. Must be an integer");
                return;
            },
        },
        None => {
            println!("Error: Can't find \"skip_days_num\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    // parsing - day_window_size
    let day_window_size_ref = match parsing_config.get("day_window_size") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"day_window_size\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    let day_window_size: i32 = match day_window_size_ref {
        Some(f) => match f.clone().parse() {
            Ok(v) => v,
            Err(e) => {
                println!("Error: Can't parse \"day_window_size\" attribute in the [parsing] section of config.ini. Must be an integer");
                return;
            },
        },
        None => {
            println!("Error: Can't find \"day_window_size\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    // parsing - epoch_seconds
    let epoch_seconds_ref = match parsing_config.get("epoch_seconds") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"epoch_seconds\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    let epoch_seconds: i32 = match epoch_seconds_ref {
        Some(f) => match f.clone().parse() {
            Ok(v) => v,
            Err(e) => {
                println!("Error: Can't parse \"epoch_seconds\" attribute in the [parsing] section of config.ini. Must be an integer");
                return;
            },
        },
        None => {
            println!("Error: Can't find \"epoch_seconds\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    
    let mut workbook: Xlsx<_> = open_workbook(input_file).expect("Cannot open input *.xlsx file");
    
    let mut state = Mode::Waiting;
    let mut first_seen_date = None;
    let mut first_parsed_date = None;
    let mut sensor_data: HashMap<NaiveDate, Vec<SensorEntry>> = HashMap::new();

    if let Some(Ok(r)) = workbook.worksheet_range(&input_file_sheet) {
        for row in r.rows() {

            if is_empty(&row) {
                state = Mode::Waiting;
                continue;
            }

            
            if state == Mode::Waiting && is_header_row(&row) {
                state = Mode::Parsing;
                continue;
            }
            
            if state == Mode::Waiting {
                continue;
            }

            let sensor_entry = if state == Mode::Parsing {
                match SensorEntry::from(row) {
                    Some(v) => v,
                    None => {
                        state = Mode::Waiting;
                        continue;
                    },
                }
            } else {
                continue;
            };

            if first_seen_date.is_none() {
                first_seen_date = Some(sensor_entry.date.clone())
            }


            let first_seen_date_value = first_seen_date.unwrap();

            // check if still skipping first X days
            if sensor_entry.date.signed_duration_since(first_seen_date_value).num_days() < skip_days_num.into() {
                continue;
            }

            if first_parsed_date.is_none() {
                first_parsed_date = Some(sensor_entry.date.clone())
            }
            let first_parsed_date_value = first_parsed_date.unwrap();


            // check if all required dates parsed
            if sensor_entry.date.signed_duration_since(first_parsed_date_value).num_days() >= day_window_size.into() {
                break;
            }

            
            let entry = sensor_data.entry(sensor_entry.date).or_insert(vec![]);
            entry.push(sensor_entry);

            
        }
    }
    
    println!("{:#?}", sensor_data.keys());

    summarize(sensor_data, output_file, &date_format, &decimals_format);

}



fn is_header_row(row: &[calamine::DataType]) -> bool {
    match &row[0] {
        calamine::DataType::String(s) => if !s.eq("Date") { return false },
        _ => return false,
    };
    match &row[1] {
        calamine::DataType::String(s) => if !s.eq("Time") { return false },
        _ => return false,
    };
    match &row[2] {
        calamine::DataType::String(s) => if !s.eq("Mag. Value") { return false },
        _ => return false,
    };
    true
}

fn is_empty(row: &[calamine::DataType]) -> bool {
    if row.is_empty() {
        return true;
    }
    for item in row.iter() {
        match item {
            calamine::DataType::Empty => continue,
            _ => return false,
        }
    }
    true
}


fn extract_date(cell: &calamine::DataType) -> Option<NaiveDate> {
    
    match cell {
        calamine::DataType::DateTime(float) => {
            // Excel's datetime is a float where the integer part is the number of days since 1900-01-01
            // and the decimal part represents the time of the day.
            let days = float.trunc() as i64;
            let naive_date = NaiveDate::from_ymd(1900, 1, 1) + Duration::days(days - 2);
            Some(naive_date)
        },
        _ => None,
    }
}

fn extract_time(cell: &calamine::DataType) -> Option<NaiveTime> {
    
    match cell {
        calamine::DataType::DateTime(float) => {
            // Excel's datetime is a float where the integer part is the number of days since 1900-01-01
            // and the decimal part represents the time of the day.
            let days_proportion = float.fract();
            let days = float.trunc() as i64;

            let naive_time = NaiveTime::from_num_seconds_from_midnight(
                (days_proportion * 24.0 * 60.0 * 60.0) as u32, 0);
            Some(naive_time)
        },
        _ => None,
    }
}

fn extract_mag_value(cell: &calamine::DataType) -> Option<i32> {
    match cell {
        calamine::DataType::Float(float) => Some(*float as i32),
        calamine::DataType::Int(val) => Some(*val as i32),
        _ => None,
    }
}


fn extract_Y_N(cell: &calamine::DataType) -> Option<bool> {
    match cell {
        calamine::DataType::String(s) => {
            if s.eq("Y") || s.eq("N") {
                Some(s.eq("Y"))
            } else {
                None
            }
        },
        _ => None,
    }
}

fn sorted_dates<T>(map: &HashMap<NaiveDate, T>) -> Vec<NaiveDate> {
    let mut dates: Vec<_> = map.keys().cloned().collect();
    dates.sort();
    dates
}

fn summarize(sensor_data: HashMap<NaiveDate, Vec<SensorEntry>>, out_file: String, date_format: &str, decimal_format: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let mut sheet =  workbook.add_worksheet();

    let bold_format = Format::new().set_bold();
    let decimal_format = Format::new().set_num_format(decimal_format);
    let date_format = Format::new().set_num_format(date_format);


    let columns = vec![
        "Day",
        "Date",
        "Weekday",
        "Total Vig.",
        "Total Mod.",
        "Total Low",
        "Total Sed.",
        "Con. Vig.",
        "Con. Mod.",
        "T. Non-zero",
        "T. Zero",
        "T. Empty",
        "Tot Counts",
        "Ave Counts/Min",
        "Ave Counts/Epoch",

    ]


    for (index, day) in sorted_dates(&sensor_data).into_iter().enumerate() {

    }

    
    Ok(())
}