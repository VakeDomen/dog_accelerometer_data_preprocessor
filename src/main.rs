use std::{collections::HashMap, error::Error};

use calamine::{open_workbook, Xlsx, Reader};
use chrono::{NaiveDate, NaiveTime, Duration, Datelike, Weekday};
use configparser::ini::Ini;
use rust_xlsxwriter::{Workbook, Format, ExcelDateTime, Color, FormatBorder};

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

        let mut vigorus = match extract_Y_N(&data[3]) {
            Some(d) => d,
            None => return None,
        };
        let mut moderate = match extract_Y_N(&data[4]) {
            Some(d) => d,
            None => return None,
        };
        let mut low = match extract_Y_N(&data[5]) {
            Some(d) => d,
            None => return None,
        };
        let mut sedentary = match extract_Y_N(&data[6]) {
            Some(d) => d,
            None => return None,
        };
        let mut con_vig = match extract_Y_N(&data[7]) {
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

    // format - time_format 
    let time_format_ref = match format_config.get("time") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"time\" attribute in the [format] section of config.ini");
            return;
        },
    };

    let time_format = match time_format_ref {
        Some(f) => f.clone(),
        None => {
            println!("Error: Can't find \"time\" attribute in the [format] section of config.ini");
            return;
        },
    };

    // format - weekend_color 
    let weekend_color_ref = match format_config.get("weekend_color") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"weekend_color\" attribute in the [format] section of config.ini");
            return;
        },
    };

    let weekend_color: u32 = match weekend_color_ref {
        Some(f) => match u32::from_str_radix(f, 16) {
            Ok(v) => v,
            Err(_) => {
                println!("Error: Can't parse \"weekend_color\" attribute in the [parsing] section of config.ini. Must be an hex RGB.");
                return;
            },
        }
        None => {
            println!("Error: Can't find \"weekend_color\" attribute in the [format] section of config.ini");
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
            Err(_) => {
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
            Err(_) => {
                println!("Error: Can't parse \"epoch_seconds\" attribute in the [parsing] section of config.ini. Must be an integer");
                return;
            },
        },
        None => {
            println!("Error: Can't find \"epoch_seconds\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    // parsing - cutpoint_low
    let cutpoint_low_ref = match parsing_config.get("cutpoint_low") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"cutpoint_low\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    let cutpoint_low: i32 = match cutpoint_low_ref {
        Some(f) => match f.clone().parse() {
            Ok(v) => v,
            Err(_) => {
                println!("Error: Can't parse \"cutpoint_low\" attribute in the [parsing] section of config.ini. Must be an integer");
                return;
            },
        },
        None => {
            println!("Error: Can't find \"cutpoint_low\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    // parsing - cutpoint_moderate
    let cutpoint_moderate_ref = match parsing_config.get("cutpoint_moderate") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"cutpoint_moderate\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    let cutpoint_moderate: i32 = match cutpoint_moderate_ref {
        Some(f) => match f.clone().parse() {
            Ok(v) => v,
            Err(_) => {
                println!("Error: Can't parse \"cutpoint_moderate\" attribute in the [parsing] section of config.ini. Must be an integer");
                return;
            },
        },
        None => {
            println!("Error: Can't find \"cutpoint_moderate\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    // parsing - cutpoint_vigorus
    let cutpoint_vigorus_ref = match parsing_config.get("cutpoint_vigorus") {
        Some(f) => f,
        None => {
            println!("Error: Can't find \"cutpoint_vigorus\" attribute in the [parsing] section of config.ini");
            return;
        },
    };

    let cutpoint_vigorus: i32 = match cutpoint_vigorus_ref {
        Some(f) => match f.clone().parse() {
            Ok(v) => v,
            Err(_) => {
                println!("Error: Can't parse \"cutpoint_vigorus\" attribute in the [parsing] section of config.ini. Must be an integer");
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

    match summarize(
        sensor_data,
        output_file, 
        &date_format, 
        &time_format,
        &decimals_format, 
        epoch_seconds,
        cutpoint_low,
        cutpoint_moderate,
        cutpoint_vigorus,
        weekend_color,
    ) {
        Ok(_) => println!("Done!"),
        Err(e) => println!("Error: {:#?}", e.to_string()),
    };

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

fn summarize(
    sensor_data: HashMap<NaiveDate, Vec<SensorEntry>>, 
    out_file: String, 
    date_format: &str, 
    time_format: &str, 
    decimal_format: &str, 
    epoch_time: i32,
    cutpoint_low: i32,
    cutpoint_moderate: i32,
    cutpoint_vigorus: i32,
    weekend_color: u32,
) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let sheet =  workbook.add_worksheet();

    let mut basic_format = Format::new();
    let bold_format = Format::new().set_bold().set_border(FormatBorder::Hair);
    let mut decimal_format = Format::new().set_num_format(decimal_format);
    let mut date_format = Format::new().set_num_format(date_format);
    let mut time_format = Format::new().set_num_format(time_format);
    let mut weekend_format = Format::new().set_background_color(Color::RGB(weekend_color));


    

    let columns = vec![
        "Day",
        "Date",
        "Weekday",
        "Total Vig.",
        "Total Mod.",
        "Total Low",
        "Total Sed.",
        // "Con. Vig.",
        // "Con. Mod.",
        "T. Non-zero",
        "T. Zero",
        "T. Empty",
        "Tot Counts",
        "Ave Counts/Min",
        "Ave Counts/Epoch",
    ];

    for i in 0..columns.len() {
        sheet.set_column_width(i as u16, 10)?;
        sheet.write_with_format(0, i as u16, columns[i], &bold_format)?;
    }

    for (index, day) in sorted_dates(&sensor_data).into_iter().enumerate() {
        let row = (index + 1) as u32;

        if day.weekday() == Weekday::Sat || day.weekday() == Weekday::Sun {
            basic_format = basic_format.set_background_color(weekend_color).set_border(FormatBorder::Hair);
            decimal_format = decimal_format.set_background_color(weekend_color).set_border(FormatBorder::Hair);
            date_format = date_format.set_background_color(weekend_color).set_border(FormatBorder::Hair);
            time_format = time_format.set_background_color(weekend_color).set_border(FormatBorder::Hair);
            weekend_format = weekend_format.set_background_color(weekend_color).set_border(FormatBorder::Hair);
        } else {
            basic_format = basic_format.set_background_color(Color::White).set_border(FormatBorder::Hair);
            decimal_format = decimal_format.set_background_color(Color::White).set_border(FormatBorder::Hair);
            date_format = date_format.set_background_color(Color::White).set_border(FormatBorder::Hair);
            time_format = time_format.set_background_color(Color::White).set_border(FormatBorder::Hair);
            weekend_format = weekend_format.set_background_color(Color::White).set_border(FormatBorder::Hair);
        }
        

        for col_name in columns.iter() {
            let position = columns.iter().position(|n| n == col_name).unwrap() as u16;

            match *col_name {
                "Day"               => sheet.write_with_format(row,position, row, &basic_format)?,
                "Date"              => sheet.write_with_format(row,position, &calc_date(&day)?, &date_format)?,
                "Weekday"           => sheet.write_with_format(row,position, calc_weekday(&day), &basic_format)?,
                "Total Vig."        => sheet.write_with_format(row,position, &calc_total_vig(sensor_data.get(&day), epoch_time, cutpoint_vigorus)?, &time_format)?,
                "Total Mod."        => sheet.write_with_format(row,position, &calc_total_mod(sensor_data.get(&day), epoch_time, cutpoint_moderate, cutpoint_vigorus)?, &time_format)?,
                "Total Low"         => sheet.write_with_format(row,position, &calc_total_low(sensor_data.get(&day), epoch_time, cutpoint_low, cutpoint_moderate)?, &time_format)?,
                "Total Sed."        => sheet.write_with_format(row,position, &calc_total_sed(sensor_data.get(&day), epoch_time, cutpoint_low)?, &time_format)?,
                // "Con. Vig."         => sheet.write(row,position, calc_con_vig(sensor_data.get(&day)))?,
                // "Con. Mod."         => sheet.write(row,position, calc_con_mod(sensor_data.get(&day)))?,
                "T. Non-zero"       => sheet.write_with_format(row,position, &calc_t_non_zero(sensor_data.get(&day), epoch_time)?, &time_format)?,
                "T. Zero"           => sheet.write_with_format(row,position, &calc_t_zero(sensor_data.get(&day), epoch_time)?, &time_format)?,
                "T. Empty"          => sheet.write_with_format(row,position, &calc_t_empty(sensor_data.get(&day), epoch_time)?, &time_format)?,
                "Tot Counts"        => sheet.write_with_format(row,position, calc_tot_counts(sensor_data.get(&day)), &basic_format)?,
                "Ave Counts/Min"    => sheet.write_with_format(row,position, calc_ave_counts_min(sensor_data.get(&day), epoch_time), &decimal_format)?,
                "Ave Counts/Epoch"  => sheet.write_with_format(row,position, calc_ave_counts_epoch(sensor_data.get(&day), epoch_time), &decimal_format)?,
                _                   => sheet.write_with_format(row,position, "Not handled!", &basic_format)?,
            };
        }
    }


    workbook.save(out_file)?;
    
    Ok(())
}

fn calc_ave_counts_min(
    day: Option<&Vec<SensorEntry>>,
    epoch_time: i32,
) -> f32 {
    let day = match day {
        Some(day) => day,
        None => return -99.,
    };
    day.iter().map(|s| s.value).sum::<i32>() as f32 / (day.len() as f32 / (60. / epoch_time as f32) )
}

fn calc_ave_counts_epoch(
    day: Option<&Vec<SensorEntry>>,
    epoch_time: i32,
) -> f32 {
    let avg = calc_ave_counts_min(day, epoch_time);
    avg / (60. / epoch_time as f32)
}

fn avg_count(day: &Vec<SensorEntry>) -> f32 {
    day
        .iter()
        .map(|s| s.value)
        .sum::<i32>() as f32 
            / (day.len() as f32)
}

fn calc_tot_counts(day: Option<&Vec<SensorEntry>>) -> String {
    let day = match day {
        Some(day) => day,
        None => return "No data".to_string(),
    };
    format!("{}", day.iter().map(|s| s.value).sum::<i32>())
}

fn calc_t_empty(
    day: Option<&Vec<SensorEntry>>,
    epoch_time: i32,
) -> Result<ExcelDateTime, Box<dyn Error>>  {
    let day = match day {
        Some(day) => day,
        None => return Err("No data".to_string().into()),
    };

    let mut count = 0;
    for entry in day.iter() {
        if entry.value == -1 {
            count += 1;
        }
    }
    count *= epoch_time;
    seconds_to_edt(count)
}

fn calc_t_zero(
    day: Option<&Vec<SensorEntry>>,
    epoch_time: i32,
) -> Result<ExcelDateTime, Box<dyn Error>> {
    let day = match day {
        Some(day) => day,
        None => return Err("No data".to_string().into()),
    };

    let mut count = 0;
    for entry in day.iter() {
        if entry.value == 0 {
            count += 1;
        }
    }
    count *= epoch_time;
    seconds_to_edt(count)
}

fn calc_t_non_zero(
    day: Option<&Vec<SensorEntry>>,
    epoch_time: i32,
) -> Result<ExcelDateTime, Box<dyn Error>> {
    let day = match day {
        Some(day) => day,
        None => return Err("No data".to_string().into()),
    };

    let mut count = 0;
    for entry in day.iter() {
        if entry.value > 0 {
            count += 1;
        }
    }
    count *= epoch_time;
    seconds_to_edt(count)
}

// fn calc_con_mod(day: Option<&Vec<SensorEntry>>) -> String {
//     "".to_string()
// }

// fn calc_con_vig(day: Option<&Vec<SensorEntry>>) -> String {
//     "".to_string()
// }

fn calc_total_sed(
    day: Option<&Vec<SensorEntry>>, 
    epoch_time: i32, 
    cutpoint_low: i32
) -> Result<ExcelDateTime, Box<dyn Error>> {
    let day = match day {
        Some(day) => day,
        None => return Err("No data".to_string().into()),
    };

    let mut count = 0;
    for entry in day.iter() {
        if entry.value < cutpoint_low {
            count += 1;
        }
    }
    count *= epoch_time;
    seconds_to_edt(count)
}

fn calc_total_low(
    day: Option<&Vec<SensorEntry>>, 
    epoch_time: i32,
    cutpoint_low: i32,
    cutpoint_moderate: i32,
) -> Result<ExcelDateTime, Box<dyn Error>> {
    let day = match day {
        Some(day) => day,
        None => return Err("No data".to_string().into()),
    };

    let mut count = 0;
    for entry in day.iter() {
        if entry.value >= cutpoint_low && entry.value < cutpoint_moderate{
            count += 1;
        }
    }
    count *= epoch_time;
    seconds_to_edt(count)
}

fn calc_total_mod(
    day: Option<&Vec<SensorEntry>>, 
    epoch_time: i32,
    cutpoint_moderate: i32,
    cutpoint_vigorus: i32,
) -> Result<ExcelDateTime, Box<dyn Error>> {
    let day = match day {
        Some(day) => day,
        None => return Err("No data".to_string().into()),
    };

    let mut count = 0;
    for entry in day.iter() {
        if entry.value >= cutpoint_moderate && entry.value < cutpoint_vigorus{
            count += 1;
        }
    }
    count *= epoch_time;
    seconds_to_edt(count)
}

fn calc_total_vig(day: Option<&Vec<SensorEntry>>, epoch_time: i32, cutpoint_vigorus: i32) -> Result<ExcelDateTime, Box<dyn Error>> {
    let day = match day {
        Some(day) => day,
        None => return Err("No data".to_string().into()),
    };

    let mut count = 0;
    for entry in day.iter() {
        if entry.value >= cutpoint_vigorus {
            count += 1;
        }
    }
    count *= epoch_time;
    seconds_to_edt(count)
}

fn calc_weekday(date: &NaiveDate) -> String {
    date.format("%a").to_string()
}

fn calc_date(day: &NaiveDate) -> Result<ExcelDateTime, rust_xlsxwriter::XlsxError> {
    ExcelDateTime::from_ymd(day.year() as u16, day.month() as u8, day.day() as u8)
}

fn seconds_to_edt(seconds: i32) -> Result<ExcelDateTime, Box<dyn Error>> {
    let hours: u16 = (seconds / 3600).try_into().unwrap();
    let remainder = seconds % 3600;
    let minutes: u8 = (remainder / 60).try_into().unwrap();
    let seconds: u8 = (remainder % 60).try_into().unwrap();
    let seconds_f = seconds as f64; // convert to f64 for ExcelDateTime
    match ExcelDateTime::from_hms(hours, minutes, seconds_f) {
        Ok(d) => Ok(d),
        Err(e) => Err(e.into()),
    }
}