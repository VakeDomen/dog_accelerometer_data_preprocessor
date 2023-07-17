use std::{path::{Path, self}, fs};

use calamine::{open_workbook, Xlsx, Reader};
use configparser::ini::Ini;


enum Mode {
    Reading,
    Waiting,
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
    // opens a new workbookž
    let mut i = 0;

    let mut state = Mode::Waiting;

    let mut workbook: Xlsx<_> = open_workbook(input_file).expect("Cannot open input *.xlsx file");
    if let Some(Ok(r)) = workbook.worksheet_range(&input_file_sheet) {
        for row in r.rows() {
            if is_empty(row) {
                state = Mode::Waiting;
            }

            if is_header_row(row) {
                println!("{:?}", row);
                state = Mode::Reading;
            }


            i+=1;
            if i > 70 {
                break;
            }
        }
    }
    

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
