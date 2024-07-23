mod utils;
use calamine::{open_workbook_auto_from_rs, Data, Reader, Sheets};
use csv::Writer;
use std::io::Cursor;
use utils::set_panic_hook;
use wasm_bindgen::prelude::*;

pub fn open_workbook_from_u8(bytes: &[u8]) -> Result<Sheets<Cursor<&[u8]>>, JsValue> {
    let cursor = Cursor::new(bytes);

    // Open the workbook using calamine
    let workbook = match open_workbook_auto_from_rs(cursor) {
        Ok(wb) => wb,
        Err(e) => {
            return Err(JsValue::from_str(&format!(
                "Failed to open workbook: {:?}",
                e
            )))
        }
    };
    Ok(workbook)
}

#[wasm_bindgen]
pub fn list_sheet_names(bytes: &[u8]) -> Result<Vec<String>, JsValue> {
    let workbook = open_workbook_from_u8(bytes)?;
    let sheet_names = workbook.sheet_names().to_vec();
    Ok(sheet_names)
}

#[wasm_bindgen]
pub fn xlsx_to_csv(bytes: &[u8], sheet_index: usize) -> Result<String, JsValue> {
    set_panic_hook();

    let mut workbook = open_workbook_from_u8(bytes)?;

    let sheet_names = workbook.sheet_names().to_vec();

    if sheet_index >= sheet_names.len() {
        return Err(JsValue::from_str("Invalid sheet index"));
    }
    let sheet_name = &sheet_names[sheet_index];

    let range = match workbook.worksheet_range(sheet_name) {
        Ok(r) => r,
        Err(e) => return Err(JsValue::from_str(&format!("Failed to read sheet: {:?}", e))),
    };

    let mut wtr = Writer::from_writer(vec![]);
    for row in range.rows() {
        let mut csv_row: Vec<String> = Vec::new();
        for cell in row.iter() {
            let cell_value = match cell {
                Data::String(s) => s.clone(),
                Data::Int(i) => i.to_string(),
                Data::Float(f) => f.to_string(),
                Data::Bool(b) => b.to_string(),
                Data::Error(e) => format!("Error: {:?}", e),
                Data::Empty => "".to_string(),
                Data::DateTime(d) => d.to_string(),
                Data::DateTimeIso(d) => d.to_string(),
                Data::DurationIso(d) => d.to_string(),
            };
            csv_row.push(cell_value);
        }
        if let Err(e) = wtr.write_record(csv_row) {
            return Err(JsValue::from_str(&format!(
                "Failed to write CSV record: {:?}",
                e
            )));
        }
    }
    let csv_data = match String::from_utf8(wtr.into_inner().unwrap()) {
        Ok(s) => s,
        Err(e) => {
            return Err(JsValue::from_str(&format!(
                "Failed to convert CSV to string: {:?}",
                e
            )))
        }
    };

    Ok(csv_data)
}

#[cfg(test)]
mod tests {
    use std::fs::File;
    use std::io::Read;
    use std::path::PathBuf;

    use crate::{list_sheet_names, xlsx_to_csv};

    #[test]
    fn test_excel_to_csv() {
        let mut d = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        d.push("tests/data/test.xlsx");
        println!("The path is: {:?}", d);
        let mut file = File::open(d).unwrap();
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer).unwrap();

        match xlsx_to_csv(&mut buffer, 0) {
            Ok(csv) => {
                assert_eq!(
                    csv,
                    "X,Y,Strings\n2039,3949,row1\n392,293,row2\n3939,2930,row3\n"
                );
            }
            Err(_e) => panic!("Error occurred"),
        }
    }

    #[test]
    fn test_list_sheet_names() {
        let mut d = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        d.push("tests/data/test.xlsx");
        println!("The path is: {:?}", d);
        let mut file = File::open(d).unwrap();
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer).unwrap();

        match list_sheet_names(&mut buffer) {
            Ok(names) => {
                assert_eq!(names.len(), 1);
                assert_eq!(names[0], "Sheet1");
            }
            Err(_e) => panic!("Error occurred"),
        }
    }
}
