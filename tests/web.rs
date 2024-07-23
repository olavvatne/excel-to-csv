//! Test suite for the Web and headless browsers.

#![cfg(target_arch = "wasm32")]

extern crate wasm_bindgen_test;
use excel_to_csv::xlsx_to_csv;
use std::fs::File;
use std::io::Read;
use std::path::PathBuf;
use wasm_bindgen_test::*;

wasm_bindgen_test_configure!(run_in_browser);

#[wasm_bindgen_test]
fn test_convert_to_csv() {
    let buffer = include_bytes!("../tests/data/test.xlsx");

    match xlsx_to_csv(buffer, 0) {
        Ok(csv) => {
            assert_eq!(
                csv,
                "X,Y,Strings\n2039,3949,row1\n392,293,row2\n3939,2930,row3\n"
            );
        }
        Err(_e) => panic!("Error occurred"),
    }
}
