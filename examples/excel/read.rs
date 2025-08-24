use calamine::{Reader, open_workbook_auto};
use std::env;
use std::path::PathBuf;
use stock::parse::convert_stocks;

/**
 * Main entry point of the program
 * This program reads an Excel file and processes its data
 */
fn main() {
    // Get the Excel file path from command line arguments
    let excel_file = env::args().nth(1).expect("请提供一个excel文件进行读取");
    // Get the worksheet name from command line arguments
    let sheet_name = env::args().nth(2).expect("请把工作表名称作为第二个参数");

    // Create a PathBuf from the file path
    let excel_path = PathBuf::from(excel_file);
    // Check if the file has a valid Excel extension
    match excel_path.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("请提供一个有效的excel文件"),
    }

    // Open the Excel workbook
    let mut workbook = open_workbook_auto(&excel_path).unwrap();
    // Get the specified worksheet range
    let range = workbook.worksheet_range(&sheet_name).unwrap();

    // Convert the worksheet data to stock data
    let socks = convert_stocks(&range).unwrap();

    // Print each stock's data in a formatted way
    socks.iter().for_each(|s| println!("{s:#?}"));
}
