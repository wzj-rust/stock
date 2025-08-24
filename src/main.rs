use calamine::{Reader, open_workbook_auto};
use std::path::PathBuf;
use stock::parse::convert_stocks;
use stock::traverse_dir;

fn main() -> std::io::Result<()> {
    let dir_path = "/Users/wu_zhijun/works/RustroverProjects/stock/examples/excel";
    traverse_dir(dir_path).iter().for_each(|path| {
        for (key, val) in path.iter() {
            let excel_file = dir_path.to_string() + "/" + val;
            let sheet = key;

            println!("Reading excel file: {excel_file}");

            let excel_path = PathBuf::from(excel_file);
            // Check if the file has a valid Excel extension
            match excel_path.extension().and_then(|s| s.to_str()) {
                Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
                _ => {
                    println!("请提供一个有效的excel文件");
                    continue;
                }
            }

            let mut workbook = open_workbook_auto(&excel_path).unwrap();

            // Get the specified worksheet range
            let range = workbook.worksheet_range(&sheet).unwrap();

            // Convert the worksheet data to stock data
            let socks = convert_stocks(&range).unwrap();

            // Print each stock's data in a formatted way
            socks.iter().for_each(|s| println!("{s:#?}"));
        }
    });
    Ok(())
}
