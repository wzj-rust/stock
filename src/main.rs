use calamine::{Reader, open_workbook_auto};
use std::path::PathBuf;
use stock::config::Config;
use stock::log::set_tracing_subscriber;
use stock::parse::convert_stocks;
use stock::traverse_dir;
use tracing::{debug, info};

fn main() -> std::io::Result<()> {
    // 读取 stock-config.yml 的配置文件，读取不到直接退出程序
    let filename = Config::read_default_config().unwrap();
    let config = Config::load(filename).unwrap();

    // 设置日志级别
    set_tracing_subscriber(config.level.get_log_level());

    let dir_path = config.excel.get_path();

    traverse_dir(&dir_path).iter().for_each(|path| {
        for (key, val) in path.iter() {
            let excel_file = dir_path.to_string() + "/" + val;
            let sheet = key;

            debug!("Reading excel file: {excel_file}");

            let excel_path = PathBuf::from(excel_file);
            // Check if the file has a valid Excel extension
            match excel_path.extension().and_then(|s| s.to_str()) {
                Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
                _ => {
                    info!("请提供一个有效的excel文件");
                    continue;
                }
            }

            let mut workbook = open_workbook_auto(&excel_path).unwrap();

            // Get the specified worksheet range
            let range = workbook.worksheet_range(sheet).unwrap();

            // Convert the worksheet data to stock data
            let socks = convert_stocks(&range).unwrap();

            // Print each stock's data in a formatted way
            socks.iter().for_each(|s| debug!("{s:#?}"));
        }
    });
    Ok(())
}
