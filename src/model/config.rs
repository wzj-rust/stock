use std::path::Path;
use std::{error::Error, fs};

/// 配置信息
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct Config {
    pub excel: ExcelConfig, // excel文件路径
    pub level: LevelConfig, // 打印日志级别
}

impl Config {
    /// 读取配置文件
    pub fn load(filename: impl AsRef<Path>) -> Result<Self, Box<dyn Error>> {
        let config = fs::read_to_string(filename)?;
        let config = serde_yaml::from_str(&config)?;
        Ok(config)
    }

    /// 读取默认的配置
    pub fn read_default_config() -> Result<String, Box<dyn Error>> {
        let filename = std::env::var("STOCK_CONFIG").unwrap_or_else(|_| {
            let first_path = Path::new("./stock-config.yml");
            let path = shellexpand::tilde("~/.config/stock-config.yml");
            let second_path = Path::new(path.as_ref());
            let third_path = Path::new(
                "/Users/wu_zhijun/works/RustroverProjects/stock/fixtures/stock-config.yml",
            );
            match (
                first_path.exists(),
                second_path.exists(),
                third_path.exists(),
            ) {
                (true, _, _) => first_path.to_str().unwrap().to_string(),
                (_, true, _) => second_path.to_str().unwrap().to_string(),
                (_, _, true) => third_path.to_str().unwrap().to_string(),
                _ => panic!("no config file found"),
            }
        });
        Ok(filename)
    }
}

/// Excel读取路径地址
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct ExcelConfig {
    pub path: String,
}

impl ExcelConfig {
    pub fn get_path(&self) -> String {
        self.path.to_string()
    }
}

/// Level控制台打印日志配置信息
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct LevelConfig {
    pub log: String,
}

impl LevelConfig {
    pub fn get_log_level(&self) -> tracing::Level {
        match self.log.as_str() {
            "trace" => tracing::Level::TRACE,
            "debug" => tracing::Level::DEBUG,
            "warn" => tracing::Level::WARN,
            "error" => tracing::Level::ERROR,
            _ => tracing::Level::INFO,
        }
    }
}
