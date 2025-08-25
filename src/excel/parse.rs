use crate::stock_excel::StockExcel;
use calamine::{Data, Range};
use tracing::debug;

pub fn convert_stocks(range: &Range<Data>) -> Result<Vec<StockExcel>, anyhow::Error> {
    let mut stocks = Vec::new();

    for rows in range.rows() {
        // 移除表头内容
        if rows[0] == Data::String("发生日期".to_string()) {
            continue;
        }

        let mut stock = StockExcel::new();
        for (col_number, cell_data) in rows.iter().enumerate() {
            match col_number {
                0 => {
                    stock.happen_date = match *cell_data {
                        Data::DateTime(ref date) => {
                            date.as_datetime().unwrap().format("%Y-%m-%d").to_string()
                        }
                        Data::DateTimeIso(ref date) => date.to_string(),
                        Data::String(ref date) => date.to_string(),
                        _ => {
                            debug!("happen_date: {cell_data:?}");
                            "".to_string()
                        }
                    }
                }
                1 => {
                    stock.transaction_time = match *cell_data {
                        Data::DateTime(ref time) => {
                            time.as_datetime().unwrap().format("%H:%M:%S").to_string()
                        }
                        Data::DateTimeIso(ref time) => {
                            let time_split: Vec<&str> = time.split('.').collect();
                            time_split[0].to_string()
                        }
                        Data::String(ref time) => time.to_string(),
                        _ => {
                            debug!("transaction_time: {cell_data:?}");
                            "".to_string()
                        }
                    }
                }
                2 => {
                    stock.business_name = match *cell_data {
                        Data::String(ref name) => name.to_string(),
                        _ => {
                            debug!("business_name: {cell_data:?}");
                            "".to_string()
                        }
                    }
                }
                3 => {
                    stock.stock_code = match *cell_data {
                        Data::Int(ref code) => format!("{:06}", code),
                        Data::Float(ref code) => format!("{:06}", code),
                        Data::String(ref code) => code.to_string(),
                        Data::Empty => "".to_string(),
                        _ => {
                            debug!("stock_code: {cell_data:?}");
                            "".to_string()
                        }
                    }
                }
                4 => {
                    stock.stock_name = match *cell_data {
                        Data::String(ref name) => name.to_string(),
                        Data::Empty => "".to_string(),
                        _ => {
                            debug!("stock_name: {cell_data:?}");
                            "".to_string()
                        }
                    }
                }
                5 => {
                    stock.transaction_price = match *cell_data {
                        Data::Float(ref price) => *price as f32,
                        _ => {
                            debug!("transaction_price: {cell_data:?}");
                            0.0
                        }
                    }
                }
                6 => {
                    stock.transaction_count = match *cell_data {
                        Data::Int(ref count) => *count as u32,
                        Data::Float(ref count) => *count as u32,
                        _ => {
                            debug!("transaction_count: {cell_data:?}");
                            0
                        }
                    }
                }
                7 => {
                    stock.transaction_balance = match *cell_data {
                        Data::Float(ref balance) => *balance as f32,
                        _ => {
                            debug!("transaction_balance: {cell_data:?}");
                            0.0
                        }
                    }
                }
                8 => {
                    stock.stock_quantity = match *cell_data {
                        Data::Int(ref quantity) => *quantity as u32,
                        Data::Float(ref quantity) => *quantity as u32,
                        _ => {
                            debug!("stock_quantity: {cell_data:?}");
                            0
                        }
                    }
                }
                9 => {
                    stock.kickback = match *cell_data {
                        Data::Float(ref kickback) => *kickback as f32,
                        _ => {
                            debug!("kickback: {cell_data:?}");
                            0.0
                        }
                    }
                }
                10 => {
                    stock.exchange_fee = match *cell_data {
                        Data::Float(ref fee) => *fee as f32,
                        _ => {
                            debug!("exchange_fee: {cell_data:?}");
                            0.0
                        }
                    }
                }
                11 => {
                    stock.stamp_duty = match *cell_data {
                        Data::Float(ref duty) => *duty as f32,
                        _ => {
                            debug!("stamp_duty: {cell_data:?}");
                            0.0
                        }
                    }
                }
                12 => {
                    stock.transfer_fee = match *cell_data {
                        Data::Float(ref fee) => *fee as f32,
                        _ => {
                            debug!("transfer_fee: {cell_data:?}");
                            0.0
                        }
                    }
                }
                13 => {
                    stock.other_fee = match *cell_data {
                        Data::Float(ref fee) => *fee as f32,
                        _ => {
                            debug!("other_fee: {cell_data:?}");
                            0.0
                        }
                    }
                }
                14 => {
                    stock.transaction_amount = match *cell_data {
                        Data::Float(ref amount) => *amount as f32,
                        _ => {
                            debug!("transaction_amount: {cell_data:?}");
                            0.0
                        }
                    }
                }
                15 => {
                    stock.funds_balance = match *cell_data {
                        Data::Float(ref balance) => *balance as f32,
                        _ => {
                            debug!("funds_balance: {cell_data:?}");
                            0.0
                        }
                    }
                }
                16 => {
                    stock.authorization_number = match *cell_data {
                        Data::Int(ref number) => number.to_string(),
                        Data::Float(ref number) => number.to_string(),
                        Data::String(ref number) => number.to_string(),
                        Data::Empty => "".to_string(),
                        _ => {
                            debug!("authorization_number: {cell_data:?}");
                            "".to_string()
                        }
                    }
                }
                17 => {
                    stock.authorization_price = match *cell_data {
                        Data::Float(ref price) => *price as f32,
                        _ => {
                            debug!("authorization_price: {cell_data:?}");
                            0.0
                        }
                    }
                }
                18 => {
                    stock.authorization_count = match *cell_data {
                        Data::Int(ref count) => *count as u32,
                        Data::Float(ref count) => *count as u32,
                        _ => {
                            debug!("authorization_count: {cell_data:?}");
                            0
                        }
                    }
                }
                19 => {
                    stock.shareholder_code = match *cell_data {
                        Data::Float(ref code) => (*code as u64).to_string(),
                        Data::String(ref code) => code.to_string(),
                        Data::Empty => "".to_string(),
                        _ => {
                            debug!("shareholder_code: {cell_data:?}");
                            "".to_string()
                        }
                    }
                }
                20 => {
                    stock.fund_account = match *cell_data {
                        Data::Int(ref account) => account.to_string(),
                        Data::Float(ref account) => (*account as u32).to_string(),
                        Data::String(ref account) => account.to_string(),
                        _ => {
                            debug!("fund_account: {cell_data:?}");
                            "".to_string()
                        }
                    }
                }
                21 => {
                    stock.currency_type = match *cell_data {
                        Data::String(ref ctype) => ctype.to_string(),
                        _ => {
                            debug!("currency_type: {cell_data:?}");
                            "".to_string()
                        }
                    }
                }
                _ => {
                    debug!("unknown column: {cell_data:?}");
                }
            }
        }
        stocks.push(stock);
    }

    Ok(stocks)
}
