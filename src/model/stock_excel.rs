#[derive(Debug, Clone, PartialEq)]
pub struct StockExcel {
    pub happen_date: String,          // 交易日期
    pub transaction_time: String,     // 成交时间
    pub business_name: String,        // 业务名称
    pub stock_code: String,           // 证券代码
    pub stock_name: String,           // 证券名称
    pub transaction_price: f32,       // 成交价格
    pub transaction_count: u32,       // 成交数量
    pub transaction_balance: f32,     // 成交金额
    pub stock_quantity: u32,          // 证券数量
    pub kickback: f32,                // 佣金
    pub exchange_fee: f32,            // 交易所规费
    pub stamp_duty: f32,              // 印花税
    pub transfer_fee: f32,            // 过户费
    pub other_fee: f32,               // 其他费用
    pub transaction_amount: f32,      // 发生金额
    pub funds_balance: f32,           // 资金本次余额
    pub authorization_number: String, // 委托编号
    pub authorization_price: f32,     // 委托价格
    pub authorization_count: u32,     // 委托数量
    pub shareholder_code: String,     // 股东代码
    pub fund_account: String,         // 资金账号
    pub currency_type: String,        // 币种
}

impl StockExcel {
    pub fn new() -> Self {
        Self {
            happen_date: "".to_string(),
            transaction_time: "".to_string(),
            business_name: "".to_string(),
            stock_code: "".to_string(),
            stock_name: "".to_string(),
            transaction_price: 0.0,
            transaction_count: 0,
            transaction_balance: 0.0,
            stock_quantity: 0,
            kickback: 0.0,
            exchange_fee: 0.0,
            stamp_duty: 0.0,
            transfer_fee: 0.0,
            other_fee: 0.0,
            transaction_amount: 0.0,
            funds_balance: 0.0,
            authorization_number: "".to_string(),
            authorization_price: 0.0,
            authorization_count: 0,
            shareholder_code: "".to_string(),
            fund_account: "".to_string(),
            currency_type: "".to_string(),
        }
    }
}
