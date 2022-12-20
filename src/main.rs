use std::fs;
use std::env;
use std::path::Path;
use std::collections::HashMap;
use std::time::Instant;
use mailparse::MailHeader;
use mailparse::parse_mail;
use xlsxwriter::{Workbook, XlsxError};

fn is_sender(h: &&MailHeader) -> bool {
    if h.get_key() == "From" { 
        return true;
    }
    if h.get_key() == "Sender" {
        return true;
    }
    if h.get_key() == "Return-Path" {
        return true;
    }
    if h.get_key() == "Reply-To" {
        return true;
    }
    false
}


fn parse_eml_files(dir: &Path, email_counts: &mut HashMap<String, u32>) -> Result<(), Box<dyn std::error::Error>> {
    for entry in fs::read_dir(dir)? {
        let entry = entry?;
        let path = entry.path();

        if path.is_file() {
             // 只处理扩展名为eml的文件
            if path.extension().and_then(|s| s.to_str()) == Some("eml") {
                // 解析eml文件
                println!("正在解析:{:?}",path);
                let raw_data = fs::read(path)?;
                let email = parse_mail(&raw_data)?;

                // 查找From头信息，并获取发件人邮箱地址
                let from_header = email.headers.iter().find(|h| is_sender(h));
                if let Some(from_header) = from_header {
                    let email_addr = from_header.get_value();
                    *email_counts.entry(email_addr).or_insert(0) += 1;
                }
            }
        }else if path.is_dir(){
            parse_eml_files(&path, email_counts)?;
        }  
    }
    Ok(())
}

fn create_excel_file(email_counts: &HashMap<String, u32>) -> Result<(), XlsxError> {
    let file = Workbook::new("email_counts.xlsx")?;
    let mut sheet = file.add_worksheet(None)?;

    // 写入表头
    sheet.write_string(0, 0, "Email 地址", None)?;
    sheet.write_string(0, 1, "计数", None)?;
    

    // 写入数据
    let mut row = 1;
    for (email_addr, count) in email_counts {
        sheet.write_string(row, 0, email_addr, None)?;
        sheet.write_string(row, 1, &count.to_string(), None)?;
        row += 1;
    }

    file.close()
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();

    // 如果命令行参数不正确，则输出错误信息并退出
    if args.len() != 2 {
        println!("Usage: emlparser <dir>");
        std::process::exit(1);
    }

    let dir = &args[1];

    // 如果文件夹不存在或不是目录，则输出错误信息并退出
    if !Path::new(dir).is_dir() {
        println!("Error: {} is not a valid directory", dir);
        std::process::exit(1);
    }

    let start = Instant::now();
    let mut email_counts = HashMap::new();
    parse_eml_files(Path::new(dir), &mut email_counts)?;
    create_excel_file(&email_counts)?;
    let elapsed = start.elapsed();

    println!("解析完成耗时: {}.{:09} 秒", elapsed.as_secs(), elapsed.subsec_nanos());
    println!("");

    Ok(())
}