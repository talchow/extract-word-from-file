// step1:get the file path from user,and check if the file exists
// step2:read the file to a hashmap, key is the word, value is the count
// step3:output the hashmap to a excel file
use std::collections::HashMap;
use std::fs;
use std::io::Write;
use std::path::Path;
use xlsxwriter::Workbook;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    loop {
        println!("Please input the file<path/name>: ");
        let _ = std::io::stdout().flush();
        let mut file_path = String::new();
        std::io::stdin().read_line(&mut file_path)?;
        file_path = file_path.trim().to_string();
        let path = Path::new(&file_path);
        if !path.exists() {
            println!("File not found!");
            return Ok(());
        }
        let word_count = count_words(&file_path)?;
        write_to_file(word_count)?;
        println!("File processed successfully!");
        return Ok(());
    }
}

fn count_words(file_path: &str) -> Result<HashMap<String, u32>, std::io::Error> {
    let data = fs::read(file_path)?;
    let mut word_count = HashMap::new();
    let lines = String::from_utf8(data).unwrap();
    // let line_convert = lines.split(" ")
    for line in lines.chars().filter(|c|c.is_ascii_alphabetic()) {
        *word_count.entry(line.to_string()).or_insert(0) += 1;
    }
    Ok(word_count)
}

fn write_to_file(data: HashMap<String, u32>) -> Result<(), std::io::Error> {
    let workbook = Workbook::new("output.xlsx").expect("failed to create workbook");
    let mut worksheet = workbook
        .add_worksheet(None)
        .expect("failed to create worksheet");
    let mut row = 0;
    let col = 0;
    for (key, value) in data {
        worksheet
            .write_string(row, col, &key, None)
            .expect("failed to write string");
        worksheet
            .write_number(row, col + 1, value.into(), None)
            .expect("failed to write number");
        row += 1;
    }
    if let Err(_) = workbook.close() {
        println!("Failed to close workbook");
    }
    Ok(())
}
