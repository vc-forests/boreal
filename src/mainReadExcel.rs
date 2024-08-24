use calamine::{open_workbook_auto, Reader, DataType};
use std::error::Error;
use std::fs::File;
use std::io::{Write, BufWriter};
use std::path::Path;

fn main() -> Result<(), Box<dyn Error>> {
    let path = Path::new("src/doc/tfwp_2024q1_neg_en.xlsx"); // Default Read in English

    /* 
    Due to large excel file, it cannot display all of value in the terminal, 
    I created the text file to store all of the components.
    */
    let output_file_path = "output.txt";
    let output_file = File::create(output_file_path)?;

    // mut: mutable
    let mut writer = BufWriter::new(output_file);

    let mut workbook = open_workbook_auto(path)?;

    /* 
    Iterate over the sheets
    to_owned: Owned copy of the list, without it, mutable borrow occurs, cannot call by references
    */
    for sheet_name in workbook.sheet_names().to_owned() {
        if let Some(Ok(range)) = workbook.worksheet_range(&sheet_name) {
            println!("Reading sheet: {}\nOutput generated in {}", sheet_name, output_file_path);
            // Instead of print! in terminal, I do write! for file 
            writeln!(writer, "Reading sheet: {}", sheet_name)?;

            /* 
            Enumerate: i = 0 
            for (int i = 0; i < row.length; i++)
            */
            for (i, row) in range.rows().enumerate() {
                write!(writer, "Row {}: ", i + 1)?; // For debugging
                for cell in row {
                    match cell {
                        DataType::Int(val) => write!(writer, "{}\t", val)?,
                        DataType::Float(val) => write!(writer, "{}\t", val)?,
                        DataType::String(val) => write!(writer, "{}\t", val)?,
                        DataType::Bool(val) => write!(writer, "{}\t", val)?,
                        DataType::Empty => write!(writer, "(empty)\t")?,
                        _ => write!(writer, "(other)\t")?,
                    }
                }
                writeln!(writer)?; // New line after each row
            }
        }
    }

    Ok(()) // Compile without errors
}
