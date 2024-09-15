use calamine::{open_workbook_auto, DataType, Reader};
use regex::Regex;
use std::collections::HashSet;
use std::error::Error;
use std::path::Path;
use std::fs::File;
use printpdf::*;
use std::io::BufWriter;
use serde::Serialize;
use csv::Writer;

#[derive(Serialize)] // When declare a struct, use derive to declare a trait of that struct
/*
    struct created
    @param: 
    row: non-negative value
    postal-code: String 
*/
struct PostalCode {
    row: usize,
    postal_code: String,
}

fn main() -> Result<(), Box<dyn Error>> {
    let path = Path::new("src/doc/tfwp_2024q1_neg_en.xlsx"); // Default Read in English

    /*
    Due to large excel file, it cannot display all of value in the terminal,
    I created the PDF file to store all of the components.
    A4 paper = 210mm x 297mm  
    */
    let (doc, page1, layer1) = PdfDocument::new("Postal Codes", Mm(210.0), Mm(297.0), "Layer 1");

    // Make current_layer mutable
    let mut current_layer = doc.get_page(page1).get_layer(layer1);
    let font = doc.add_builtin_font(BuiltinFont::Helvetica)?;

    /* Canadian Postal Code:
        D: Digits represents as \d in Rust
        C: Characters represents as [A-Z] in Rust
        ?: Mark space is optional, Some of Postal Code have space and some are not.
        unwrap(): If run fails, terminate the program
        Ex. Format: CDC DCD
    */
    let postal_code_regex = Regex::new(r"[A-Z]\d[A-Z] ?\d[A-Z]\d").unwrap();

    // Store a collection of unique items like duplicate value
    let mut unique_postal_codes = HashSet::new();

    /*
    Iterate over the sheets
    to_owned: Owned copy of the list, without it, mutable borrow occurs, cannot call by references
    */
    let mut workbook = open_workbook_auto(path)?;

    /*
        Vector created
    */
    let mut postal_codes = Vec::new();
    let mut y_position = Mm(285.0); // Start from top of the page

    for sheet_name in workbook.sheet_names().to_owned() {
        if let Some(Ok(range)) = workbook.worksheet_range(&sheet_name) {
            println!("Reading sheet: {}\nOutput generated in PDF, CSV, JSON", sheet_name);

            /* Write the sheet name at the top
                @param: follow by sheetname, font size, x-intercept, y-intercept, font style
            */
            current_layer.use_text(format!("Reading sheet: {}", sheet_name), 12.0, Mm(10.0), y_position, &font);

            y_position -= Mm(10.0); // Move down for next line

            /*
            Enumerate: i = 0
            for (int i = 0; i < row.length; i++)
            */
            for (i, row) in range.rows().enumerate() {
                if let Some(cell) = row.get(3) { // Column D is at index 3, index starts from 0
                    if let DataType::String(val) = cell {
                        if let Some(postal_code) = postal_code_regex.find(val) {
                            // Store it in a HashSet if there are no duplicate value, if duplicate value occurs, it will skip writing in the PDF
                            if unique_postal_codes.insert(postal_code.as_str().to_string()) {
                                // Push the postal code inside the vector by using the struct PostalCode
                                postal_codes.push(PostalCode {
                                    row: i + 1,
                                    postal_code: postal_code.as_str().to_string(),
                                });

                                // Write the postal code and row number to the PDF
                                current_layer.use_text(
                                    format!("Row {}: {}", i + 1, postal_code.as_str()), 
                                    10.0, 
                                    Mm(10.0), 
                                    y_position, 
                                    &font
                                );
                                y_position -= Mm(10.0); 

                                // Start a new page if the current page is full
                                if y_position < Mm(10.0) {
                                    y_position = Mm(285.0); // Reset y when move to new page
                                    let (new_page, new_layer) = doc.add_page(Mm(210.0), Mm(297.0), "New Layer");
                                    current_layer = doc.get_page(new_page).get_layer(new_layer);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Save the PDF document
    let mut file = BufWriter::new(std::fs::File::create("postal_codes.pdf")?);
    doc.save(&mut file)?;

    // Save as CSV31
    let mut csv_writer = Writer::from_path("postal_codes.csv")?;
    for entry in &postal_codes {
        csv_writer.serialize(entry)?;
    }

    // Ensure all buffered data is written
    csv_writer.flush()?;

    // Save as JSON
    let json_file = File::create("postal_codes.json")?;
    serde_json::to_writer_pretty(json_file, &postal_codes)?;

    Ok(()) // Compile without errors
}
