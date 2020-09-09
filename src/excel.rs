use crate::Thing;
use chrono::prelude::*;
use std::collections::HashMap;
use xlsxwriter::{DateTime, Format, Workbook, Worksheet};

const FILE_PATH: &str = "report.xlsx";
const FONT_SIZE: f64 = 12.0;

pub fn create_xlsx(values: Vec<Thing>) {
    let workbook = Workbook::new(FILE_PATH);
    let mut sheet = workbook.add_worksheet(None).expect("can add sheet");

    let mut width_map: HashMap<usize, usize> = HashMap::new();

    create_headers(&mut sheet, &mut width_map);

    let fmt = workbook
        .add_format()
        .set_text_wrap()
        .set_font_size(FONT_SIZE);

    let date_fmt = workbook
        .add_format()
        .set_num_format("dd/mm/yyyy hh:mm:ss AM/PM")
        .set_font_size(FONT_SIZE);

    for (i, v) in values.iter().enumerate() {
        create_row(i as u32, &v, &mut sheet, &date_fmt, &mut width_map);
    }

    width_map.iter().for_each(|(k, v)| {
        let _ = sheet.set_column(*k as u16, *k as u16, *v as f64 * 1.2, Some(&fmt));
    });

    workbook.close().expect("workbook can be closed");
}

fn create_row(
    row: u32,
    thing: &Thing,
    sheet: &mut Worksheet,
    date_fmt: &Format,
    mut width_map: &mut HashMap<usize, usize>,
) {
    let _ = sheet.write_string(row + 1, 0, &thing.id, None);
    set_new_max_width(0, thing.id.len(), &mut width_map);

    let start_date = DateTime::new(
        thing.start_date.year() as i16,
        thing.start_date.month() as i8,
        thing.start_date.day() as i8,
        thing.start_date.hour() as i8,
        thing.start_date.minute() as i8,
        thing.start_date.second() as f64,
    );
    let _ = sheet.write_datetime(row + 1, 1, &start_date, Some(date_fmt));
    set_new_max_width(1, 26, &mut width_map);

    let end_date = DateTime::new(
        thing.end_date.year() as i16,
        thing.end_date.month() as i8,
        thing.end_date.day() as i8,
        thing.end_date.hour() as i8,
        thing.end_date.minute() as i8,
        thing.end_date.second() as f64,
    );
    let _ = sheet.write_datetime(row + 1, 2, &end_date, None);
    set_new_max_width(2, 26, &mut width_map);

    let _ = sheet.write_string(row + 1, 8, &thing.project, None);
    set_new_max_width(3, thing.project.len(), &mut width_map);

    let _ = sheet.write_string(row + 1, 10, &thing.name, None);
    set_new_max_width(4, thing.name.len(), &mut width_map);

    let _ = sheet.write_string(row + 1, 11, &thing.text, None);
    set_new_max_width(5, thing.text.len(), &mut width_map);

    let _ = sheet.set_row(row, FONT_SIZE, None);
}

fn set_new_max_width(col: usize, new: usize, width_map: &mut HashMap<usize, usize>) {
    match width_map.get(&col) {
        Some(max) => {
            if new > *max {
                width_map.insert(col, new);
            }
        }
        None => {
            width_map.insert(col, new);
        }
    };
}

fn create_headers(sheet: &mut Worksheet, mut width_map: &mut HashMap<usize, usize>) {
    let _ = sheet.write_string(0, 0, "Id", None);
    let _ = sheet.write_string(0, 1, "StartDate", None);
    let _ = sheet.write_string(0, 2, "EndDate", None);
    let _ = sheet.write_string(0, 4, "Project", None);
    let _ = sheet.write_string(0, 6, "Name", None);
    let _ = sheet.write_string(0, 7, "Text", None);

    set_new_max_width(0, "Id".len(), &mut width_map);
    set_new_max_width(1, "StartDate".len(), &mut width_map);
    set_new_max_width(2, "EndDate".len(), &mut width_map);
    set_new_max_width(3, "Project".len(), &mut width_map);
    set_new_max_width(4, "Name".len(), &mut width_map);
    set_new_max_width(5, "Text".len(), &mut width_map);
}
