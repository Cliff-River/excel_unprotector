use std::fmt;
use std::io::{Cursor, Read, Seek};

use quick_xml::events::Event;
use quick_xml::{Reader, Writer};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[derive(Debug)]
pub enum ExcelError {
    InvalidZip,
    InvalidXml,
    NoWorksheet,
    Io,
}

impl fmt::Display for ExcelError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ExcelError::InvalidZip => write!(f, "文件处理失败：文件格式无效或已损坏"),
            ExcelError::InvalidXml => write!(f, "文件处理失败：无法读取工作表内容"),
            ExcelError::NoWorksheet => write!(f, "文件处理失败：文件格式无效或已损坏"),
            ExcelError::Io => write!(f, "文件处理失败"),
        }
    }
}

impl std::error::Error for ExcelError {}

fn is_worksheet_entry(name: &str) -> bool {
    let Some(rest) = name.strip_prefix("xl/worksheets/sheet") else {
        return false;
    };
    let Some(digits) = rest.strip_suffix(".xml") else {
        return false;
    };
    !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit())
}

/// 剥离单个工作表 XML 中的 <sheetProtection> 元素，其余事件原样透传
fn strip_sheet_protection(xml: &[u8]) -> Result<Vec<u8>, ExcelError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut buf = Vec::new();
    let mut skip_depth: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Start(e)) => {
                if skip_depth > 0 {
                    if e.local_name().into_inner() == "sheetProtection" {
                        skip_depth += 1;
                    }
                } else if e.local_name().into_inner() == "sheetProtection" {
                    skip_depth = 1;
                } else {
                    writer
                        .write_event(Event::Start(e.into_owned()))
                        .map_err(|_| ExcelError::InvalidXml)?;
                }
            }
            Ok(Event::Empty(e)) => {
                if skip_depth == 0 && e.local_name().into_inner() != "sheetProtection" {
                    writer
                        .write_event(Event::Empty(e.into_owned()))
                        .map_err(|_| ExcelError::InvalidXml)?;
                }
            }
            Ok(Event::End(e)) => {
                if skip_depth > 0 {
                    if e.local_name().into_inner() == "sheetProtection" {
                        skip_depth -= 1;
                    }
                } else {
                    writer
                        .write_event(Event::End(e.into_owned()))
                        .map_err(|_| ExcelError::InvalidXml)?;
                }
            }
            Ok(other) => {
                if skip_depth == 0 {
                    writer
                        .write_event(other.into_owned())
                        .map_err(|_| ExcelError::InvalidXml)?;
                }
            }
            Err(_) => return Err(ExcelError::InvalidXml),
        }
        buf.clear();
    }

    if skip_depth > 0 {
        return Err(ExcelError::InvalidXml);
    }
    Ok(writer.into_inner())
}

/// 输入 xlsx 字节，输出解除工作表保护后的 xlsx 字节
pub fn unprotect(input: &[u8]) -> Result<Vec<u8>, ExcelError> {
    let cursor = Cursor::new(input);
    let mut archive = ZipArchive::new(cursor).map_err(|_| ExcelError::InvalidZip)?;

    let out_buf: Cursor<Vec<u8>> = Cursor::new(Vec::with_capacity(input.len()));
    let mut writer = ZipWriter::new(out_buf);
    let mut saw_worksheet = false;

    for i in 0..archive.len() {
        let name = {
            let entry = archive.by_index_raw(i).map_err(|_| ExcelError::InvalidZip)?;
            entry
                .name()
                .map_err(|_| ExcelError::InvalidZip)?
                .into_owned()
        };

        if is_worksheet_entry(&name) {
            saw_worksheet = true;
            let mut entry = archive.by_index(i).map_err(|_| ExcelError::InvalidZip)?;
            let mut xml = Vec::with_capacity(entry.size() as usize);
            entry
                .read_to_end(&mut xml)
                .map_err(|_| ExcelError::Io)?;
            let stripped = strip_sheet_protection(&xml)?;
            writer
                .start_file(
                    &name,
                    SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
                )
                .map_err(|_| ExcelError::Io)?;
            use std::io::Write as _;
            writer
                .write_all(&stripped)
                .map_err(|_| ExcelError::Io)?;
        } else {
            let entry = archive.by_index_raw(i).map_err(|_| ExcelError::InvalidZip)?;
            writer
                .raw_copy_file(entry)
                .map_err(|_| ExcelError::Io)?;
        }
    }

    if !saw_worksheet {
        return Err(ExcelError::NoWorksheet);
    }

    let mut finished = writer.finish().map_err(|_| ExcelError::Io)?;
    finished
        .rewind()
        .map_err(|_| ExcelError::Io)?;
    Ok(finished.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture() -> Vec<u8> {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../excel-unprotector-server/data/protected_file.xlsx"
        );
        std::fs::read(path).expect("read fixture")
    }

    fn entry_names(data: &[u8]) -> Vec<String> {
        let mut archive = ZipArchive::new(Cursor::new(data)).expect("open output zip");
        (0..archive.len())
            .map(|i| {
                archive
                    .by_index_raw(i)
                    .unwrap()
                    .name()
                    .unwrap()
                    .into_owned()
            })
            .collect()
    }

    fn read_entry(data: &[u8], name: &str) -> Vec<u8> {
        let mut archive = ZipArchive::new(Cursor::new(data)).expect("open zip");
        let mut entry = archive.by_name(name).expect("entry exists");
        let mut buf = Vec::new();
        entry.read_to_end(&mut buf).unwrap();
        buf
    }

    #[test]
    fn removes_protection_and_preserves_entries() {
        let input = fixture();
        let output = unprotect(&input).expect("unprotect");

        // 输出可打开、条目数量与名称集合一致
        assert_eq!(entry_names(&input), entry_names(&output));

        // 工作表不再包含 sheetProtection
        for name in entry_names(&input) {
            if is_worksheet_entry(&name) {
                let xml = read_entry(&output, &name);
                let text = String::from_utf8_lossy(&xml);
                assert!(!text.contains("sheetProtection"), "{name} still protected");
            }
        }
    }

    #[test]
    fn non_worksheet_entries_are_byte_identical() {
        let input = fixture();
        let output = unprotect(&input).expect("unprotect");
        let mut in_zip = ZipArchive::new(Cursor::new(&input)).unwrap();
        let mut out_zip = ZipArchive::new(Cursor::new(&output)).unwrap();
        assert_eq!(in_zip.len(), out_zip.len());
        for i in 0..in_zip.len() {
            let in_entry = in_zip.by_index_raw(i).unwrap();
            let out_entry = out_zip.by_index_raw(i).unwrap();
            let in_name = in_entry.name().unwrap().into_owned();
            let out_name = out_entry.name().unwrap().into_owned();
            assert_eq!(in_name, out_name);
            assert_eq!(in_entry.compression(), out_entry.compression());
            if !is_worksheet_entry(&in_name) {
                assert_eq!(
                    read_entry(&input, &in_name),
                    read_entry(&output, &out_name),
                    "entry {in_name} changed"
                );
            }
        }
    }

    #[test]
    fn worksheet_xml_events_preserved_except_protection() {
        let input = fixture();
        let output = unprotect(&input).expect("unprotect");

        fn local_events(xml: &[u8]) -> Vec<(u8, String)> {
            let mut reader = Reader::from_reader(xml);
            reader.config_mut().trim_text(false);
            let mut buf = Vec::new();
            let mut events = Vec::new();
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Eof) => break,
                    Ok(Event::Start(e)) => {
                        if e.local_name().into_inner() != "sheetProtection" {
                            events.push((b'S', e.local_name().into_inner().to_string()));
                        }
                    }
                    Ok(Event::Empty(e)) => {
                        if e.local_name().into_inner() != "sheetProtection" {
                            events.push((b'E', e.local_name().into_inner().to_string()));
                        }
                    }
                    Ok(Event::End(e)) => {
                        if e.local_name().into_inner() != "sheetProtection" {
                            events.push((b'/', e.local_name().into_inner().to_string()));
                        }
                    }
                    Ok(Event::Text(t)) => {
                        events.push((b'T', t.as_ref().to_string()));
                    }
                    Ok(_) => {}
                    Err(e) => panic!("xml parse error: {e}"),
                }
                buf.clear();
            }
            events
        }

        for name in entry_names(&input) {
            if is_worksheet_entry(&name) {
                assert_eq!(
                    local_events(&read_entry(&input, &name)),
                    local_events(&read_entry(&output, &name)),
                    "worksheet {name} content changed"
                );
            }
        }
    }

    #[test]
    fn idempotent_on_unprotected_file() {
        let input = fixture();
        let once = unprotect(&input).expect("first pass");
        let twice = unprotect(&once).expect("second pass");
        for name in entry_names(&twice) {
            if is_worksheet_entry(&name) {
                let bytes = read_entry(&twice, &name);
                let text = String::from_utf8_lossy(&bytes);
                assert!(!text.contains("sheetProtection"));
            }
        }
    }

    #[test]
    fn rejects_non_zip_bytes() {
        let garbage = vec![0u8; 256];
        assert!(matches!(unprotect(&garbage), Err(ExcelError::InvalidZip)));
    }

    #[test]
    fn rejects_zip_without_worksheets() {
        let buf: Cursor<Vec<u8>> = Cursor::new(Vec::new());
        let mut writer = ZipWriter::new(buf);
        writer
            .start_file("docProps/core.xml", SimpleFileOptions::default())
            .unwrap();
        use std::io::Write as _;
        writer.write_all(b"<core/>").unwrap();
        let zip_bytes = writer.finish().unwrap().into_inner();
        assert!(matches!(
            unprotect(&zip_bytes),
            Err(ExcelError::NoWorksheet)
        ));
    }

    #[test]
    fn rejects_invalid_worksheet_xml() {
        let buf: Cursor<Vec<u8>> = Cursor::new(Vec::new());
        let mut writer = ZipWriter::new(buf);
        writer
            .start_file("xl/worksheets/sheet1.xml", SimpleFileOptions::default())
            .unwrap();
        use std::io::Write as _;
        // 未闭合的 sheetProtection
        writer
            .write_all(b"<worksheet><sheetProtection><x></worksheet>")
            .unwrap();
        let zip_bytes = writer.finish().unwrap().into_inner();
        assert!(matches!(unprotect(&zip_bytes), Err(ExcelError::InvalidXml)));
    }
}
