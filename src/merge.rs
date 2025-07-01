use crate::docx::writer_file;
use quick_xml::events::Event;
use quick_xml::{Reader, Writer};
use std::collections::HashMap;
use std::fs;
use std::fs::File;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use zip::write::{FileOptions, SimpleFileOptions};
use zip::{ZipArchive, ZipWriter};

// 分页符
const PAGE_BREAK: &str = r#"<w:p>
    <w:r>
        <w:br w:type="page"/>
    </w:r>
</w:p>"#;

// 换行符
const EMPTY_LINE: &str = r#"<w:p>
    <w:r>
        <w:t></w:t>
    </w:r>
</w:p>"#;

pub fn merge_docx(
    input_paths: &[impl AsRef<Path>],
    output_path: impl AsRef<Path>,
    concat: DocxConcat,
) -> Result<(), Box<dyn std::error::Error>> {
    merge_docx_with_page_breaks(input_paths, output_path, concat)
}

fn merge_docx_with_page_breaks(
    input_paths: &[impl AsRef<Path>],
    output_path: impl AsRef<Path>,
    concat: DocxConcat,
) -> Result<(), Box<dyn std::error::Error>> {
    // 1. 准备临时工作目录
    let temp_dir = PathBuf::from("_temp_merge");
    fs::create_dir_all(&temp_dir)?;

    // 2. 处理第一个文件作为基础模板
    let base_template = &input_paths[0];
    let (mut zip_writer, file_option) =
        initialize_output_from_template(base_template, &output_path)?;

    // 3. 合并所有文档内容
    let mut combined_content = String::new();
    let mut media_files = Vec::new();
    let mut rels_map = HashMap::new();

    for (i, path) in input_paths.iter().enumerate() {
        // 提取文档内容
        let (content, media) = extract_document_content(path)?;
        media_files.extend(media);

        // 添加文档内容（跳过第一个文件的重复处理）
        if i > 0 {
            match concat {
                DocxConcat::PAGE => {
                    combined_content.push_str(PAGE_BREAK); // 添加分页符
                }
                DocxConcat::CRLF(num) => {
                    for _ in 0..num {
                        combined_content.push_str(EMPTY_LINE); // 添加空行
                    }
                }
            }
        }
        combined_content.push_str(&extract_body_content(&content)?);

        // 处理资源关系
        process_relationships(path, &mut rels_map)?;
    }

    // 4. 构建最终document.xml
    let merged_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
        <w:document 
            xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData"
            mc:Ignorable="w14 w15 wp14">
            <w:body>{}</w:body>
        </w:document>"#,
        combined_content
    );

    // 5. 写入合并后的内容
    update_zip_content(
        &mut zip_writer,
        file_option,
        &merged_xml,
        &media_files,
        &rels_map,
    )?;

    // 清理临时文件
    fs::remove_dir_all(temp_dir)?;
    Ok(())
}

// 从模板初始化输出文件
fn initialize_output_from_template(
    template_path: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
) -> Result<(ZipWriter<File>, FileOptions<'static, ()>), Box<dyn std::error::Error>> {
    let template_file = File::open(template_path)?;
    let mut template_archive = ZipArchive::new(template_file)?;
    let output_file = File::create(output_path)?;
    let mut zip_writer = ZipWriter::new(output_file);
    let mut file_option: Option<FileOptions<()>> = None;

    // 复制模板文件结构（跳过后面要替换的文件）
    for i in 0..template_archive.len() {
        let mut file = template_archive.by_index(i)?;
        let name = file.name().to_string();

        if name != "word/document.xml"
            && !name.starts_with("word/media/")
            && name != "word/_rels/document.xml.rels"
        {
            // 文件内容
            let mut contents = Vec::new();
            // 读取文件内容到数组中
            file.read_to_end(&mut contents)?;
            writer_file(&mut zip_writer, &file, &contents)?;
        } else if name == "word/document.xml" {
            // 写入新文件
            let option = SimpleFileOptions::default()
                .compression_method(file.compression())
                .unix_permissions(file.unix_mode().unwrap_or(0o644));
            file_option = Some(option);
        }
    }

    Ok((zip_writer, file_option.unwrap()))
}

// 提取文档内容和媒体文件
fn extract_document_content(
    docx_path: impl AsRef<Path>,
) -> Result<(String, Vec<PathBuf>), Box<dyn std::error::Error>> {
    let file = File::open(docx_path)?;
    let mut archive = ZipArchive::new(file)?;
    let mut document_xml = String::new();
    let mut media_files = Vec::new();

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.name() == "word/document.xml" {
            file.read_to_string(&mut document_xml)?;
        } else if file.name().starts_with("word/media/") {
            let out_path = PathBuf::from("_temp_media").join(file.name());
            fs::create_dir_all(out_path.parent().unwrap())?;
            let mut out_file = File::create(&out_path)?;
            std::io::copy(&mut file, &mut out_file)?;
            media_files.push(out_path);
        }
    }

    Ok((document_xml, media_files))
}

// 提取body内容
fn extract_body_content(xml: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::new());
    let mut in_body = false;
    let mut depth = 0;
    let mut result = String::new();

    loop {
        match reader.read_event()? {
            Event::Start(e) if e.name().as_ref() == b"w:body" => {
                in_body = true;
                depth = 1;
            }
            Event::End(e) if in_body && e.name().as_ref() == b"w:body" => {
                in_body = false;
                break;
            }
            Event::Eof => break,
            e if in_body => {
                if let Event::Start(_) = e {
                    depth += 1;
                } else if let Event::End(_) = e {
                    depth -= 1;
                }

                if depth > 0 {
                    writer.write_event(e)?;
                }
            }
            _ => {}
        }
    }

    result = String::from_utf8(writer.into_inner())?;
    Ok(result)
}

// 处理资源关系
fn process_relationships(
    docx_path: impl AsRef<Path>,
    rels_map: &mut HashMap<String, Relationship>,
) -> Result<(), Box<dyn std::error::Error>> {
    let file = File::open(docx_path)?;
    let mut archive = ZipArchive::new(file)?;

    if let Ok(mut rels_file) = archive.by_name("word/_rels/document.xml.rels") {
        let mut content = String::new();
        rels_file.read_to_string(&mut content)?;

        let mut reader = Reader::from_str(&content);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                    let mut id = String::new();
                    let mut r#type = String::new();
                    let mut target = String::new();

                    for attr in e.attributes() {
                        let attr = attr?;
                        match attr.key.as_ref() {
                            b"Id" => id = attr.unescape_value()?.to_string(),
                            b"Type" => r#type = attr.unescape_value()?.to_string(),
                            b"Target" => target = attr.unescape_value()?.to_string(),
                            _ => {}
                        }
                    }

                    if !id.is_empty() {
                        rels_map.insert(id.to_string(), Relationship { id, r#type, target });
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
    }

    Ok(())
}

// 更新ZIP内容
fn update_zip_content(
    zip_writer: &mut ZipWriter<File>,
    file_options: FileOptions<'static, ()>,
    merged_xml: &str,
    media_files: &[PathBuf],
    rels_map: &HashMap<String, Relationship>,
) -> Result<(), Box<dyn std::error::Error>> {
    // 写入合并后的document.xml
    zip_writer.start_file("word/document.xml", file_options)?;
    zip_writer.write_all(merged_xml.as_bytes())?;

    // 写入媒体文件
    for media_path in media_files {
        let rel_path = media_path.strip_prefix("_temp_media")?;
        zip_writer.start_file(rel_path.to_str().unwrap(), SimpleFileOptions::default())?;
        let mut media_file = File::open(media_path)?;
        std::io::copy(&mut media_file, zip_writer)?;
    }

    // 写入合并后的relationships文件
    zip_writer.start_file("word/_rels/document.xml.rels", SimpleFileOptions::default())?;
    let mut rels_content = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#
        .to_string();

    for (key, relationship) in rels_map {
        rels_content.push_str(&format!(
            r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#,
            key, relationship.r#type, relationship.target
        ));
    }

    rels_content.push_str("</Relationships>");
    zip_writer.write_all(rels_content.as_bytes())?;
    Ok(())
}

struct Relationship {
    id: String,
    r#type: String,
    target: String,
}

pub enum DocxConcat {
    PAGE,
    CRLF(u32),
}
