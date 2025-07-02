use crate::docx::template::create_drawing_element;
use crate::docx::word::*;
use crate::error::DocxError;
use crate::image::{DOCX_EMU, DocxImage};
use crate::request::request_image_data;
use log::debug;
use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use reqwest::Client;
use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::time::Duration;
use zip::read::ZipFile;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

mod template;
mod word;

static PREFIX_TAG: &str = "{{";
static SUFFIX_TAG: &str = "}}";

pub struct DocxTemplate {
    // 待替换的字符串
    text_replacements: HashMap<String, String>,
    // 待替换的图片
    image_replacements: HashMap<String, Option<DocxImage>>,
    // 已经添加的图片路径
    images_map: HashMap<String, String>,
    // 请求对象
    client: Client,
}

impl DocxTemplate {
    pub fn new() -> Self {
        DocxTemplate {
            text_replacements: HashMap::new(),
            image_replacements: HashMap::new(),
            images_map: HashMap::new(),
            client: Client::builder()
                .timeout(Duration::from_secs(100)) // 设置超时
                .build()
                .unwrap(),
        }
    }

    /// 添加待替换的字符以及对应的值
    /// @param placeholder 待替换的字符串
    /// @param value 替换的值
    pub fn add_text_replacement(&mut self, placeholder: &str, value: &str) {
        self.text_replacements
            .insert(placeholder.to_string(), value.to_string());
    }

    /// 添加待替换的图片
    /// @param placeholder 待替换的字符串
    /// @param image_path 图片路径
    pub fn add_image_file_replacement(
        &mut self,
        placeholder: &str,
        image_path: Option<&str>,
    ) -> Result<(), DocxError> {
        match image_path {
            None => {
                // 插入图片到属性中
                self.image_replacements
                    .insert(placeholder.to_string(), None);
            }
            Some(file_path) => {
                // 判断是否添加过该图片
                if self.images_map.contains_key(file_path) {
                    let old_placeholder = &self.images_map[file_path];
                    let image_option = &self.image_replacements[old_placeholder];
                    // 插入图片到属性中
                    self.image_replacements
                        .insert(placeholder.to_string(), image_option.clone());
                } else {
                    // 收集添加的图片路径
                    self.images_map
                        .insert(file_path.to_string(), placeholder.to_string());
                    // 插入图片到属性中
                    self.image_replacements
                        .insert(placeholder.to_string(), Some(DocxImage::new(file_path)?));
                }
            }
        }

        Ok(())
    }

    /// 添加待替换的图片
    /// @param placeholder 替换的字符串
    /// @param image_path 图片路径
    /// @param width 图片的宽度(厘米)
    /// @param height 图片的高度(厘米)
    pub fn add_image_file_size_replacement(
        &mut self,
        placeholder: &str,
        image_path: Option<&str>,
        width: f32,
        height: f32,
    ) -> Result<(), DocxError> {
        match image_path {
            None => {
                // 插入图片到属性中
                self.image_replacements
                    .insert(placeholder.to_string(), None);
            }
            Some(file_path) => {
                // 将厘米单位换算成emu
                let width_emu = (width * DOCX_EMU) as u64;
                let height_emu = (height * DOCX_EMU) as u64;
                // 判断是否添加过该图片
                if self.images_map.contains_key(file_path) {
                    let old_placeholder = &self.images_map[file_path];
                    let image_option = &self.image_replacements[old_placeholder];

                    if let Some(image) = image_option {
                        let docx_image =
                            DocxImage::clone_image_reset_size(image, width_emu, height_emu);
                        // 插入图片到属性中
                        self.image_replacements
                            .insert(placeholder.to_string(), Some(docx_image));
                    }
                } else {
                    // 收集添加的图片路径
                    self.images_map
                        .insert(file_path.to_string(), placeholder.to_string());
                    // 插入图片到属性中
                    self.image_replacements.insert(
                        placeholder.to_string(),
                        Some(DocxImage::new_size(file_path, width_emu, height_emu)?),
                    );
                }
            }
        }

        Ok(())
    }

    /// 添加待替换的图片，替换的图片大小默认6.09*5.9厘米
    /// @param placeholder 替换的字符串
    /// @param image_url 图片路径
    pub async fn add_image_url_replacement(
        &mut self,
        placeholder: &str,
        image_url: Option<&str>,
    ) -> Result<(), DocxError> {
        match image_url {
            None => {
                // 插入图片到属性中
                self.image_replacements
                    .insert(placeholder.to_string(), None);
            }
            Some(url) => {
                // 判断是否添加过该图片
                if self.images_map.contains_key(url) {
                    let old_placeholder = &self.images_map[url];
                    let image_option = &self.image_replacements[old_placeholder];
                    // 插入图片到属性中
                    self.image_replacements
                        .insert(placeholder.to_string(), image_option.clone());
                } else {
                    // 收集添加的图片路径
                    self.images_map
                        .insert(url.to_string(), placeholder.to_string());
                    // 发送请求
                    let (image_data, image_ext) = request_image_data(&self.client, url).await?;
                    // 插入图片到属性中
                    self.image_replacements.insert(
                        placeholder.to_string(),
                        Some(DocxImage::new_image_data(url, image_data, &image_ext)?),
                    );
                }
            }
        }

        Ok(())
    }

    /// 添加待替换的图片
    /// @param placeholder 替换的字符串
    /// @param image_url 图片路径
    /// @param width 图片的宽度(厘米)
    /// @param height 图片的高度(厘米)
    pub async fn add_image_url_size_replacement(
        &mut self,
        placeholder: &str,
        image_url: Option<&str>,
        width: f32,
        height: f32,
    ) -> Result<(), DocxError> {
        match image_url {
            None => {
                // 插入图片到属性中
                self.image_replacements
                    .insert(placeholder.to_string(), None);
            }
            Some(url) => {
                // 将厘米单位换算成emu
                let width_emu = (width * DOCX_EMU) as u64;
                let height_emu = (height * DOCX_EMU) as u64;
                // 判断是否添加过该图片
                if self.images_map.contains_key(url) {
                    let old_placeholder = &self.images_map[url];
                    let image_option = &self.image_replacements[old_placeholder];
                    if let Some(image) = image_option {
                        let docx_image =
                            DocxImage::clone_image_reset_size(image, width_emu, height_emu);
                        // 插入图片到属性中
                        self.image_replacements
                            .insert(placeholder.to_string(), Some(docx_image));
                    }
                } else {
                    self.images_map
                        .insert(url.to_string(), placeholder.to_string());
                    // 发送请求
                    let (image_data, image_ext) = request_image_data(&self.client, url).await?;
                    // 插入图片到属性中
                    self.image_replacements.insert(
                        placeholder.to_string(),
                        Some(DocxImage::new_image_data_size(
                            url, image_data, &image_ext, width_emu, height_emu,
                        )?),
                    );
                }
            }
        }

        Ok(())
    }

    /// 处理模板
    /// @param template_path 模板路径
    /// @param output_path 输出路径
    pub fn process_template(
        &self,
        template_path: &str,
        output_path: &str,
    ) -> Result<(), DocxError> {
        // 1. 打开模板文件
        let template_file = File::open(template_path)?;
        let mut archive = ZipArchive::new(template_file)?;

        // 2. 创建输出文件
        let output_file = File::create(output_path)?;
        let mut zip_writer = ZipWriter::new(output_file);

        // 3. 遍历ZIP中的文件
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            // 文件内容
            let mut contents = Vec::new();
            // 读取文件内容到数组中
            file.read_to_end(&mut contents)?;
            // 匹配文件类型
            match file.name() {
                x if x == WORD_DOCUMENT => {
                    // 处理文档主内容,替换模板内容
                    contents = self.process_document_xml(&contents)?;
                }
                x if x == WORD_RELS_DOCUMENT => {
                    // 处理关系文件
                    contents = self.process_rels_xml(&contents)?;
                }
                &_ => {}
            }
            // 写入新文件
            writer_file(&mut zip_writer, &file, &contents)?
        }

        // 4. 添加新的图片文件
        for replacement in self.images_map.values() {
            if let Some(Some(replacement)) = self.image_replacements.get(replacement) {
                writer_image(&mut zip_writer, replacement)?;
            }
        }
        // 将内容写入压缩文件（docx）
        zip_writer.finish()?;
        Ok(())
    }

    fn process_element(&self, _element: &mut BytesStart) -> Result<(), DocxError> {
        let tag = String::from_utf8_lossy(_element.name().as_ref()).to_string();
        debug!("{:?}", tag);
        Ok(())
    }

    /// 处理文件内容
    /// @param contents 文件内容数组
    fn process_document_xml(&self, contents: &[u8]) -> Result<Vec<u8>, DocxError> {
        // 创建xml写对象
        let mut xml_writer = Writer::new(Cursor::new(Vec::new()));
        // 读取xml文件的内容
        let mut reader = quick_xml::Reader::from_reader(contents);
        // reader.config_mut().trim_text(true);
        // 缓存数组
        let mut buf = Vec::new();
        // 图片对应的字符串占位符
        let mut current_placeholder = String::new();
        // 循环读取xml数据
        loop {
            // 读取数据
            match reader.read_event_into(&mut buf)? {
                Event::Start(e) => {
                    let mut element = e.to_owned();
                    self.process_element(&mut element)?;
                    // 如果为空，写入标签头
                    if current_placeholder.is_empty() {
                        xml_writer.write_event(Event::Start(element))?;
                    }
                }
                Event::Text(e) => {
                    // 读取标签的内容
                    let mut text = e.unescape()?.into_owned();
                    // 判断是否有替换字符串开头内容"{{"
                    if text.contains(PREFIX_TAG) {
                        // 判断是否包含结束字符串}}
                        if text.contains(SUFFIX_TAG) {
                            // 1、替换文本占位符操作
                            self.process_text(&mut text);
                            // 2、替换图片占位符操作
                            if self.image_replacements.contains_key(&text) {
                                current_placeholder.push_str(&text);
                            } else {
                                xml_writer
                                    .write_event(Event::Text(BytesText::new(text.as_str())))?;
                            }
                        } else {
                            // 将字符串保存
                            current_placeholder.push_str(&text);
                        }
                    } else {
                        // 判断current_placeholder字符串是否有内容
                        if current_placeholder.is_empty() {
                            // 将原有字符串写入文档
                            xml_writer.write_event(Event::Text(BytesText::new(text.as_str())))?;
                        } else {
                            // 将字符串写入
                            current_placeholder.push_str(text.as_str());
                            // 判断是否有结束字符串}}
                            if current_placeholder.contains(PREFIX_TAG)
                                && current_placeholder.contains(SUFFIX_TAG)
                            {
                                // 1、替换文本占位符操作
                                self.process_text(&mut current_placeholder);
                                // 2、如果不包含写入数据
                                if !self.image_replacements.contains_key(&current_placeholder) {
                                    xml_writer.write_event(Event::Text(BytesText::new(
                                        current_placeholder.as_str(),
                                    )))?;
                                    // 清理数据
                                    current_placeholder.clear();
                                }
                            }
                        }
                    }
                }
                Event::End(e) => {
                    // 判断是否为空，为空，直接添加结尾标签
                    if current_placeholder.is_empty() {
                        xml_writer.write_event(Event::End(e))?;
                    } else if current_placeholder.contains(PREFIX_TAG)
                        && current_placeholder.contains(SUFFIX_TAG)
                    {
                        // 判断是否为段落
                        if e.name().as_ref() == WORD_PARAGRAPH_TAG {
                            // 判断是否为完整替换字符串
                            if let Some(Some(docx_image)) =
                                self.image_replacements.get(&current_placeholder)
                            {
                                // 替换占位符为图片
                                create_drawing_element(
                                    &mut xml_writer,
                                    &docx_image.relation_id,
                                    docx_image.width,
                                    docx_image.height,
                                )?;
                            }
                            // 清除字符串
                            current_placeholder.clear();
                        }
                        // 写入结尾标签
                        xml_writer.write_event(Event::End(e))?;
                    }
                }
                Event::Eof => break,
                Event::Empty(e) => {
                    // 如果为空写入文档
                    if current_placeholder.is_empty() {
                        xml_writer.write_event(Event::Empty(e))?;
                    }
                }
                e => {
                    xml_writer.write_event(e)?;
                }
            }
            buf.clear();
        }
        // 返回文件数组
        Ok(xml_writer.into_inner().into_inner())
    }

    fn process_rels_xml(&self, xml_data: &[u8]) -> Result<Vec<u8>, DocxError> {
        // 创建xml写对象
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        // 写入xml标签头
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // 写入XML根元素
        writer.write_event(Event::Start(
            BytesStart::new("Relationships").with_attributes([(
                "xmlns",
                "http://schemas.openxmlformats.org/package/2006/relationships",
            )]),
        ))?;

        // 读取原始数据
        let mut reader = quick_xml::Reader::from_reader(xml_data);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            // 读取关系文件
            match reader.read_event_into(&mut buf)? {
                // 判断关系文件内容是否为关联标签
                Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                    // 写入关系标签内容
                    writer.write_event(Event::Empty(e))?;
                }
                // 文件读取完毕
                Event::Eof => break,
                _ => {}
            }
            // 清理内容
            buf.clear();
        }

        // 添加新的图片关系
        for placeholder in self.images_map.values() {
            if let Some(Some(docx_image)) = self.image_replacements.get(placeholder) {
                // 创建图片路径
                let image_path = format!(
                    "media/image_{}.{}",
                    docx_image.relation_id, docx_image.image_ext
                );
                // 创建图片关系标签
                let relationship = BytesStart::new("Relationship").with_attributes([
                    ("Id", docx_image.relation_id.as_str()),
                    (
                        "Type",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                    ),
                    ("Target", &image_path),
                ]);
                // 写入关系标签数据
                writer.write_event(Event::Empty(relationship))?;
            }
        }

        // 结束根元素
        writer.write_event(Event::End(BytesEnd::new("Relationships")))?;
        // 输出关系文件内容
        Ok(writer.into_inner().into_inner())
    }

    // 替换模板属性
    fn process_text(&self, text: &mut String) {
        for (placeholder, value) in &self.text_replacements {
            *text = text.replace(placeholder, value);
        }
    }
}

impl Default for DocxTemplate {
    fn default() -> Self {
        Self::new()
    }
}

/// 写入图片  
/// @param zip_writer 写入对象  
/// @param replacement 图片对象  
fn writer_image(
    zip_writer: &mut ZipWriter<File>,
    replacement: &DocxImage,
) -> Result<(), DocxError> {
    let image_path = format!(
        "{}{}.{}",
        WORD_MEDIA_IMAGE, replacement.relation_id, replacement.image_ext,
    );
    // 写入图片到word压缩文件中
    zip_writer.start_file(&image_path, SimpleFileOptions::default())?;
    zip_writer.write_all(&replacement.image_data)?;
    Ok(())
}

pub fn writer_file(
    zip_writer: &mut ZipWriter<File>,
    file: &ZipFile<File>,
    contents: &[u8],
) -> Result<(), DocxError> {
    // 写入新文件
    let option = SimpleFileOptions::default()
        .compression_method(file.compression())
        .unix_permissions(file.unix_mode().unwrap_or(0o644));
    // 写入内容
    zip_writer.start_file(file.name(), option)?;
    zip_writer.write_all(contents)?;

    Ok(())
}
