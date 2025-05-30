use quick_xml::Writer;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use reqwest::Client;
use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::path::Path;
use std::time::Duration;
use thiserror::Error;
use uuid::Uuid;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

pub struct DocxTemplate {
    // 待替换的字符串
    text_replacements: HashMap<String, String>,
    // 待替换的图片
    image_replacements: HashMap<String, Option<DocxImage>>,
    // 请求对象
    client: Client,
}

impl DocxTemplate {
    pub fn new() -> Self {
        DocxTemplate {
            text_replacements: HashMap::new(),
            image_replacements: HashMap::new(),
            client: Client::builder()
                .timeout(Duration::from_secs(10)) // 设置超时
                .build()
                .unwrap(),
        }
    }

    /// 添加待替换的字符以及对应的值
    /// @param placeholder 替换的字符串
    /// @param value 替换的值
    pub fn add_text_replacement(&mut self, placeholder: &str, value: &str) {
        self.text_replacements
            .insert(placeholder.to_string(), value.to_string());
    }

    /// 添加要替换的图片
    /// @param placeholder 替换的字符串
    /// @param image_path 图片路径
    pub fn add_image_replacement(
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
            Some(data) => {
                // 插入图片到属性中
                self.image_replacements
                    .insert(placeholder.to_string(), Some(DocxImage::new(data)?));
            }
        }

        Ok(())
    }

    /// 添加要替换的图片
    /// @param placeholder 替换的字符串
    /// @param image_path 图片路径
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
                // 发送请求
                let response = self.client.get(url).send().await?;
                // 检查状态码
                if response.status().is_success() {
                    // 读取字节
                    let image_data = response.bytes().await?.to_vec();
                    // 插入图片到属性中
                    self.image_replacements.insert(
                        placeholder.to_string(),
                        Some(DocxImage::new_image_data(url, image_data)?),
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
                "word/document.xml" => {
                    // 处理文档主内容,替换模板内容
                    contents = self.process_document_xml(&contents)?;
                }
                "word/_rels/document.xml.rels" => {
                    // 处理关系文件
                    contents = self.process_rels_xml(&contents)?;
                }
                &_ => {}
            }

            // 写入新文件
            let option = SimpleFileOptions::default()
                .compression_method(file.compression())
                .unix_permissions(file.unix_mode().unwrap_or(0o644));
            // 写入内容
            zip_writer.start_file(file.name(), option)?;
            zip_writer.write_all(&contents)?;
        }

        // 4. 添加新的图片文件
        for (_, replacement) in &self.image_replacements {
            if let Some(replacement) = replacement {
                let image_path = format!(
                    "word/media/image_{}.{}",
                    replacement.relation_id,
                    DocxTemplate::get_extension(&replacement.image_path)?
                );
                // 写入图片到word压缩文件中
                zip_writer.start_file(&image_path, SimpleFileOptions::default())?;
                zip_writer.write_all(&replacement.image_data)?;
            }
        }
        // 将内容写入压缩文件（docx）
        zip_writer.finish()?;
        Ok(())
    }

    fn process_element(&self, element: &mut BytesStart) -> Result<(), DocxError> {
        // println!("{:?}", String::from_utf8_lossy(element.name().as_ref()));
        Ok(())
    }

    /// 处理文件内容
    /// @param contents 文件内容数组
    fn process_document_xml(&self, contents: &[u8]) -> Result<Vec<u8>, DocxError> {
        // 创建xml写对象
        let mut xml_writer = Writer::new(Cursor::new(Vec::new()));
        // 写入xml文件头
        // xml_writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;
        // 读取xml文件的内容
        let mut reader = quick_xml::Reader::from_reader(&contents[..]);
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
                    if e.name().as_ref() == b"w:p" {
                        current_placeholder.clear();
                    }
                    xml_writer.write_event(Event::Start(element))?;
                }
                Event::Text(e) => {
                    // 读取标签的内容
                    let mut text = e.unescape()?.into_owned();
                    // 替换文本占位符操作
                    self.process_text(&mut text);
                    // 判断图片占位符是否包含在内容
                    if self.image_replacements.contains_key(&text) {
                        current_placeholder.push_str(&text);
                    } else {
                        xml_writer.write_event(Event::Text(BytesText::new(text.as_str())))?;
                    }
                }
                Event::End(e) => {
                    // 判断标签是w:p，并且判断当前待替换的图片字符串是否为空
                    if e.name().as_ref() == b"w:p" && !current_placeholder.is_empty() {
                        if let Some(Some(docx_image)) =
                            self.image_replacements.get(&current_placeholder)
                        {
                            // 替换占位符为图片
                            DocxTemplate::create_drawing_element(
                                &mut xml_writer,
                                &docx_image.relation_id,
                                docx_image.width,
                                docx_image.height,
                            )?;
                        } else {
                            // 保留原始占位符文本
                            xml_writer.write_event(Event::Text(BytesText::from_escaped(
                                // current_placeholder.as_str(),
                                "",
                            )))?;
                        }
                        current_placeholder.clear();
                    }
                    xml_writer.write_event(Event::End(e))?;
                }
                Event::Eof => break,
                e => {
                    // 写入原有信息
                    xml_writer.write_event(e)?
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
        for (_, value) in &self.image_replacements {
            if let Some(docx_image) = value {
                // 获取图片扩展名
                let extension = DocxTemplate::get_extension(&docx_image.image_path)?;
                // 创建图片路径
                let image_path = format!("media/image_{}.{}", docx_image.relation_id, extension);
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

    fn get_extension(image_path: &str) -> Result<&str, DocxError> {
        Path::new(image_path)
            .extension()
            .and_then(|s| s.to_str())
            .ok_or_else(|| {
                DocxError::ImageNotFound("Could not determine image extension".to_string())
            })
    }
    // 替换模板属性
    fn process_text(&self, text: &mut String) {
        for (placeholder, value) in &self.text_replacements {
            *text = text.replace(placeholder, value);
        }
    }

    fn create_drawing_element<T>(
        writer: &mut Writer<T>,
        relation_id: &str,
        width: u64,
        height: u64,
    ) -> Result<(), DocxError>
    where
        T: Write,
    {
        let drawing = format!(
            r#"
        <w:drawing>
            <wp:inline distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="{}" cy="{}"/>
                <wp:docPr id="1" name="Picture 1" descr="Generated image"/>
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:nvPicPr>
                                <pic:cNvPr id="0" name="Picture 1" descr="Generated image"/>
                                <pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed="{}"/>
                                <a:stretch>
                                    <a:fillRect/>
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0"/>
                                    <a:ext cx="{}" cy="{}"/>
                                </a:xfrm>
                                <a:prstGeom prst="rect">
                                    <a:avLst/>
                                </a:prstGeom>
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:inline>
        </w:drawing>
    "#,
            width, height, relation_id, width, height,
        );

        let mut reader = quick_xml::Reader::from_str(&drawing);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Eof => break,
                e => {
                    writer.write_event(e)?;
                }
            }
        }
        Ok(())
    }
}

// 添加的图标对象
struct DocxImage {
    // 图片路径
    pub image_path: String,
    // 图片数据
    pub image_data: Vec<u8>,
    // 关联id
    pub relation_id: String,
    // 图片高度
    pub width: u64,
    // 图片高度
    pub height: u64,
}

impl DocxImage {
    /// 创建图片对象
    /// @param image_path 图片路径
    pub fn new(image_path: &str) -> Result<Self, DocxError> {
        Self::new_size(image_path, 6.09, 5.9)
    }
    /// 设置图片大小
    /// @param image_path 图片路径
    /// @param width 图片宽度
    /// @param height 图片高度
    pub fn new_size(image_path: &str, width: f32, height: f32) -> Result<Self, DocxError> {
        // 打开文件读取数据到数组中
        let mut file = File::open(image_path)?;
        let mut image_data = Vec::new();
        file.read_to_end(&mut image_data)?;
        DocxImage::new_image_data_size(image_path, image_data, width, height)
    }

    /// 设置图片大小
    /// @param image_url 图片路径
    /// @param image_data 图片数据
    pub fn new_image_data(image_url: &str, image_data: Vec<u8>) -> Result<Self, DocxError> {
        DocxImage::new_image_data_size(image_url, image_data, 6.09, 5.9)
    }

    /// 设置图片大小
    /// @param image_url 图片路径
    /// @param image_data 图片数据
    /// @param width 图片宽度
    /// @param height 图片高度
    pub fn new_image_data_size(
        image_url: &str,
        image_data: Vec<u8>,
        width: f32,
        height: f32,
    ) -> Result<Self, DocxError> {
        Ok(DocxImage {
            image_path: image_url.to_string(),
            relation_id: format!("rId{}", Uuid::new_v4().simple()),
            width: (width * 360000.0) as u64,
            height: (height * 360000.0) as u64,
            image_data,
        })
    }
}

#[derive(Error, Debug)]
pub enum DocxError {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("Zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("UTF-8 error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),
    #[error("Image not found: {0}")]
    ImageNotFound(String),
    #[error("Image url not found: {0}")]
    ImageUrlFound(#[from] reqwest::Error),
}
