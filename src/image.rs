use crate::error::DocxError;
use image::{GenericImageView, load_from_memory};
use std::fs::File;
use std::io::Read;
use std::path::Path;
use uuid::Uuid;

// 1厘米约等于360000EMU
pub static DOCX_EMU: f32 = 360000.0;
pub static DOCX_MAX_EMU: u64 = (21.0 * 360000.0 / 2.0) as u64;
// 1英寸=96像素
static DPI: f64 = 96f64;
// 1英寸=914400 EMU
static EMU: f64 = 914400f64;

// 添加的图标对象
pub struct DocxImage {
    // 图片路径
    pub image_path: String,
    // 图片扩展名
    pub image_ext: String,
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
    /// @param image_path 本地图片路径  
    pub fn new(image_path: &str) -> Result<Self, DocxError> {
        // 打开文件读取数据到数组中
        let mut file = File::open(image_path)?;
        let mut image_data = Vec::new();
        file.read_to_end(&mut image_data)?;
        // 获取扩展名
        let ext = get_extension(image_path)?;
        let (width_emu, height_emu) = get_image_size(&image_data)?;
        Self::new_image_data_size(image_path, image_data, ext, width_emu, height_emu)
    }

    /// 设置图片大小  
    /// @param image_path 图片路径  
    /// @param width 图片宽度（emu）
    /// @param height 图片高度 （emu）
    pub fn new_size(image_path: &str, width: u64, height: u64) -> Result<Self, DocxError> {
        // 打开文件读取数据到数组中
        let mut file = File::open(image_path)?;
        let mut image_data = Vec::new();
        file.read_to_end(&mut image_data)?;
        // 获取扩展名
        let ext = get_extension(image_path)?;
        DocxImage::new_image_data_size(image_path, image_data, ext, width, height)
    }

    /// 设置图片大小  
    /// @param image_url 图片路径  
    /// @param image_data 图片数据  
    pub fn new_image_data(
        image_url: &str,
        image_data: Vec<u8>,
        image_ext: &str,
    ) -> Result<Self, DocxError> {
        let (width_emu, height_emu) = get_image_size(&image_data)?;
        DocxImage::new_image_data_size(image_url, image_data, image_ext, width_emu, height_emu)
    }

    /// 设置图片大小  
    /// @param image_url 图片路径  
    /// @param image_data 图片数据  
    /// @param image_ext 图片扩展名  
    /// @param width 图片宽度（emu） 
    /// @param height 图片高度（emu） 
    pub fn new_image_data_size(
        image_url: &str,
        image_data: Vec<u8>,
        image_ext: &str,
        width: u64,
        height: u64,
    ) -> Result<Self, DocxError> {
        Ok(DocxImage {
            image_path: image_url.to_string(),
            image_ext: image_ext.to_string(),
            relation_id: format!("rId{}", Uuid::new_v4().simple()),
            width,
            height,
            image_data,
        })
    }
}

fn get_image_size(image_data: &[u8]) -> Result<(u64, u64), DocxError> {
    let img = load_from_memory(image_data)?;
    let (width_px, height_px) = img.dimensions();
    let mut width_emu = (width_px as f64 * EMU / DPI) as u64;
    let mut height_emu = (height_px as f64 * EMU / DPI) as u64;
    // 判断图片是否大于文档宽度
    if width_emu > DOCX_MAX_EMU {
        height_emu = DOCX_MAX_EMU * height_emu / width_emu;
        width_emu = DOCX_MAX_EMU;
        Ok((width_emu, height_emu))
    } else {
        Ok((width_emu, height_emu))
    }
}

/// 获取本地图片的扩展名
/// @param image_path 本地图片路径
/// @return 图片扩展名
fn get_extension(image_path: &str) -> Result<&str, DocxError> {
    Path::new(image_path)
        .extension()
        .and_then(|s| s.to_str())
        .ok_or_else(|| DocxError::ImageNotFound("Could not determine image extension".to_string()))
}
