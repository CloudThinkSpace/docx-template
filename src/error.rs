use thiserror::Error;

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
    InvalidImageUrl(#[from] reqwest::Error),
    #[error("read image size error: {0}")]
    ReadImageSize(#[from] image::ImageError),
}