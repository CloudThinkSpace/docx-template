use crate::error::DocxError;
use reqwest::Client;

/// 获取图片数据   
/// @param client 请求客户端  
/// @param url 图片url路径  
/// @return (data, ext) 返回 Vec<u8>和图片扩展名  
pub async fn request_image_data(
    client: &Client,
    url: &str,
) -> Result<(Vec<u8>, String), DocxError> {
    // 发送请求
    let response = client.get(url).send().await?;
    // 检查状态码
    if !response.status().is_success() {
        return Err(DocxError::NotImage("请求图片错误".to_string()));
    }
    // 获取请求头
    let headers = response.headers().clone();
    // 先读取头信息
    let content_type = headers
        .get(reqwest::header::CONTENT_TYPE)
        .map(|h| h.to_str().unwrap_or(""))
        .unwrap_or("");
    // 获取内容类型以验证是否为图片
    if !content_type.starts_with("image/") {
        return Err(DocxError::NotImage(content_type.to_string()));
    }
    // 读取字节
    let image_data = response.bytes().await?.to_vec();
    let extension = &content_type[6..];
    Ok((image_data, extension.to_string()))
}
