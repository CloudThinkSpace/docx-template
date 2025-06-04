pub mod docx;
pub mod error;
pub mod image;
pub mod request;
mod docx1;

#[cfg(test)]
mod tests {
    use crate::docx::DocxTemplate;

    #[tokio::test] // 使用 tokio 运行时
    async fn test_replacement() {
        // 1. 创建模板处理器
        let mut docx_template = DocxTemplate::new();

        // 2. 添加要替换的内容
        docx_template.add_text_replacement("{{groupLeader}}", "Acme 公司");
        docx_template.add_text_replacement("{{groupMembers}}", "张三");
        docx_template.add_text_replacement("{{city}}", "2023-11-25");
        docx_template.add_text_replacement("{{town}}", "¥10,000");
        docx_template.add_text_replacement("{{county}}", "30天内付清");

        docx_template.add_image_url_replacement("{{photo1}}", Some("http://gips3.baidu.com/it/u=100751361,1567855012&fm=3028&app=3028&f=JPEG&fmt=auto?w=960&h=1280")).await.expect("msg");
        docx_template.add_image_url_size_replacement("{{photo2}}", Some("http://gips3.baidu.com/it/u=100751361,1567855012&fm=3028&app=3028&f=JPEG&fmt=auto?w=960&h=1280"),5.0,5.0).await.expect("msg");
        docx_template
            .add_image_file_replacement("{{photo3}}", Some("./data/bgImg.png"))
            .expect("添加本地图片失败");
        docx_template
            .add_image_file_replacement("{{photo4}}", Some("./data/bgImg.png"))
            .expect("添加图片失败");

        // 3. 处理模板并生成新文档
        docx_template.process_template("./data/西藏自治区严格管控核查表单.docx", "output.docx").expect("");

        println!("文档生成成功!");
    }
}
