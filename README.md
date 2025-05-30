# 可以修改docx中的内容：
> 可以提供两类数据
- 字符串替换：{{name}} to text
- 图片替换： {{image1}} to image（png/jpeg）
``` rust
// 1. 创建模板处理器
let mut docx_template = DocxTemplate::new();

// 2. 添加要替换的内容
docx_template.add_text_replacement("{{groupLeader}}", "Acme 公司");
docx_template.add_text_replacement("{{groupMembers}}", "张三");
docx_template.add_text_replacement("{{city}}", "2023-11-25");
docx_template.add_text_replacement("{{town}}", "¥10,000");
docx_template.add_text_replacement("{{county}}", "30天内付清");

// 添加在线图片，大小默认6.09*5.9
docx_template.add_image_url_replacement("{{photo1}}", Some("http://xxxxx/4da6f0c9-2610-4574-8f0a-638f9f5eb1d7.png")).await?;
// 大小为厘米，根据文档实际需求设置大小
docx_template.add_image_size_url_replacement("{{photo2}}", Some("http://xxxxx/5d3c83de-99e1-4081-a4ee-ba4925d1d3a5.png"),5.0,5.0).await?;
// 替换图片占位符为空
docx_template.add_image_replacement("{{photo3}}", None).expect("添加图片失败");
docx_template.add_image_replacement("{{photo4}}", None).expect("添加图片失败");
```
其中替换图片可以替换本地土和在线图片