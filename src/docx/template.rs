use std::io::Write;
use quick_xml::events::Event;
use quick_xml::Writer;
use crate::error::DocxError;


/// 创建图片标签  
/// @param writer 写入对象  
/// @param relation_id 关联图片编号  
/// @param width 图片宽度  
/// @param height 图片高度  
pub fn create_drawing_element<T>(
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