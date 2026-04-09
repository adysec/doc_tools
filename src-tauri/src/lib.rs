use image::codecs::jpeg::JpegEncoder;
use image::codecs::png::PngEncoder;
use image::{ColorType, ImageEncoder, ImageReader};
use regex::Regex;
use serde::Serialize;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::time::Instant;
use tauri::Emitter;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[derive(Serialize)]
struct CompressionResult {
    input_path: String,
    output_path: String,
    original_size: u64,
    compressed_size: u64,
    compressed_images: u32,
    elapsed_seconds: f32,
}

#[derive(Serialize)]
struct UnlockResult {
    input_path: String,
    output_path: String,
    removed_rules: u32,
    elapsed_seconds: f32,
}

#[derive(Serialize)]
struct WatermarkResult {
    input_path: String,
    output_path: String,
    cleared_headers: u32,
    removed_backgrounds: u32,
    removed_shapes: u32,
    elapsed_seconds: f32,
}

#[derive(Serialize)]
struct OutlineResult {
    input_path: String,
    output_path: String,
    removed_heading_styles: u32,
    replaced_outline_levels: u32,
    inserted_outline_levels: u32,
    removed_toc_index_blocks: u32,
    removed_toc_index_fields: u32,
    elapsed_seconds: f32,
}

fn extract_heading_style_ids(styles_xml: &str) -> Result<HashSet<String>, String> {
    let mut ids = HashSet::new();
    for lvl in 1..=9 {
        ids.insert(format!("Heading{lvl}"));
    }

    let style_block = Regex::new(r#"(?s)<w:style\b([^>]*)>(.*?)</w:style>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let style_type_attr = Regex::new(r#"\bw:type="([^"]+)""#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let style_id_attr = Regex::new(r#"\bw:styleId="([^"]+)""#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let based_on = Regex::new(r#"<w:basedOn\s+w:val="([^"]+)"\s*/>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let style_name = Regex::new(r#"<w:name\s+w:val="([^"]+)"\s*/>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;

    for cap in style_block.captures_iter(styles_xml) {
        let attrs = cap.get(1).map(|m| m.as_str()).unwrap_or_default();
        let style_type = style_type_attr
            .captures(attrs)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str())
            .unwrap_or_default();
        if style_type != "paragraph" {
            continue;
        }

        let style_id = style_id_attr
            .captures(attrs)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str())
            .unwrap_or_default()
            .to_string();
        if style_id.is_empty() {
            continue;
        }

        let block = cap.get(2).map(|m| m.as_str()).unwrap_or_default();

        let has_outline_level = block.contains("<w:outlineLvl");
        let based_on_heading = based_on
            .captures(block)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str())
            .map(|v| {
                let upper = v.to_ascii_uppercase();
                upper.starts_with("HEADING") || upper.starts_with("TITLE")
            })
            .unwrap_or(false);
        let name_heading = style_name
            .captures(block)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str().to_lowercase())
            .map(|v| v.contains("heading") || v.contains("标题"))
            .unwrap_or(false);

        if has_outline_level || based_on_heading || name_heading {
            ids.insert(style_id);
        }
    }

    Ok(ids)
}

#[derive(Default, Clone)]
struct StyleDirectFormat {
    ppr_inner: String,
    rpr_inner: String,
}

fn extract_default_paragraph_style_id(styles_xml: &str) -> Result<Option<String>, String> {
    let style_block = Regex::new(r#"(?s)<w:style\b([^>]*)>(.*?)</w:style>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let style_type_attr = Regex::new(r#"\bw:type="([^"]+)""#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let style_id_attr = Regex::new(r#"\bw:styleId="([^"]+)""#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let default_attr = Regex::new(r#"\bw:default="1""#)
        .map_err(|e| format!("正则编译失败: {e}"))?;

    for cap in style_block.captures_iter(styles_xml) {
        let attrs = cap.get(1).map(|m| m.as_str()).unwrap_or_default();
        let style_type = style_type_attr
            .captures(attrs)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str())
            .unwrap_or_default();
        if style_type != "paragraph" || !default_attr.is_match(attrs) {
            continue;
        }
        let style_id = style_id_attr
            .captures(attrs)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str().to_string());
        if style_id.is_some() {
            return Ok(style_id);
        }
    }
    Ok(None)
}

fn extract_heading_style_direct_formats(
    styles_xml: &str,
    heading_style_ids: &HashSet<String>,
) -> Result<HashMap<String, StyleDirectFormat>, String> {
    let style_block = Regex::new(r#"(?s)<w:style\b([^>]*)>(.*?)</w:style>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let style_type_attr = Regex::new(r#"\bw:type="([^"]+)""#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let style_id_attr = Regex::new(r#"\bw:styleId="([^"]+)""#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let ppr_block = Regex::new(r#"(?s)<w:pPr\b[^>]*>(.*?)</w:pPr>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let rpr_block = Regex::new(r#"(?s)<w:rPr\b[^>]*>(.*?)</w:rPr>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;

    let mut map = HashMap::new();
    for cap in style_block.captures_iter(styles_xml) {
        let attrs = cap.get(1).map(|m| m.as_str()).unwrap_or_default();
        let style_type = style_type_attr
            .captures(attrs)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str())
            .unwrap_or_default();
        if style_type != "paragraph" {
            continue;
        }

        let Some(style_id) = style_id_attr
            .captures(attrs)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str().to_string())
        else {
            continue;
        };

        if !heading_style_ids.contains(&style_id) {
            continue;
        }

        let inner = cap.get(2).map(|m| m.as_str()).unwrap_or_default();
        let ppr_inner = ppr_block
            .captures(inner)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str().to_string())
            .unwrap_or_default();
        let rpr_inner = rpr_block
            .captures(inner)
            .and_then(|c| c.get(1))
            .map(|m| m.as_str().to_string())
            .unwrap_or_default();

        map.insert(style_id, StyleDirectFormat { ppr_inner, rpr_inner });
    }
    Ok(map)
}

#[derive(Serialize, Clone)]
struct CompressionProgress {
    processed: usize,
    total: usize,
    percent: f32,
    current_file: String,
    compressed_images: u32,
}

#[tauri::command]
async fn compress_docx(
    app_handle: tauri::AppHandle,
    input_path: String,
    output_path: String,
    quality: u8,
    max_width: u32,
) -> Result<CompressionResult, String> {
    tauri::async_runtime::spawn_blocking(move || {
        compress_docx_impl(app_handle, input_path, output_path, quality, max_width)
    })
    .await
    .map_err(|e| format!("后台任务执行失败: {e}"))?
}

fn compress_docx_impl(
    app_handle: tauri::AppHandle,
    input_path: String,
    output_path: String,
    quality: u8,
    max_width: u32,
) -> Result<CompressionResult, String> {
    if !input_path.to_lowercase().ends_with(".docx") {
        return Err("输入文件必须是 .docx".to_string());
    }

    if !(1..=100).contains(&quality) {
        return Err("压缩质量必须在 1-100 之间".to_string());
    }

    if max_width < 300 || max_width > 4000 {
        return Err("最大宽度必须在 300-4000 之间".to_string());
    }

    let timer = Instant::now();
    let original_size = std::fs::metadata(&input_path)
        .map_err(|e| format!("读取输入文件信息失败: {e}"))?
        .len();

    let file = File::open(&input_path).map_err(|e| format!("无法打开输入文件: {e}"))?;
    let mut archive = ZipArchive::new(file).map_err(|e| format!("DOCX 解析失败: {e}"))?;
    let total_files = archive.len();

    let mut buffer = Cursor::new(Vec::new());
    let mut compressed_images = 0u32;

    let _ = app_handle.emit(
        "compress-progress",
        CompressionProgress {
            processed: 0,
            total: total_files,
            percent: 0.0,
            current_file: "初始化中".to_string(),
            compressed_images,
        },
    );

    {
        let mut zip_writer = ZipWriter::new(&mut buffer);

        for i in 0..total_files {
            let mut entry = archive
                .by_index(i)
                .map_err(|e| format!("读取 DOCX 条目失败: {e}"))?;
            let name = entry.name().to_string();
            let mut data = Vec::new();
            entry
                .read_to_end(&mut data)
                .map_err(|e| format!("读取文件数据失败: {e}"))?;

            if name.starts_with("word/media/")
                && (name.ends_with(".png")
                    || name.ends_with(".jpg")
                    || name.ends_with(".jpeg")
                    || name.ends_with(".PNG")
                    || name.ends_with(".JPG")
                    || name.ends_with(".JPEG"))
            {
                if let Ok(reader) = ImageReader::new(Cursor::new(&data)).with_guessed_format() {
                    if let Ok(decoded) = reader.decode() {
                        let resized = decoded.resize(
                            max_width.min(decoded.width()),
                            max_width.min(decoded.height()),
                            image::imageops::FilterType::Lanczos3,
                        );

                        let mut img_buf = Vec::new();
                        let lower = name.to_ascii_lowercase();

                        if lower.ends_with(".png") {
                            let encoder = PngEncoder::new(&mut img_buf);
                            encoder
                                .write_image(
                                    &resized.to_rgba8(),
                                    resized.width(),
                                    resized.height(),
                                    ColorType::Rgba8.into(),
                                )
                                .map_err(|e| format!("PNG 编码失败: {e}"))?;
                        } else {
                            let mut encoder = JpegEncoder::new_with_quality(&mut img_buf, quality);
                            encoder
                                .encode_image(&resized)
                                .map_err(|e| format!("JPEG 编码失败: {e}"))?;
                        }

                        if img_buf.len() < data.len() {
                            data = img_buf;
                            compressed_images += 1;
                        }
                    }
                }
            }

            zip_writer
                .start_file(
                    name.clone(),
                    zip::write::FileOptions::default()
                        .compression_method(CompressionMethod::Deflated)
                        .compression_level(Some(6)),
                )
                .map_err(|e| format!("写入 ZIP 条目失败: {e}"))?;
            zip_writer
                .write_all(&data)
                .map_err(|e| format!("写入压缩结果失败: {e}"))?;

            let processed = i + 1;
            let percent = if total_files == 0 {
                100.0
            } else {
                (processed as f32 / total_files as f32) * 100.0
            };
            let _ = app_handle.emit(
                "compress-progress",
                CompressionProgress {
                    processed,
                    total: total_files,
                    percent,
                    current_file: name,
                    compressed_images,
                },
            );
        }

        zip_writer
            .finish()
            .map_err(|e| format!("结束 ZIP 写入失败: {e}"))?;
    }

    let output_bytes = buffer.into_inner();
    let compressed_size = output_bytes.len() as u64;
    let mut out_file = File::create(&output_path).map_err(|e| format!("创建输出文件失败: {e}"))?;
    out_file
        .write_all(&output_bytes)
        .map_err(|e| format!("写入输出文件失败: {e}"))?;

    let _ = app_handle.emit(
        "compress-progress",
        CompressionProgress {
            processed: total_files,
            total: total_files,
            percent: 100.0,
            current_file: "完成".to_string(),
            compressed_images,
        },
    );

    Ok(CompressionResult {
        input_path,
        output_path,
        original_size,
        compressed_size,
        compressed_images,
        elapsed_seconds: timer.elapsed().as_secs_f32(),
    })
}

#[tauri::command]
async fn unlock_docx(input_path: String, output_path: String) -> Result<UnlockResult, String> {
    tauri::async_runtime::spawn_blocking(move || unlock_docx_impl(input_path, output_path))
        .await
        .map_err(|e| format!("后台任务执行失败: {e}"))?
}

fn unlock_docx_impl(input_path: String, output_path: String) -> Result<UnlockResult, String> {
    if !input_path.to_lowercase().ends_with(".docx") {
        return Err("输入文件必须是 .docx".to_string());
    }
    if !output_path.to_lowercase().ends_with(".docx") {
        return Err("输出文件必须是 .docx".to_string());
    }

    let timer = Instant::now();
    let source = File::open(&input_path).map_err(|e| format!("无法打开输入文件: {e}"))?;
    let mut archive = ZipArchive::new(source).map_err(|e| format!("DOCX 解析失败: {e}"))?;

    let single_tag = Regex::new(r#"(?s)<w:documentProtection[^>]*/>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let block_tag = Regex::new(r#"(?s)<w:documentProtection[^>]*>.*?</w:documentProtection>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;

    let mut removed_rules = 0u32;
    let mut buffer = Cursor::new(Vec::new());

    {
        let mut zip_writer = ZipWriter::new(&mut buffer);

        for i in 0..archive.len() {
            let mut entry = archive
                .by_index(i)
                .map_err(|e| format!("读取 DOCX 条目失败: {e}"))?;
            let name = entry.name().to_string();
            let mut data = Vec::new();
            entry
                .read_to_end(&mut data)
                .map_err(|e| format!("读取文件数据失败: {e}"))?;

            if name == "word/settings.xml" {
                let content = String::from_utf8(data)
                    .map_err(|e| format!("读取 settings.xml 失败: {e}"))?;
                let single_count = single_tag.find_iter(&content).count() as u32;
                let block_count = block_tag.find_iter(&content).count() as u32;
                let cleaned = single_tag.replace_all(&content, "");
                let cleaned = block_tag.replace_all(&cleaned, "");
                removed_rules += single_count + block_count;
                data = cleaned.into_owned().into_bytes();
            }

            zip_writer
                .start_file(
                    name,
                    zip::write::FileOptions::default()
                        .compression_method(CompressionMethod::Deflated)
                        .compression_level(Some(6)),
                )
                .map_err(|e| format!("写入 ZIP 条目失败: {e}"))?;
            zip_writer
                .write_all(&data)
                .map_err(|e| format!("写入解锁结果失败: {e}"))?;
        }

        zip_writer
            .finish()
            .map_err(|e| format!("结束 ZIP 写入失败: {e}"))?;
    }

    let output_bytes = buffer.into_inner();
    let mut out_file = File::create(&output_path).map_err(|e| format!("创建输出文件失败: {e}"))?;
    out_file
        .write_all(&output_bytes)
        .map_err(|e| format!("写入输出文件失败: {e}"))?;

    Ok(UnlockResult {
        input_path,
        output_path,
        removed_rules,
        elapsed_seconds: timer.elapsed().as_secs_f32(),
    })
}

#[tauri::command]
async fn remove_docx_watermark(
    input_path: String,
    output_path: String,
) -> Result<WatermarkResult, String> {
    tauri::async_runtime::spawn_blocking(move || remove_docx_watermark_impl(input_path, output_path))
        .await
        .map_err(|e| format!("后台任务执行失败: {e}"))?
}

fn remove_docx_watermark_impl(
    input_path: String,
    output_path: String,
) -> Result<WatermarkResult, String> {
    if !input_path.to_lowercase().ends_with(".docx") {
        return Err("输入文件必须是 .docx".to_string());
    }
    if !output_path.to_lowercase().ends_with(".docx") {
        return Err("输出文件必须是 .docx".to_string());
    }

    let timer = Instant::now();
    let source = File::open(&input_path).map_err(|e| format!("无法打开输入文件: {e}"))?;
    let mut archive = ZipArchive::new(source).map_err(|e| format!("DOCX 解析失败: {e}"))?;

    let background_tag = Regex::new(r#"(?s)<w:background[^>]*>.*?</w:background>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let shape_tag = Regex::new(r#"(?s)<v:shape[^>]*>.*?</v:shape>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;

    let mut buffer = Cursor::new(Vec::new());
    let mut cleared_headers = 0u32;
    let mut removed_backgrounds = 0u32;
    let mut removed_shapes = 0u32;

    {
        let mut zip_writer = ZipWriter::new(&mut buffer);

        for i in 0..archive.len() {
            let mut entry = archive
                .by_index(i)
                .map_err(|e| format!("读取 DOCX 条目失败: {e}"))?;
            let name = entry.name().to_string();
            let mut data = Vec::new();
            entry
                .read_to_end(&mut data)
                .map_err(|e| format!("读取文件数据失败: {e}"))?;

            if name.starts_with("word/header") && name.ends_with(".xml") {
                data = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
                    .to_vec();
                cleared_headers += 1;
            }

            if name == "word/document.xml" {
                let content = String::from_utf8(data)
                    .map_err(|e| format!("读取 document.xml 失败: {e}"))?;
                removed_backgrounds += background_tag.find_iter(&content).count() as u32;
                let cleaned = background_tag.replace_all(&content, "");
                removed_shapes += shape_tag.find_iter(&cleaned).count() as u32;
                let cleaned = shape_tag.replace_all(&cleaned, "");
                data = cleaned.into_owned().into_bytes();
            }

            zip_writer
                .start_file(
                    name,
                    zip::write::FileOptions::default()
                        .compression_method(CompressionMethod::Deflated)
                        .compression_level(Some(6)),
                )
                .map_err(|e| format!("写入 ZIP 条目失败: {e}"))?;
            zip_writer
                .write_all(&data)
                .map_err(|e| format!("写入去水印结果失败: {e}"))?;
        }

        zip_writer
            .finish()
            .map_err(|e| format!("结束 ZIP 写入失败: {e}"))?;
    }

    let output_bytes = buffer.into_inner();
    let mut out_file = File::create(&output_path).map_err(|e| format!("创建输出文件失败: {e}"))?;
    out_file
        .write_all(&output_bytes)
        .map_err(|e| format!("写入输出文件失败: {e}"))?;

    Ok(WatermarkResult {
        input_path,
        output_path,
        cleared_headers,
        removed_backgrounds,
        removed_shapes,
        elapsed_seconds: timer.elapsed().as_secs_f32(),
    })
}

#[tauri::command]
async fn remove_docx_outline(input_path: String, output_path: String) -> Result<OutlineResult, String> {
    tauri::async_runtime::spawn_blocking(move || remove_docx_outline_impl(input_path, output_path))
        .await
        .map_err(|e| format!("后台任务执行失败: {e}"))?
}

fn remove_docx_outline_impl(input_path: String, output_path: String) -> Result<OutlineResult, String> {
    if !input_path.to_lowercase().ends_with(".docx") {
        return Err("输入文件必须是 .docx".to_string());
    }
    if !output_path.to_lowercase().ends_with(".docx") {
        return Err("输出文件必须是 .docx".to_string());
    }

    let timer = Instant::now();
    let source = File::open(&input_path).map_err(|e| format!("无法打开输入文件: {e}"))?;
    let mut archive = ZipArchive::new(source).map_err(|e| format!("DOCX 解析失败: {e}"))?;

    let mut heading_style_ids = if let Ok(mut styles_entry) = archive.by_name("word/styles.xml") {
        let mut styles_data = Vec::new();
        styles_entry
            .read_to_end(&mut styles_data)
            .map_err(|e| format!("读取 styles.xml 失败: {e}"))?;
        let styles_content = String::from_utf8(styles_data)
            .map_err(|e| format!("读取 styles.xml 失败: {e}"))?;
        extract_heading_style_ids(&styles_content)?
    } else {
        HashSet::new()
    };
    for lvl in 1..=9 {
        heading_style_ids.insert(format!("Heading{lvl}"));
    }
    let mut default_paragraph_style_id: Option<String> = None;
    let mut heading_style_formats: HashMap<String, StyleDirectFormat> = HashMap::new();
    if let Ok(mut styles_entry) = archive.by_name("word/styles.xml") {
        let mut styles_data = Vec::new();
        styles_entry
            .read_to_end(&mut styles_data)
            .map_err(|e| format!("读取 styles.xml 失败: {e}"))?;
        let styles_content = String::from_utf8(styles_data)
            .map_err(|e| format!("读取 styles.xml 失败: {e}"))?;
        default_paragraph_style_id = extract_default_paragraph_style_id(&styles_content)?;
        heading_style_formats =
            extract_heading_style_direct_formats(&styles_content, &heading_style_ids)?;
    }

    let mut escaped_style_ids: Vec<String> = heading_style_ids
        .iter()
        .map(|id| regex::escape(id))
        .collect();
    escaped_style_ids.sort();
    let heading_style_ref = if escaped_style_ids.is_empty() {
        Regex::new(r#"$^"#).map_err(|e| format!("正则编译失败: {e}"))?
    } else {
        Regex::new(&format!(
            r#"<w:pStyle\b[^>]*\bw:val="(?:{})"[^>]*/?>"#,
            escaped_style_ids.join("|")
        ))
        .map_err(|e| format!("正则编译失败: {e}"))?
    };

    let outline_lvl = Regex::new(r#"<w:outlineLvl\s+w:val="[0-9]+"\s*/>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let paragraph_block = Regex::new(r#"(?s)<w:p\b[^>]*>.*?</w:p>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let p_open = Regex::new(r#"<w:p\b[^>]*>"#).map_err(|e| format!("正则编译失败: {e}"))?;
    let ppr_block = Regex::new(r#"(?s)<w:pPr\b[^>]*>(.*?)</w:pPr>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let rpr_block = Regex::new(r#"(?s)<w:rPr\b[^>]*>(.*?)</w:rPr>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let run_block = Regex::new(r#"(?s)<w:r\b[^>]*>.*?</w:r>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let run_open = Regex::new(r#"<w:r\b[^>]*>"#).map_err(|e| format!("正则编译失败: {e}"))?;
    let text_run = Regex::new(r#"(?s)<w:t\b[^>]*>.*?</w:t>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let pstyle_tag = Regex::new(r#"<w:pStyle\b[^>]*\bw:val="[^"]+"[^>]*/?>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let pstyle_val = Regex::new(r#"<w:pStyle\b[^>]*\bw:val="([^"]+)"[^>]*/?>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let toc_index_field_simple = Regex::new(
        r#"(?is)<w:fldSimple\b[^>]*w:instr="[^"]*(?:TOC|INDEX|TC|XE)[^"]*"[^>]*>(.*?)</w:fldSimple>"#,
    )
    .map_err(|e| format!("正则编译失败: {e}"))?;
    let toc_index_instr_text = Regex::new(
        r#"(?is)<w:instrText\b[^>]*>[^<]*(?:TOC|INDEX|TC|XE)[^<]*</w:instrText>"#,
    )
    .map_err(|e| format!("正则编译失败: {e}"))?;
    let fld_char = Regex::new(r#"<w:fldChar\b[^>]*/>"#).map_err(|e| format!("正则编译失败: {e}"))?;
    let sdt_pr = Regex::new(r#"(?is)<w:sdtPr\b[^>]*>.*?</w:sdtPr>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let sdt_open_close = Regex::new(r#"</?w:sdt\b[^>]*>"#).map_err(|e| format!("正则编译失败: {e}"))?;
    let sdt_content_open_close = Regex::new(r#"</?w:sdtContent\b[^>]*>"#)
    .map_err(|e| format!("正则编译失败: {e}"))?;

    let mut removed_heading_styles = 0u32;
    let mut replaced_outline_levels = 0u32;
    let mut inserted_outline_levels = 0u32;
    let mut removed_toc_index_blocks = 0u32;
    let mut removed_toc_index_fields = 0u32;
    let mut buffer = Cursor::new(Vec::new());

    {
        let mut zip_writer = ZipWriter::new(&mut buffer);

        for i in 0..archive.len() {
            let mut entry = archive
                .by_index(i)
                .map_err(|e| format!("读取 DOCX 条目失败: {e}"))?;
            let name = entry.name().to_string();
            let mut data = Vec::new();
            entry
                .read_to_end(&mut data)
                .map_err(|e| format!("读取文件数据失败: {e}"))?;

            if name == "word/document.xml" || name == "word/styles.xml" {
                let content = String::from_utf8(data).map_err(|e| format!("读取 {name} 失败: {e}"))?;

                let cleaned = if name == "word/styles.xml" {
                    content
                } else {
                    paragraph_block
                        .replace_all(&content, |caps: &regex::Captures| {
                            let para = caps.get(0).map(|m| m.as_str()).unwrap_or_default();
                            if !heading_style_ref.is_match(para) {
                                return para.to_string();
                            }

                            let Some(style_id) = pstyle_val
                                .captures(para)
                                .and_then(|c| c.get(1))
                                .map(|m| m.as_str().to_string())
                            else {
                                return para.to_string();
                            };
                            if !heading_style_ids.contains(&style_id) {
                                return para.to_string();
                            }

                            removed_heading_styles += 1;
                            let mut updated = para.to_string();

                            if let Some(default_id) = default_paragraph_style_id.as_ref() {
                                updated = pstyle_tag
                                    .replace_all(
                                        &updated,
                                        format!("<w:pStyle w:val=\"{default_id}\"/>").as_str(),
                                    )
                                    .into_owned();
                            } else {
                                updated = pstyle_tag.replace_all(&updated, "").into_owned();
                            }

                            let fmt = heading_style_formats.get(&style_id).cloned().unwrap_or_default();
                            let existing_outline = outline_lvl.find_iter(&updated).count() as u32;
                            replaced_outline_levels += existing_outline;

                            let mut merged_ppr = fmt.ppr_inner;
                            let existing_ppr = ppr_block
                                .captures(&updated)
                                .and_then(|c| c.get(1))
                                .map(|m| m.as_str().to_string())
                                .unwrap_or_default();
                            if !existing_ppr.is_empty() {
                                merged_ppr.push_str(&existing_ppr);
                            }
                            merged_ppr = pstyle_tag.replace_all(&merged_ppr, "").into_owned();
                            merged_ppr = outline_lvl.replace_all(&merged_ppr, "").into_owned();

                            // Keep numbering label appearance by applying heading rPr at paragraph level too.
                            if !fmt.rpr_inner.is_empty() {
                                let existing_para_rpr = rpr_block
                                    .captures(&merged_ppr)
                                    .and_then(|c| c.get(1))
                                    .map(|m| m.as_str().to_string())
                                    .unwrap_or_default();
                                let mut merged_para_rpr = fmt.rpr_inner.clone();
                                if !existing_para_rpr.is_empty() {
                                    merged_para_rpr.push_str(&existing_para_rpr);
                                }
                                if rpr_block.is_match(&merged_ppr) {
                                    merged_ppr = rpr_block
                                        .replacen(
                                            &merged_ppr,
                                            1,
                                            format!("<w:rPr>{merged_para_rpr}</w:rPr>").as_str(),
                                        )
                                        .into_owned();
                                } else {
                                    merged_ppr.push_str(&format!("<w:rPr>{merged_para_rpr}</w:rPr>"));
                                }
                            }

                            merged_ppr.push_str("<w:outlineLvl w:val=\"9\"/>");
                            inserted_outline_levels += 1;

                            if ppr_block.is_match(&updated) {
                                updated = ppr_block
                                    .replacen(
                                        &updated,
                                        1,
                                        format!("<w:pPr>{merged_ppr}</w:pPr>").as_str(),
                                    )
                                    .into_owned();
                            } else {
                                updated = p_open
                                    .replacen(
                                        &updated,
                                        1,
                                        format!("$0<w:pPr>{merged_ppr}</w:pPr>").as_str(),
                                    )
                                    .into_owned();
                            }

                            if !fmt.rpr_inner.is_empty() {
                                updated = run_block
                                    .replace_all(&updated, |run_caps: &regex::Captures| {
                                        let run = run_caps
                                            .get(0)
                                            .map(|m| m.as_str())
                                            .unwrap_or_default();
                                        if !text_run.is_match(run) {
                                            return run.to_string();
                                        }

                                        if rpr_block.is_match(run) {
                                            let existing_rpr = rpr_block
                                                .captures(run)
                                                .and_then(|c| c.get(1))
                                                .map(|m| m.as_str())
                                                .unwrap_or_default();
                                            let merged_rpr = format!("{}{}", fmt.rpr_inner, existing_rpr);
                                            rpr_block
                                                .replacen(
                                                    run,
                                                    1,
                                                    format!("<w:rPr>{merged_rpr}</w:rPr>").as_str(),
                                                )
                                                .into_owned()
                                        } else {
                                            run_open
                                                .replacen(
                                                    run,
                                                    1,
                                                    format!("$0<w:rPr>{}</w:rPr>", fmt.rpr_inner).as_str(),
                                                )
                                                .into_owned()
                                        }
                                    })
                                    .into_owned();
                            }

                            updated
                        })
                        .into_owned()
                };

                let cleaned = if name == "word/document.xml" {
                    removed_toc_index_fields += toc_index_field_simple.find_iter(&cleaned).count() as u32;
                    let cleaned = toc_index_field_simple.replace_all(&cleaned, "$1");

                    removed_toc_index_fields += toc_index_instr_text.find_iter(&cleaned).count() as u32;
                    let cleaned = toc_index_instr_text.replace_all(&cleaned, "");

                    removed_toc_index_fields += fld_char.find_iter(&cleaned).count() as u32;
                    let cleaned = fld_char.replace_all(&cleaned, "");

                    removed_toc_index_blocks += sdt_pr.find_iter(&cleaned).count() as u32;
                    let cleaned = sdt_pr.replace_all(&cleaned, "");

                    removed_toc_index_blocks += sdt_open_close.find_iter(&cleaned).count() as u32;
                    let cleaned = sdt_open_close.replace_all(&cleaned, "");

                    removed_toc_index_blocks += sdt_content_open_close.find_iter(&cleaned).count() as u32;
                    sdt_content_open_close.replace_all(&cleaned, "").into_owned()
                } else {
                    cleaned
                };

                data = cleaned.into_bytes();
            }

            zip_writer
                .start_file(
                    name,
                    zip::write::FileOptions::default()
                        .compression_method(CompressionMethod::Deflated)
                        .compression_level(Some(6)),
                )
                .map_err(|e| format!("写入 ZIP 条目失败: {e}"))?;
            zip_writer
                .write_all(&data)
                .map_err(|e| format!("写入去大纲结果失败: {e}"))?;
        }

        zip_writer
            .finish()
            .map_err(|e| format!("结束 ZIP 写入失败: {e}"))?;
    }

    let output_bytes = buffer.into_inner();
    let mut out_file = File::create(&output_path).map_err(|e| format!("创建输出文件失败: {e}"))?;
    out_file
        .write_all(&output_bytes)
        .map_err(|e| format!("写入输出文件失败: {e}"))?;

    Ok(OutlineResult {
        input_path,
        output_path,
        removed_heading_styles,
        replaced_outline_levels,
        inserted_outline_levels,
        removed_toc_index_blocks,
        removed_toc_index_fields,
        elapsed_seconds: timer.elapsed().as_secs_f32(),
    })
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            compress_docx,
            unlock_docx,
            remove_docx_watermark,
            remove_docx_outline
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
