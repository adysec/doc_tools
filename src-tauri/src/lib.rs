use image::codecs::jpeg::JpegEncoder;
use image::codecs::png::PngEncoder;
use image::{ColorType, ImageEncoder, ImageReader};
use regex::Regex;
use serde::Serialize;
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
    elapsed_seconds: f32,
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

    let heading_style = Regex::new(r#"<w:pStyle\s+w:val="Heading[0-9]+"\s*/>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let outline_lvl = Regex::new(r#"<w:outlineLvl\s+w:val="[0-9]+"\s*/>"#)
        .map_err(|e| format!("正则编译失败: {e}"))?;
    let ppr_open = Regex::new(r#"(<w:pPr>)"#).map_err(|e| format!("正则编译失败: {e}"))?;

    let mut removed_heading_styles = 0u32;
    let mut replaced_outline_levels = 0u32;
    let mut inserted_outline_levels = 0u32;
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

            if name == "word/document.xml" {
                let content = String::from_utf8(data)
                    .map_err(|e| format!("读取 document.xml 失败: {e}"))?;

                removed_heading_styles += heading_style.find_iter(&content).count() as u32;
                let cleaned = heading_style.replace_all(&content, "");

                replaced_outline_levels += outline_lvl.find_iter(&cleaned).count() as u32;
                let cleaned = outline_lvl.replace_all(&cleaned, "<w:outlineLvl w:val=\"9\"/>");

                inserted_outline_levels += ppr_open.find_iter(&cleaned).count() as u32;
                let cleaned = ppr_open.replace_all(&cleaned, "$1<w:outlineLvl w:val=\"9\"/>");

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
