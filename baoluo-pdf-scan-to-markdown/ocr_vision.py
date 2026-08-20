#!/usr/bin/env python3
"""macOS 原生 Vision OCR：PDF 扫描件 → 带坐标的 JSON（供后处理清理）"""
import sys, json, re
import Quartz
import Vision
from Foundation import NSURL

def page_to_cgimage(page, scale=2.0):
    bounds = page.boundsForBox_(Quartz.kPDFDisplayBoxCropBox)
    w, h = bounds.size.width, bounds.size.height
    pw, ph = int(w * scale), int(h * scale)
    cs = Quartz.CGColorSpaceCreateDeviceRGB()
    ctx = Quartz.CGBitmapContextCreate(None, pw, ph, 8, 0, cs,
                                       Quartz.kCGImageAlphaPremultipliedLast)
    Quartz.CGContextSetRGBFillColor(ctx, 1, 1, 1, 1)
    Quartz.CGContextFillRect(ctx, Quartz.CGRectMake(0, 0, pw, ph))
    Quartz.CGContextScaleCTM(ctx, scale, scale)
    Quartz.CGContextDrawPDFPage(ctx, page.pageRef())
    return Quartz.CGBitmapContextCreateImage(ctx)

def recognize_items(cgimage):
    """返回 [(text, x, y, w, h)]，归一化坐标，原点左下，y 向上"""
    handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(cgimage, None)
    req = Vision.VNRecognizeTextRequest.alloc().init()
    req.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    req.setRecognitionLanguages_(['zh-Hans', 'en-US'])
    req.setUsesLanguageCorrection_(True)
    handler.performRequests_error_([req], None)
    results = req.results() or []
    items = []
    for obs in results:
        text = obs.topCandidates_(1)[0].string()
        b = obs.boundingBox()
        items.append({"t": text,
                      "x": round(b.origin.x, 4),
                      "y": round(b.origin.y, 4),
                      "w": round(b.size.width, 4),
                      "h": round(b.size.height, 4)})
    return items

def ocr_pdf_json(pdf_path, out_json, page_range=None):
    from Quartz import PDFDocument
    doc = PDFDocument.alloc().initWithURL_(NSURL.fileURLWithPath_(pdf_path))
    if doc is None:
        print("无法打开 PDF"); sys.exit(1)
    total = doc.pageCount()
    start, end = 1, total
    if page_range:
        m = re.match(r"^(\d+)(?:-(\d+))?$", page_range)
        if not m:
            print(f"页码范围格式错误: {page_range}（示例: 27 或 27-29）"); sys.exit(1)
        start = int(m.group(1))
        end = int(m.group(2)) if m.group(2) else start
    print(f"总页数: {total}, 处理 {start}-{end}", file=sys.stderr)
    pages = []
    for i in range(start, min(end, total) + 1):
        page = doc.pageAtIndex_(i - 1)
        img = page_to_cgimage(page)
        items = recognize_items(img)
        pages.append({"page": i, "items": items})
        if i % 25 == 0:
            print(f"进度: {i}/{total}", file=sys.stderr)
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(pages, f, ensure_ascii=False)
    print(f"完成: {out_json}", file=sys.stderr)

if __name__ == "__main__":
    pdf = sys.argv[1]
    out = sys.argv[2]
    rng = sys.argv[3] if len(sys.argv) > 3 else None
    ocr_pdf_json(pdf, out, rng)
