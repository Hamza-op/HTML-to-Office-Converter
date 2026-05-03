import fitz
from typing import List, Callable, Optional, Tuple
from scene_graph import (
    Group, Image, Page, Table, TableCell, TextBlock, TextParagraph, TextRun,
    UnknownRegion, VectorPath,
)

def _as_fitz_rect(bbox):
    if not bbox:
        return fitz.Rect(0, 0, 0, 0)
    if isinstance(bbox, fitz.Rect):
        return bbox
    if isinstance(bbox, (list, tuple)) and len(bbox) == 4:
        return fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])
    return fitz.Rect(0, 0, 0, 0)

def _float_to_rgb(color):
    if not color:
        return None
    if isinstance(color, int):
        b = color & 255
        g = (color >> 8) & 255
        r = (color >> 16) & 255
        return (r, g, b)
    if isinstance(color, (list, tuple)) and len(color) >= 3:
        vals = []
        for c in color[:3]:
            c = 0.0 if c is None else float(c)
            c = max(0.0, min(1.0, c))
            vals.append(int(round(c * 255)))
        return tuple(vals)
    return None

def extract_pdf_to_scene_graph(
    pdf_path: str,
    on_status: Optional[Callable[[str], None]] = None
) -> Tuple[List[Page], List[str]]:
    status = on_status or (lambda _: None)
    status("Extracting PDF to scene graph...")
    
    pages: List[Page] = []
    warnings: List[str] = []
    
    with fitz.open(pdf_path) as doc:
        for i in range(len(doc)):
            status(f"Parsing PDF page {i + 1}/{len(doc)}...")
            fitz_page = doc[i]
            page_rect = fitz_page.rect
            
            page = Page(page_rect.width, page_rect.height, i + 1)
            try:
                pix = fitz_page.get_pixmap(dpi=150)
                page.fallback_bytes = pix.tobytes("png")
            except Exception as e:
                warnings.append(f"Page {i+1}: Snapshot caching failed ({e})")
            
            # Extract tables
            extracted_tables = []
            try:
                finder = fitz_page.find_tables()
                page_area = max(page_rect.get_area(), 1.0)
                for t in finder.tables:
                    try:
                        tb = _as_fitz_rect(t.bbox)
                        if tb.width <= 0 or tb.height <= 0:
                            continue
                        if tb.get_area() / page_area > 0.65:
                            continue
                        extracted = t.extract() or []
                        if extracted:
                            extracted_tables.append((tb, extracted))
                    except Exception as e:
                        warnings.append(f"Page {i+1}: Failed to parse a table ({e})")
                        continue
            except Exception as e:
                warnings.append(f"Page {i+1}: Table detection failed ({e})")
                
            for table_rect, rows in extracted_tables:
                row_count = len(rows)
                col_count = max(len(r) for r in rows) if row_count else 0
                if row_count > 0 and col_count > 0:
                    table = Table(
                        rows=row_count, cols=col_count,
                        bbox=(table_rect.x0, table_rect.y0, table_rect.x1, table_rect.y1),
                        provenance="fitz_tables"
                    )
                    for ridx, row in enumerate(rows):
                        for cidx in range(col_count):
                            txt = row[cidx] if cidx < len(row) and row[cidx] is not None else ""
                            cell = TableCell(ridx, cidx)
                            cell.text_block = TextBlock(bbox=(0,0,0,0)) # Dummy bbox
                            paragraph = TextParagraph()
                            paragraph.runs.append(TextRun(text=str(txt), bold=(ridx==0)))
                            cell.text_block.paragraphs.append(paragraph)
                            table.cells.append(cell)
                    page.root.children.append(table)
            
            # Extract drawings
            try:
                drawings = fitz_page.get_drawings()
                for path in drawings:
                    rect = _as_fitz_rect(path.get("rect"))
                    stroke_rgb = _float_to_rgb(path.get("color"))
                    fill_rgb = _float_to_rgb(path.get("fill"))
                    width_pt = path.get("width", 0.75) or 0.75
                    
                    if not stroke_rgb and not fill_rgb:
                        continue
                        
                    items = path.get("items", [])
                    if not items:
                        continue
                        
                    vp = VectorPath(
                        bbox=(rect.x0, rect.y0, rect.x1, rect.y1),
                        provenance="fitz_drawings"
                    )
                    vp.fill_color = fill_rgb
                    vp.stroke_color = stroke_rgb
                    vp.stroke_width = width_pt
                    
                    # Simply store the raw fitz drawing ops for now
                    # We map them to PPTX constructs in the writer
                    for item in items:
                        vp.segments.append({
                            "type": item[0],
                            "args": item[1:]
                        })
                    
                    page.root.children.append(vp)
            except Exception as e:
                warnings.append(f"Page {i+1}: Drawing extraction failed ({e})")

            # High-res images
            high_res_images = []
            for img in fitz_page.get_images(full=True):
                xref = img[0]
                try:
                    extracted = doc.extract_image(xref)
                    if not extracted: continue
                    image_bytes = extracted.get("image")
                    if not image_bytes: continue
                    for rect in fitz_page.get_image_rects(img):
                        if rect.width > 0 and rect.height > 0:
                            high_res_images.append({"rect": rect, "bytes": image_bytes})
                except Exception as e:
                    warnings.append(f"Page {i+1}: High-res image extraction failed ({e})")
                    continue

            # Text blocks and block-level images
            text_dict = fitz_page.get_text("dict")
            blocks = text_dict.get("blocks", [])
            
            used_h_imgs = set()
            for block in blocks:
                if block.get("type") == 1:
                    brect = _as_fitz_rect(block.get("bbox"))
                    for idx, h_img in enumerate(high_res_images):
                        h_rect = h_img["rect"]
                        if brect.intersect(h_rect).get_area() > 0.5 * brect.get_area():
                            used_h_imgs.add(idx)
                            
            for idx, h_img in enumerate(high_res_images):
                if idx not in used_h_imgs:
                    rect = h_img["rect"]
                    image_node = Image(
                        image_bytes=h_img["bytes"],
                        bbox=(rect.x0, rect.y0, rect.x1, rect.y1),
                        provenance="fitz_high_res_image"
                    )
                    page.root.children.append(image_node)

            for block in blocks:
                try:
                    brect = _as_fitz_rect(block.get("bbox"))
                    
                    skip = False
                    for table_rect, _ in extracted_tables:
                        if brect.intersects(table_rect):
                            if brect.intersect(table_rect).get_area() > 0.5 * brect.get_area():
                                skip = True
                                break
                    if skip: continue

                    if block.get("type") == 1:
                        image_bytes = block.get("image")
                        best_idx, max_overlap = -1, 0
                        for idx, h_img in enumerate(high_res_images):
                            overlap = brect.intersect(h_img["rect"]).get_area()
                            if overlap > max_overlap and overlap > 0.5 * brect.get_area():
                                max_overlap = overlap
                                best_idx = idx
                                
                        if best_idx != -1:
                            image_bytes = high_res_images[best_idx]["bytes"]
                            
                        if image_bytes:
                            page.root.children.append(Image(
                                image_bytes=image_bytes,
                                bbox=(brect.x0, brect.y0, brect.x1, brect.y1),
                                provenance="fitz_block_image"
                            ))
                    elif block.get("type") == 0:
                        lines = block.get("lines", [])
                        if not lines: continue
                        
                        tb = TextBlock(
                            bbox=(brect.x0, brect.y0, brect.x1, brect.y1),
                            provenance="fitz_text"
                        )
                        
                        for line in lines:
                            paragraph = TextParagraph()
                            for span in line.get("spans", []):
                                text = span.get("text", "")
                                if not text: continue
                                
                                color = _float_to_rgb(span.get("color"))
                                flags = int(span.get("flags", 0) or 0)
                                
                                tr = TextRun(
                                    text=text,
                                    font_name=span.get("font"),
                                    font_size=span.get("size"),
                                    color=color,
                                    bold=bool(flags & 16),
                                    italic=bool(flags & 2)
                                )
                                paragraph.runs.append(tr)
                                
                            if paragraph.runs:
                                tb.paragraphs.append(paragraph)
                                
                        if tb.paragraphs:
                            page.root.children.append(tb)
                except Exception as e:
                    warnings.append(f"Page {i+1}: Block extraction failed ({e}), using region fallback.")
                    try:
                        brect = _as_fitz_rect(block.get("bbox", [0,0,0,0]))
                        if brect.width > 0 and brect.height > 0:
                            pix = fitz_page.get_pixmap(clip=brect, dpi=150)
                            fallback = UnknownRegion(
                                image_fallback=pix.tobytes("png"),
                                bbox=(brect.x0, brect.y0, brect.x1, brect.y1),
                                provenance="region_raster_fallback"
                            )
                            page.root.children.append(fallback)
                    except Exception as fallback_err:
                        warnings.append(f"Page {i+1}: Region fallback generation failed ({fallback_err})")

            # Group consecutive VectorPaths
            new_children = []
            current_group = None
            
            for child in page.root.children:
                if isinstance(child, VectorPath):
                    if current_group is None:
                        current_group = Group(provenance="consecutive_vectors")
                        current_group.children.append(child)
                    else:
                        current_group.children.append(child)
                else:
                    if current_group is not None:
                        if len(current_group.children) > 1:
                            new_children.append(current_group)
                        else:
                            new_children.append(current_group.children[0])
                        current_group = None
                    new_children.append(child)
                    
            if current_group is not None:
                if len(current_group.children) > 1:
                    new_children.append(current_group)
                else:
                    new_children.append(current_group.children[0])
                    
            page.root.children = new_children

            if len(page.root.children) == 0:
                try:
                    if page.fallback_bytes:
                        image_bytes = page.fallback_bytes
                    else:
                        pix = fitz_page.get_pixmap(dpi=150)
                        image_bytes = pix.tobytes("png")
                    fallback = UnknownRegion(
                        image_fallback=image_bytes,
                        bbox=(page_rect.x0, page_rect.y0, page_rect.x1, page_rect.y1),
                        provenance="raster_fallback"
                    )
                    page.root.children.append(fallback)
                    warnings.append(f"Page {i+1}: Used full-page raster fallback")
                except Exception as e:
                    warnings.append(f"Page {i+1}: Fallback generation failed ({e})")

            pages.append(page)
            
    return pages, warnings
