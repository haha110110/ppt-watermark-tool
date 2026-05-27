import os
import fitz  # PyMuPDF

class PdfEngine:
    def __init__(self):
        pass
        
    def process_presentation(self, input_path, output_path, temp_folder, watermark_engine, resolution_mode="1080p", progress_callback=None):
        input_path = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)
        
        doc = None
        new_doc = None
        
        try:
            doc = fitz.open(input_path)
            total_pages = len(doc)
            image_paths = []
            
            target_height = 1080 if resolution_mode == "1080p" else 720
            
            for i in range(total_pages):
                if progress_callback:
                    progress_callback(f"导出并处理 PDF 第 {i+1}/{total_pages} 页...")
                    
                page = doc[i]
                
                # Calculate zoom matrix to hit target_height
                zoom = target_height / page.rect.height
                mat = fitz.Matrix(zoom, zoom)
                
                # Render page to an image
                pix = page.get_pixmap(matrix=mat, alpha=False)
                
                img_path = os.path.join(temp_folder, f"page_{i+1}.png")
                pix.save(img_path)
                
                # Apply watermark directly to the exported PNG and get back the compressed JPG path
                final_img_path = watermark_engine.apply_watermark(img_path)
                image_paths.append(final_img_path)
                
            doc.close()
            doc = None
            
            if progress_callback:
                progress_callback("正在拼合全新的 PDF...")
                
            new_doc = fitz.open()
            
            for i, img_path in enumerate(image_paths):
                # Open the JPG image as a document to easily insert it into the new PDF
                img_doc = fitz.open(img_path)
                # Convert the image document to a PDF document bytes, then open as pdf
                pdfbytes = img_doc.convert_to_pdf()
                img_pdf = fitz.open("pdf", pdfbytes)
                
                # Insert the single page into our new document
                new_doc.insert_pdf(img_pdf)
                img_doc.close()
                img_pdf.close()
                
            # Save the new PDF with garbage collection and deflation for optimization
            new_doc.save(output_path, garbage=3, deflate=True)
            new_doc.close()
            new_doc = None
            
        finally:
            if doc is not None:
                try: doc.close()
                except: pass
            if new_doc is not None:
                try: new_doc.close()
                except: pass

    def quit(self):
        # PyMuPDF doesn't leave ghost processes, so no cleanup needed here
        pass
