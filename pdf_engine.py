import os
import fitz  # PyMuPDF
import win32com.client
from pythoncom import CoInitialize, CoUninitialize

class PdfEngine:
    def __init__(self):
        CoInitialize()
        try:
            self.app = win32com.client.Dispatch("PowerPoint.Application")
        except Exception as e:
            CoUninitialize()
            raise e
        
    def process_presentation(self, input_path, output_path, temp_folder, watermark_engine, resolution_mode="1080p", progress_callback=None):
        input_path = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)
        
        doc = None
        new_presentation = None
        
        try:
            doc = fitz.open(input_path)
            total_pages = len(doc)
            image_paths = []
            
            target_height = 1080 if resolution_mode == "1080p" else 720
            
            # Detect aspect ratio from the first page
            first_page_rect = doc[0].rect
            ratio = first_page_rect.width / first_page_rect.height
            midpoint = (16/9 + 4/3) / 2
            is_16_9 = ratio >= midpoint
            
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
                progress_callback("正在将处理好的图片拼合成全新 PPT...")
                
            new_presentation = self.app.Presentations.Add(WithWindow=False)
            
            if is_16_9:
                new_presentation.PageSetup.SlideWidth = 960
                new_presentation.PageSetup.SlideHeight = 540
            else:
                new_presentation.PageSetup.SlideWidth = 720
                new_presentation.PageSetup.SlideHeight = 540
                
            ppLayoutBlank = 12
            
            for i, img_path in enumerate(image_paths):
                new_slide = new_presentation.Slides.Add(i + 1, ppLayoutBlank) 
                new_slide.Shapes.AddPicture(
                    img_path, 
                    LinkToFile=0, 
                    SaveWithDocument=-1, 
                    Left=0, 
                    Top=0, 
                    Width=new_presentation.PageSetup.SlideWidth, 
                    Height=new_presentation.PageSetup.SlideHeight
                )
                
            new_presentation.SaveAs(output_path)
            new_presentation.Close()
            new_presentation = None
            
        finally:
            if doc is not None:
                try: doc.close()
                except: pass
            if new_presentation is not None:
                try: new_presentation.Close()
                except: pass

    def quit(self):
        try:
            self.app.Quit()
        except:
            pass
        CoUninitialize()
