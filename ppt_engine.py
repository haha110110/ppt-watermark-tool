import os
import win32com.client
from pythoncom import CoInitialize, CoUninitialize

class PptEngine:
    def __init__(self):
        # Essential for COM objects running in a background thread
        CoInitialize()
        try:
            self.app = win32com.client.Dispatch("PowerPoint.Application")
            # Some versions of PowerPoint fail if run completely hidden, but we'll try to keep it minimal
            # self.app.Visible = True 
        except Exception as e:
            CoUninitialize()
            raise e

    def _get_target_resolution(self, width, height, resolution_mode="1080p"):
        ratio = width / height
        target_16_9 = 16 / 9  # ~1.777
        target_4_3 = 4 / 3    # ~1.333
        
        # Midpoint to decide which it is closer to
        midpoint = (target_16_9 + target_4_3) / 2
        
        # Decide target height
        target_height = 1080 if resolution_mode == "1080p" else 720
        
        # If it's closer to 16:9
        if ratio >= midpoint:
            target_width = 1920 if target_height == 1080 else 1280
            return target_width, target_height, True
        else:
            target_width = 1440 if target_height == 1080 else 960
            return target_width, target_height, False
            
    def process_presentation(self, input_path, output_path, temp_folder, watermark_engine, resolution_mode="1080p", output_format="pptx", progress_callback=None):
        input_path = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)
        
        presentation = None
        new_presentation = None
        
        try:
            # 1. Open original presentation
            # WithWindow=False ensures it runs in the background without stealing focus
            presentation = self.app.Presentations.Open(input_path, ReadOnly=True, WithWindow=False)
            
            orig_width = presentation.PageSetup.SlideWidth
            orig_height = presentation.PageSetup.SlideHeight
            
            target_px_width, target_px_height, is_16_9 = self._get_target_resolution(orig_width, orig_height, resolution_mode)
            
            total_slides = presentation.Slides.Count
            image_paths = []
            
            # 2. Export each slide and watermark
            for i in range(1, total_slides + 1):
                if progress_callback:
                    progress_callback(f"导出并处理幻灯片 {i}/{total_slides}...")
                
                slide = presentation.Slides(i)
                img_path = os.path.join(temp_folder, f"slide_{i}.png")
                
                # Export to exact target size. PowerPoint will stretch if the native ratio differs.
                slide.Export(img_path, "PNG", target_px_width, target_px_height)
                
                # Apply watermark directly to the exported PNG and get back the compressed JPG path
                final_img_path = watermark_engine.apply_watermark(img_path)
                image_paths.append(final_img_path)
                
            # Close original
            presentation.Close()
            presentation = None
            
            # 3. Create new presentation/pdf for reassembly
            if output_format == "pdf":
                import fitz
                if progress_callback:
                    progress_callback("正在将处理好的图片拼合成全新 PDF...")
                new_doc = fitz.open()
                for img_path in image_paths:
                    img_doc = fitz.open(img_path)
                    pdfbytes = img_doc.convert_to_pdf()
                    img_pdf = fitz.open("pdf", pdfbytes)
                    new_doc.insert_pdf(img_pdf)
                    img_doc.close()
                    img_pdf.close()
                new_doc.save(output_path, garbage=3, deflate=True)
                new_doc.close()
            else:
                if progress_callback:
                    progress_callback("正在拼合新的PPT...")
                    
                new_presentation = self.app.Presentations.Add(WithWindow=False)
                
                # Set to standard 16:9 or 4:3 in points (1 inch = 72 points)
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
                    
                # Save new presentation
                new_presentation.SaveAs(output_path)
                new_presentation.Close()
                new_presentation = None
            
        finally:
            if presentation is not None:
                try: presentation.Close()
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
