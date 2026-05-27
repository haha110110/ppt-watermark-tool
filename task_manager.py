import threading
import tempfile
import os
import shutil
from ppt_engine import PptEngine
from pdf_engine import PdfEngine
from watermark_engine import WatermarkEngine

class TaskManager:
    def __init__(self, ui_callback, finish_callback):
        self.ui_callback = ui_callback
        self.finish_callback = finish_callback
        self.is_running = False
        
    def start_processing(self, files, output_dir, watermark_path, scale_percent, resolution_mode, quality):
        if self.is_running:
            return
        self.is_running = True
        
        # Run in a background thread to prevent UI freezing
        thread = threading.Thread(target=self._run, args=(files, output_dir, watermark_path, scale_percent, resolution_mode, quality))
        thread.daemon = True
        thread.start()
        
    def _run(self, files, output_dir, watermark_path, scale_percent, resolution_mode, quality):
        total_files = len(files)
        temp_dir = tempfile.mkdtemp()
        
        ppt_engine = None
        pdf_engine = None
        
        try:
            # Initialize common engine
            wm_engine = WatermarkEngine(watermark_path, scale_percent=scale_percent, quality=quality)
            
            for i, input_file in enumerate(files):
                filename = os.path.basename(input_file)
                name, ext = os.path.splitext(filename)
                self.ui_callback(f"({i+1}/{total_files}) 开始处理: {filename}", i, total_files)
                
                # Force output to be .pptx
                out_filename = f"{name}.pptx"
                out_path = os.path.join(output_dir, out_filename)
                
                # Check for output naming conflict
                if os.path.abspath(out_path) == os.path.abspath(input_file):
                    # Append _wm to prevent overwrite
                    out_path = os.path.join(output_dir, f"{name}_wm.pptx")
                    
                # Clean temp dir for this run
                for f in os.listdir(temp_dir):
                    os.remove(os.path.join(temp_dir, f))
                    
                # Setup localized progress callback for the engine
                def progress_cb(msg):
                    self.ui_callback(f"({i+1}/{total_files}) {filename} - {msg}", i, total_files)
                    
                try:
                    ext = os.path.splitext(filename)[1].lower()
                    
                    # Dynamically instantiate and route to the correct engine
                    if ext == ".pdf":
                        if pdf_engine is None:
                            pdf_engine = PdfEngine()
                        engine = pdf_engine
                    else:
                        if ppt_engine is None:
                            ppt_engine = PptEngine()
                        engine = ppt_engine
                        
                    engine.process_presentation(
                        input_file, 
                        out_path, 
                        temp_dir, 
                        wm_engine, 
                        resolution_mode=resolution_mode,
                        progress_callback=progress_cb
                    )
                except Exception as e:
                    self.ui_callback(f"错误 - {filename}: {str(e)}", i, total_files)
                    
            self.ui_callback("全部处理完成！", total_files, total_files)
            
        except Exception as e:
            self.ui_callback(f"致命错误: {str(e)}", 0, total_files)
        finally:
            # Cleanup COM objects and temp files
            if ppt_engine:
                ppt_engine.quit()
            if pdf_engine:
                pdf_engine.quit()
            shutil.rmtree(temp_dir, ignore_errors=True)
            self.is_running = False
            self.finish_callback()
