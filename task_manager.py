import threading
import tempfile
import os
import shutil
from ppt_engine import PptEngine
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
        engine = None
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Initialize engines
            wm_engine = WatermarkEngine(watermark_path, scale_percent=scale_percent, quality=quality)
            engine = PptEngine()
            
            for i, input_file in enumerate(files):
                filename = os.path.basename(input_file)
                self.ui_callback(f"({i+1}/{total_files}) 开始处理: {filename}", i, total_files)
                
                # Check for output naming conflict
                out_path = os.path.join(output_dir, filename)
                if os.path.abspath(out_path) == os.path.abspath(input_file):
                    # Append _wm to prevent overwrite
                    name, ext = os.path.splitext(filename)
                    out_path = os.path.join(output_dir, f"{name}_wm{ext}")
                    
                # Clean temp dir for this run
                for f in os.listdir(temp_dir):
                    os.remove(os.path.join(temp_dir, f))
                    
                # Setup localized progress callback for the engine
                def progress_cb(msg):
                    self.ui_callback(f"({i+1}/{total_files}) {filename} - {msg}", i, total_files)
                    
                try:
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
            if engine:
                engine.quit()
            shutil.rmtree(temp_dir, ignore_errors=True)
            self.is_running = False
            self.finish_callback()
