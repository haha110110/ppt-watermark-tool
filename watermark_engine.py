import os
from PIL import Image

class WatermarkEngine:
    def __init__(self, watermark_path, scale_percent=15, padding=20, quality=80):
        self.watermark_path = watermark_path
        self.scale_percent = scale_percent / 100.0
        self.padding = padding
        self.quality = quality
        self.watermark_image = None
        
        if self.watermark_path and os.path.exists(self.watermark_path):
            # Open and convert to RGBA for transparency support
            self.watermark_image = Image.open(self.watermark_path).convert("RGBA")
            
    def apply_watermark(self, image_path):
        with Image.open(image_path) as base_image:
            # Convert base to RGBA just in case
            base_image = base_image.convert("RGBA")
            bg_width, bg_height = base_image.size
            
            # 1. Apply watermark if it exists
            if self.watermark_image:
                wm_width, wm_height = self.watermark_image.size
                ratio = wm_width / wm_height
                
                new_wm_height = int(bg_height * self.scale_percent)
                new_wm_width = int(new_wm_height * ratio)
                
                # Resize watermark (using LANCZOS for high quality)
                resized_wm = self.watermark_image.resize((new_wm_width, new_wm_height), Image.Resampling.LANCZOS)
                
                # Calculate position (Bottom Right)
                x = bg_width - new_wm_width - self.padding
                y = bg_height - new_wm_height - self.padding
                
                # Paste watermark using its own alpha channel as the mask
                base_image.paste(resized_wm, (x, y), resized_wm)
            
            # 2. Prevent transparent backgrounds from turning black when saving as JPG
            # We create a pure white background image
            white_bg = Image.new("RGBA", base_image.size, "WHITE")
            # Composite the (possibly transparent) slide onto the white background
            final_image = Image.alpha_composite(white_bg, base_image)
            
            # 3. Convert to RGB and save as JPG to compress size
            final_rgb = final_image.convert("RGB")
            
            # Change extension to .jpg
            jpg_path = image_path.rsplit(".", 1)[0] + ".jpg"
            final_rgb.save(jpg_path, "JPEG", quality=self.quality)
            
            # Return the new jpg path so the PPT engine knows what to insert
            return jpg_path
