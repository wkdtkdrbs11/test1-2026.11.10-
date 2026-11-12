import subprocess
import os
import shutil
import zipfile
from PIL import Image

BASE_DIR = r"C:\Users\ABC\.gemini\antigravity\scratch\portfolio-profile"
OUT_DIR = os.path.join(BASE_DIR, "포트폴리오_슬라이드_JPG")
ZIP_PATH = os.path.join(BASE_DIR, "마케터_장상균_포트폴리오_JPG전체.zip")
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

os.makedirs(OUT_DIR, exist_ok=True)

titles = [
    "01_표지",
    "02_이력사항",
    "03_목차_INDEX",
    "04_Project1_문제정의_피벗",
    "05_Project1_솔루션_성과",
    "06_Project2_운영총괄_포지셔닝",
    "07_Project2_캠페인성과_지표",
    "08_Project3_문제정의_가설검증",
    "09_Contact_마무리"
]

print("Starting high-resolution 1920x1080 16:9 JPG capture for all 9 slides...")

jpg_files = []

for idx in range(1, 10):
    url = f"http://localhost:8080/?capture={idx}"
    temp_png = os.path.join(OUT_DIR, f"temp_{idx}.png")
    out_jpg_name = f"장상균_포트폴리오_{titles[idx-1]}.jpg"
    out_jpg_path = os.path.join(OUT_DIR, out_jpg_name)
    
    cmd = [
        CHROME_PATH,
        "--headless=new",
        "--disable-gpu",
        "--window-size=1920,1080",
        f"--screenshot={temp_png}",
        url
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    if os.path.exists(temp_png):
        with Image.open(temp_png) as im:
            rgb_im = im.convert('RGB')
            # Save as 1920x1080 high quality JPG
            rgb_im.save(out_jpg_path, 'JPEG', quality=95, optimize=True)
        os.remove(temp_png)
        jpg_files.append(out_jpg_path)
        print(f"Captured: {out_jpg_name} (1920x1080, 16:9)")
    else:
        print(f"Failed to capture slide {idx}")

# Also create a zip archive for one-click download
with zipfile.ZipFile(ZIP_PATH, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for f in jpg_files:
        zipf.write(f, os.path.basename(f))

print(f"SUCCESS: All 9 slides exported to {OUT_DIR}")
print(f"SUCCESS: ZIP package created at {ZIP_PATH}")
