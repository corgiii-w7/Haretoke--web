#!/usr/bin/env python3
"""
劇団ハレトケ 画像最適化スクリプト
- フォルダ内の大きい画像（JPEG / PNG）を検出し、Web表示に適したサイズ（長辺最大1400px・高品質圧縮）に変換します。
- 元のオリジナル画像は自動的に original_images/ フォルダに安全にバックアップされます。
- 最適化後の方がサイズが大きくなる場合は元のファイルを保持します。
"""

import os
import subprocess
import glob
import shutil

# 対象とする最大長辺ピクセル数（Web表示・スマホ・PC拡大表示に十分な高画質）
MAX_DIMENSION = 1400
# JPEG圧縮品質 (0-100)
JPEG_QUALITY = 80

def get_dimensions(filepath):
    """sips を使って画像の幅・高さを取得"""
    try:
        w_out = subprocess.check_output(["sips", "-g", "pixelWidth", filepath], text=True)
        h_out = subprocess.check_output(["sips", "-g", "pixelHeight", filepath], text=True)
        w = int(w_out.strip().split()[-1])
        h = int(h_out.strip().split()[-1])
        return w, h
    except Exception as e:
        return 0, 0

def optimize_image(src_path, dest_path=None, max_dim=MAX_DIMENSION, quality=JPEG_QUALITY):
    if dest_path is None:
        dest_path = src_path
    
    orig_size = os.path.getsize(src_path)
    w, h = get_dimensions(src_path)
    
    # 既に十分小さく、リサイズも不要ならスキップ
    if orig_size < 350 * 1024 and max(w, h) <= max_dim and src_path == dest_path:
        print(f"⏩ スキップ (既に最適サイズ): {os.path.basename(src_path)} ({orig_size // 1024} KB)")
        return
    
    temp_dest = dest_path + ".tmp.jpg"
    cmd = [
        "sips",
        "-s", "format", "jpeg",
        "-s", "formatOptions", str(quality),
        "-Z", str(max_dim),
        src_path,
        "--out", temp_dest
    ]
    
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        new_size = os.path.getsize(temp_dest)
        
        # 最適化後の方が小さい場合のみ置き換え
        if new_size < orig_size:
            os.replace(temp_dest, dest_path)
            reduction = 100 - (new_size * 100 // orig_size)
            print(f"✅ 最適化完了: {os.path.basename(src_path)} ({orig_size // 1024} KB -> {new_size // 1024} KB, -{reduction}%)")
        else:
            if os.path.exists(temp_dest):
                os.remove(temp_dest)
            print(f"⏩ スキップ (元画像の方が軽量): {os.path.basename(src_path)} ({orig_size // 1024} KB)")
    except Exception as e:
        if os.path.exists(temp_dest):
            os.remove(temp_dest)
        print(f"❌ エラー ({src_path}): {e}")

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)
    
    print("========================================")
    print("🎭 劇団ハレトケ 画像最適化ツール")
    print(f"設定: 最大解像度={MAX_DIMENSION}px, JPEG品質={JPEG_QUALITY}%")
    print("========================================")
    
    patterns = ["*.jpg", "*.jpeg", "*.JPG", "*.JPEG", "*.png", "*.PNG"]
    files = []
    for p in patterns:
        files.extend(glob.glob(p))
    
    for f in sorted(files):
        # original_images フォルダ内のファイルは除外
        if "original_images" in f:
            continue
        size = os.path.getsize(f)
        if size > 300 * 1024:
            # バックアップ作成
            backup_dir = os.path.join(base_dir, "original_images")
            os.makedirs(backup_dir, exist_ok=True)
            backup_path = os.path.join(backup_dir, f)
            if not os.path.exists(backup_path):
                shutil.copy2(f, backup_path)
                print(f"📦 バックアップ保存: original_images/{f}")
            
            optimize_image(f, f)
            
    print("========================================")
    print("✨ 最適化処理が完了しました！")
    print("========================================")

if __name__ == "__main__":
    main()
