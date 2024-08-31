from PIL import Image
import os
import zipfile

def extract_zip(zip_path, extract_to):
    """
    ZIPファイルを指定されたディレクトリに解凍します。
    """
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
        print(f"Extracted ZIP file to: {extract_to}")

def convert_psd_to_png(source_dir, target_dir):
    """
    PSDファイルをPNGに変換し、指定のディレクトリに保存します。
    """
    # 指定されたディレクトリに存在するPSDファイルを取得
    psd_files = [f for f in os.listdir(source_dir) if f.endswith('.psd')]
    
    # ターゲットディレクトリが存在しない場合は作成
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
    
    # PSDファイルをPNG形式に変換し、指定のディレクトリに保存
    for psd_file in psd_files:
        psd_path = os.path.join(source_dir, psd_file)
        png_path = os.path.join(target_dir, os.path.splitext(psd_file)[0] + '.png')
        
        try:
            # PSDファイルを開く
            with Image.open(psd_path) as img:
                # PNG形式で保存
                img.save(png_path, 'PNG')
                print(f"Converted: {psd_path} -> {png_path}")
        except Exception as e:
            print(f"Failed to convert {psd_path}: {e}")

if __name__ == "__main__":
    # ZIPファイルのパス
    zip_file_path = "./postcard_psd.zip"
    
    # ZIPファイルの解凍先ディレクトリ
    extract_directory = "./posdcard_psd"
    
    # PNGファイルを保存するディレクトリ
    target_directory = "../NengaMaker/figures"
    
    # ZIPファイルを解凍
    extract_zip(zip_file_path, extract_directory)
    
    # 解凍されたPSDファイルをPNGに変換
    convert_psd_to_png(extract_directory, target_directory)
