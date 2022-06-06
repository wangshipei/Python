import zipfile
import os
from pathlib import Path
import glob
from PIL import Image
import shutil
from tqdm import tqdm


new_zip_dir = r'C:\Users\24910\PycharmProjects\resize_pics\data2'
os.chdir(r'C:\Users\24910\PycharmProjects\resize_pics\data')
name_list = []
zip_list = glob.glob('*.*')
for zp in tqdm(zip_list, position=0, leave=True, desc=f"正在解压zip文件"):
    zFile = zipfile.ZipFile(zp, 'r')
    for fileM in zFile.namelist():
        if type(fileM) == int:
            pass
        else:
            try:
                extracted_path = Path(zFile.extract(fileM, new_zip_dir))
                extracted_path.rename(new_zip_dir + '//' + fileM.encode('cp437').decode('gbk'))
            except UnicodeEncodeError:
                pass
    zFile.close()

for dirpath, dirnames, filenames in os.walk(new_zip_dir, topdown=False):
    if not dirnames and not filenames:
        os.rmdir(dirpath)


pic_dirlist = os.listdir(new_zip_dir)
for p in tqdm(pic_dirlist, position=0, leave=True, desc=f"正在压缩图片"):
    dirname = new_zip_dir + '\\' + p
    os.chdir(dirname)
    pic_list = glob.glob('*.jpg')
    for pic in pic_list:
        image = Image.open(pic)
        image.thumbnail((1000, 1000))
        image.save(pic)

for zip_name in tqdm(pic_dirlist, position=0, leave=True, desc=f"正在压缩文件"):
    shutil.make_archive(new_zip_dir + '\\' + zip_name, 'zip', new_zip_dir + '\\' + zip_name)

for dname in pic_dirlist:
    shutil.rmtree(new_zip_dir + '\\' + dname, ignore_errors=True)
