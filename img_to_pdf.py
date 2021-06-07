from glob import glob
from os import path
from PIL import Image


def img_to_pdf(imgs, to_file):
    imgs_ = [Image.open(img).convert('RGB') for img in imgs]
    imgs_[0].save(to_file, save_all=True, append_images=imgs_[1:])
    print(f'img_to_pdf {to_file} done.')


# img_to_pdf(['0.jpg', '1.jpg', '2.jpg'], 'out_2.pdf')

img_dirs = glob(r"data\rongchuang\*\*\*_pdf") + glob(r"data/rongchuang/*/*/*_ppt") + glob(r"data/rongchuang/*/*/*_pptx")
for img_dir in img_dirs:
    # r = path.join(img_dir, r"\*.jpg")
    img_files_ = glob(img_dir + r"\*.jpg")
    img_files = sorted(img_files_, key=path.getmtime)
    if len(img_files) == 0:
        print(f"{img_dir} not glob!")
        continue
    basename_dir = path.basename(img_dir)
    # print(path.basename(html_dir))
    # print(html_files, len(html_files))
    output_file = path.join(img_dir, f"{basename_dir}.pdf")
    # print(output_file)
    # input()
    img_to_pdf(img_files, output_file)
