# coding=utf-8
# Kindle to PDF Converter
# Launch kindle for PC
# Require installing Ghostscript and ExifTool
# pip install pyautogui fpdf pillow
from dataclasses import dataclass
import os
import pyautogui
from fpdf import FPDF
import subprocess
import argparse
import time
from PIL import Image
import pygetwindow as gw
import sys

@dataclass
class K2pConfig:
    border_color: tuple = (0xE7, 0xE7, 0xE7)
    background_color: tuple = (0xFF, 0xFF, 0xFF)


class kindle2pdf():
    PAGE_NUMBER_MAX = 500  # safety limit
    OUTPUT_FOLDER = 'output'
    TEMP_BOOK_NAME = 'temp_book.pdf'
    TEMP_CMP_BOOK_NAME = 'temp_cmp_book.pdf'

    def __init__(self, config: K2pConfig = None):
        self.config = config or K2pConfig()
        self.prev_img = None
        self.image_width = None
        self.image_height = None
        self.image_top:tuple = None
        self.image_bottom:tuple = None
        self.kindle_window = None

    # save image
    def _save_image(self, im, page_num):
        # create "output" folder if not exists
        if not os.path.exists(self.OUTPUT_FOLDER):
            os.makedirs(self.OUTPUT_FOLDER)
        im.save(f'{self.OUTPUT_FOLDER}/page_{page_num:04d}.png')

    def _next_page(self, right=False):
        # Get Kindle window coordinates and click relative to its top-left corner
        if self.kindle_window is None:
            raise ValueError("Kindle window not set.")

        if right:
            x = self.kindle_window.right - 20
        else:
            x = self.kindle_window.left + 110

        y = self.kindle_window.top + self.kindle_window.height//2

        pyautogui.click(x, y)
        time.sleep(0.5)  # wait for page to load

    def _is_last_page(self, im):
        # check if previous image is same as current image
        if self.prev_img:
            # compare image bytes
            return self.prev_img.tobytes() == im.tobytes()

    def _capture_all_pages(self, right=False):
        page_number = 0
        while True:
            print(f'Capturing page {page_number + 1}...')
            im = self._capture_kindle_window()

            if self._is_last_page(im):
                print('Last page reached.')
                break

            self._save_image(im, page_number + 1)
            page_number += 1

            self.prev_img = im

            self._next_page(right)

            if page_number >= self.PAGE_NUMBER_MAX:
                print('Reached maximum page limit.')
                break
        return page_number

    def _find_left_border(self, im, sample_y):
        width, height = im.size
        border_find = False
        prev_pixel = None
        for x in range(width):
            pixel = im.getpixel((x, sample_y))
            if border_find == False:
                if pixel == self.config.border_color:
                    border_find = True
            elif (prev_pixel == self.config.background_color and
                  pixel      != self.config.background_color  ):
                return x

            prev_pixel = pixel
        return None
    
    def _find_right_border(self, im, sample_y):
        width, height = im.size
        prev_pixel = None
        for x in range(width-20, -1, -1): # skip 20 pixels from right edge because of black border
            pixel = im.getpixel((x, sample_y))
            if (prev_pixel == self.config.background_color and
                pixel      != self.config.background_color  ):
                return x
            prev_pixel = pixel
        return None

    def _detect_crop_border_x(self, im):
        left = None
        right = None
        # get width and height
        width, height = im.size
        y_block = height//3
        sample_y_list = (y_block, y_block*2)
        for sample_y in sample_y_list:
            left_temp = self._find_left_border(im, sample_y)
            right_temp = self._find_right_border(im, sample_y)
            if left_temp is not None:
                left = left_temp if left is None else min(left, left_temp)
            if right_temp is not None:
                right = right_temp if right is None else max(right, right_temp)
        return left, right

    def _find_top_border(self, im, sample_x):
        width, height = im.size
        for y in range(height):
            pixel = im.getpixel((sample_x, y))
            if pixel == self.config.border_color:
                return y + 1        
        return None

    def _find_bottom_border(self, im, sample_x):
        width, height = im.size
        border_find = False
        for y in range(height-1, -1, -1):
            pixel = im.getpixel((sample_x, y))
            if border_find == False:
                if pixel == self.config.border_color:
                    border_find = True
            else:
                if pixel == self.config.border_color:
                    return y - 1
        return None

    def _detect_crop_border_y(self, im):
        top = None
        bottom = None
        # get width and height
        width, height = im.size
        x_block = width//3
        sample_x_list = (x_block, x_block*2)

        for x in sample_x_list:
            top_temp = self._find_top_border(im, x)
            bottom_temp = self._find_bottom_border(im, x)
            if top_temp is not None:
                top = top_temp if top is None else min(top, top_temp)
            if bottom_temp is not None:
                bottom = bottom_temp if bottom is None else max(bottom, bottom_temp)
        return top, bottom

    def _calc_image_size(self, page_number):
        left    = None
        right   = None
        top     = None
        bottom  = None
        for img_idx in range(page_number):
            img_path = f'{self.OUTPUT_FOLDER}/page_{img_idx + 1:04d}.png'
            with Image.open(img_path) as im:
                left_temp, right_temp = self._detect_crop_border_x(im)
                top_temp, bottom_temp = self._detect_crop_border_y(im)
                if (left_temp   is not None and 
                    right_temp  is not None and 
                    top_temp    is not None and 
                    bottom_temp is not None    ):
                    left    = left_temp    if left   is None else min(left, left_temp)
                    right   = right_temp   if right  is None else max(right, right_temp)
                    top     = top_temp     if top    is None else min(top, top_temp)
                    bottom  = bottom_temp  if bottom is None else max(bottom, bottom_temp)
                else:
                    print(f'Could not find crop offsets for image {img_path}. Skipping size calculation.')
                    print(f'start_offset_x: {left_temp}, end_offset_x: {right_temp}, start_offset_y: {top_temp}, end_offset_y: {bottom_temp}')

        if (left   is not None and 
            right  is not None and 
            top    is not None and 
            bottom is not None    ):
            self.image_width = right - left + 1
            self.image_height = bottom - top + 1
            self.image_top = (left, top)
            self.image_bottom = (right, bottom)
            print(f'Calculated image size: {self.image_width}x{self.image_height}')
            print(f'Crop offsets: top-left({self.image_top[0]}, {self.image_top[1]}), bottom-right({self.image_bottom[0]}, {self.image_bottom[1]})')
        else:
            print('Could not calculate image size due to missing offsets.')
            # throw error
            raise ValueError('Could not calculate image size due to missing offsets.')


    def _crop_images(self, page_number):
        print('Cropping images...')
        for i in range(page_number):
            img_path = f'{self.OUTPUT_FOLDER}/page_{i + 1:04d}.png'
            with Image.open(img_path) as im:
                cropped_im = im.crop((self.image_top[0], self.image_top[1], self.image_bottom[0], self.image_bottom[1]))
                cropped_im.save(img_path)
        print('Cropping completed.')

    def _paste_right_half(self, pdf, img_path):
        half_width = self.image_width // 2
        with Image.open(img_path) as im:
            right_half = im.crop((half_width, 0, self.image_width, self.image_height))
            temp_path = img_path.replace('.png', '_right.png')
            right_half.save(temp_path)
        pdf.image(temp_path, 0, 0, half_width, self.image_height)
        os.remove(temp_path)

    def _paste_left_half(self, pdf, img_path):
        half_width = self.image_width // 2
        with Image.open(img_path) as im:
            left_half = im.crop((0, 0, half_width, self.image_height))
            temp_path = img_path.replace('.png', '_left.png')
            left_half.save(temp_path)
        pdf.image(temp_path, 0, 0, half_width, self.image_height)
        os.remove(temp_path)

    def _create_pdf(self, page_number, comic):
        print('Creating PDF...')
        if comic:
            half_width = self.image_width // 2
            pdf = FPDF(unit='pt', format=[half_width, self.image_height])
            for i in range(page_number):
                img_path = f'{self.OUTPUT_FOLDER}/page_{i + 1:04d}.png'
                pdf.add_page()
                self._paste_right_half(pdf, img_path)
                pdf.add_page()
                self._paste_left_half(pdf, img_path)
        else:
            pdf = FPDF(unit='pt', format=[self.image_width, self.image_height])
            for i in range(page_number):
                pdf.add_page()
                pdf.image(f'{self.OUTPUT_FOLDER}/page_{i + 1:04d}.png', 0, 0, self.image_width, self.image_height)

        pdf.output(f'{self.OUTPUT_FOLDER}/{self.TEMP_BOOK_NAME}', 'F')
        print('PDF created.')

    def _compress_pdf(self):
        print('Compressing PDF...')
        gs_command = [
        'gswin64c',
        '-sDEVICE=pdfwrite',
        '-dCompatibilityLevel=1.4',
        '-dPDFSETTINGS=/ebook',
        '-dNOPAUSE',
        '-dQUIET',
        '-dBATCH',
        f'-sOutputFile={self.OUTPUT_FOLDER}/{self.TEMP_CMP_BOOK_NAME}',
        f'{self.OUTPUT_FOLDER}/{self.TEMP_BOOK_NAME}'
        ]
        try:        
          subprocess.run(gs_command, check=True)
          print('PDF compressed.')
        except subprocess.CalledProcessError as e:
          print('Error during PDF compression:', e)

    def _inject_metadata(self, input_pdf, output_pdf):
        print('Injecting metadata...')
        command = f'exiftool '\
                f'-Creator="PFU ScanSnap Organizer 4.1.30 #S1500" '\
                f'-CreatorTool="PFU ScanSnap Organizer 4.1.30 #S1500" '\
                f'-Producer="Adobe PDF Scan Library 3.2" '\
                f'-CreationDate="D:20231016222044+09\'00\'" '\
                f'-ModDate="D:20231016222044+09\'00\'" '\
                f'-Author="" '\
                f'-Subject="" '\
                f'-Title="" '\
                f'-Keywords="" '\
                f'-o "{output_pdf}" "{input_pdf}"'
        print("Running command:", command)
        os.system(command)
        print(f"Metadata injected. Output saved to {output_pdf}")

    def _clean_up(self, page_number):
        for i in range(page_number):
            img_path = f'{self.OUTPUT_FOLDER}/page_{i + 1:04d}.png'
            if os.path.exists(img_path):
                os.remove(img_path)

        tmp_book_path = f'{self.OUTPUT_FOLDER}/{self.TEMP_BOOK_NAME}'
        if os.path.exists(tmp_book_path):
          os.remove(tmp_book_path)

        tmp_cmp_book_path = f'{self.OUTPUT_FOLDER}/{self.TEMP_CMP_BOOK_NAME}'
        if os.path.exists(tmp_cmp_book_path):
            os.remove(tmp_cmp_book_path)

    def _get_kindle_window(self):
        kindle_windows = [w for w in gw.getWindowsWithTitle('Kindle for PC') if w.visible]
        if kindle_windows:
            self.kindle_window = kindle_windows[0]
        else:
            print("Kindle window not found.")
            raise ValueError("Kindle window not found.")

    def _capture_kindle_window(self):
        if self.kindle_window is None:
            raise ValueError("Kindle window not set.")
        im = pyautogui.screenshot(region=(self.kindle_window.left, self.kindle_window.top, self.kindle_window.width-1, self.kindle_window.height-1))
        return im

    def _maximize_kindle_window(self):
        if(self.kindle_window is None):
            raise ValueError("Kindle window not set.")

        print("Activating and maximizing Kindle window...")
        self.kindle_window.activate()
        time.sleep(0.5)
        self.kindle_window.maximize()
        time.sleep(1)

    def main_process(self, comic, output_book_name, right):
        self._get_kindle_window()
        self._maximize_kindle_window()
        page_num = self._capture_all_pages(right)
        self._calc_image_size(page_num)
        self._crop_images(page_num)
        self._create_pdf(page_num, comic)
        self._compress_pdf()
        self._inject_metadata(f'{self.OUTPUT_FOLDER}/{self.TEMP_CMP_BOOK_NAME}', output_book_name)
        self._clean_up(page_num)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Kindle to PDF Converter')
    parser.add_argument('-name', nargs='?', default='kindle_book', help='Book name')
    parser.add_argument('-comic', action='store_true', help='Capture comic pages')
    parser.add_argument('-right', action='store_true', help='')
    args = parser.parse_args()
    output_book_name = args.name + '.pdf'
    comic = args.comic
    right = args.right

    config = K2pConfig(
        border_color = (0xE7, 0xE7, 0xE7),
        background_color = (0xFF, 0xFF, 0xFF)
    )

    if(comic):
        config.background_color=(0, 0, 0)

    k2p = kindle2pdf(config)
    k2p.main_process(comic, output_book_name, right)

