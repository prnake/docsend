from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
from pathlib import Path

from PIL import Image
from requests_html import HTMLSession
import fitz


class DocSend:

    def __init__(self, doc_id):
        self.doc_id = doc_id.rpartition('/')[-1]
        self.url = f'https://docsend.com/view/{doc_id}'
        self.s = HTMLSession()

    def fetch_meta(self):
        r = self.s.get(self.url)
        r.raise_for_status()
        self.auth_token = None
        if r.html.find('input[@name="authenticity_token"]'):
            self.auth_token = r.html.find('input[@name="authenticity_token"]')[0].attrs['value']
        self.pages = int(r.html.find('.document-thumb-container')[-1].attrs['data-page-num'])

    def authorize(self, email, passcode=None):
        form = {
            'utf8': '✓',
            '_method': 'patch',
            'authenticity_token': self.auth_token,
            'link_auth_form[email]': email,
            'link_auth_form[passcode]': passcode,
            'commit': 'Continue',
        }
        f = self.s.post(self.url, data=form)
        f.raise_for_status()

    def fetch_images(self):
        self.image_urls = []
        pool = ThreadPoolExecutor(self.pages)
        results = list(pool.map(self._fetch_image, range(1, self.pages + 1)))
        self.images = [r[0] for r in results]
        self.page_links = [r[1] for r in results]

    def _fetch_image(self, page):
        meta = self.s.get(f'{self.url}/page_data/{page}')
        meta.raise_for_status()
        page_data = meta.json()
        data = self.s.get(page_data['imageUrl'])
        data.raise_for_status()
        rgba = Image.open(BytesIO(data.content))
        rgb = Image.new('RGB', rgba.size, (255, 255, 255))
        rgb.paste(rgba)
        # 保存 documentLinks 数据
        links = page_data.get('documentLinks', [])
        return rgb, links

    def save_pdf(self, name=None):
        self._save_pdf_with_links(name)
        # self.images[0].save(
        #     name,
        #     format='PDF',
        #     append_images=self.images[1:],
        #     save_all=True
        # )

    def _save_pdf_with_links(self, name=None):
        """使用 PyMuPDF 创建 PDF 并添加超链接"""
        doc = fitz.open()  # 创建新的 PDF 文档
        
        for page_num, (image, links) in enumerate(zip(self.images, self.page_links)):
            # 将 PIL 图片转换为字节
            img_bytes = BytesIO()
            image.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            # 创建新页面，使用图片尺寸
            page = doc.new_page(width=image.width, height=image.height)
            
            # 插入图片
            img_rect = fitz.Rect(0, 0, image.width, image.height)
            page.insert_image(img_rect, stream=img_bytes.getvalue())
            
            # 添加超链接
            for link in links:
                # 坐标是相对于页面尺寸的比例（0-1）
                x = link['x'] * image.width
                y = link['y'] * image.height
                width = link['width'] * image.width
                height = link['height'] * image.height
                
                # 创建链接矩形区域
                link_rect = fitz.Rect(x, y, x + width, y + height)
                
                # 获取 URI（优先使用原始 URI，如果没有则使用 trackedUrl）
                uri = link.get('uri', '')
                if not uri and link.get('trackedUrl'):
                    # 如果是相对路径，转换为完整 URL
                    tracked_url = link['trackedUrl']
                    if tracked_url.startswith('/'):
                        uri = f'https://docsend.com{tracked_url}'
                    else:
                        uri = tracked_url
                
                if uri:
                    # 添加超链接注释
                    page.insert_link({
                        'kind': fitz.LINK_URI,
                        'from': link_rect,
                        'uri': uri
                    })
        
        # 保存 PDF
        doc.save(name)
        doc.close()

    def save_images(self, name):
        path = Path(name)
        path.mkdir(exist_ok=True)
        for page, image in enumerate(self.images, start=1):
            image.save(path / f'{page}.png', format='PNG')
