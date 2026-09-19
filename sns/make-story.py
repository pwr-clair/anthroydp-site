# 인스타 스토리(1080x1920) 카드 배경 - 포스터 톤(블랙/블루/화이트 디더링)
# 사용법: python3 make-story.py <원본이미지> <출력png> [블루강도 0~1, 기본 0.55]
import sys
from PIL import Image, ImageOps, ImageEnhance, ImageChops

src, out = sys.argv[1], sys.argv[2]
blue_amt = float(sys.argv[3]) if len(sys.argv) > 3 else 0.55
W, H = 1080, 1920
im = ImageOps.exif_transpose(Image.open(src).convert("RGB"))
im = ImageOps.fit(im, (W, H), method=Image.LANCZOS, centering=(0.5, 0.5))

g = ImageOps.autocontrast(ImageOps.grayscale(im), cutoff=1)
g = ImageEnhance.Contrast(g).enhance(1.3)

# 블루 마스크: 중간톤 이상이 파랑으로 찍히도록 감마를 낮춤
blue_src = g.point(lambda v: int(255 * ((v / 255) ** (1.0 / (0.35 + blue_amt))) ))
blue_mask = blue_src.convert("1", dither=Image.FLOYDSTEINBERG)
# 화이트 마스크: 밝은 하이라이트만 드문드문 흰 점
white_src = g.point(lambda v: int(255 * max(0, (v - 190) / 65) ** 2.2))
white_mask = white_src.convert("1", dither=Image.FLOYDSTEINBERG)

black = Image.new("RGB", (W, H), (5, 8, 20))
blue = Image.new("RGB", (W, H), (28, 88, 255))
white = Image.new("RGB", (W, H), (240, 245, 255))
rgb = Image.composite(blue, black, blue_mask)
rgb = Image.composite(white, rgb, white_mask)

# 위·아래 어둡게(텍스트 가독성)
grad = Image.new("L", (W, H), 0)
px = grad.load()
for y in range(H):
    v = 0
    if y < 760: v = int(225 * (1 - y / 760) ** 1.1)
    elif y > 1200: v = int(215 * ((y - 1200) / (H - 1200)) ** 1.1)
    for x in range(W): px[x, y] = v
rgb = Image.composite(black, rgb, grad)
rgb.save(out)
print("saved", out, rgb.size)
