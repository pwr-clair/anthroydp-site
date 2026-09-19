# SNS 카드 (인스타그램 스토리 1080×1920)

- `closing-party-story.png` — 클로징 파티 <D.C : 맨 처음으로 다시> 스토리 카드 (완성본)
- `make-story.py` — 원본 사진을 포스터 톤(블랙/블루/화이트 디더링)으로 가공해 배경 생성
- `story.html` — 텍스트 레이아웃 (BG_SRC, FONT_DIR 자리는 렌더 시 치환)
- `shot.js` — Playwright로 1080×1920 PNG 렌더

다른 사진으로 다시 만들기:

```
python3 sns/make-story.py <사진> bg.png
sed -e "s#BG_SRC#file://$PWD/bg.png#" -e "s#FONT_DIR#file://<폰트폴더>#g" sns/story.html > story.html
node sns/shot.js $PWD/story.html sns/closing-party-story.png
```
폰트: Noto Sans KR(가변, notofonts/noto-cjk), Archivo Black(Google Fonts).
