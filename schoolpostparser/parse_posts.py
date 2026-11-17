#!/usr/bin/env python3
"""Download a VK community wall archive and prepare two editable A4 DOCX files."""
from __future__ import annotations
import argparse, json, os, re, sys, time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

import requests
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

API = "https://api.vk.com/method/"
VERSION = "5.199"

# Transparent first-pass rules. Unclear content stays in the review file.
USEFUL = re.compile(r"\b(школ|учен|обуча|класс|педагог|учител|образован|олимпиад|конкурс|побед|приз|наград|достижен|мероприят|урок|встреч|акци|проект|библиотек|спорт|команд|экскурс|безопасност|первокласс|выпуск|линейк|творчеств|профориент|грамотност|памят)\w*", re.I)
AD = re.compile(r"\b(реклама|купить|заказать|скидк|промокод|акци[яи] недели|подписывайтесь|переходите по ссылке|партн[её]рск|розыгрыш|продаж|доставка|цены от)\w*", re.I)
VK_NOISE = re.compile(r"(?:\n?\s*(?:👉|➡️|🔗)?\s*(?:комментарии|обсуждение|подробнее|ссылка на запись)\s*:?\s*https?://\S+)|\[club\d+\|([^\]]+)\]|\[id\d+\|([^\]]+)\]", re.I)

def wall_id(url: str, token: str) -> int:
    """Resolve public slug to numeric VK group owner_id using VK API."""
    slug = urlparse(url if "://" in url else "https://" + url).path.strip("/").split("/")[0]
    if not slug: raise ValueError("Не удалось определить адрес паблика")
    r = requests.get(API + "utils.resolveScreenName", params={"screen_name": slug, "access_token": token, "v": VERSION}, timeout=30)
    r.raise_for_status(); data = r.json()
    if data.get("error") or not data.get("response") or data["response"].get("type") != "group":
        raise RuntimeError(f"VK не разрешил адрес паблика: {data.get('error', data)}")
    return -int(data["response"]["object_id"])

def fetch_wall(owner_id: int, token: str, start: datetime, end: datetime) -> list[dict]:
    posts=[]; offset=0
    while True:
        r=requests.get(API+"wall.get", params={"owner_id":owner_id,"count":100,"offset":offset,"access_token":token,"v":VERSION},timeout=40)
        r.raise_for_status(); data=r.json()
        if data.get("error"): raise RuntimeError(f"VK API: {data['error'].get('error_msg','ошибка')}")
        items=data["response"]["items"]
        if not items: break
        stop=False
        for p in items:
            dt=datetime.fromtimestamp(p["date"], timezone.utc)
            if dt < start: stop=True; break
            if start <= dt < end and p.get("text","").strip(): posts.append(p)
        if stop or len(items)<100: break
        offset+=len(items); time.sleep(.35)
    return sorted(posts,key=lambda p:p["date"])

def post_photos(post: dict) -> list[str]:
    urls=[]
    for a in post.get("attachments",[]):
        if a.get("type")=="photo":
            sizes=a["photo"].get("sizes",[])
            if sizes: urls.append(max(sizes,key=lambda s:s.get("width",0)*s.get("height",0))["url"])
    return urls

def clean(text: str) -> str:
    text=VK_NOISE.sub(lambda m: next((g for g in m.groups() if g), ""), text)
    text=re.sub(r"https?://\S+", "", text)
    text=re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}","\n\n",text).strip()

def category(text: str) -> str:
    if AD.search(text): return "review"
    # Explicit school-news signal required; no match goes to review, never discarded.
    return "useful" if USEFUL.search(text) else "review"

def download(url: str, path: Path) -> bool:
    try:
        r=requests.get(url,timeout=40); r.raise_for_status()
        path.write_bytes(r.content); return True
    except Exception as e:
        print(f"Не удалось скачать фото: {e}",file=sys.stderr); return False

def make_doc(posts: list[dict], dest: Path, label: str, photo_dir: Path) -> None:
    doc=Document(); sec=doc.sections[0]
    sec.page_width=Cm(21); sec.page_height=Cm(29.7)
    sec.top_margin=Cm(1.4); sec.bottom_margin=Cm(1.4); sec.left_margin=Cm(1.6); sec.right_margin=Cm(1.6)
    normal=doc.styles["Normal"]; normal.font.name="Arial"; normal.font.size=Pt(10.5)
    normal.paragraph_format.space_after=Pt(5)
    for i,p in enumerate(posts):
        dt=datetime.fromtimestamp(p["date"], timezone.utc).astimezone()
        title=doc.add_paragraph(); title.alignment=WD_ALIGN_PARAGRAPH.CENTER
        run=title.add_run(dt.strftime("%d.%m.%Y")); run.bold=True; run.font.size=Pt(11)
        for para in clean(p.get("text","")).split("\n"):
            if para.strip(): doc.add_paragraph(para.strip())
        # Keep image area modest; user can adjust in Word. Max two images to favor legibility.
        saved=[]
        for j,url in enumerate(post_photos(p)[:2]):
            fp=photo_dir/f"{abs(p.get('owner_id',0))}_{p['id']}_{j}.jpg"
            if download(url,fp): saved.append(fp)
        for fp in saved:
            try:
                para=doc.add_paragraph(); para.alignment=WD_ALIGN_PARAGRAPH.CENTER
                para.add_run().add_picture(str(fp),width=Cm(8.4 if len(saved)>1 else 12.5))
            except Exception as e: print(f"Пропускаю неподдерживаемое фото {fp.name}: {e}",file=sys.stderr)
        if i<len(posts)-1: doc.add_page_break()
    doc.core_properties.title=label
    doc.save(dest)

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--url",default="https://vk.ru/schoolkng1")
    ap.add_argument("--token",default=os.getenv("VK_TOKEN"),help="VK API access token (or set VK_TOKEN)")
    ap.add_argument("--month",default="2026-10",help="Month in YYYY-MM format")
    ap.add_argument("--json",help="Use previously exported VK wall JSON instead of VK API")
    ap.add_argument("--out",default="output")
    args=ap.parse_args()
    try: start=datetime.strptime(args.month,"%Y-%m").replace(tzinfo=timezone.utc)
    except ValueError: ap.error("--month ожидает YYYY-MM")
    end=datetime(start.year+ (start.month==12), start.month%12+1,1,tzinfo=timezone.utc)
    out=Path(args.out); photos=out/"photos"; photos.mkdir(parents=True,exist_ok=True)
    if args.json:
        raw=json.loads(Path(args.json).read_text(encoding="utf-8"))
        posts=raw.get("response",{}).get("items",raw if isinstance(raw,list) else [])
        posts=[p for p in posts if start.timestamp()<=p.get("date",0)<end.timestamp()]
    else:
        if not args.token: ap.error("Нужен --token или переменная VK_TOKEN; альтернативно --json с экспортом стены")
        posts=fetch_wall(wall_id(args.url,args.token),args.token,start,end)
    good=[p for p in posts if category(p.get("text",""))=="useful"]
    review=[p for p in posts if category(p.get("text",""))!="useful"]
    out.mkdir(parents=True,exist_ok=True)
    make_doc(good,out/"useful.docx","Полезные публикации",photos)
    make_doc(review,out/"review.docx","Публикации для проверки",photos)
    (out/"posts.json").write_text(json.dumps(posts,ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"Всего: {len(posts)}; в полезные: {len(good)}; на проверку: {len(review)}")
    print(f"Файлы: {out/'useful.docx'}, {out/'review.docx'}; исходные данные: {out/'posts.json'}")
if __name__=="__main__": main()
