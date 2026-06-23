# -*- coding: utf-8 -*-
"""재계산+오류검사+시트별 PNG 추출 통합."""
import win32com.client as win32, win32clipboard
from PIL import Image
import io, os, struct, time

src = r"C:\Users\Osstem\우리집_가계부.xlsx"
outdir = r"C:\Users\Osstem\preview"
os.makedirs(outdir, exist_ok=True)
area = {
    "1_시작하기": ("시작하기", "B1:C28"),
    "2_설정": ("설정", "A1:K35"),
    "3_거래내역": ("거래내역", "A1:M38"),
    "4_대시보드": ("대시보드", "A1:N58"),
    "5_기간리포트": ("기간리포트", "A1:L50"),
    "6_진단": ("진단&인사이트", "A1:E33"),
}

def clip_to_png(path):
    win32clipboard.OpenClipboard()
    try:
        data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
    finally:
        win32clipboard.CloseClipboard()
    bmih = struct.unpack('<I', data[:4])[0]
    bpp = struct.unpack('<H', data[14:16])[0]
    comp = struct.unpack('<I', data[16:20])[0]
    clr = struct.unpack('<I', data[32:36])[0]
    palette = (clr if clr else (1 << bpp)) if bpp <= 8 else clr
    masks = 12 if comp == 3 else 0
    offset = 14 + bmih + masks + palette * 4
    fh = b'BM' + struct.pack('<I', 14 + len(data)) + b'\x00\x00\x00\x00' + struct.pack('<I', offset)
    Image.open(io.BytesIO(fh + data)).save(path, 'PNG')

excel = win32.Dispatch("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(src)
    time.sleep(2.0)
    try: excel.Calculate()
    except Exception: pass
    time.sleep(1.0)
    for fn, (sheet, rng) in area.items():
        sh = wb.Worksheets(sheet)
        try: sh.Activate()
        except Exception: pass
        try: sh.Range("A1").Select()
        except Exception: pass
        time.sleep(0.5)
        r = sh.Range(rng)
        for _ in range(4):
            try:
                excel.CutCopyMode = False; time.sleep(0.3)
                r.CopyPicture(1, 2); time.sleep(0.6)
                clip_to_png(os.path.join(outdir, fn + ".png"))
                break
            except Exception:
                time.sleep(1.0)
    print("PNG done")
    wb.Close(False)
finally:
    try: excel.Quit()
    except Exception: pass
